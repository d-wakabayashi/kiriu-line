#!/usr/bin/env python3
"""
KIRIU ライン負荷最適化システム - メインエントリーポイント

使用方法:
    python main.py                              # デフォルト設定で実行
    python main.py --template input.xlsx        # テンプレートから設定を読み込んで実行
    python main.py --generate-template          # 入力テンプレートを生成
    python main.py --capacities caps.json       # カスタム能力値を使用
    python main.py --time-limit 600             # ソルバー制限時間を600秒に設定
    python main.py --output-dir ./results       # 出力先ディレクトリを指定
"""

import argparse
import json
import sys
from pathlib import Path

from config import (
    DEFAULT_CAPACITIES,
    DEFAULT_TIME_LIMIT_SECONDS,
    OUTPUT_DIR,
    SPEC_FILE,
    PLAN_FILE,
    PLAN_SHEET,
)
from data_loader import load_all_data, load_equipment_spec, load_production_plan, merge_data, PartSpec
from model import optimize
from visualize import generate_all_outputs, generate_text_report
from excel_output import export_to_excel


def parse_args() -> argparse.Namespace:
    """コマンドライン引数を解析"""
    parser = argparse.ArgumentParser(
        description='KIRIU ライン負荷最適化システム',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    # テンプレート関連
    parser.add_argument(
        '--generate-template',
        action='store_true',
        help='入力テンプレートExcelを生成して終了',
    )

    parser.add_argument(
        '--template',
        type=str,
        default=None,
        help='入力テンプレートExcelファイルパス（設定をテンプレートから読み込む）',
    )

    parser.add_argument(
        '--template-output',
        type=str,
        default='input_template.xlsx',
        help='生成するテンプレートのファイル名 (デフォルト: input_template.xlsx)',
    )

    # 入力ファイル設定
    parser.add_argument(
        '--spec-file',
        type=str,
        default=SPEC_FILE,
        help=f'設備仕様ファイルパス (デフォルト: {SPEC_FILE})',
    )

    parser.add_argument(
        '--plan-file',
        type=str,
        default=PLAN_FILE,
        help=f'生産計画ファイルパス (デフォルト: {PLAN_FILE})',
    )

    parser.add_argument(
        '--plan-sheet',
        type=str,
        default=PLAN_SHEET,
        help=f'生産計画シート名 (デフォルト: {PLAN_SHEET})',
    )

    parser.add_argument(
        '--capacities',
        type=str,
        default=None,
        help='ライン能力設定JSONファイルパス',
    )

    parser.add_argument(
        '--time-limit',
        type=int,
        default=DEFAULT_TIME_LIMIT_SECONDS,
        help=f'ソルバー制限時間（秒） (デフォルト: {DEFAULT_TIME_LIMIT_SECONDS})',
    )

    parser.add_argument(
        '--output-dir',
        type=str,
        default=OUTPUT_DIR,
        help=f'出力ディレクトリ (デフォルト: {OUTPUT_DIR})',
    )

    parser.add_argument(
        '--no-visualize',
        action='store_true',
        help='可視化出力をスキップ',
    )

    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='データ読み込みのみ実行（最適化はスキップ）',
    )

    return parser.parse_args()


def load_capacities(filepath: str | None) -> dict[str, int]:
    """ライン能力設定を読み込む"""
    if filepath is None:
        return DEFAULT_CAPACITIES.copy()

    print(f"能力設定ファイル読み込み: {filepath}")
    with open(filepath, 'r', encoding='utf-8') as f:
        caps = json.load(f)

    # 値が整数であることを確認
    return {k: int(v) for k, v in caps.items()}


def run_with_template(template_path: str) -> int:
    """テンプレートから設定を読み込んで実行"""
    from input_template import load_input_config, load_parts_master, get_monthly_capacities
    from output_handler import process_outputs, create_result_email_body

    print("=" * 60)
    print("KIRIU ライン負荷最適化システム")
    print("=" * 60)
    print()
    print(f"【テンプレート読み込み】: {template_path}")

    # テンプレートから設定読み込み
    try:
        config = load_input_config(template_path)
        parts_master = load_parts_master(template_path)
    except Exception as e:
        print(f"エラー: テンプレートの読み込みに失敗しました - {e}")
        return 1

    print(f"  設備仕様ファイル: {config.spec_file}")
    print(f"  生産計画ファイル: {config.plan_file}")
    print(f"  生産計画シート: {config.plan_sheet}")
    print(f"  制限時間: {config.time_limit}秒")
    print(f"  出力ディレクトリ: {config.output_dir}")

    # 出力ディレクトリを設定
    import config as app_config
    app_config.OUTPUT_DIR = config.output_dir
    Path(config.output_dir).mkdir(parents=True, exist_ok=True)

    # データ読み込み
    try:
        print("\n【データ読み込み】")
        specs = load_equipment_spec(config.spec_file)
        demands, plan_infos = load_production_plan(config.plan_file, config.plan_sheet)

        # 部品マスタからの追加・上書き
        if parts_master:
            print(f"  部品マスタから追加/上書き: {len(parts_master)}件")
            for part_num, info in parts_master.items():
                if info['main_line']:
                    specs[part_num] = PartSpec(
                        part_number=part_num,
                        part_name=info['part_name'],
                        main_line=info['main_line'],
                        sub1_line=info['sub1_line'],
                        sub2_line=info['sub2_line'],
                    )

        specs, demands = merge_data(specs, demands, plan_infos)

        if not specs or not demands:
            print("エラー: 有効なデータがありません")
            return 1

    except FileNotFoundError as e:
        print(f"エラー: ファイルが見つかりません - {e}")
        return 1
    except Exception as e:
        print(f"エラー: データ読み込みに失敗しました - {e}")
        raise

    # ライン能力設定（月別対応）
    monthly_caps = get_monthly_capacities(config.capacities)
    # 最適化には各月の最大値を使用（または平均値）
    capacities = {line: max(caps) if isinstance(caps, list) else caps
                  for line, caps in config.capacities.items()}

    print("\n【ライン能力設定】")
    for line, cap in sorted(capacities.items()):
        print(f"  {line}: {cap:,}")

    # 最適化実行
    print("\n【最適化実行】")
    result = optimize(specs, demands, capacities, config.time_limit)

    if result.status not in ('OPTIMAL', 'FEASIBLE'):
        print(f"\nエラー: 最適化に失敗しました - ステータス: {result.status}")
        return 1

    # 結果出力
    print("\n【結果出力】")
    output_dir = Path(config.output_dir)
    generate_all_outputs(result, capacities)
    excel_path = export_to_excel(result, specs, capacities, str(output_dir / 'optimization_result.xlsx'))

    # 出力ファイル一覧
    output_files = [
        str(Path(config.output_dir) / 'optimization_result.xlsx'),
        str(Path(config.output_dir) / 'optimization_report.txt'),
        str(Path(config.output_dir) / 'line_loads.png'),
        str(Path(config.output_dir) / 'load_summary.png'),
    ]

    # Google Drive / メール送信
    if config.output_to_gdrive or config.send_email:
        # レポートテキストを生成
        report_text = generate_text_report(result, capacities)

        email_body = create_result_email_body(
            status=result.status,
            objective_value=result.objective_value,
            solve_time=result.solve_time,
            summary=report_text[:2000],  # 長すぎる場合は切り詰め
        )

        process_outputs(
            files=output_files,
            output_to_gdrive=config.output_to_gdrive,
            gdrive_folder_id=config.gdrive_folder_id,
            send_email_flag=config.send_email,
            email_to=config.email_to,
            email_subject=config.email_subject,
            email_body=email_body,
        )

    print("\n完了しました。")
    return 0


def main() -> int:
    """メイン処理"""
    args = parse_args()

    # テンプレート生成モード
    if args.generate_template:
        from input_template import generate_input_template
        print("入力テンプレートを生成します...")
        generate_input_template(args.template_output)
        print(f"\n使用方法:")
        print(f"  1. {args.template_output} を開いて設定を入力")
        print(f"  2. python main.py --template {args.template_output} で実行")
        return 0

    # テンプレートモード
    if args.template:
        return run_with_template(args.template)

    # 従来モード（コマンドライン引数）
    print("=" * 60)
    print("KIRIU ライン負荷最適化システム")
    print("=" * 60)
    print()

    # 出力ディレクトリを設定
    import config
    config.OUTPUT_DIR = args.output_dir

    # データ読み込み
    try:
        print("【データ読み込み】")
        specs = load_equipment_spec(args.spec_file)
        demands, plan_infos = load_production_plan(args.plan_file, args.plan_sheet)
        specs, demands = merge_data(specs, demands, plan_infos)

        if not specs or not demands:
            print("エラー: 有効なデータがありません")
            return 1

    except FileNotFoundError as e:
        print(f"エラー: ファイルが見つかりません - {e}")
        return 1
    except Exception as e:
        print(f"エラー: データ読み込みに失敗しました - {e}")
        raise

    # ドライラン時はここで終了
    if args.dry_run:
        print("\n【ドライラン完了】")
        print(f"  部品数: {len(specs)}")
        print(f"  総需要: {sum(sum(d.monthly_demand) for d in demands.values()):,}")
        return 0

    # ライン能力読み込み
    try:
        capacities = load_capacities(args.capacities)
        print("\n【ライン能力設定】")
        for line, cap in sorted(capacities.items()):
            print(f"  {line}: {cap:,}")
    except Exception as e:
        print(f"エラー: 能力設定の読み込みに失敗しました - {e}")
        return 1

    # 最適化実行
    print("\n【最適化実行】")
    result = optimize(specs, demands, capacities, args.time_limit)

    if result.status not in ('OPTIMAL', 'FEASIBLE'):
        print(f"\nエラー: 最適化に失敗しました - ステータス: {result.status}")
        return 1

    # 結果出力
    if not args.no_visualize:
        generate_all_outputs(result, capacities)
        # Excel出力
        export_to_excel(result, specs, capacities)
    else:
        print("\n【最適化完了】")
        print(f"  ステータス: {result.status}")
        print(f"  目的関数値: {result.objective_value:,.0f}")
        print(f"  実行時間: {result.solve_time:.2f}秒")

    print("\n完了しました。")
    return 0


if __name__ == '__main__':
    sys.exit(main())
