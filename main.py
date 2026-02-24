#!/usr/bin/env python3
"""
KIRIU ライン負荷最適化システム - メインエントリーポイント

使用方法:
    python main.py                              # デフォルト設定で実行
    python main.py --template input.xlsx        # テンプレートから設定を読み込んで実行
    python main.py --generate-template          # 入力テンプレートを生成
    python main.py --spreadsheet                # Google Spreadsheetから読み書き
    python main.py --setup-sheets               # Spreadsheetにテンプレートをセットアップ
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
    DEFAULT_SPREADSHEET_ID,
    DEFAULT_TIME_LIMIT_SECONDS,
    DISC_LINES,
    OUTPUT_DIR,
    SPEC_FILE,
    PLAN_FILE,
    PLAN_SHEET,
)
from data_loader import load_all_data, load_equipment_spec, load_production_plan, merge_data, PartSpec
from model import optimize
from visualize import generate_all_outputs, generate_text_report
from excel_output import export_to_excel

# 負荷率パターン（100%, 90%, 80%）
LOAD_RATE_PATTERNS = [1.0, 0.9, 0.8]


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

    # Google Spreadsheet 連携
    parser.add_argument(
        '--spreadsheet',
        nargs='?',
        const=DEFAULT_SPREADSHEET_ID,
        default=None,
        help=f'Google Spreadsheetから読み書き (デフォルトID: {DEFAULT_SPREADSHEET_ID})',
    )

    parser.add_argument(
        '--setup-sheets',
        nargs='?',
        const=DEFAULT_SPREADSHEET_ID,
        default=None,
        help=f'Spreadsheetにテンプレートをセットアップ (デフォルトID: {DEFAULT_SPREADSHEET_ID})',
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

    # 複数負荷率パターンで最適化実行
    output_base = Path(config.output_dir)
    results_summary = []
    all_output_files = []

    for rate in LOAD_RATE_PATTERNS:
        pct_label = f"{int(rate * 100)}pct"
        pattern_dir = output_base / pct_label

        print(f"\n{'=' * 60}")
        print(f"【最適化実行】負荷率上限: {int(rate * 100)}%")
        print(f"{'=' * 60}")

        result = optimize(specs, demands, capacities, config.time_limit, load_rate_limit=rate)

        if result.status not in ('OPTIMAL', 'FEASIBLE'):
            print(f"  エラー: 最適化に失敗しました - ステータス: {result.status}")
            results_summary.append((rate, pct_label, result.status, None, None, None))
            continue

        # 結果サマリーを収集
        total_load = sum(sum(loads) for loads in result.line_loads.values())
        total_cap = sum(capacities.get(line, 0) for line in DISC_LINES) * 12
        avg_rate_pct = total_load / total_cap * 100 if total_cap > 0 else 0
        total_unmet = sum(sum(u) for u in result.unmet_demand.values()) if result.unmet_demand else 0
        results_summary.append((rate, pct_label, result.status, result.solve_time, avg_rate_pct, total_unmet))

        # 結果出力
        pattern_dir.mkdir(parents=True, exist_ok=True)
        generate_all_outputs(result, capacities, output_dir=str(pattern_dir))
        export_to_excel(result, specs, capacities, str(pattern_dir / 'optimization_result.xlsx'))

        all_output_files.extend([
            str(pattern_dir / 'optimization_result.xlsx'),
            str(pattern_dir / 'optimization_report.txt'),
            str(pattern_dir / 'line_loads.png'),
            str(pattern_dir / 'load_summary.png'),
        ])

    # パターン比較サマリー
    print(f"\n{'=' * 60}")
    print("【パターン比較サマリー】")
    print(f"{'=' * 60}")
    print(f"{'負荷率上限':>12} {'ステータス':>10} {'実行時間':>10} {'平均負荷率':>10} {'未割当合計':>10}")
    print("-" * 56)
    for rate, label, status, solve_time, avg_r, unmet in results_summary:
        time_str = f"{solve_time:.2f}s" if solve_time is not None else "-"
        avg_str = f"{avg_r:.1f}%" if avg_r is not None else "-"
        unmet_str = f"{unmet:,}" if unmet is not None else "-"
        print(f"{int(rate * 100)}% ({label}){status:>10} {time_str:>10} {avg_str:>10} {unmet_str:>10}")

    # Google Drive / メール送信（最初の成功結果のレポートを使用）
    if config.output_to_gdrive or config.send_email:
        # 100%パターンのレポートを基にメール本文を作成
        first_success = next(
            ((r, l) for r, l, s, *_ in results_summary if s in ('OPTIMAL', 'FEASIBLE')),
            None,
        )
        if first_success:
            report_path = output_base / f"{first_success[1]}" / 'optimization_report.txt'
            report_text = report_path.read_text(encoding='utf-8') if report_path.exists() else ""
        else:
            report_text = "全パターンで最適化に失敗しました。"

        email_body = create_result_email_body(
            status=results_summary[0][2] if results_summary else "UNKNOWN",
            objective_value=None,
            solve_time=0,
            summary=report_text[:2000],
        )

        process_outputs(
            files=all_output_files,
            output_to_gdrive=config.output_to_gdrive,
            gdrive_folder_id=config.gdrive_folder_id,
            send_email_flag=config.send_email,
            email_to=config.email_to,
            email_subject=config.email_subject,
            email_body=email_body,
        )

    print("\n完了しました。")
    return 0


def run_with_spreadsheet(spreadsheet_id: str, time_limit: int = DEFAULT_TIME_LIMIT_SECONDS) -> int:
    """Google Spreadsheetから読み書きして最適化を実行"""
    from sheets_io import (
        read_input_sheet, read_line_capacities, write_results,
        has_work_pattern_sheets, read_work_patterns, read_line_jph, read_monthly_working_days,
    )
    from config import calculate_monthly_capacities
    from data_loader import merge_data

    print("=" * 60)
    print("KIRIU ライン負荷最適化システム (Google Spreadsheet モード)")
    print("=" * 60)
    print()
    print(f"スプレッドシートID: {spreadsheet_id}")
    print(f"URL: https://docs.google.com/spreadsheets/d/{spreadsheet_id}")
    print()

    # 入力シートから部品データ読み込み
    try:
        print("【入力シート読み込み】")
        specs, demands = read_input_sheet(spreadsheet_id)

        if not specs or not demands:
            print("エラー: 入力シートに有効なデータがありません")
            return 1

    except Exception as e:
        print(f"エラー: 入力シートの読み込みに失敗しました - {e}")
        raise

    print(f"\n  部品数: {len(specs)}")
    print(f"  総需要: {sum(sum(d.monthly_demand) for d in demands.values()):,}")

    # 新シートがあるか確認 → 勤務体制パターン方式 / 旧方式フォールバック
    use_work_patterns = False
    try:
        use_work_patterns = has_work_pattern_sheets(spreadsheet_id)
    except Exception:
        pass

    if use_work_patterns:
        return _run_spreadsheet_work_patterns(spreadsheet_id, specs, demands, time_limit)
    else:
        print("\n※ 勤務体制パターンシートが未設定のため、旧方式（負荷率パターン）で実行します。")
        return _run_spreadsheet_load_rate(spreadsheet_id, specs, demands, time_limit)


def _run_spreadsheet_work_patterns(spreadsheet_id, specs, demands, time_limit):
    """勤務体制パターン方式でSpreadsheet最適化を実行"""
    from sheets_io import read_work_patterns, read_line_jph, read_monthly_working_days, write_results
    from config import calculate_monthly_capacities

    # 新シート3枚を読み込み
    try:
        print("\n【勤務体制パターン読み込み】")
        work_patterns = read_work_patterns(spreadsheet_id)
        jph = read_line_jph(spreadsheet_id)
        working_days = read_monthly_working_days(spreadsheet_id)
    except Exception as e:
        print(f"エラー: 勤務体制パターンの読み込みに失敗しました - {e}")
        raise

    # パターン別の月別能力を計算
    pattern_capacities = calculate_monthly_capacities(jph, work_patterns, working_days)

    print("\n【パターン別月間能力（平均）】")
    for pattern_name, caps in pattern_capacities.items():
        print(f"  {pattern_name}:")
        for line in DISC_LINES:
            avg = sum(caps[line]) // 12
            print(f"    {line}: {avg:,}/月")

    # 勤務体制パターンごとに最適化実行
    results_summary = []

    for pattern_name, capacities in pattern_capacities.items():
        sheet_suffix = f"_{pattern_name}"

        print(f"\n{'=' * 60}")
        print(f"【最適化実行】勤務体制: {pattern_name}")
        print(f"{'=' * 60}")

        result = optimize(specs, demands, capacities, time_limit)

        if result.status not in ('OPTIMAL', 'FEASIBLE'):
            print(f"  エラー: 最適化に失敗しました - ステータス: {result.status}")
            results_summary.append((pattern_name, result.status, None, None, None))
            continue

        # 結果サマリーを収集
        total_load = sum(sum(loads) for loads in result.line_loads.values())
        total_cap = sum(sum(capacities.get(line, [0] * 12)) for line in DISC_LINES)
        avg_rate_pct = total_load / total_cap * 100 if total_cap > 0 else 0
        total_unmet = sum(sum(u) for u in result.unmet_demand.values()) if result.unmet_demand else 0
        results_summary.append((pattern_name, result.status, result.solve_time, avg_rate_pct, total_unmet))

        # 結果をスプレッドシートに書き込み
        try:
            print(f"\n  結果をスプレッドシートに書き込み中...")
            write_results(spreadsheet_id, result, specs, capacities, sheet_suffix)
        except Exception as e:
            print(f"  警告: 結果の書き込みに失敗しました - {e}")

    # パターン比較サマリー
    print(f"\n{'=' * 60}")
    print("【勤務体制パターン比較サマリー】")
    print(f"{'=' * 60}")
    print(f"{'勤務体制':>14} {'ステータス':>10} {'実行時間':>10} {'平均負荷率':>10} {'未割当合計':>10}")
    print("-" * 58)
    for entry in results_summary:
        name = entry[0]
        status = entry[1]
        solve_time = entry[2]
        avg_r = entry[3]
        unmet = entry[4]
        time_str = f"{solve_time:.2f}s" if solve_time is not None else "-"
        avg_str = f"{avg_r:.1f}%" if avg_r is not None else "-"
        unmet_str = f"{unmet:,}" if unmet is not None else "-"
        print(f"{name:>14} {status:>10} {time_str:>10} {avg_str:>10} {unmet_str:>10}")

    print(f"\n結果は以下のスプレッドシートに書き込まれました:")
    print(f"  https://docs.google.com/spreadsheets/d/{spreadsheet_id}")
    print("\n完了しました。")
    return 0


def _run_spreadsheet_load_rate(spreadsheet_id, specs, demands, time_limit):
    """旧方式（負荷率パターン）でSpreadsheet最適化を実行"""
    from sheets_io import read_line_capacities, write_results

    # ライン能力読み込み
    try:
        print("\n【ライン能力読み込み】")
        capacities = read_line_capacities(spreadsheet_id)
    except Exception as e:
        print(f"エラー: ライン能力の読み込みに失敗しました - {e}")
        raise

    # 複数負荷率パターンで最適化実行
    results_summary = []

    for rate in LOAD_RATE_PATTERNS:
        pct_label = f"{int(rate * 100)}pct"
        sheet_suffix = f"_{pct_label}"

        print(f"\n{'=' * 60}")
        print(f"【最適化実行】負荷率上限: {int(rate * 100)}%")
        print(f"{'=' * 60}")

        result = optimize(specs, demands, capacities, time_limit, load_rate_limit=rate)

        if result.status not in ('OPTIMAL', 'FEASIBLE'):
            print(f"  エラー: 最適化に失敗しました - ステータス: {result.status}")
            results_summary.append((rate, pct_label, result.status, None, None, None))
            continue

        # 結果サマリーを収集
        total_load = sum(sum(loads) for loads in result.line_loads.values())
        total_cap = sum(sum(capacities.get(line, [0] * 12)) for line in DISC_LINES)
        avg_rate_pct = total_load / total_cap * 100 if total_cap > 0 else 0
        total_unmet = sum(sum(u) for u in result.unmet_demand.values()) if result.unmet_demand else 0
        results_summary.append((rate, pct_label, result.status, result.solve_time, avg_rate_pct, total_unmet))

        # 結果をスプレッドシートに書き込み
        try:
            print(f"\n  結果をスプレッドシートに書き込み中...")
            write_results(spreadsheet_id, result, specs, capacities, sheet_suffix)
        except Exception as e:
            print(f"  警告: 結果の書き込みに失敗しました - {e}")

    # パターン比較サマリー
    print(f"\n{'=' * 60}")
    print("【パターン比較サマリー】")
    print(f"{'=' * 60}")
    print(f"{'負荷率上限':>12} {'ステータス':>10} {'実行時間':>10} {'平均負荷率':>10} {'未割当合計':>10}")
    print("-" * 56)
    for rate, label, status, solve_time, avg_r, unmet in results_summary:
        time_str = f"{solve_time:.2f}s" if solve_time is not None else "-"
        avg_str = f"{avg_r:.1f}%" if avg_r is not None else "-"
        unmet_str = f"{unmet:,}" if unmet is not None else "-"
        print(f"{int(rate * 100)}% ({label}){status:>10} {time_str:>10} {avg_str:>10} {unmet_str:>10}")

    print(f"\n結果は以下のスプレッドシートに書き込まれました:")
    print(f"  https://docs.google.com/spreadsheets/d/{spreadsheet_id}")
    print("\n完了しました。")
    return 0


def main() -> int:
    """メイン処理"""
    args = parse_args()

    # スプレッドシート テンプレートセットアップモード
    if args.setup_sheets:
        from sheets_io import setup_template
        print("Google Spreadsheetにテンプレートをセットアップします...")
        setup_template(args.setup_sheets)
        return 0

    # スプレッドシートモード
    if args.spreadsheet:
        return run_with_spreadsheet(args.spreadsheet, args.time_limit)

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

    # 複数負荷率パターンで最適化実行
    output_base = Path(args.output_dir)
    results_summary = []

    for rate in LOAD_RATE_PATTERNS:
        pct_label = f"{int(rate * 100)}pct"
        pattern_dir = output_base / pct_label

        print(f"\n{'=' * 60}")
        print(f"【最適化実行】負荷率上限: {int(rate * 100)}%")
        print(f"{'=' * 60}")

        result = optimize(specs, demands, capacities, args.time_limit, load_rate_limit=rate)

        if result.status not in ('OPTIMAL', 'FEASIBLE'):
            print(f"  エラー: 最適化に失敗しました - ステータス: {result.status}")
            results_summary.append((rate, pct_label, result.status, None, None, None))
            continue

        # 結果サマリーを収集
        total_load = sum(sum(loads) for loads in result.line_loads.values())
        total_cap = sum(capacities.get(line, 0) for line in DISC_LINES) * 12
        avg_rate_pct = total_load / total_cap * 100 if total_cap > 0 else 0
        total_unmet = sum(sum(u) for u in result.unmet_demand.values()) if result.unmet_demand else 0
        results_summary.append((rate, pct_label, result.status, result.solve_time, avg_rate_pct, total_unmet))

        # 結果出力
        if not args.no_visualize:
            pattern_dir.mkdir(parents=True, exist_ok=True)
            generate_all_outputs(result, capacities, output_dir=str(pattern_dir))
            export_to_excel(result, specs, capacities, str(pattern_dir / 'optimization_result.xlsx'))
        else:
            print(f"\n  ステータス: {result.status}")
            print(f"  目的関数値: {result.objective_value:,.0f}")
            print(f"  実行時間: {result.solve_time:.2f}秒")

    # パターン比較サマリー
    print(f"\n{'=' * 60}")
    print("【パターン比較サマリー】")
    print(f"{'=' * 60}")
    print(f"{'負荷率上限':>12} {'ステータス':>10} {'実行時間':>10} {'平均負荷率':>10} {'未割当合計':>10}")
    print("-" * 56)
    for rate, label, status, solve_time, avg_r, unmet in results_summary:
        time_str = f"{solve_time:.2f}s" if solve_time is not None else "-"
        avg_str = f"{avg_r:.1f}%" if avg_r is not None else "-"
        unmet_str = f"{unmet:,}" if unmet is not None else "-"
        print(f"{int(rate * 100)}% ({label}){status:>10} {time_str:>10} {avg_str:>10} {unmet_str:>10}")

    print("\n完了しました。")
    return 0


if __name__ == '__main__':
    sys.exit(main())
