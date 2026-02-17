"""
KIRIU ライン負荷最適化システム - 可視化モジュール
"""

import json
from pathlib import Path

import matplotlib
matplotlib.use('Agg')  # GUIバックエンドを使用しない
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

from config import DISC_LINES, DEFAULT_CAPACITIES, MONTHS, OUTPUT_DIR
from model import OptimizationResult


# 日本語フォント対応
_JP_AVAILABLE = False

def _setup_japanese_font():
    """日本語フォントを設定"""
    global _JP_AVAILABLE

    # japanize_matplotlibを試す
    try:
        import japanize_matplotlib
        _JP_AVAILABLE = True
        return
    except ImportError:
        pass

    # 日本語フォントを探す
    import matplotlib.font_manager as fm
    jp_font_names = ['Noto Sans CJK JP', 'IPAGothic', 'IPAPGothic', 'TakaoPGothic',
                     'Yu Gothic', 'MS Gothic', 'Hiragino Sans', 'Meiryo']

    available_fonts = {f.name for f in fm.fontManager.ttflist}

    for font_name in jp_font_names:
        if font_name in available_fonts:
            plt.rcParams['font.family'] = font_name
            _JP_AVAILABLE = True
            return

    # 日本語フォントが見つからない場合
    plt.rcParams['font.family'] = ['DejaVu Sans', 'sans-serif']
    _JP_AVAILABLE = False

_setup_japanese_font()


# 英語ラベル（日本語フォントがない場合に使用）
MONTHS_EN = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']


def get_month_labels():
    """月ラベルを取得（日本語フォントの有無に応じて切り替え）"""
    return MONTHS if _JP_AVAILABLE else MONTHS_EN


def get_label(jp: str, en: str) -> str:
    """日本語/英語ラベルを切り替え"""
    return jp if _JP_AVAILABLE else en


def ensure_output_dir() -> Path:
    """出力ディレクトリを作成"""
    output_path = Path(OUTPUT_DIR)
    output_path.mkdir(parents=True, exist_ok=True)
    return output_path


def plot_line_loads(
    result: OptimizationResult,
    capacities: dict[str, int] | None = None,
    output_path: str | None = None,
) -> None:
    """
    ライン別月別負荷グラフを作成

    Args:
        result: 最適化結果
        capacities: ライン能力辞書
        output_path: 出力ファイルパス
    """
    caps = capacities or DEFAULT_CAPACITIES
    month_labels = get_month_labels()

    fig, axes = plt.subplots(3, 3, figsize=(15, 12))
    axes = axes.flatten()

    colors = plt.cm.tab10(np.linspace(0, 1, 12))

    for idx, line in enumerate(DISC_LINES):
        ax = axes[idx]
        loads = result.line_loads.get(line, [0] * 12)
        cap = caps.get(line, 0)

        # 棒グラフ
        bars = ax.bar(range(12), loads, color='steelblue', alpha=0.8)

        # オーバーフロー部分を赤色で表示
        for i, load in enumerate(loads):
            if load > cap:
                ax.bar(i, load - cap, bottom=cap, color='red', alpha=0.8)

        # 能力線
        cap_label = get_label(f'能力: {cap:,}', f'Cap: {cap:,}')
        ax.axhline(y=cap, color='red', linestyle='--', linewidth=2, label=cap_label)

        ax.set_title(f'Line {line}', fontsize=12, fontweight='bold')
        ax.set_xticks(range(12))
        ax.set_xticklabels(month_labels, rotation=45, ha='right', fontsize=8)
        ax.set_ylabel(get_label('生産数量', 'Production'))
        ax.legend(loc='upper right', fontsize=8)

        # Y軸の範囲を設定
        max_val = max(max(loads) if loads else 0, cap)
        ax.set_ylim(0, max_val * 1.2)

        # 負荷率を表示
        avg_load = sum(loads) / 12
        avg_rate = avg_load / cap * 100 if cap > 0 else 0
        ax.text(0.02, 0.98, f'Avg: {avg_rate:.1f}%',
                transform=ax.transAxes, va='top', fontsize=9,
                bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))

    plt.tight_layout()

    if output_path:
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        print(f"グラフ保存: {output_path}")
    else:
        output_dir = ensure_output_dir()
        plt.savefig(output_dir / 'line_loads.png', dpi=150, bbox_inches='tight')
        print(f"グラフ保存: {output_dir / 'line_loads.png'}")

    plt.close()


def plot_load_summary(
    result: OptimizationResult,
    capacities: dict[str, int] | None = None,
    output_path: str | None = None,
) -> None:
    """
    全ライン負荷サマリーグラフを作成

    Args:
        result: 最適化結果
        capacities: ライン能力辞書
        output_path: 出力ファイルパス
    """
    caps = capacities or DEFAULT_CAPACITIES
    month_labels = get_month_labels()

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

    # 左: 平均負荷率の棒グラフ
    avg_rates = []
    for line in DISC_LINES:
        loads = result.line_loads.get(line, [0] * 12)
        cap = caps.get(line, 0)
        avg_load = sum(loads) / 12
        avg_rate = avg_load / cap * 100 if cap > 0 else 0
        avg_rates.append(avg_rate)

    colors = ['red' if rate > 100 else 'steelblue' for rate in avg_rates]
    bars = ax1.bar(DISC_LINES, avg_rates, color=colors, alpha=0.8)

    ax1.axhline(y=100, color='red', linestyle='--', linewidth=2, label=get_label('能力100%', 'Capacity 100%'))
    ax1.set_title(get_label('ライン別平均負荷率', 'Average Load Rate by Line'), fontsize=14, fontweight='bold')
    ax1.set_xlabel(get_label('ライン', 'Line'))
    ax1.set_ylabel(get_label('負荷率 (%)', 'Load Rate (%)'))
    ax1.legend()

    # 値ラベルを追加
    for bar, rate in zip(bars, avg_rates):
        ax1.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 2,
                 f'{rate:.1f}%', ha='center', va='bottom', fontsize=9)

    # 右: 月別総負荷の推移
    monthly_totals = [0] * 12
    for line in DISC_LINES:
        loads = result.line_loads.get(line, [0] * 12)
        for i, load in enumerate(loads):
            monthly_totals[i] += load

    total_cap = sum(caps.values())

    ax2.bar(range(12), monthly_totals, color='steelblue', alpha=0.8, label=get_label('総生産量', 'Total Production'))
    ax2.axhline(y=total_cap, color='red', linestyle='--', linewidth=2, label=get_label(f'総能力: {total_cap:,}', f'Total Cap: {total_cap:,}'))

    ax2.set_title(get_label('月別総生産量', 'Monthly Total Production'), fontsize=14, fontweight='bold')
    ax2.set_xticks(range(12))
    ax2.set_xticklabels(month_labels, rotation=45, ha='right')
    ax2.set_ylabel(get_label('生産数量', 'Production'))
    ax2.legend()

    plt.tight_layout()

    if output_path:
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        print(f"グラフ保存: {output_path}")
    else:
        output_dir = ensure_output_dir()
        plt.savefig(output_dir / 'load_summary.png', dpi=150, bbox_inches='tight')
        print(f"グラフ保存: {output_dir / 'load_summary.png'}")

    plt.close()


def generate_text_report(
    result: OptimizationResult,
    capacities: dict[str, int] | None = None,
    output_path: str | None = None,
) -> str:
    """
    テキストレポートを生成

    Args:
        result: 最適化結果
        capacities: ライン能力辞書
        output_path: 出力ファイルパス

    Returns:
        レポート文字列
    """
    caps = capacities or DEFAULT_CAPACITIES

    lines = []
    lines.append("=" * 60)
    lines.append("KIRIU ライン負荷最適化レポート")
    lines.append("=" * 60)
    lines.append("")

    # 基本情報
    lines.append("【最適化結果】")
    lines.append(f"  ステータス: {result.status}")
    if result.objective_value is not None:
        lines.append(f"  目的関数値: {result.objective_value:,.0f}")
    lines.append(f"  実行時間: {result.solve_time:.2f}秒")
    lines.append("")

    # ライン別負荷サマリー
    lines.append("【ライン別負荷サマリー】")
    lines.append("-" * 50)
    lines.append(f"{'ライン':>8} {'能力':>10} {'平均負荷':>10} {'負荷率':>8} {'最大負荷':>10}")
    lines.append("-" * 50)

    total_cap = 0
    total_load = 0
    overflow_lines = []

    for line in DISC_LINES:
        loads = result.line_loads.get(line, [0] * 12)
        cap = caps.get(line, 0)
        avg_load = sum(loads) / 12
        max_load = max(loads)
        rate = avg_load / cap * 100 if cap > 0 else 0

        total_cap += cap
        total_load += avg_load

        status = "⚠️" if max_load > cap else ""
        if max_load > cap:
            overflow_lines.append(line)

        lines.append(f"{line:>8} {cap:>10,} {avg_load:>10,.0f} {rate:>7.1f}% {max_load:>10,} {status}")

    lines.append("-" * 50)
    total_rate = total_load / total_cap * 100 if total_cap > 0 else 0
    lines.append(f"{'合計':>8} {total_cap:>10,} {total_load:>10,.0f} {total_rate:>7.1f}%")
    lines.append("")

    # オーバーフロー詳細
    if overflow_lines:
        lines.append("【オーバーフロー詳細】")
        lines.append("※ 以下のラインで能力超過が発生しています:")
        for line in overflow_lines:
            loads = result.line_loads.get(line, [0] * 12)
            cap = caps.get(line, 0)
            for month_idx, load in enumerate(loads):
                if load > cap:
                    overflow_qty = load - cap
                    lines.append(f"  {line} / {MONTHS[month_idx]}: {load:,} (超過: {overflow_qty:,})")
        lines.append("")

    # サブライン使用状況
    sub_usage_count = 0
    sub_usage_qty = 0
    for part_num, monthly in result.sub_line_usage.items():
        for month_lines in monthly:
            if len(month_lines) > 1:  # メイン以外も使用
                sub_usage_count += 1

    for part_num, line_data in result.allocation.items():
        for line, monthly in line_data.items():
            # メインラインでない場合
            spec = None  # 仕様情報がない場合の処理
            for _, spec_data in result.sub_line_usage.items():
                pass  # 実際のspec情報は別途取得が必要
            sub_usage_qty += sum(monthly)  # 簡易計算

    lines.append("【サブライン使用状況】")
    lines.append(f"  サブライン使用月数: {sub_usage_count}")
    lines.append("")

    # 月別総生産量
    lines.append("【月別総生産量】")
    lines.append("-" * 40)
    lines.append(f"{'月':>6} {'生産量':>12} {'能力':>12} {'負荷率':>8}")
    lines.append("-" * 40)

    monthly_totals = [0] * 12
    for line in DISC_LINES:
        loads = result.line_loads.get(line, [0] * 12)
        for i, load in enumerate(loads):
            monthly_totals[i] += load

    for month_idx, total in enumerate(monthly_totals):
        rate = total / total_cap * 100 if total_cap > 0 else 0
        lines.append(f"{MONTHS[month_idx]:>6} {total:>12,} {total_cap:>12,} {rate:>7.1f}%")

    lines.append("-" * 40)
    lines.append(f"{'年間計':>6} {sum(monthly_totals):>12,}")
    lines.append("")

    lines.append("=" * 60)

    report = "\n".join(lines)

    # ファイル出力
    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f"レポート保存: {output_path}")
    else:
        output_dir = ensure_output_dir()
        report_path = output_dir / 'optimization_report.txt'
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f"レポート保存: {report_path}")

    return report


def save_solution_json(
    result: OptimizationResult,
    output_path: str | None = None,
) -> None:
    """
    最適化結果をJSONで保存

    Args:
        result: 最適化結果
        output_path: 出力ファイルパス
    """
    data = {
        'status': result.status,
        'objective_value': result.objective_value,
        'solve_time': result.solve_time,
        'allocation': result.allocation,
        'line_loads': result.line_loads,
        'overflow': result.overflow,
    }

    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"JSON保存: {output_path}")
    else:
        output_dir = ensure_output_dir()
        json_path = output_dir / 'solution.json'
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"JSON保存: {json_path}")


def generate_all_outputs(
    result: OptimizationResult,
    capacities: dict[str, int] | None = None,
    output_dir: str | None = None,
) -> None:
    """
    全ての出力ファイルを生成

    Args:
        result: 最適化結果
        capacities: ライン能力辞書
        output_dir: 出力ディレクトリパス（指定時は各ファイルをこのディレクトリに出力）
    """
    print("\n出力ファイル生成中...")

    if output_dir:
        out = Path(output_dir)
        out.mkdir(parents=True, exist_ok=True)
        plot_line_loads(result, capacities, output_path=str(out / 'line_loads.png'))
        plot_load_summary(result, capacities, output_path=str(out / 'load_summary.png'))
        report = generate_text_report(result, capacities, output_path=str(out / 'optimization_report.txt'))
        save_solution_json(result, output_path=str(out / 'solution.json'))
    else:
        plot_line_loads(result, capacities)
        plot_load_summary(result, capacities)
        report = generate_text_report(result, capacities)
        save_solution_json(result)

    print("\n" + report)


if __name__ == '__main__':
    # テスト用のダミーデータで動作確認
    from model import OptimizationResult

    # ダミー結果を作成
    dummy_loads = {
        '4915': [65000, 70000, 68000, 72000, 65000, 60000, 55000, 70000, 75000, 68000, 62000, 58000],
        '4919': [75000, 80000, 78000, 82000, 75000, 70000, 65000, 80000, 85000, 78000, 72000, 68000],
        '4927': [35000, 38000, 40000, 42000, 38000, 35000, 32000, 40000, 43000, 38000, 35000, 32000],
        '4928': [35000, 38000, 40000, 42000, 38000, 35000, 32000, 40000, 43000, 38000, 35000, 32000],
        '4934': [45000, 48000, 50000, 52000, 48000, 45000, 42000, 50000, 53000, 48000, 45000, 42000],
        '4935': [80000, 85000, 83000, 87000, 80000, 75000, 70000, 85000, 90000, 83000, 77000, 73000],
        '4945': [45000, 48000, 50000, 52000, 48000, 45000, 42000, 50000, 53000, 48000, 45000, 42000],
        '4G01': [45000, 48000, 50000, 52000, 48000, 45000, 42000, 50000, 53000, 48000, 45000, 42000],
        '4J01': [8000, 9000, 10000, 11000, 9000, 8000, 7000, 10000, 12000, 9000, 8000, 7000],
    }

    dummy_result = OptimizationResult(
        status='OPTIMAL',
        objective_value=12345.0,
        allocation={},
        line_loads=dummy_loads,
        overflow={line: [0] * 12 for line in DISC_LINES},
        sub_line_usage={},
        solve_time=30.5,
    )

    generate_all_outputs(dummy_result)
