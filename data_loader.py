"""
KIRIU ライン負荷最適化システム - データ読み込みモジュール
"""

import re
from dataclasses import dataclass

import pandas as pd

from config import DISC_LINES, SPEC_FILE, PLAN_FILE, PLAN_SHEET


@dataclass
class PartSpec:
    """部品仕様情報"""
    part_number: str          # 部品番号（正規化済み）
    part_name: str            # 部品名称
    main_line: str | None     # メインライン
    sub1_line: str | None     # サブ1ライン
    sub2_line: str | None     # サブ2ライン


@dataclass
class PartDemand:
    """部品需要情報"""
    part_number: str          # 部品番号（正規化済み）
    part_name: str            # 部品名称
    monthly_demand: list[int] # 月別需要（4月〜3月の12ヶ月）


def normalize_part_number(part_num: str) -> str:
    """部品番号を正規化（ハイフン除去、空白除去、全角→半角）"""
    if pd.isna(part_num):
        return ''
    s = str(part_num).strip()
    # 全角英数を半角に
    s = s.translate(str.maketrans(
        'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ０１２３４５６７８９',
        'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
    ))
    return re.sub(r'[-\s]', '', s).upper()


def normalize_line_name(line_name: str) -> str | None:
    """ライン名を正規化"""
    if pd.isna(line_name):
        return None
    s = str(line_name).strip()
    # 全角を半角に
    s = s.translate(str.maketrans(
        '０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ',
        '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    ))
    # 数字とアルファベットのみ抽出
    s = re.sub(r'[^0-9A-Z]', '', s.upper())
    # 先頭のMを除去（Ｍ4927 → 4927）
    if s.startswith('M') and len(s) > 1:
        s = s[1:]
    # 末尾の.0を除去（4935.0 → 4935）
    if s.endswith('0') and '.' in str(line_name):
        # 元の値に小数点があった場合
        pass
    # ライン名として有効か確認
    if s in DISC_LINES:
        return s
    # 数値のみの場合、前に4を付けてチェック
    if s.isdigit() and len(s) == 3:
        candidate = '4' + s
        if candidate in DISC_LINES:
            return candidate
    return s if s else None


def load_equipment_spec(filepath: str = SPEC_FILE) -> dict[str, PartSpec]:
    """
    設備仕様ファイルから部品仕様を読み込む

    ファイル構造:
    - 列1: メインライン
    - 列2: サブ1ライン
    - 列3: サブ2ライン
    - 列6: 部品番号

    Returns:
        dict[部品番号, PartSpec]
    """
    print(f"設備仕様ファイル読み込み: {filepath}")

    df = pd.read_excel(filepath, sheet_name=0, header=None)

    parts = {}

    for idx, row in df.iterrows():
        # 行8以降がデータ行（ライン仕様行はスキップ）
        if idx < 8:
            continue

        # 列1: メインライン
        main_raw = row.iloc[1] if len(row) > 1 else None
        # 列2: サブ1ライン
        sub1_raw = row.iloc[2] if len(row) > 2 else None
        # 列3: サブ2ライン
        sub2_raw = row.iloc[3] if len(row) > 3 else None
        # 列6: 部品番号
        part_raw = row.iloc[6] if len(row) > 6 else None

        part_num = normalize_part_number(part_raw)
        if not part_num:
            continue

        # 「ライン仕様」や「加工ﾗｲﾝ計」などの行をスキップ
        if 'ライン' in str(main_raw) or '仕様' in str(main_raw) or '計' in part_num:
            continue

        main_line = normalize_line_name(main_raw)
        sub1_line = normalize_line_name(sub1_raw)
        sub2_line = normalize_line_name(sub2_raw)

        # メインラインが有効なディスクラインの場合のみ登録
        if main_line and main_line in DISC_LINES:
            # サブラインもディスクラインのもののみ
            if sub1_line and sub1_line not in DISC_LINES:
                sub1_line = None
            if sub2_line and sub2_line not in DISC_LINES:
                sub2_line = None

            parts[part_num] = PartSpec(
                part_number=part_num,
                part_name='',  # 仕様ファイルに名称列がない
                main_line=main_line,
                sub1_line=sub1_line,
                sub2_line=sub2_line,
            )

    print(f"  読み込み部品数: {len(parts)}")

    # ライン別部品数を表示
    line_counts = {line: 0 for line in DISC_LINES}
    for spec in parts.values():
        if spec.main_line:
            line_counts[spec.main_line] = line_counts.get(spec.main_line, 0) + 1
    print("  ライン別メイン割当部品数:")
    for line in DISC_LINES:
        if line_counts[line] > 0:
            print(f"    {line}: {line_counts[line]}件")

    return parts


@dataclass
class PartPlanInfo:
    """生産計画から取得した部品情報（仕様補完用）"""
    part_number: str
    part_name: str
    main_line: str | None
    sub1_line: str | None
    sub2_line: str | None


def load_production_plan(
    filepath: str = PLAN_FILE,
    sheet_name: str = PLAN_SHEET
) -> tuple[dict[str, PartDemand], dict[str, PartPlanInfo]]:
    """
    生産計画ファイルから部品需要を読み込む

    ファイル構造（赤堀分工場2029上期シート）:
    - 行17: ヘッダー
    - 行18以降: データ
    - 列7: 分類名（5:ディスク）
    - 列8,9,10: メインライン、サブ1、サブ2
    - 列17: 部品番号
    - 列18: 部品名
    - 列28-39: 月別数量（4月〜3月）

    Returns:
        (dict[部品番号, PartDemand], dict[部品番号, PartPlanInfo])
    """
    print(f"生産計画ファイル読み込み: {filepath}")
    print(f"  シート: {sheet_name}")

    df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)

    # ヘッダー行（行17, 0-indexed）を確認
    header_row = 17

    # 月別数量の列位置を特定（「数量4月」～「数量3月」）
    header = df.iloc[header_row]
    month_cols = []
    months_to_find = ['数量4月', '数量5月', '数量6月', '数量7月', '数量8月', '数量9月',
                      '数量10月', '数量11月', '数量12月', '数量1月', '数量2月', '数量3月']

    for month in months_to_find:
        for col_idx, val in enumerate(header):
            if pd.notna(val) and month in str(val):
                month_cols.append(col_idx)
                break

    if len(month_cols) != 12:
        print(f"  警告: 月別列が{len(month_cols)}個しか見つかりませんでした")
        # フォールバック: 列28-39を使用
        month_cols = list(range(28, 40))

    print(f"  月別数量列: {month_cols}")

    # 分類名の列位置を特定
    category_col = None
    part_col = None
    name_col = None
    line_cols = []  # 加工ライン列（メイン、サブ1、サブ2）

    for col_idx, val in enumerate(header):
        val_str = str(val) if pd.notna(val) else ''
        if '分類名' in val_str:
            category_col = col_idx
        elif '部品番号' in val_str and part_col is None:
            part_col = col_idx
        elif '部品名' in val_str and name_col is None:
            name_col = col_idx
        elif val_str == '加工ﾗｲﾝ':
            line_cols.append(col_idx)  # すべての加工ライン列を収集

    # デフォルト位置
    if category_col is None:
        category_col = 7
    if part_col is None:
        part_col = 17
    if name_col is None:
        name_col = 18
    if not line_cols:
        line_cols = [8, 9, 10]  # デフォルト: メイン、サブ1、サブ2

    main_line_col = line_cols[0] if len(line_cols) > 0 else 8
    sub1_line_col = line_cols[1] if len(line_cols) > 1 else 9
    sub2_line_col = line_cols[2] if len(line_cols) > 2 else 10

    print(f"  分類名列: {category_col}")
    print(f"  部品番号列: {part_col}")
    print(f"  ライン列: メイン={main_line_col}, サブ1={sub1_line_col}, サブ2={sub2_line_col}")

    demands = {}
    plan_infos = {}  # 生産計画から取得したライン情報
    disc_keywords = ['ディスク', 'ﾃﾞｨｽｸ', '5:ディスク', '5：ディスク', '5:ﾃﾞｨｽｸ']

    for idx, row in df.iterrows():
        if idx <= header_row:
            continue

        # 分類名でフィルタ（ディスクのみ）
        category = str(row.iloc[category_col]) if pd.notna(row.iloc[category_col]) else ''
        is_disc = any(kw in category for kw in disc_keywords) or category.startswith('5:') or category == '5'

        if not is_disc:
            continue

        # ライン情報を取得
        main_line_raw = row.iloc[main_line_col] if pd.notna(row.iloc[main_line_col]) else ''
        main_line = normalize_line_name(main_line_raw)

        # メインラインがディスクラインかチェック
        if main_line not in DISC_LINES:
            continue

        # サブライン情報も取得
        sub1_line_raw = row.iloc[sub1_line_col] if sub1_line_col < len(row) and pd.notna(row.iloc[sub1_line_col]) else ''
        sub2_line_raw = row.iloc[sub2_line_col] if sub2_line_col < len(row) and pd.notna(row.iloc[sub2_line_col]) else ''
        sub1_line = normalize_line_name(sub1_line_raw)
        sub2_line = normalize_line_name(sub2_line_raw)

        # サブラインがディスクラインでない場合はNone
        if sub1_line and sub1_line not in DISC_LINES:
            sub1_line = None
        if sub2_line and sub2_line not in DISC_LINES:
            sub2_line = None

        # 部品番号
        part_raw = row.iloc[part_col] if pd.notna(row.iloc[part_col]) else ''
        part_num = normalize_part_number(part_raw)

        if not part_num or '計' in str(part_raw):
            continue

        # 部品名
        part_name = str(row.iloc[name_col]) if pd.notna(row.iloc[name_col]) else ''

        # ライン情報を保存（最初に見つかったものを使用）
        if part_num not in plan_infos:
            plan_infos[part_num] = PartPlanInfo(
                part_number=part_num,
                part_name=part_name,
                main_line=main_line,
                sub1_line=sub1_line,
                sub2_line=sub2_line,
            )

        # 月別需要を取得
        monthly = []
        for col_idx in month_cols:
            val = row.iloc[col_idx] if col_idx < len(row) else 0
            if pd.isna(val):
                val = 0
            try:
                val = int(float(val))
            except (ValueError, TypeError):
                val = 0
            monthly.append(max(0, val))

        # 需要がゼロでない場合のみ追加
        if sum(monthly) > 0:
            # 同一部品番号が複数行ある場合は合算
            if part_num in demands:
                existing = demands[part_num]
                for i in range(12):
                    existing.monthly_demand[i] += monthly[i]
            else:
                demands[part_num] = PartDemand(
                    part_number=part_num,
                    part_name=part_name,
                    monthly_demand=monthly,
                )

    print(f"  読み込み部品数: {len(demands)}")
    total_demand = sum(sum(d.monthly_demand) for d in demands.values())
    print(f"  年間総需要: {total_demand:,}")

    return demands, plan_infos


def merge_data(
    specs: dict[str, PartSpec],
    demands: dict[str, PartDemand],
    plan_infos: dict[str, PartPlanInfo] | None = None,
) -> tuple[dict[str, PartSpec], dict[str, PartDemand]]:
    """
    仕様データと需要データをマージし、不整合を報告

    Args:
        specs: 設備仕様から読み込んだ部品仕様
        demands: 生産計画から読み込んだ需要
        plan_infos: 生産計画から取得したライン情報（仕様補完用）

    Returns:
        (フィルタ済み仕様, フィルタ済み需要)
    """
    print("\nデータマージ処理:")

    # 需要があるが仕様がない部品
    demand_only = set(demands.keys()) - set(specs.keys())

    # 生産計画のライン情報から仕様を自動補完
    auto_generated = 0
    if plan_infos and demand_only:
        for pn in demand_only:
            if pn in plan_infos:
                info = plan_infos[pn]
                if info.main_line:
                    specs[pn] = PartSpec(
                        part_number=pn,
                        part_name=info.part_name,
                        main_line=info.main_line,
                        sub1_line=info.sub1_line,
                        sub2_line=info.sub2_line,
                    )
                    auto_generated += 1

    if auto_generated > 0:
        print(f"  生産計画から仕様を自動補完: {auto_generated}件")

    # 補完後に再計算
    demand_only = set(demands.keys()) - set(specs.keys())
    if demand_only:
        print(f"  警告: 需要はあるが仕様がない部品: {len(demand_only)}件")
        for pn in sorted(list(demand_only))[:10]:
            print(f"    - {pn}")
        if len(demand_only) > 10:
            print(f"    ... 他{len(demand_only) - 10}件")

    # 仕様はあるが需要がない部品
    spec_only = set(specs.keys()) - set(demands.keys())
    if spec_only:
        print(f"  情報: 仕様はあるが需要がない部品: {len(spec_only)}件")

    # 両方あるものだけ使用
    common = set(specs.keys()) & set(demands.keys())
    print(f"  マッチした部品数: {len(common)}")

    filtered_specs = {k: v for k, v in specs.items() if k in common}
    filtered_demands = {k: v for k, v in demands.items() if k in common}

    return filtered_specs, filtered_demands


def load_all_data() -> tuple[dict[str, PartSpec], dict[str, PartDemand]]:
    """
    全データを読み込んでマージ

    Returns:
        (部品仕様辞書, 部品需要辞書)
    """
    specs = load_equipment_spec()
    demands, plan_infos = load_production_plan()
    return merge_data(specs, demands, plan_infos)


if __name__ == '__main__':
    # テスト実行
    specs, demands = load_all_data()

    print("\n=== データサマリー ===")
    print(f"部品数: {len(specs)}")

    # ライン別部品数
    line_counts = {line: 0 for line in DISC_LINES}
    for spec in specs.values():
        if spec.main_line:
            line_counts[spec.main_line] = line_counts.get(spec.main_line, 0) + 1

    print("\nライン別メイン割当部品数:")
    for line in DISC_LINES:
        print(f"  {line}: {line_counts[line]}件")

    # 月別総需要
    monthly_totals = [0] * 12
    for demand in demands.values():
        for i, qty in enumerate(demand.monthly_demand):
            monthly_totals[i] += qty

    print("\n月別総需要:")
    month_names = ['4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月', '3月']
    for name, total in zip(month_names, monthly_totals):
        print(f"  {name}: {total:,}")

    print(f"\n年間合計: {sum(monthly_totals):,}")
