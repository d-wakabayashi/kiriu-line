"""
KIRIU ライン負荷最適化システム - 設定ファイル
"""

from dataclasses import dataclass

# ディスクライン一覧
DISC_LINES = ['4915', '4919', '4927', '4928', '4934', '4935', '4945', '4G01', '4J01']

# デフォルトライン能力（月間生産可能数量）
# 注: 実際の能力値は後で調整可能
DEFAULT_CAPACITIES = {
    '4915': 70000,
    '4919': 80000,
    '4927': 40000,
    '4928': 40000,
    '4934': 50000,
    '4935': 85000,
    '4945': 50000,
    '4G01': 50000,
    '4J01': 10000,
}

# デフォルト JPH（Jobs Per Hour: 時間あたり生産数）
DEFAULT_JPH = {
    '4915': 350,
    '4919': 400,
    '4927': 200,
    '4928': 200,
    '4934': 250,
    '4935': 425,
    '4945': 250,
    '4G01': 250,
    '4J01': 50,
}


@dataclass
class WorkPattern:
    """勤務体制パターン"""
    name: str           # 例: "2直2交替"
    formula: str        # 例: "{月間稼働日数} * 7.5 * 2 - {月除外時間}"
    exclusion_hours: float  # 月除外時間（例: 5）


# デフォルト勤務体制パターン
DEFAULT_WORK_PATTERNS = [
    WorkPattern(name='2直2交替', formula='{月間稼働日数} * 7.5 * 2 - {月除外時間}', exclusion_hours=5),
    WorkPattern(name='3直3交替', formula='{月間稼働日数} * 7.5 * 3 - {月除外時間}', exclusion_hours=8),
]

# デフォルト月間稼働日数（4月〜3月）
DEFAULT_MONTHLY_WORKING_DAYS = [20, 19, 21, 22, 21, 20, 22, 19, 21, 20, 18, 21]


def evaluate_work_formula(formula: str, days: float, exclusion: float) -> float:
    """
    勤務体制の月稼働時間数式を安全に評価する。

    Args:
        formula: 数式文字列（例: "{月間稼働日数} * 7.5 * 2 - {月除外時間}"）
        days: 月間稼働日数
        exclusion: 月除外時間

    Returns:
        月稼働時間
    """
    expr = formula.replace('{月間稼働日数}', str(float(days)))
    expr = expr.replace('{月除外時間}', str(float(exclusion)))

    # 安全な評価: 数字、演算子、空白、小数点のみ許可
    allowed = set('0123456789.+-*/ ()')
    if not all(c in allowed for c in expr):
        raise ValueError(f'数式に不正な文字が含まれています: {formula}')

    return float(eval(expr))


def calculate_monthly_capacities(
    jph: dict[str, float],
    patterns: list[WorkPattern],
    monthly_working_days: list[float],
) -> dict[str, dict[str, list[int]]]:
    """
    勤務体制パターンごとにライン別月別能力を計算する。

    Args:
        jph: {ライン名: JPH値}
        patterns: 勤務体制パターンのリスト
        monthly_working_days: 12ヶ月分の月間稼働日数

    Returns:
        {パターン名: {ライン名: [月別能力 x12]}}
    """
    result = {}
    for pattern in patterns:
        caps = {}
        for line in DISC_LINES:
            line_jph = jph.get(line, DEFAULT_JPH.get(line, 0))
            monthly = []
            for month_idx in range(12):
                days = monthly_working_days[month_idx] if month_idx < len(monthly_working_days) else 20
                hours = evaluate_work_formula(pattern.formula, days, pattern.exclusion_hours)
                capacity = int(line_jph * hours)
                monthly.append(max(0, capacity))
            caps[line] = monthly
        result[pattern.name] = caps
    return result

# 目的関数の重み係数
WEIGHT_OVERFLOW = 10000   # オーバーフロー最小化（最優先）
WEIGHT_SUB_USE = 100      # サブライン使用回数最小化（第2優先）
WEIGHT_SUB_QTY = 1        # サブラインへの移動量最小化（第3優先）

# ソルバー設定
DEFAULT_TIME_LIMIT_SECONDS = 300  # 5分

# 月名（4月始まり）
MONTHS = ['4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月', '3月']

# データファイルパス
SPEC_FILE = '/home/user/kiriu-line/KRA部品-設備仕様確認.xlsx'
PLAN_FILE = '/home/user/kiriu-line/2026通期事計加工計画20251104 165220(配付） .xlsm'
PLAN_SHEET = '赤堀分工場2029上期'

# 出力ディレクトリ
OUTPUT_DIR = '/home/user/kiriu-line/output'

# Google Spreadsheet ID
DEFAULT_SPREADSHEET_ID = '1xBDJkTcQmR0vzuupD36slF8TziTc3EiNzLvArCIMBRM'
