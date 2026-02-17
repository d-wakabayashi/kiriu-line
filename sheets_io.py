"""
KIRIU ライン負荷最適化システム - Google Sheets I/Oモジュール

~/.clasprc.json から OAuth client_id/secret を取得し、
gspread + google-auth-oauthlib で認証してスプレッドシートを読み書きする。
"""

import json
from pathlib import Path

import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

from config import DISC_LINES, DEFAULT_CAPACITIES, MONTHS, DEFAULT_SPREADSHEET_ID
from data_loader import PartSpec, PartDemand, normalize_part_number, normalize_line_name

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
]

CLASPRC_PATH = Path.home() / '.clasprc.json'
TOKEN_PATH = Path.home() / '.kiriu-line-token.json'

INPUT_SHEET_NAME = '入力シート'
LINE_CAPACITY_SHEET_NAME = 'ライン能力'


def get_client() -> gspread.Client:
    """
    OAuth2 認証して gspread クライアントを返す。

    1. ~/.kiriu-line-token.json にキャッシュがあれば再利用
    2. なければ ~/.clasprc.json から client_id/secret を読み、認証フローを実行
    3. トークンを ~/.kiriu-line-token.json に保存
    """
    creds = None

    # キャッシュされたトークンを試す
    if TOKEN_PATH.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
        except Exception:
            creds = None

    # トークンがない or 期限切れ
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception:
            creds = None

    if not creds or not creds.valid:
        # ~/.clasprc.json から client_id/secret を取得
        if not CLASPRC_PATH.exists():
            raise FileNotFoundError(
                f'{CLASPRC_PATH} が見つかりません。\n'
                'clasp login を実行するか、OAuth クライアント情報を設定してください。'
            )

        with open(CLASPRC_PATH, 'r', encoding='utf-8') as f:
            clasp_config = json.load(f)

        # clasprc.json の oauth2 credentials を取得
        oauth2 = clasp_config.get('oauth2ClientSettings') or clasp_config.get('token', {})
        client_id = oauth2.get('clientId') or oauth2.get('client_id')
        client_secret = oauth2.get('clientSecret') or oauth2.get('client_secret')

        if not client_id or not client_secret:
            raise ValueError(
                f'{CLASPRC_PATH} に clientId/clientSecret が見つかりません。'
            )

        client_config = {
            'installed': {
                'client_id': client_id,
                'client_secret': client_secret,
                'auth_uri': 'https://accounts.google.com/o/oauth2/auth',
                'token_uri': 'https://oauth2.googleapis.com/token',
                'redirect_uris': ['http://localhost'],
            }
        }

        flow = InstalledAppFlow.from_client_config(client_config, SCOPES)

        print('ブラウザで Google 認証を行います...')
        print('ブラウザが開かない場合は、以下のURLにアクセスしてください。')
        creds = flow.run_local_server(port=0)

        # トークンを保存
        TOKEN_PATH.write_text(creds.to_json(), encoding='utf-8')
        print(f'トークンを保存しました: {TOKEN_PATH}')

    return gspread.authorize(creds)


def read_input_sheet(
    spreadsheet_id: str = DEFAULT_SPREADSHEET_ID,
) -> tuple[dict[str, PartSpec], dict[str, PartDemand]]:
    """
    入力シートから部品仕様と需要を読み込む。

    シート構造:
    | 部品番号 | 部品名 | メインライン | サブ1ライン | サブ2ライン | 4月 | ... | 3月 |

    同一部品番号+同一ラインの行は需要を合算する。

    Returns:
        (specs辞書, demands辞書)
    """
    client = get_client()
    sh = client.open_by_key(spreadsheet_id)
    ws = sh.worksheet(INPUT_SHEET_NAME)

    rows = ws.get_all_values()
    if len(rows) < 2:
        print('入力シートにデータがありません。')
        return {}, {}

    header = rows[0]
    print(f'入力シート読み込み: {len(rows) - 1}行')

    specs: dict[str, PartSpec] = {}
    demands: dict[str, PartDemand] = {}
    row_tracking: dict[str, list[tuple[str, int]]] = {}

    for row_idx, row in enumerate(rows[1:], start=2):
        if len(row) < 17:
            row += [''] * (17 - len(row))

        part_num = normalize_part_number(row[0])
        if not part_num:
            continue

        part_name = str(row[1]).strip()
        main_line = normalize_line_name(row[2])
        sub1_line = normalize_line_name(row[3])
        sub2_line = normalize_line_name(row[4])

        # メインラインがディスクラインか確認
        if not main_line or main_line not in DISC_LINES:
            continue

        # サブラインがディスクラインでなければ None
        if sub1_line and sub1_line not in DISC_LINES:
            sub1_line = None
        if sub2_line and sub2_line not in DISC_LINES:
            sub2_line = None

        # 仕様を登録（最初に見つかったものを使用）
        if part_num not in specs:
            specs[part_num] = PartSpec(
                part_number=part_num,
                part_name=part_name,
                main_line=main_line,
                sub1_line=sub1_line,
                sub2_line=sub2_line,
            )

        # 月別需要（列5〜16）
        monthly = []
        for col_idx in range(5, 17):
            val = row[col_idx] if col_idx < len(row) else ''
            try:
                val = int(float(str(val).replace(',', '').strip())) if val else 0
            except (ValueError, TypeError):
                val = 0
            monthly.append(max(0, val))

        if sum(monthly) > 0:
            # 行トラッキング
            if part_num not in row_tracking:
                row_tracking[part_num] = []
            row_tracking[part_num].append((main_line, row_idx))

            # 同一部品番号の複数行は需要を合算
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

    # 合算された部品を表示
    consolidated = {pn: entries for pn, entries in row_tracking.items() if len(entries) > 1}
    if consolidated:
        print(f'  同一部品番号の複数行を合算: {len(consolidated)}件')
        for pn, entries in sorted(consolidated.items()):
            lines = set(e[0] for e in entries)
            line_str = ', '.join(sorted(lines))
            row_nums = ', '.join(str(e[1]) for e in entries)
            total = sum(demands[pn].monthly_demand)
            print(f'    {pn} (ライン: {line_str}): {len(entries)}行(行{row_nums}) → 合算後年間需要 {total:,}')

    print(f'  読み込み部品数: {len(demands)}')
    total_demand = sum(sum(d.monthly_demand) for d in demands.values())
    print(f'  年間総需要: {total_demand:,}')

    return specs, demands


def read_line_capacities(
    spreadsheet_id: str = DEFAULT_SPREADSHEET_ID,
) -> dict[str, list[int]]:
    """
    ライン能力シートからライン能力を読み込む。

    シート構造:
    | ライン | 4月 | 5月 | ... | 3月 |

    Returns:
        {ライン名: [月別能力 x12]}
    """
    client = get_client()
    sh = client.open_by_key(spreadsheet_id)
    ws = sh.worksheet(LINE_CAPACITY_SHEET_NAME)

    rows = ws.get_all_values()
    if len(rows) < 2:
        print('ライン能力シートにデータがありません。デフォルト値を使用します。')
        return {line: [cap] * 12 for line, cap in DEFAULT_CAPACITIES.items()}

    capacities: dict[str, list[int]] = {}

    for row in rows[1:]:
        if len(row) < 13:
            row += [''] * (13 - len(row))

        line_name = str(row[0]).strip()
        if line_name not in DISC_LINES:
            continue

        monthly = []
        for col_idx in range(1, 13):
            val = row[col_idx] if col_idx < len(row) else ''
            try:
                val = int(float(str(val).replace(',', '').strip())) if val else 0
            except (ValueError, TypeError):
                val = 0
            monthly.append(max(0, val))

        capacities[line_name] = monthly

    # 不足ラインにデフォルト値を補完
    for line in DISC_LINES:
        if line not in capacities:
            default_cap = DEFAULT_CAPACITIES.get(line, 0)
            capacities[line] = [default_cap] * 12

    print('ライン能力読み込み:')
    for line in DISC_LINES:
        avg = sum(capacities[line]) // 12
        print(f'  {line}: 平均 {avg:,}/月')

    return capacities


def setup_template(spreadsheet_id: str = DEFAULT_SPREADSHEET_ID) -> None:
    """
    スプレッドシートにテンプレート（ヘッダー行とデフォルトライン能力）をセットアップする。

    - 入力シート: ヘッダー行を書き込み
    - ライン能力シート: ヘッダー行 + 9ライン分のデフォルト値を書き込み
    """
    client = get_client()
    sh = client.open_by_key(spreadsheet_id)

    # --- 入力シート ---
    try:
        ws_input = sh.worksheet(INPUT_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws_input = sh.add_worksheet(title=INPUT_SHEET_NAME, rows=500, cols=17)

    input_header = ['部品番号', '部品名', 'メインライン', 'サブ1ライン', 'サブ2ライン'] + MONTHS
    ws_input.update('A1', [input_header])
    print(f'入力シートにヘッダーを書き込みました: {INPUT_SHEET_NAME}')

    # --- ライン能力シート ---
    try:
        ws_cap = sh.worksheet(LINE_CAPACITY_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws_cap = sh.add_worksheet(title=LINE_CAPACITY_SHEET_NAME, rows=20, cols=13)

    cap_header = ['ライン'] + MONTHS
    cap_data = [cap_header]
    for line in DISC_LINES:
        default_cap = DEFAULT_CAPACITIES.get(line, 0)
        cap_data.append([line] + [default_cap] * 12)

    ws_cap.update('A1', cap_data)
    print(f'ライン能力シートにデフォルト値を書き込みました: {LINE_CAPACITY_SHEET_NAME}')

    print('\nセットアップ完了。')
    print(f'スプレッドシート: https://docs.google.com/spreadsheets/d/{spreadsheet_id}')
    print('入力シートにデータを入力してから --spreadsheet で最適化を実行してください。')


def write_results(
    spreadsheet_id: str,
    result,
    specs: dict[str, PartSpec],
    capacities: dict[str, list[int]],
    sheet_suffix: str = '',
) -> None:
    """
    最適化結果をスプレッドシートの新シートに書き込む。

    Args:
        spreadsheet_id: スプレッドシートID
        result: OptimizationResult
        specs: 部品仕様辞書
        capacities: ライン能力辞書 {ライン: [月別能力 x12]}
        sheet_suffix: シート名のサフィックス（例: '_100pct'）
    """
    from model import OptimizationResult
    assert isinstance(result, OptimizationResult)

    client = get_client()
    sh = client.open_by_key(spreadsheet_id)

    sheet_name = f'最適化結果{sheet_suffix}'

    # 既存シートがあれば削除
    try:
        existing = sh.worksheet(sheet_name)
        sh.del_worksheet(existing)
    except gspread.WorksheetNotFound:
        pass

    ws = sh.add_worksheet(title=sheet_name, rows=2000, cols=20)

    all_rows: list[list] = []

    # --- サマリー ---
    all_rows.append(['KIRIU ライン負荷最適化結果' + sheet_suffix])
    all_rows.append([])
    all_rows.append(['ステータス', result.status])
    all_rows.append(['目的関数値', result.objective_value if result.objective_value else 'N/A'])
    all_rows.append(['実行時間（秒）', f'{result.solve_time:.2f}'])

    total_unmet = 0
    if result.unmet_demand:
        total_unmet = sum(sum(m) for m in result.unmet_demand.values())
    all_rows.append(['未割当数量合計', total_unmet])
    all_rows.append([])

    # --- 部品別割当表 ---
    all_rows.append(['【部品別ライン別割当】'])
    alloc_header = ['部品番号', '部品名', 'メインライン', '割当ライン'] + MONTHS + ['年間計']
    all_rows.append(alloc_header)

    for part_num in sorted(result.allocation.keys()):
        part_data = result.allocation[part_num]
        spec = specs.get(part_num)
        part_name = spec.part_name if spec else ''
        main_line = spec.main_line if spec else ''

        for line, monthly in part_data.items():
            if sum(monthly) == 0:
                continue
            row = [part_num, part_name, main_line, line] + list(monthly) + [sum(monthly)]
            all_rows.append(row)

    all_rows.append([])

    # --- 未割当 ---
    if result.unmet_demand:
        unmet_parts = {pn: m for pn, m in result.unmet_demand.items() if sum(m) > 0}
        if unmet_parts:
            all_rows.append(['【未割当一覧】'])
            unmet_header = ['部品番号', '部品名', 'メインライン'] + MONTHS + ['年間計']
            all_rows.append(unmet_header)
            for pn in sorted(unmet_parts.keys()):
                monthly_unmet = unmet_parts[pn]
                spec = specs.get(pn)
                part_name = spec.part_name if spec else ''
                main_line = spec.main_line if spec else ''
                row = [pn, part_name, main_line] + list(monthly_unmet) + [sum(monthly_unmet)]
                all_rows.append(row)
            all_rows.append([])

    # --- ライン別負荷率表 ---
    all_rows.append(['【ライン別負荷率】'])
    load_header = ['ライン', '項目'] + MONTHS + ['年間計', '平均']
    all_rows.append(load_header)

    for line in DISC_LINES:
        line_caps = capacities.get(line, [0] * 12)
        loads = result.line_loads.get(line, [0] * 12)

        # キャパシティ行
        all_rows.append([line, 'キャパシティ'] + line_caps + [sum(line_caps), sum(line_caps) // 12])

        # 生産数行
        all_rows.append(['', '生産数'] + list(loads) + [sum(loads), sum(loads) // 12])

        # 負荷率行
        rates = []
        for load, cap in zip(loads, line_caps):
            rates.append(f'{load / cap * 100:.1f}%' if cap > 0 else '0.0%')
        total_cap = sum(line_caps)
        avg_rate = f'{sum(loads) / total_cap * 100:.1f}%' if total_cap > 0 else '0.0%'
        all_rows.append(['', '負荷率'] + rates + ['', avg_rate])

        # 空行
        all_rows.append([])

    # 一括書き込み
    ws.update('A1', all_rows)

    print(f'結果シートを書き込みました: {sheet_name}')
