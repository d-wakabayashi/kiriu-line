"""
KIRIU ライン負荷最適化システム - 出力ハンドラモジュール

各種出力先（ローカル、Google Drive、メール）への結果送信を管理
"""

import os
import smtplib
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from dataclasses import dataclass
from typing import Callable


@dataclass
class OutputResult:
    """出力結果"""
    success: bool
    message: str
    details: dict | None = None


class OutputHandler:
    """出力ハンドラ"""

    def __init__(self, output_dir: str = './output'):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.generated_files: list[str] = []

    def add_file(self, filepath: str):
        """生成されたファイルを登録"""
        self.generated_files.append(filepath)

    def get_files(self) -> list[str]:
        """生成されたファイル一覧を取得"""
        return self.generated_files.copy()


# ============================================================
# Google Drive 出力
# ============================================================

def upload_to_google_drive(
    files: list[str],
    folder_id: str,
    credentials_path: str | None = None,
) -> OutputResult:
    """
    Google Driveにファイルをアップロード

    Args:
        files: アップロードするファイルパスのリスト
        folder_id: アップロード先のGoogleドライブフォルダID
        credentials_path: サービスアカウント認証情報JSONのパス

    Returns:
        OutputResult

    必要な環境変数:
        GOOGLE_APPLICATION_CREDENTIALS: 認証情報JSONのパス（credentials_pathが指定されていない場合）
    """
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
    except ImportError:
        return OutputResult(
            success=False,
            message='Google Drive API ライブラリがインストールされていません。'
                    '\n以下のコマンドでインストールしてください:'
                    '\npip install google-api-python-client google-auth',
        )

    # 認証情報の取得
    creds_path = credentials_path or os.environ.get('GOOGLE_APPLICATION_CREDENTIALS')
    if not creds_path:
        return OutputResult(
            success=False,
            message='Google Drive認証情報が設定されていません。'
                    '\nGOOGLE_APPLICATION_CREDENTIALS環境変数を設定するか、'
                    '\n認証情報ファイルパスを指定してください。',
        )

    if not Path(creds_path).exists():
        return OutputResult(
            success=False,
            message=f'認証情報ファイルが見つかりません: {creds_path}',
        )

    try:
        # サービスアカウント認証
        credentials = service_account.Credentials.from_service_account_file(
            creds_path,
            scopes=['https://www.googleapis.com/auth/drive.file']
        )

        service = build('drive', 'v3', credentials=credentials)

        uploaded_files = []
        for filepath in files:
            if not Path(filepath).exists():
                continue

            filename = Path(filepath).name

            # MIMEタイプの判定
            ext = Path(filepath).suffix.lower()
            mime_types = {
                '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                '.pdf': 'application/pdf',
                '.json': 'application/json',
                '.txt': 'text/plain',
                '.png': 'image/png',
            }
            mime_type = mime_types.get(ext, 'application/octet-stream')

            file_metadata = {
                'name': filename,
                'parents': [folder_id],
            }
            media = MediaFileUpload(filepath, mimetype=mime_type)

            file = service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id, name, webViewLink'
            ).execute()

            uploaded_files.append({
                'name': file.get('name'),
                'id': file.get('id'),
                'link': file.get('webViewLink'),
            })

        return OutputResult(
            success=True,
            message=f'{len(uploaded_files)}件のファイルをGoogle Driveにアップロードしました。',
            details={'uploaded_files': uploaded_files},
        )

    except Exception as e:
        return OutputResult(
            success=False,
            message=f'Google Driveへのアップロードに失敗しました: {str(e)}',
        )


# ============================================================
# メール送信
# ============================================================

@dataclass
class EmailConfig:
    """メール設定"""
    smtp_server: str
    smtp_port: int
    username: str
    password: str
    from_address: str
    use_tls: bool = True


def load_email_config_from_env() -> EmailConfig | None:
    """環境変数からメール設定を読み込む"""
    required = ['SMTP_SERVER', 'SMTP_USERNAME', 'SMTP_PASSWORD', 'SMTP_FROM']
    missing = [k for k in required if not os.environ.get(k)]

    if missing:
        return None

    return EmailConfig(
        smtp_server=os.environ.get('SMTP_SERVER', ''),
        smtp_port=int(os.environ.get('SMTP_PORT', '587')),
        username=os.environ.get('SMTP_USERNAME', ''),
        password=os.environ.get('SMTP_PASSWORD', ''),
        from_address=os.environ.get('SMTP_FROM', ''),
        use_tls=os.environ.get('SMTP_USE_TLS', 'true').lower() == 'true',
    )


def send_email(
    to_addresses: list[str],
    subject: str,
    body: str,
    attachments: list[str] | None = None,
    config: EmailConfig | None = None,
) -> OutputResult:
    """
    メールを送信

    Args:
        to_addresses: 送信先メールアドレスのリスト
        subject: 件名
        body: 本文
        attachments: 添付ファイルパスのリスト
        config: メール設定（省略時は環境変数から取得）

    Returns:
        OutputResult

    必要な環境変数（configが指定されていない場合）:
        SMTP_SERVER: SMTPサーバーアドレス
        SMTP_PORT: SMTPポート（デフォルト: 587）
        SMTP_USERNAME: 認証ユーザー名
        SMTP_PASSWORD: 認証パスワード
        SMTP_FROM: 送信元アドレス
        SMTP_USE_TLS: TLS使用（デフォルト: true）
    """
    email_config = config or load_email_config_from_env()

    if not email_config:
        return OutputResult(
            success=False,
            message='メール設定が見つかりません。'
                    '\n以下の環境変数を設定してください:'
                    '\nSMTP_SERVER, SMTP_USERNAME, SMTP_PASSWORD, SMTP_FROM',
        )

    try:
        # メッセージ作成
        msg = MIMEMultipart()
        msg['From'] = email_config.from_address
        msg['To'] = ', '.join(to_addresses)
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        # 添付ファイル
        if attachments:
            for filepath in attachments:
                if not Path(filepath).exists():
                    continue

                with open(filepath, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename="{Path(filepath).name}"'
                    )
                    msg.attach(part)

        # 送信
        with smtplib.SMTP(email_config.smtp_server, email_config.smtp_port) as server:
            if email_config.use_tls:
                server.starttls()
            server.login(email_config.username, email_config.password)
            server.send_message(msg)

        return OutputResult(
            success=True,
            message=f'メールを送信しました: {", ".join(to_addresses)}',
            details={
                'to': to_addresses,
                'subject': subject,
                'attachments': [Path(f).name for f in (attachments or []) if Path(f).exists()],
            },
        )

    except smtplib.SMTPAuthenticationError:
        return OutputResult(
            success=False,
            message='SMTP認証に失敗しました。ユーザー名とパスワードを確認してください。',
        )
    except smtplib.SMTPException as e:
        return OutputResult(
            success=False,
            message=f'メール送信に失敗しました: {str(e)}',
        )
    except Exception as e:
        return OutputResult(
            success=False,
            message=f'予期しないエラーが発生しました: {str(e)}',
        )


def create_result_email_body(
    status: str,
    objective_value: float | None,
    solve_time: float,
    summary: str,
) -> str:
    """最適化結果のメール本文を生成"""
    body = f"""
KIRIU ライン負荷最適化 - 実行結果レポート
{'=' * 50}

【最適化結果】
  ステータス: {status}
  目的関数値: {objective_value:,.0f if objective_value else 'N/A'}
  実行時間: {solve_time:.2f}秒

{summary}

詳細は添付のExcelファイルをご確認ください。

---
このメールはKIRIU ライン負荷最適化システムから自動送信されています。
"""
    return body


# ============================================================
# 統合出力関数
# ============================================================

def process_outputs(
    files: list[str],
    output_to_gdrive: bool = False,
    gdrive_folder_id: str = '',
    send_email_flag: bool = False,
    email_to: str = '',
    email_subject: str = 'ライン負荷最適化結果',
    email_body: str = '',
) -> dict[str, OutputResult]:
    """
    各出力先への処理を実行

    Args:
        files: 出力ファイルのリスト
        output_to_gdrive: Google Driveに出力するか
        gdrive_folder_id: Google DriveフォルダID
        send_email_flag: メール送信するか
        email_to: 送信先メールアドレス（カンマ区切りで複数指定可）
        email_subject: メール件名
        email_body: メール本文

    Returns:
        {出力先: OutputResult}
    """
    results = {}

    # Google Drive出力
    if output_to_gdrive and gdrive_folder_id:
        print("\nGoogle Driveへアップロード中...")
        results['google_drive'] = upload_to_google_drive(files, gdrive_folder_id)
        if results['google_drive'].success:
            print(f"  {results['google_drive'].message}")
        else:
            print(f"  エラー: {results['google_drive'].message}")

    # メール送信
    if send_email_flag and email_to:
        print("\nメール送信中...")
        to_list = [addr.strip() for addr in email_to.split(',') if addr.strip()]
        results['email'] = send_email(
            to_addresses=to_list,
            subject=email_subject,
            body=email_body,
            attachments=files,
        )
        if results['email'].success:
            print(f"  {results['email'].message}")
        else:
            print(f"  エラー: {results['email'].message}")

    return results


if __name__ == '__main__':
    # テスト
    print("出力ハンドラモジュール")
    print("\n環境変数の設定状況:")
    print(f"  GOOGLE_APPLICATION_CREDENTIALS: {'設定済み' if os.environ.get('GOOGLE_APPLICATION_CREDENTIALS') else '未設定'}")
    print(f"  SMTP_SERVER: {'設定済み' if os.environ.get('SMTP_SERVER') else '未設定'}")
