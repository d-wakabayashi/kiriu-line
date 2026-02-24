# KIRIU ライン負荷最適化API - Dockerfile
FROM python:3.12-slim

WORKDIR /app

# 依存関係のインストール
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt uvicorn fastapi

# アプリケーションコードのコピー
COPY config.py .
COPY data_loader.py .
COPY model.py .
COPY sheets_io.py .
COPY api.py .
COPY excel_output.py .
COPY visualize.py .
COPY main.py .
COPY input_template.py .
COPY output_handler.py .

# ポート設定
ENV PORT=8080
EXPOSE 8080

# 起動コマンド
CMD ["python", "-m", "uvicorn", "api:app", "--host", "0.0.0.0", "--port", "8080"]
