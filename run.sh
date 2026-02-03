#!/bin/bash
#
# KIRIU ライン負荷最適化システム - 実行スクリプト
#
# 使用方法:
#   ./run.sh                    # デフォルト設定で実行
#   ./run.sh template           # テンプレートモードで実行
#   ./run.sh generate           # テンプレート生成
#   ./run.sh --help             # ヘルプ表示
#

set -e

# スクリプトのディレクトリに移動
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# 仮想環境のパス
VENV_DIR=".venv"
PYTHON="$VENV_DIR/bin/python"

# 仮想環境の確認
if [ ! -f "$PYTHON" ]; then
    echo "エラー: 仮想環境が見つかりません"
    echo "以下のコマンドでセットアップしてください:"
    echo "  uv venv .venv && uv pip install -r requirements.txt"
    exit 1
fi

# 引数に応じた処理
case "${1:-}" in
    template)
        # テンプレートモードで実行
        TEMPLATE_FILE="${2:-input_template.xlsx}"
        if [ ! -f "$TEMPLATE_FILE" ]; then
            echo "エラー: テンプレートファイルが見つかりません: $TEMPLATE_FILE"
            echo "先に ./run.sh generate で生成してください"
            exit 1
        fi
        echo "テンプレートモードで実行: $TEMPLATE_FILE"
        "$PYTHON" main.py --template "$TEMPLATE_FILE"
        ;;

    generate)
        # テンプレート生成
        OUTPUT_FILE="${2:-input_template.xlsx}"
        echo "テンプレートを生成: $OUTPUT_FILE"
        "$PYTHON" main.py --generate-template --template-output "$OUTPUT_FILE"
        ;;

    dry-run)
        # ドライラン
        echo "ドライラン（データ読み込みのみ）"
        "$PYTHON" main.py --dry-run
        ;;

    --help|-h|help)
        echo "KIRIU ライン負荷最適化システム"
        echo ""
        echo "使用方法:"
        echo "  ./run.sh                     デフォルト設定で実行"
        echo "  ./run.sh template [file]     テンプレートモードで実行"
        echo "  ./run.sh generate [file]     テンプレート生成"
        echo "  ./run.sh dry-run             ドライラン（データ確認のみ）"
        echo "  ./run.sh --help              このヘルプを表示"
        echo ""
        echo "詳細オプション:"
        "$PYTHON" main.py --help
        ;;

    *)
        # デフォルト実行（引数をそのまま渡す）
        "$PYTHON" main.py "$@"
        ;;
esac
