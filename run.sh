#!/usr/bin/env bash
# Công cụ hỗ trợ kiểm tra thuế — script chạy cho macOS/Linux
set -e
cd "$(dirname "$0")"

echo "============================================"
echo "  CÔNG CỤ HỖ TRỢ KIỂM TRA THUẾ"
echo "============================================"

# 1) Kiểm tra Python
if ! command -v python3 >/dev/null 2>&1; then
    echo "[LỖI] Chưa cài đặt Python3."
    echo "Tải tại: https://www.python.org/downloads/"
    exit 1
fi

# 2) Tạo môi trường ảo lần đầu
if [ ! -d ".venv" ]; then
    echo "Đang chuẩn bị lần đầu, vui lòng đợi..."
    python3 -m venv .venv
fi

# 3) Kích hoạt và cài thư viện
source .venv/bin/activate
echo "Đang kiểm tra / cài đặt thư viện..."
python -m pip install --quiet --upgrade pip
python -m pip install --quiet -r requirements.txt

# 4) Chạy ứng dụng
echo ""
echo "Đang mở ứng dụng trên trình duyệt tại http://localhost:8502 (đóng: Ctrl+C)"
echo ""
streamlit run app.py --server.port 8502
