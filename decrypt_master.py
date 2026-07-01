#!/usr/bin/env python3
"""Giải mã file Excel HOYA - kéo thả file vào đây"""
import msoffcrypto, sys, os

# Đường dẫn mặc định
src = os.path.expanduser("~/Downloads/Detail Spare parts FY2026 new.xlsx")
dst = os.path.expanduser("~/Desktop/Detail_Spare_parts_DECRYPTED.xlsx")
password = "sp"

# Nếu có argument
if len(sys.argv) > 1:
    src = sys.argv[1]
if len(sys.argv) > 2:
    dst = sys.argv[2]
if len(sys.argv) > 3:
    password = sys.argv[3]

try:
    with open(src, 'rb') as f:
        file = msoffcrypto.OfficeFile(f)
        file.load_key(password=password)
        with open(dst, 'wb') as out:
            file.decrypt(out)
    print(f"✅ Giải mã thành công!")
    print(f"📁 File gốc: {src}")
    print(f"📁 File mới: {dst}")
    print(f"📏 Dung lượng: {os.path.getsize(dst)/1024:.0f} KB")
except Exception as e:
    print(f"❌ Lỗi: {e}")
