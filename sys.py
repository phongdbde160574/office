import subprocess
import ctypes
import os
from pathlib import Path
import time
import json
import sys

TARGET_LABEL = "HDD-DATA-DE160574"
ISO_NAME = "Office_professional_plus_2021_x86_x64_dvd_c6dd6dc6.iso"


# ==============================
# 1️⃣ Nâng quyền Admin
# ==============================
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


if not is_admin():
    print("🔒 Yêu cầu quyền Admin…")
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, __file__, None, 1
    )
    sys.exit()


# ==============================
# 2️⃣ Tìm ổ theo Label
# ==============================
def find_drive_by_label():
    print(f"🔍 Đang tìm phân vùng có tên: {TARGET_LABEL}\n")

    cmd = 'Get-Volume | Select-Object DriveLetter, FileSystemLabel | ConvertTo-Json'
    result = subprocess.check_output(["powershell", "-Command", cmd], text=True)

    try:
        volumes = json.loads(result)
        if isinstance(volumes, dict):
            volumes = [volumes]
    except:
        print("❌ Lỗi đọc Volume!")
        return None

    for vol in volumes:
        drive = vol.get("DriveLetter")
        label = vol.get("FileSystemLabel")

        if drive:
            drive_path = f"{drive}:\\"
            print(f"📌 {drive_path}  →  Label: {label}")
            if label == TARGET_LABEL:
                print(f"\n🎯 Tìm thấy phân vùng: {drive_path}\n")
                return drive_path

    print(f"❌ Không tìm thấy phân vùng: {TARGET_LABEL}\n")
    return None


# ==============================
# 3️⃣ Kiểm tra & Mount ISO
# ==============================
def check_and_mount_iso(drive_path):
    iso_path = Path(drive_path) / ISO_NAME

    if not iso_path.exists():
        print(f"❌ Không tìm thấy file ISO tại:\n{iso_path}")
        input("Nhấn Enter để thoát…")
        sys.exit()

    print(f"📌 Tìm thấy ISO:\n{iso_path}")
    print("📀 Đang mount ISO…")

    cmd = f'Mount-DiskImage -ImagePath "{iso_path}"'
    subprocess.run(["powershell", "-Command", cmd])
    time.sleep(2)
    return iso_path


# ==============================
# 4️⃣ Tìm ổ Mount & chạy setup.exe
# ==============================
def run_office_setup():
    print("🔍 Đang tìm setup.exe…")

    cmd = '(Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object {$_.DriveType -eq 5}).DeviceID'
    result = subprocess.check_output(["powershell", "-Command", cmd], text=True)
    drives = [d.strip() for d in result.splitlines() if d.strip()]

    for d in drives:
        setup_path = Path(d) / "setup.exe"
        if setup_path.exists():
            print(f"🚀 Chạy setup tại: {setup_path}")
            os.chdir(d)
            subprocess.Popen(str(setup_path))
            print("👉 Office setup đã chạy! Cài đặt tiếp trên cửa sổ mới.")
            return True

    print("⚠ Không tìm thấy setup.exe → Kích hoạt bằng script online!")
    return False


# ==============================
# 5️⃣ Dự phòng: Active Office
# ==============================
def run_activation_script():
    print("🔑 Đang chạy script kích hoạt Office…")

    cmd = 'iex (curl.exe -s --doh-url https://1.1.1.1/dns-query https://get.activated.win | Out-String)'
    subprocess.run(["powershell", "-Command", cmd])
    print("✔ Hoàn tất kích hoạt!")


# ==============================
# MAIN
# ==============================
if __name__ == "__main__":
    print("===== 🚀 AUTO INSTALL OFFICE 2021 🚀 =====\n")

    drive = find_drive_by_label()
    if not drive:
        input("Nhấn Enter để thoát…")
        sys.exit()

    check_and_mount_iso(drive)

    if not run_office_setup():
        run_activation_script()

    input("\n🎯 Xong! Nhấn Enter để thoát…")
