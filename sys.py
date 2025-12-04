import subprocess
import ctypes
import os
from pathlib import Path
import time
import json
import sys
import threading

# Mặc định
DEFAULT_LABEL = "HDD-DATA-DE160574"
ISO_NAME = "Office_professional_plus_2021_x86_x64_dvd_c6dd6dc6.iso"

# Đường dẫn installer
FOXIT_EXE = Path(r"C:\WINDOWS UPDATE BLOCKER NEW\FoxitReader501.0523_enu_Setup.exe")
CHROME_EXE = Path(r"C:\WINDOWS UPDATE BLOCKER NEW\ChromeSetup.exe")


# ========================================
# 1️⃣ Kiểm tra & nâng quyền Admin
# ========================================
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


if not is_admin():
    print("🔒 Đang yêu cầu quyền Administrator…")
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, __file__, None, 1
    )
    sys.exit()


# ========================================
# 2️⃣ Tìm phân vùng theo Label
# ========================================
def find_drive_by_label(label):
    print(f"🔍 Đang tìm phân vùng có tên: {label}\n")

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
        fs_label = vol.get("FileSystemLabel")

        if drive:
            drive_path = f"{drive}:\\"
            print(f"📌 {drive_path}  →  Label: {fs_label}")

            if fs_label.lower() == label.lower():
                print(f"\n🎯 Đã tìm thấy phân vùng: {drive_path}\n")
                return drive_path

    print(f"❌ Không tìm thấy phân vùng: {label}")
    return None


# ========================================
# 3️⃣ Cho người dùng chọn ổ
# ========================================
def choose_drive():
    print("Chọn ổ muốn dùng:")
    print("1️⃣  Dùng mặc định:", DEFAULT_LABEL)
    print("2️⃣  Nhập tên ổ/label khác")

    choice = input("Nhập 1 hoặc 2: ").strip()
    if choice == "1":
        label = DEFAULT_LABEL
    elif choice == "2":
        label = input("Nhập tên ổ/label: ").strip()
    else:
        print("❌ Lựa chọn không hợp lệ!")
        sys.exit()

    drive = find_drive_by_label(label)
    if not drive:
        input("Nhấn Enter để thoát…")
        sys.exit()

    return drive


# ========================================
# 4️⃣ Mount file ISO
# ========================================
def check_and_mount_iso(drive_path):
    iso_path = Path(drive_path) / ISO_NAME

    if not iso_path.exists():
        print(f"❌ Không tìm thấy ISO tại:\n{iso_path}")
        input("Nhấn Enter để thoát…")
        sys.exit()

    print(f"📌 Tìm thấy ISO:\n{iso_path}")
    print("📀 Đang mount ISO…")

    cmd = f'Mount-DiskImage -ImagePath "{iso_path}"'
    subprocess.run(["powershell", "-Command", cmd])

    time.sleep(2)
    return iso_path


# ========================================
# 5️⃣ Chạy setup Office từ ổ mount
# ========================================
def run_office_setup():
    print("🔍 Đang tìm setup.exe trong ổ mount…")

    cmd = '(Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object {$_.DriveType -eq 5}).DeviceID'
    result = subprocess.check_output(["powershell", "-Command", cmd], text=True)
    drives = [d.strip() for d in result.splitlines() if d.strip()]

    for d in drives:
        setup_path = Path(d) / "setup.exe"

        if setup_path.exists():
            print(f"🚀 Đang chạy Office setup tại: {setup_path}")
            os.chdir(d)
            subprocess.Popen(str(setup_path))
            print("👉 Office installer đang chạy…")
            return True

    print("⚠ Không tìm thấy setup.exe trong ổ mount!")
    return False


# ========================================
# 6️⃣ Active Office
# ========================================
def run_activation_script():
    print("🔑 Đang kích hoạt Office…")

    cmd = 'iex (curl.exe -s --doh-url https://1.1.1.1/dns-query https://get.activated.win | Out-String)'
    subprocess.run(["powershell", "-Command", cmd])

    print("✔ Office đã kích hoạt xong!")


# ========================================
# 7️⃣ Hiện biểu tượng This PC
# ========================================
def enable_this_pc_icon():
    try:
        cmds = [
            r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" '
            r'/v "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" /t REG_DWORD /d 0 /f',

            r'reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu" '
            r'/v "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" /t REG_DWORD /d 0 /f'
        ]

        for c in cmds:
            subprocess.run(c, shell=True)

        subprocess.run("RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", shell=True)
        subprocess.run("taskkill /f /im explorer.exe && start explorer.exe", shell=True)

        print("✔ Đã hiện This PC trên Desktop")

    except Exception as e:
        print(f"❌ Lỗi This PC: {e}")


# ========================================
# 8️⃣ Cài Foxit Silent
# ========================================
def install_foxit():
    if not FOXIT_EXE.exists():
        print("⚠ Không tìm thấy Foxit installer!")
        return False

    print("🚫 Tắt PUAProtection…")
    subprocess.run(["powershell", "-Command", "Set-MpPreference -PUAProtection 0"])

    print("📦 Cài Foxit PDF Reader (silent)…")

    try:
        subprocess.run(
            [
                str(FOXIT_EXE),
                "/silent",
                "/install",
                "/norestart"
            ],
            check=True
        )
        print("✔ Foxit đã cài xong!")

    except:
        print("❌ Lỗi cài Foxit")
        return False

    print("🔒 Bật lại PUAProtection…")
    subprocess.run(["powershell", "-Command", "Set-MpPreference -PUAProtection 1"])

    return True


# ========================================
# 9️⃣ Cài Google Chrome Silent
# ========================================
def install_chrome():
    if not CHROME_EXE.exists():
        print("⚠ Không tìm thấy Chrome installer!")
        return False

    print("🌐 Cài Google Chrome (silent)…")

    try:
        subprocess.run(
            [
                str(CHROME_EXE),
                "/install",
                "--do-not-launch-chrome"
            ],
            check=True
        )
        print("✔ Chrome đã cài xong!")
        return True

    except:
        print("❌ Lỗi cài Chrome")
        return False


# ========================================
# 10️⃣ MAIN – CHẠY SONG SONG
# ========================================
if __name__ == "__main__":
    print("===== 🚀 AUTO INSTALL OFFICE 2021 PRO 🚀 =====\n")

    drive = choose_drive()

    check_and_mount_iso(drive)

    print("\n🚀 Đang cài song song: Chrome + Foxit + Office…\n")

    # Tạo thread
    threads = [
        threading.Thread(target=install_chrome, name="Chrome"),
        threading.Thread(target=install_foxit, name="Foxit"),
        threading.Thread(target=run_office_setup, name="Office")
    ]

    # Chạy thread
    for t in threads:
        t.start()

    # Chỉ chờ Chrome và Foxit – Office chạy cửa sổ riêng
    for t in threads:
        if t.name != "Office":
            t.join()

    # Kích hoạt Office
    run_activation_script()

    enable_this_pc_icon()

    input("\n🎯 Hoàn tất cài đặt! Nhấn Enter để thoát…")
