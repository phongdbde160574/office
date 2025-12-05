import subprocess
import ctypes
import os
import json
import sys
import shutil

# ========================================
# ⚙️ CẤU HÌNH
# ========================================
DEFAULT_LABEL = "HDD-DATA-DE160574"

INSTALL_FOLDER = r"\AutoCAD 2021"
INSTALL_FOLDER1 = r"\Autodesk\AutoCAD_2021_English_Win_64bit_dlm"
SETUP_EXE = "Setup.exe"

TARGET_DIR = r"C:\Program Files\Autodesk\AutoCAD 2021"
PATCHED_ACAD_RELATIVE_PATH = "acad.exe"


# ========================================
# 1️⃣ KIỂM TRA QUYỀN ADMIN
# ========================================
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


if not is_admin():
    print("Yêu cầu quyền Administrator…")
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    sys.exit()


# ========================================
# 2️⃣ TÌM PHÂN VÙNG THEO LABEL
# ========================================
def find_drive_by_label(label):
    cmd = 'Get-Volume | Select-Object DriveLetter, FileSystemLabel | ConvertTo-Json'
    try:
        result = subprocess.check_output(["powershell", "-Command", cmd], text=True, timeout=10)
        volumes = json.loads(result)
        if isinstance(volumes, dict):
            volumes = [volumes]
    except:
        print("Lỗi khi tìm phân vùng!")
        return None

    for vol in volumes:
        drive = vol.get("DriveLetter")
        fs_label = vol.get("FileSystemLabel")
        if drive:
            if fs_label and fs_label.lower() == label.lower():
                return f"{drive}:\\"
    return None


# ========================================
# 3️⃣ CHỌN PHÂN VÙNG
# ========================================
def choose_drive():
    print("Chọn ổ muốn dùng:")
    print(f"1. Dùng mặc định: {DEFAULT_LABEL}")
    print("2. Nhập tên ổ/label khác")

    choice = input("Nhập 1 hoặc 2: ").strip()
    label = DEFAULT_LABEL
    if choice == "2":
        label = input("Nhập tên ổ/label: ").strip()
        if not label:
            print("Label không hợp lệ.")
            sys.exit()

    drive = find_drive_by_label(label)
    if not drive:
        print("Không tìm thấy phân vùng!")
        sys.exit()
    return drive


# ========================================
# 4️⃣ COPY ACAD.EXE
# ========================================
def copy_acad(drive_path):
    install_dir = os.path.join(drive_path, INSTALL_FOLDER)
    source_acad_patch = os.path.join(install_dir, PATCHED_ACAD_RELATIVE_PATH)
    target_acad = os.path.join(TARGET_DIR, "acad.exe")
    print("install_dir: ", install_dir)
    print("source_acad_patch: ", source_acad_patch)
    print("target_acad: ", target_acad)
    if not os.path.exists(source_acad_patch):
        print("Không tìm thấy file acad.exe trong bộ cài!")
        return

    try:
        shutil.copy2(source_acad_patch, target_acad)
        print("Copy acad.exe thành công!")
    except Exception as e:
        print(f"Lỗi copy acad.exe: {e}")


# ========================================
# 5️⃣ CÀI ĐẶT AUTOCAD (KHÔNG COPY)
# ========================================
def install_autocad(drive_path):
    install_dir = os.path.join(drive_path, INSTALL_FOLDER1)
    setup_path = os.path.join(install_dir, SETUP_EXE)

    if not os.path.exists(setup_path):
        print("Không tìm thấy Setup.exe!")
        sys.exit()

    subprocess.run([setup_path, "/q"], check=False)
    print("Cài đặt AutoCAD đã chạy xong (không copy).")


# ========================================
# MENU CHÍNH
# ========================================
def show_menu():
    while True:
        print("\n==== MENU AUTO ====")
        print("1. Chỉ cài AutoCAD")
        print("2. Chỉ copy acad.exe")
        print("3. Cài AutoCAD xong rồi copy acad.exe")
        print("4. Thoát")
        choice = input("Chọn: ").strip()

        if choice == "1":
            drive = choose_drive()
            install_autocad(drive)

        elif choice == "2":
            drive = choose_drive()
            copy_acad(drive)

        elif choice == "3":
            drive = choose_drive()
            install_autocad(drive)   # chạy blocking → đợi xong
            print("\nĐang copy file acad.exe…")
            copy_acad(drive)

        elif choice == "4":
            sys.exit()

        else:
            print("Lựa chọn không hợp lệ.")

# ========================================
# MAIN
# ========================================
if __name__ == "__main__":
    show_menu()
