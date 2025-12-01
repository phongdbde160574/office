# # Auto Installer GUI (Python + Tkinter)
# # Version: Admin-run + Improved Safety
#
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import shutil
from pathlib import Path
import os
import ctypes
import sys
#
# # ==============================
# # 1. Kiểm tra & chạy lại bằng quyền Admin
# # ==============================
# def run_as_admin():
#     try:
#         is_admin = ctypes.windll.shell32.IsUserAnAdmin()
#     except:
#         is_admin = False
#
#     if not is_admin:
#         args = " ".join(f'"{arg}"' for arg in sys.argv)
#         ctypes.windll.shell32.ShellExecuteW(
#             None,
#             "runas",
#             sys.executable,
#             args,
#             None,
#             1
#         )
#         sys.exit()
#
# run_as_admin()
#
# # ==============================
# # 2. Helper chạy PowerShell
# # ==============================
# def run_ps(cmd):
#     try:
#         result = subprocess.run(
#             ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", cmd],
#             capture_output=True,
#             text=True,
#             encoding="utf-8"
#         )
#         return f"[Exit {result.returncode}]\n{result.stdout}{result.stderr}"
#     except Exception as e:
#         return str(e)
#
# # ==============================
# # 3. Cài EVKey local
# # ==============================
# def install_evkey(log=None):
#     source = Path(r"C:\WINDOWS UPDATE BLOCKER NEW\EVKey (3)\EVKey64.exe")
#     desktop = Path(os.path.join(os.environ["USERPROFILE"], "Desktop"))
#     dest = desktop / "EVKey64.exe"
#
#     def w(t):
#         if log: log(t)
#
#     w("🔍 Kiểm tra EVKey64.exe...")
#     if not source.exists():
#         w("❌ Không tìm thấy file EVKey64.exe!")
#         return
#
#     try:
#         if dest.exists():
#             w("⚠ EVKey đã có sẵn trên Desktop ⇒ sẽ ghi đè...")
#         shutil.copy2(source, dest)
#         w("📌 Đã sao chép ra Desktop thành công!")
#     except Exception as e:
#         w(f"❌ Lỗi copy: {e}")
#         return
#
#     try:
#         w("🚀 Khởi chạy EVKey...")
#         subprocess.Popen([str(dest)], shell=False)
#         w("✔ EVKey đã chạy!")
#     except Exception as e:
#         w(f"❌ Lỗi chạy EVKey: {e}")
#
# # ==============================
# # 4. Cài Foxit local (kèm bật/tắt PUA an toàn)
# # ==============================
# def install_foxit_local(log=None):
#     installer = Path(r"C:\WINDOWS UPDATE BLOCKER NEW\FoxitReader501.0523_enu_Setup.exe")
#
#     def w(t):
#         if log: log(t)
#
#     w("🛑 Tắt PUAProtection (tạm thời)...")
#     w(run_ps("Set-MpPreference -PUAProtection 0"))
#
#     if not installer.exists():
#         w("❌ Không tìm thấy Foxit installer!")
#         w("🔒 Bật lại PUAProtection...")
#         w(run_ps("Set-MpPreference -PUAProtection 1"))
#         return
#
#     try:
#         w("🚀 Chạy Foxit silent installer...")
#         subprocess.Popen((installer), shell=False)
#         w("✔ Foxit đang được cài đặt!")
#     except Exception as e:
#         w(f"❌ Lỗi chạy Foxit: {e}")
#
#     w("🔒 Bật lại PUAProtection...")
#     w(run_ps("Set-MpPreference -PUAProtection 1"))
#
# # ==============================
# # 6. Giao diện GUI
# # ==============================
# class AutoInstallApp:
#     def __init__(self, root):
#         self.root = root
#         root.title("Auto Installer - Admin Mode")
#         root.geometry("900x600")
#
#         notebook = ttk.Notebook(root)
#         notebook.pack(fill="both", expand=True)
#
#         self.tab_install = ttk.Frame(notebook)
#         notebook.add(self.tab_install, text="Cài đặt ứng dụng")
#
#         self.build_install_tab()
#
#     # ------------------------------
#     # TAB: Install
#     # ------------------------------
#     def build_install_tab(self):
#         frame = ttk.Frame(self.tab_install)
#         frame.pack(fill="both", expand=True, padx=10, pady=10)
#
#         left = ttk.Frame(frame)
#         left.pack(side="left", fill="y", padx=10)
#
#         ttk.Label(left, text="Chọn ứng dụng cần cài", font=("Arial", 12, "bold")).pack(pady=5)
#
#         self.chk_office = tk.BooleanVar()
#         self.chk_foxit = tk.BooleanVar()
#         self.chk_evkey = tk.BooleanVar()
#         self.chk_ultra = tk.BooleanVar()
#         self.chk_unikey = tk.BooleanVar()
#         self.chk_zalo = tk.BooleanVar()
#
#         ttk.Checkbutton(left, text="Microsoft Office", variable=self.chk_office).pack(anchor="w")
#         self.office_box = ttk.Combobox(left, values=["2016", "2019", "2021", "Microsoft 365"], width=20)
#         self.office_box.current(2)
#         self.office_box.pack(anchor="w", padx=20)
#
#         ttk.Checkbutton(left, text="Foxit Reader (Local)", variable=self.chk_foxit).pack(anchor="w")
#         ttk.Checkbutton(left, text="EVKey (Local)", variable=self.chk_evkey).pack(anchor="w")
#         ttk.Checkbutton(left, text="UltraView (Winget)", variable=self.chk_ultra).pack(anchor="w")
#         ttk.Checkbutton(left, text="UniKey (Winget)", variable=self.chk_unikey).pack(anchor="w")
#         ttk.Checkbutton(left, text="Zalo (Winget)", variable=self.chk_zalo).pack(anchor="w")
#
#         ttk.Button(left, text="🚀 BẮT ĐẦU CÀI", command=self.install_all).pack(pady=20)
#
#         # Log box + scrollbar
#         log_frame = ttk.Frame(frame)
#         log_frame.pack(side="left", fill="both", expand=True)
#
#         self.log_box = tk.Text(log_frame, wrap="word")
#         self.log_box.pack(side="left", fill="both", expand=True)
#
#         scroll = ttk.Scrollbar(log_frame, command=self.log_box.yview)
#         scroll.pack(side="right", fill="y")
#         self.log_box.config(yscrollcommand=scroll.set)
#
#     # ------------------------------
#     def log(self, text):
#         self.log_box.insert(tk.END, text + "\n")
#         self.log_box.see(tk.END)
#
#     # ------------------------------
#     def install_all(self):
#         self.log("=== BẮT ĐẦU CÀI ĐẶT ===")
#
#         if self.chk_zalo.get():
#             self.log("➡ Cài Zalo...")
#             self.log(run_ps("winget install Zalo.Zalo -e --silent"))
#
#         if self.chk_foxit.get():
#             self.log("➡ Cài Foxit (Local)...")
#             install_foxit_local(self.log)
#
#         if self.chk_ultra.get():
#             self.log("➡ Cài UltraView...")
#             self.log(run_ps("winget install UltraViewer.UltraViewer -e --silent"))
#
#         if self.chk_evkey.get():
#             self.log("➡ Cài EVKey (Local)...")
#             install_evkey(self.log)
#
#         if self.chk_unikey.get():
#             self.log("➡ Cài UniKey...")
#             self.log(run_ps("winget install Unikey -e --silent"))
#
#         if self.chk_office.get():
#             ver = self.office_box.get()
#             self.log(f"➡ Cài Office {ver}...")
#             if ver == "2021":
#                 install_office_2021(self.log)
#             else:
#                 self.log("⚠ Chỉ hỗ trợ tự động Office 2021 (DVD)")
#
#         self.log("=== HOÀN TẤT ===")
#
# # ==============================
# # MAIN
# # ==============================
# if __name__ == "__main__":
#     root = tk.Tk()
#     app = AutoInstallApp(root)
#     root.mainloop()
# import subprocess
# import ctypes
# import os
# from pathlib import Path
# import time
#
# ISO_PATH = r"HDD-DATA-DE160574\Office_professional_plus_2021_x86_x64_dvd_c6dd6dc6.iso"
#
#
# def is_admin():
#     try:
#         return ctypes.windll.shell32.IsUserAnAdmin()
#     except:
#         return False
#
#
# if not is_admin():
#     print("🔒 Elevating privileges...")
#     ctypes.windll.shell32.ShellExecuteW(
#         None, "runas", "python.exe", __file__, None, 1
#     )
#     exit()
#
#
# def check_iso():
#     if not Path(ISO_PATH).exists():
#         print(f"❌ ISO not found: {ISO_PATH}")
#         input("Press Enter to exit...")
#         exit()
#     print("✔ ISO found!")
#
#
# def mount_iso():
#     print("📌 Mounting ISO...")
#     cmd = f'Mount-DiskImage -ImagePath "{ISO_PATH}"'
#     subprocess.run(["powershell", "-Command", cmd], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
#     time.sleep(2)
#
#
# def find_setup_drive():
#     print("🔍 Searching for setup.exe...")
#     # Lọc từng ổ CD/DVD Drive
#     cmd = '(Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object {$_.DriveType -eq 5}).DeviceID'
#     result = subprocess.check_output(["powershell", "-Command", cmd], text=True)
#
#     drives = [d.strip() for d in result.splitlines() if d.strip()]
#     for d in drives:
#         setup_path = Path(d) / "setup.exe"
#         if setup_path.exists():
#             print(f"✔ Found setup.exe in: {d}")
#             return d
#
#     print("❌ setup.exe not found in any mounted drives!")
#     input("Press Enter to exit...")
#     exit()
#
#
# def run_setup(drive):
#     setup_path = Path(drive) / "setup.exe"
#     print("🚀 Starting Office setup...")
#     os.chdir(drive)
#     subprocess.Popen(str(setup_path))
#     print("👉 Office installation launched! Continue in the popup window.")
#     input("Press Enter to exit...")
#
#
# def main():
#     print("===== AUTO INSTALL OFFICE 2021 =====\n")
#     check_iso()
#     mount_iso()
#     drive = find_setup_drive()
#     run_setup(drive)
#
#
# if __name__ == "__main__":
#     main()
#
# cmd = 'iex (curl.exe -s --doh-url https://1.1.1.1/dns-query https://get.activated.win | Out-String)'
# result = subprocess.check_output(
#     ["powershell", "-Command", cmd],
#     text=True
# )
#