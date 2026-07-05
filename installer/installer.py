# -*- coding: utf-8 -*-
"""Windows 安装向导。

该脚本会被 PyInstaller 打包成 setup.exe。运行后将主程序安装到当前用户
LocalAppData 目录，创建桌面快捷方式、开始菜单快捷方式和卸载入口。
"""
import os
import shutil
import subprocess
import sys
import textwrap
import winreg
from pathlib import Path
import tkinter as tk
from tkinter import messagebox


APP_NAME = "财务第三方支付核对"
APP_VERSION = "1.0.3"
APP_EXE_NAME = "财务第三方支付核对-Qt版-v1.0.3.exe"
APP_FOLDER_NAME = "PaymentReconciliationQt"
PUBLISHER = "wss434631143"


def resource_path(relative):
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


def default_install_dir():
    return Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / APP_FOLDER_NAME


def desktop_dir():
    return Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Desktop"


def start_menu_dir():
    return Path(os.environ.get("APPDATA", str(Path.home()))) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / APP_NAME


def run_powershell(script):
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        check=True,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )


def create_shortcut(shortcut_path, target_path, icon_path, arguments=""):
    shortcut_path.parent.mkdir(parents=True, exist_ok=True)
    ps = f"""
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut('{str(shortcut_path).replace("'", "''")}')
    $shortcut.TargetPath = '{str(target_path).replace("'", "''")}'
    $shortcut.WorkingDirectory = '{str(target_path.parent).replace("'", "''")}'
    $shortcut.IconLocation = '{str(icon_path).replace("'", "''")}'
    $shortcut.Arguments = '{str(arguments).replace("'", "''")}'
    $shortcut.Save()
    """
    run_powershell(ps)


def write_uninstaller(install_dir):
    uninstall_ps1 = install_dir / "uninstall.ps1"
    desktop_link = desktop_dir() / f"{APP_NAME}.lnk"
    menu_dir = start_menu_dir()
    script = f"""
    $ErrorActionPreference = 'SilentlyContinue'
    Remove-Item -LiteralPath '{str(desktop_link).replace("'", "''")}' -Force
    Remove-Item -LiteralPath '{str(menu_dir).replace("'", "''")}' -Recurse -Force
    Remove-Item -Path 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\PaymentReconciliationQt' -Recurse -Force
    Start-Process -WindowStyle Hidden -FilePath cmd.exe -ArgumentList '/c timeout /t 1 /nobreak > nul & rmdir /s /q "{install_dir}"'
    """
    uninstall_ps1.write_text(textwrap.dedent(script).strip(), encoding="utf-8")
    return uninstall_ps1


def register_uninstall(install_dir, app_exe, uninstall_ps1):
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\PaymentReconciliationQt"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
        winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, APP_NAME)
        winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, APP_VERSION)
        winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, PUBLISHER)
        winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, str(install_dir))
        winreg.SetValueEx(key, "DisplayIcon", 0, winreg.REG_SZ, str(app_exe))
        winreg.SetValueEx(
            key,
            "UninstallString",
            0,
            winreg.REG_SZ,
            f'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{uninstall_ps1}"',
        )
        winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)


def install():
    install_dir = default_install_dir()
    install_dir.mkdir(parents=True, exist_ok=True)

    source_exe = resource_path(f"payload/{APP_EXE_NAME}")
    source_icon = resource_path("payload/app_icon.ico")
    if not source_exe.exists():
        raise FileNotFoundError(f"安装包缺少主程序：{source_exe}")

    app_exe = install_dir / APP_EXE_NAME
    icon_path = install_dir / "app_icon.ico"
    shutil.copy2(source_exe, app_exe)
    shutil.copy2(source_icon, icon_path)

    uninstall_ps1 = write_uninstaller(install_dir)
    create_shortcut(desktop_dir() / f"{APP_NAME}.lnk", app_exe, icon_path)

    menu_dir = start_menu_dir()
    create_shortcut(menu_dir / f"{APP_NAME}.lnk", app_exe, icon_path)
    create_shortcut(
        menu_dir / "卸载.lnk",
        Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe",
        icon_path,
        f'-NoProfile -ExecutionPolicy Bypass -File "{uninstall_ps1}"',
    )
    register_uninstall(install_dir, app_exe, uninstall_ps1)
    return app_exe


class InstallerWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} 安装向导")
        self.geometry("560x360")
        self.resizable(False, False)
        try:
            self.iconbitmap(resource_path("payload/app_icon.ico"))
        except Exception:
            pass

        self.step = 0
        self.title_label = tk.Label(self, text=f"欢迎安装 {APP_NAME}", font=("Microsoft YaHei", 18, "bold"))
        self.body = tk.Label(self, text="", font=("Microsoft YaHei", 11), justify="left", wraplength=500)
        self.path_label = tk.Label(self, text=f"安装位置：{default_install_dir()}", font=("Microsoft YaHei", 9), fg="#4b5563")
        self.status = tk.Label(self, text="", font=("Microsoft YaHei", 10), fg="#2563eb")
        self.back_btn = tk.Button(self, text="上一步", width=10, command=self.back)
        self.next_btn = tk.Button(self, text="下一步", width=10, command=self.next)
        self.cancel_btn = tk.Button(self, text="退出", width=10, command=self.destroy)

        self.title_label.pack(anchor="w", padx=28, pady=(28, 12))
        self.body.pack(anchor="w", padx=30, pady=6)
        self.path_label.pack(anchor="w", padx=30, pady=8)
        self.status.pack(anchor="w", padx=30, pady=8)
        btn_frame = tk.Frame(self)
        btn_frame.pack(side="bottom", fill="x", padx=24, pady=20)
        self.cancel_btn.pack(in_=btn_frame, side="right", padx=4)
        self.next_btn.pack(in_=btn_frame, side="right", padx=4)
        self.back_btn.pack(in_=btn_frame, side="right", padx=4)
        self.render()

    def render(self):
        self.back_btn.config(state=("normal" if self.step else "disabled"))
        if self.step == 0:
            self.body.config(text="本向导将把程序安装到当前用户目录，并自动创建桌面图标和开始菜单入口。")
            self.next_btn.config(text="下一步")
        elif self.step == 1:
            self.body.config(text="准备开始安装。安装完成后可直接从桌面图标启动，也可以在 Windows 应用和功能中卸载。")
            self.next_btn.config(text="开始安装")
        else:
            self.body.config(text="安装已完成。桌面和开始菜单中已经创建快捷方式。")
            self.next_btn.config(text="完成")
            self.back_btn.config(state="disabled")

    def back(self):
        self.step = max(0, self.step - 1)
        self.status.config(text="")
        self.render()

    def next(self):
        if self.step == 0:
            self.step = 1
            self.render()
            return
        if self.step == 1:
            try:
                self.status.config(text="正在安装，请稍候...")
                self.update_idletasks()
                app_exe = install()
                self.status.config(text=f"安装完成：{app_exe}")
                self.step = 2
                self.render()
            except Exception as exc:
                messagebox.showerror("安装失败", str(exc))
            return
        self.destroy()


if __name__ == "__main__":
    InstallerWindow().mainloop()
