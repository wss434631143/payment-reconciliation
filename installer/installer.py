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
from tkinter import filedialog, messagebox


APP_NAME = "财务第三方支付核对"
APP_VERSION = "1.0.4"
APP_EXE_NAME = "财务第三方支付核对-Qt版-v1.0.4.exe"
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
    """写入可被 Windows 程序和功能直接调用的卸载脚本。"""
    uninstall_cmd = install_dir / "uninstall.cmd"
    desktop_link = desktop_dir() / f"{APP_NAME}.lnk"
    menu_dir = start_menu_dir()
    cleanup_cmd = Path(os.environ.get("TEMP", str(Path.home()))) / "PaymentReconciliationQt_cleanup.cmd"
    script = f"""
@echo off
setlocal
chcp 65001 >nul
taskkill /IM "{APP_EXE_NAME}" /F >nul 2>nul
del /F /Q "{desktop_link}" >nul 2>nul
rmdir /S /Q "{menu_dir}" >nul 2>nul
reg delete "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\PaymentReconciliationQt" /f >nul 2>nul
set "CLEANUP={cleanup_cmd}"
set "INSTALL_DIR={install_dir}"
(
  echo @echo off
  echo timeout /t 2 /nobreak ^>nul
  echo rmdir /s /q "%%INSTALL_DIR%%"
  echo del /f /q "%%~f0" ^>nul 2^>nul
) > "%%CLEANUP%%"
start "" /min cmd /c "%%CLEANUP%%"
exit /b 0
"""
    uninstall_cmd.write_text(textwrap.dedent(script).strip() + "\n", encoding="utf-8-sig")
    return uninstall_cmd

def register_uninstall(install_dir, app_exe, uninstall_cmd):
    """注册当前用户级卸载入口，供 Windows 程序和功能识别。"""
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\PaymentReconciliationQt"
    cmd_exe = Path(os.environ.get("SystemRoot", r"C:\\Windows")) / "System32" / "cmd.exe"
    uninstall_string = f'"{cmd_exe}" /c ""{uninstall_cmd}""'
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
        winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, APP_NAME)
        winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, APP_VERSION)
        winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, PUBLISHER)
        winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, str(install_dir))
        winreg.SetValueEx(key, "DisplayIcon", 0, winreg.REG_SZ, f"{app_exe},0")
        winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ, uninstall_string)
        winreg.SetValueEx(key, "QuietUninstallString", 0, winreg.REG_SZ, uninstall_string)
        winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)

def install(install_dir, create_desktop=True, create_start_menu=True):
    install_dir = Path(install_dir).expanduser()
    install_dir.mkdir(parents=True, exist_ok=True)

    source_exe = resource_path(f"payload/{APP_EXE_NAME}")
    source_icon = resource_path("payload/app_icon.ico")
    if not source_exe.exists():
        raise FileNotFoundError(f"安装包缺少主程序：{source_exe}")

    app_exe = install_dir / APP_EXE_NAME
    icon_path = install_dir / "app_icon.ico"
    shutil.copy2(source_exe, app_exe)
    shutil.copy2(source_icon, icon_path)

    uninstall_cmd = write_uninstaller(install_dir)
    if create_desktop:
        create_shortcut(desktop_dir() / f"{APP_NAME}.lnk", app_exe, icon_path)

    menu_dir = start_menu_dir()
    if create_start_menu:
        create_shortcut(menu_dir / f"{APP_NAME}.lnk", app_exe, icon_path)
        create_shortcut(
            menu_dir / "卸载.lnk",
            Path(os.environ.get("SystemRoot", r"C:\\Windows")) / "System32" / "cmd.exe",
            icon_path,
            f'/c ""{uninstall_cmd}""',
        )
    register_uninstall(install_dir, app_exe, uninstall_cmd)
    return app_exe


class InstallerWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} 安装向导")
        self.geometry("640x420")
        self.resizable(False, False)
        try:
            self.iconbitmap(resource_path("payload/app_icon.ico"))
        except Exception:
            pass

        self.step = 0
        self.install_dir = tk.StringVar(value=str(default_install_dir()))
        self.create_desktop = tk.BooleanVar(value=True)
        self.create_start_menu = tk.BooleanVar(value=True)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = tk.Frame(self, padx=28, pady=22)
        header.grid(row=0, column=0, sticky="ew")
        self.title_label = tk.Label(header, text=f"欢迎安装 {APP_NAME}", font=("Microsoft YaHei", 18, "bold"))
        self.title_label.pack(anchor="w")
        self.body = tk.Label(header, text="", font=("Microsoft YaHei", 10), justify="left", wraplength=560, fg="#374151")
        self.body.pack(anchor="w", pady=(10, 0))

        self.content = tk.Frame(self, padx=30, pady=4)
        self.content.grid(row=1, column=0, sticky="nsew")
        self.content.columnconfigure(1, weight=1)

        self.status = tk.Label(self, text="", font=("Microsoft YaHei", 10), fg="#2563eb", anchor="w")
        self.status.grid(row=2, column=0, sticky="ew", padx=30, pady=(0, 6))

        btn_frame = tk.Frame(self, padx=24, pady=14, bg="#f3f4f6")
        btn_frame.grid(row=3, column=0, sticky="ew")
        self.cancel_btn = tk.Button(btn_frame, text="退出", width=10, command=self.destroy)
        self.next_btn = tk.Button(btn_frame, text="下一步", width=10, command=self.next)
        self.back_btn = tk.Button(btn_frame, text="上一步", width=10, command=self.back)
        self.cancel_btn.pack(side="right", padx=(8, 0))
        self.next_btn.pack(side="right", padx=(8, 0))
        self.back_btn.pack(side="right")
        self.render()

    def clear_content(self):
        for child in self.content.winfo_children():
            child.destroy()

    def render(self):
        self.clear_content()
        self.back_btn.config(state=("normal" if self.step else "disabled"))
        if self.step == 0:
            self.title_label.config(text=f"欢迎安装 {APP_NAME}")
            self.body.config(text="本向导将引导你选择安装位置和快捷方式选项。安装后可从桌面或开始菜单启动，也可以在 Windows 应用和功能中卸载。")
            self.next_btn.config(text="下一步")
        elif self.step == 1:
            self.title_label.config(text="选择安装选项")
            self.body.config(text="请选择安装位置，并决定是否创建桌面图标和开始菜单入口。")
            self.next_btn.config(text="开始安装")
            tk.Label(self.content, text="安装位置", font=("Microsoft YaHei", 10)).grid(row=0, column=0, sticky="w", pady=8)
            path_entry = tk.Entry(self.content, textvariable=self.install_dir, font=("Microsoft YaHei", 10))
            path_entry.grid(row=0, column=1, sticky="ew", padx=(12, 8), pady=8)
            tk.Button(self.content, text="浏览...", command=self.choose_dir, width=9).grid(row=0, column=2, sticky="e", pady=8)
            tk.Checkbutton(self.content, text="创建桌面快捷方式", variable=self.create_desktop, font=("Microsoft YaHei", 10)).grid(row=1, column=1, sticky="w", pady=(16, 4))
            tk.Checkbutton(self.content, text="创建开始菜单快捷方式", variable=self.create_start_menu, font=("Microsoft YaHei", 10)).grid(row=2, column=1, sticky="w", pady=4)
        else:
            self.title_label.config(text="安装完成")
            self.body.config(text="安装已完成。桌面和开始菜单中已经创建快捷方式。")
            self.next_btn.config(text="完成")
            self.back_btn.config(state="disabled")

    def choose_dir(self):
        path = filedialog.askdirectory(title="选择安装位置", initialdir=str(Path(self.install_dir.get()).parent))
        if path:
            self.install_dir.set(str(Path(path) / APP_FOLDER_NAME if Path(path).name != APP_FOLDER_NAME else Path(path)))

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
                if not self.install_dir.get().strip():
                    messagebox.showwarning("请选择安装位置", "安装位置不能为空。")
                    return
                app_exe = install(self.install_dir.get().strip(), self.create_desktop.get(), self.create_start_menu.get())
                self.status.config(text=f"安装完成：{app_exe}")
                self.step = 2
                self.render()
            except Exception as exc:
                messagebox.showerror("安装失败", str(exc))
            return
        self.destroy()


if __name__ == "__main__":
    InstallerWindow().mainloop()
