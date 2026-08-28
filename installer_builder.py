from __future__ import annotations
import os
import shutil
import subprocess
import sys
import winreg
import tkinter as tk
from pathlib import Path
from tkinter import messagebox

APP_NAME = "Claviculario ECL"
INSTALL_DIR = Path(r"C:\Claviculario_ECL")
EXE_NAME = "Claviculario_ECL.exe"
UNINSTALL_KEY = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\Claviculario_ECL"


def resource_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS"))
    return Path(__file__).resolve().parent


def create_shortcut(shortcut: Path, target: Path, icon: Path) -> None:
    shortcut.parent.mkdir(parents=True, exist_ok=True)
    ps = (
        "$ws=New-Object -ComObject WScript.Shell;"
        f"$s=$ws.CreateShortcut('{str(shortcut).replace(chr(39), chr(39)*2)}');"
        f"$s.TargetPath='{str(target).replace(chr(39), chr(39)*2)}';"
        f"$s.WorkingDirectory='{str(INSTALL_DIR).replace(chr(39), chr(39)*2)}';"
        f"$s.IconLocation='{str(icon).replace(chr(39), chr(39)*2)}';"
        "$s.Save();"
    )
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
        check=False,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def register_uninstaller(uninstaller: Path, exe: Path) -> None:
    with winreg.CreateKeyEx(
        winreg.HKEY_LOCAL_MACHINE, UNINSTALL_KEY, 0,
        winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY
    ) as key:
        winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, "Claviculário ECL")
        winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, "0.9")
        winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, "Polo Caruaru")
        winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, str(INSTALL_DIR))
        winreg.SetValueEx(key, "DisplayIcon", 0, winreg.REG_SZ, str(exe))
        winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ, f'"{uninstaller}"')
        winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)


def install():
    payload = resource_dir() / "payload"
    if not payload.exists():
        messagebox.showerror("Instalador", "Arquivos internos não encontrados.")
        return

    try:
        INSTALL_DIR.mkdir(parents=True, exist_ok=True)
        protected = {"data", "backups", "protocolos", "relatorios", "logs"}

        for item in payload.iterdir():
            dst = INSTALL_DIR / item.name
            if item.name.lower() in protected:
                dst.mkdir(parents=True, exist_ok=True)
                # Só copia conteúdo inicial quando ainda não existir.
                if item.is_dir():
                    for child in item.rglob("*"):
                        if child.is_file():
                            out = dst / child.relative_to(item)
                            if not out.exists():
                                out.parent.mkdir(parents=True, exist_ok=True)
                                shutil.copy2(child, out)
                continue

            if item.is_dir():
                if dst.exists():
                    shutil.rmtree(dst)
                shutil.copytree(item, dst)
            else:
                shutil.copy2(item, dst)

        exe = INSTALL_DIR / EXE_NAME
        uninstaller = INSTALL_DIR / "Desinstalar_Claviculario_ECL.exe"
        if not exe.exists() or not uninstaller.exists():
            raise FileNotFoundError("Executável principal ou desinstalador não foi instalado.")

        user_profile = Path(os.environ.get("USERPROFILE", str(Path.home())))
        desktop = user_profile / "Desktop" / "Claviculario ECL.lnk"
        start_menu = Path(os.environ["APPDATA"]) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Claviculario ECL" / "Claviculario ECL.lnk"

        create_shortcut(desktop, exe, exe)
        create_shortcut(start_menu, exe, exe)
        register_uninstaller(uninstaller, exe)

        messagebox.showinfo(
            "Claviculário ECL",
            "Instalação concluída com sucesso!\n\n"
            "Instalado em:\nC:\\Claviculario_ECL\n\n"
            "O programa também aparece em Aplicativos instalados do Windows."
        )
        os.startfile(str(exe))
        root.destroy()
    except Exception as exc:
        messagebox.showerror("Erro na instalação", str(exc))


root = tk.Tk()
root.title("Instalar Claviculário ECL")
root.geometry("560x330")
root.resizable(False, False)
root.configure(bg="#F3F6F9")

tk.Label(root, text="CLAVICULÁRIO ECL", font=("Segoe UI",22,"bold"), bg="#F3F6F9", fg="#17324D").pack(pady=(38,8))
tk.Label(root, text="Instalação profissional • Polo Caruaru", font=("Segoe UI",11), bg="#F3F6F9", fg="#4C5B68").pack()
tk.Label(
    root,
    text="O sistema será instalado em C:\\Claviculario_ECL.\n"
         "Atalho na Área de Trabalho e Menu Iniciar.\n"
         "Atualizações preservam data, backups, protocolos e relatórios.",
    justify="center", font=("Segoe UI",10), bg="#F3F6F9", fg="#4C5B68"
).pack(pady=28)

tk.Button(
    root, text="INSTALAR", command=install,
    font=("Segoe UI",11,"bold"), bg="#17324D", fg="white",
    activebackground="#244A6D", activeforeground="white",
    relief="flat", padx=35, pady=12, cursor="hand2"
).pack()

tk.Label(root, text="Versão 0.9", font=("Segoe UI",9), bg="#F3F6F9", fg="#6B7783").pack(side="bottom", pady=18)
root.mainloop()
