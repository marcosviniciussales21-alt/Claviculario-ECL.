from __future__ import annotations
import os
import shutil
import winreg
import tkinter as tk
from pathlib import Path
from tkinter import messagebox

INSTALL_DIR = Path(r"C:\Claviculario_ECL")
UNINSTALL_KEY = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\Claviculario_ECL"


def remove_shortcuts():
    user = Path(os.environ.get("USERPROFILE", str(Path.home())))
    shortcuts = [
        user / "Desktop" / "Claviculario ECL.lnk",
        Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Claviculario ECL",
    ]
    for item in shortcuts:
        try:
            if item.is_dir():
                shutil.rmtree(item)
            elif item.exists():
                item.unlink()
        except Exception:
            pass


def remove_registry():
    try:
        winreg.DeleteKeyEx(
            winreg.HKEY_LOCAL_MACHINE, UNINSTALL_KEY,
            access=winreg.KEY_WOW64_64KEY
        )
    except Exception:
        pass


def uninstall():
    keep = messagebox.askyesno(
        "Desinstalar Claviculário ECL",
        "Deseja PRESERVAR seus dados, backups, protocolos e relatórios?\n\n"
        "Sim = remove somente o programa e mantém os dados.\n"
        "Não = remove tudo."
    )
    if not messagebox.askyesno("Confirmar", "Deseja realmente desinstalar o Claviculário ECL?"):
        return

    try:
        protected = {"data", "backups", "protocolos", "relatorios", "logs"} if keep else set()
        if INSTALL_DIR.exists():
            for item in list(INSTALL_DIR.iterdir()):
                if item.name in protected:
                    continue
                try:
                    if item.is_dir():
                        shutil.rmtree(item)
                    else:
                        item.unlink()
                except Exception:
                    pass

        remove_shortcuts()
        remove_registry()

        messagebox.showinfo(
            "Desinstalação",
            "Claviculário ECL removido.\n\n"
            + ("Seus dados foram preservados em C:\\Claviculario_ECL." if keep else "Os arquivos foram removidos.")
        )
        root.destroy()
    except Exception as exc:
        messagebox.showerror("Erro", str(exc))


root=tk.Tk()
root.title("Desinstalar Claviculário ECL")
root.geometry("480x240")
root.resizable(False,False)
root.configure(bg="#F3F6F9")
tk.Label(root,text="DESINSTALAR CLAVICULÁRIO ECL",font=("Segoe UI",16,"bold"),bg="#F3F6F9",fg="#17324D").pack(pady=(40,14))
tk.Button(root,text="DESINSTALAR",command=uninstall,font=("Segoe UI",11,"bold"),bg="#9B2C2C",fg="white",relief="flat",padx=28,pady=10).pack()
root.mainloop()
