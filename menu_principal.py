# -*- coding: utf-8 -*-
from __future__ import annotations

import sys, subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path

HERE = Path(__file__).resolve().parent
CORE = HERE / "concilia_core.py"

PROVIDERS = [
    ("ALELO", "Alelo"),
    ("TICKET", "Ticket"),
    ("FARMACIASAPP", "FarmaciasApp"),
]

class Menu(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Concilia - Menu Principal (Multi-janelas)")
        self.geometry("520x260")
        self.resizable(False, False)

        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Selecione a bandeira para abrir em uma janela separada:", font=("Segoe UI", 11, "bold")).pack(anchor="w")

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(14, 10))

        for prov, label in PROVIDERS:
            ttk.Button(btns, text=f"Abrir {label}", command=lambda p=prov: self.open_provider(p)).pack(fill="x", pady=4)

        ttk.Separator(frm).pack(fill="x", pady=10)

        ttk.Label(frm, text="Dica: cada janela abre em um processo separado (evita conflito de Tk).").pack(anchor="w")

        ttk.Button(frm, text="Sair", command=self.destroy).pack(anchor="e", pady=(10,0))

    def open_provider(self, prov: str):
        if not CORE.exists():
            messagebox.showerror("Erro", f"Não encontrei o core: {CORE}")
            return
        try:
            subprocess.Popen([sys.executable, str(CORE), "--provider", prov], cwd=str(HERE))
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir {prov}: {e}")

def main():
    Menu().mainloop()

if __name__ == "__main__":
    main()
