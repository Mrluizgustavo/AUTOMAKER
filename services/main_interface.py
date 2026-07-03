import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys
import os

from services.despesas.services.dashboard_despesas import tela_dashboard_despesas, tela_excel_despesa
from services.telegrama.main_telegrama import tela_telegrama
from services.contratos.main_contratos import tela_contratos

from services.ui_theme import BG, FONT_TITLE, TEXT, ACCENT3, ACCENT4, FONT_BODY, TEXT_SUB, ACCENT, SURFACE, FONT_HEAD, FONT_SMALL, SURFACE2, FONT_BTN, _set_bg_recursive, badge, status_bar

# ── Tela: Menu Principal ─────────────────────────────────────────────────────
def tela_menu_principal(parent_frame, roteador):

    # Cabeçalho
    header = tk.Frame(parent_frame, bg=BG)
    header.pack(fill="x", padx=30, pady=(28, 0))
    tk.Label(header, text="Central de Automações", font=FONT_TITLE,
             bg=BG, fg=TEXT).pack(side="left")
    
    badge_frame = tk.Frame(header, bg=BG)
    badge_frame.pack(side="right", anchor="s", pady=6)
    badge(badge_frame, "5 módulos", ACCENT3)

    tk.Label(parent_frame, text="Selecione um módulo para iniciar o processamento",
             font=FONT_BODY, bg=BG, fg=TEXT_SUB).pack(anchor="w", padx=30, pady=(4, 18))


    # ── Grid de cards ──
    
    automacoes = [
        ("📊", "Relatório Despesas",  "Simplificação e montagem de relatório",ACCENT, lambda p: tela_excel_despesa(p, roteador)),
        ("📈", "Dashboard Despesas",  "Visualize KPIs e métricas cruzadas\ndo relatório de despesas", ACCENT4, lambda p: tela_dashboard_despesas(p, roteador)),
        ("📃", "Telegrama CORREIOS",  "Montagem de telegramas", ACCENT, lambda p: tela_telegrama(p, roteador)),
        ("📄", "Contratos",  "Processamento de contratos em PDF", ACCENT3, lambda p: tela_contratos(p, roteador)),
        ]

    grid = tk.Frame(parent_frame, bg=BG)
    grid.pack(fill="both", expand=True, padx=24, pady=0)

    for i, (icon, titulo, desc, cor, tela_fn) in enumerate(automacoes):
        row, col = divmod(i, 2)

        card = tk.Frame(grid, bg=SURFACE, relief="flat", bd=0)
        card.grid(row=row, column=col, padx=8, pady=8, sticky="news")
        card.configure(cursor="hand2")
        grid.columnconfigure(col, weight=1)
        grid.rowconfigure(row, minsize=160)

        accent_bar = tk.Frame(card, bg=cor, height=4)
        accent_bar.pack(fill="x")


        inner = tk.Frame(card, bg=SURFACE, padx=18, pady=14)
        inner.pack(fill="both", expand=True)


        top_row = tk.Frame(inner, bg=SURFACE)
        top_row.pack(fill="x")
        tk.Label(top_row, text=icon, font=("Segoe UI", 22), bg=SURFACE, fg=cor).pack(side="left")

        tk.Label(inner, text=titulo, font=FONT_HEAD, bg=SURFACE, fg=TEXT, anchor="w").pack(fill="x", pady=(6, 2))
        tk.Label(inner, text=desc, font=FONT_SMALL, bg=SURFACE, fg=TEXT_SUB,
                 justify="left", anchor="w").pack(fill="x")

        
        btn = tk.Button(
            inner, text="Abrir  →",
            bg=cor, fg=TEXT,
            activebackground=SURFACE2, activeforeground=TEXT,
            relief="flat", cursor="hand2",
            font=FONT_BTN, padx=14, pady=6,
            command=lambda f=tela_fn: roteador(parent_frame,f)
        )

        btn.pack(anchor="e", pady=(12, 0))

        def _on_enter(e, c=card, b=accent_bar, clr=cor):
            c.config(bg=SURFACE2)
            for ch in c.winfo_children():
                _set_bg_recursive(ch, SURFACE2)
            b.config(bg=clr)

        def _on_leave(e, c=card, b=accent_bar, clr=cor):
            c.config(bg=SURFACE)
            for ch in c.winfo_children():
                _set_bg_recursive(ch, SURFACE)
            b.config(bg=clr)

        card.bind("<Enter>", _on_enter)
        card.bind("<Leave>", _on_leave)

    status_bar(parent_frame)


