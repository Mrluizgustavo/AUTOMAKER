import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys
import os

from services.despesas.services.dashboard_despesas import tela_dashboard_despesas, tela_excel_despesa
from services.telegrama.services.tela_telegrama import tela_telegrama

from services import ui_theme


# ── Tela: Menu Principal ─────────────────────────────────────────────────────
def tela_menu_principal(root):
    ui_theme.limpar_janela(root)
    root.geometry("960x660")
    root.configure(bg=ui_theme.BG)

    # ── Sidebar ──
    sidebar = tk.Frame(root, bg=ui_theme.SURFACE, width=220)
    sidebar.pack(side="left", fill="y")
    sidebar.pack_propagate(False)

    tk.Label(sidebar, text="⬡", font=("Segoe UI", 28), bg=ui_theme.SURFACE, fg=ui_theme.ACCENT).pack(pady=(32, 4))
    tk.Label(sidebar, text="AutoPanel", font=("Segoe UI", 13, "bold"), bg=ui_theme.SURFACE, fg=ui_theme.TEXT).pack()
    tk.Label(sidebar, text="Gerenciador de Automação", font=ui_theme.FONT_SMALL, bg=ui_theme.SURFACE, fg=ui_theme.TEXT_SUB).pack(pady=(2, 24))

    ui_theme.divider(sidebar)

    nav_items = [
        ("🗂  Automações",    True),
        ("📊  Dashboards",    False),
        ("📋  Outros",        False),
        ("⚙  Configurações",  False),
    ]

    for label, active in nav_items:
        f_color = ui_theme.ACCENT if active else ui_theme.SURFACE
        t_color = ui_theme.TEXT if active else ui_theme.TEXT_SUB
        item = tk.Frame(sidebar, bg=f_color, cursor="hand2")
        item.pack(fill="x", padx=12, pady=2)
        tk.Label(item, text=label, bg=f_color, fg=t_color,
                 font=ui_theme.FONT_BODY, padx=14, pady=9, anchor="w").pack(fill="x")

    # ── Área principal ──
    main = tk.Frame(root, bg=ui_theme.BG)
    main.pack(side="left", fill="both", expand=True)

    # Cabeçalho
    header = tk.Frame(main, bg=ui_theme.BG)
    header.pack(fill="x", padx=30, pady=(28, 0))
    tk.Label(header, text="Central de Automações", font=ui_theme.FONT_TITLE,
             bg=ui_theme.BG, fg=ui_theme.TEXT).pack(side="left")
    badge_frame = tk.Frame(header, bg=ui_theme.BG)
    badge_frame.pack(side="right", anchor="s", pady=6)
    ui_theme.badge(badge_frame, "5 módulos", ui_theme.ACCENT3)

    tk.Label(main, text="Selecione um módulo para iniciar o processamento",
             font=ui_theme.FONT_BODY, bg=ui_theme.BG, fg=ui_theme.TEXT_SUB).pack(anchor="w", padx=30, pady=(4, 18))

    # ── Grid de cards ──
    
    automacoes = [
        ("📊", "Relatório Despesas",  "Simplificação e montagem de relatório",ui_theme.ACCENT, tela_excel_despesa),
        ("📈", "Dashboard Despesas",  "Visualize KPIs e métricas cruzadas\ndo relatório de despesas", "#0891B2", tela_dashboard_despesas),
        ("📃", "Telegrama CORREIOS",  "Montagem de telegramas", ui_theme.ACCENT, tela_telegrama),
        ]

    grid = tk.Frame(main, bg=ui_theme.BG)
    grid.pack(fill="both", expand=True, padx=24, pady=0)

    for i, (icon, titulo, desc, cor, tela_fn) in enumerate(automacoes):
        row, col = divmod(i, 2)

        card = tk.Frame(grid, bg=ui_theme.SURFACE, relief="flat", bd=0)
        card.grid(row=row, column=col, padx=8, pady=8, sticky="news")
        card.configure(cursor="hand2")
        grid.columnconfigure(col, weight=1)
        grid.rowconfigure(row, minsize=160)

        accent_bar = tk.Frame(card, bg=cor, height=4)
        accent_bar.pack(fill="x")


        inner = tk.Frame(card, bg=ui_theme.SURFACE, padx=18, pady=14)
        inner.pack(fill="both", expand=True)


        top_row = tk.Frame(inner, bg=ui_theme.SURFACE)
        top_row.pack(fill="x")
        tk.Label(top_row, text=icon, font=("Segoe UI", 22), bg=ui_theme.SURFACE, fg=cor).pack(side="left")

        tk.Label(inner, text=titulo, font=ui_theme.FONT_HEAD, bg=ui_theme.SURFACE, fg=ui_theme.TEXT, anchor="w").pack(fill="x", pady=(6, 2))
        tk.Label(inner, text=desc, font=ui_theme.FONT_SMALL, bg=ui_theme.SURFACE, fg=ui_theme.TEXT_SUB,
                 justify="left", anchor="w").pack(fill="x")

        fn = tela_fn
        
        btn = tk.Button(
            inner, text="Abrir  →",
            bg=cor, fg=ui_theme.TEXT,
            activebackground=ui_theme.SURFACE2, activeforeground=ui_theme.TEXT,
            relief="flat", cursor="hand2",
            font=ui_theme.FONT_BTN, padx=14, pady=6,
            command=lambda r=root, f=fn: f(r)
        )
        btn.pack(anchor="e", pady=(12, 0))

        def _on_enter(e, c=card, b=accent_bar, clr=cor):
            c.config(bg=ui_theme.SURFACE2)
            for ch in c.winfo_children():
                ui_theme._set_bg_recursive(ch, ui_theme.SURFACE2)
            b.config(bg=clr)

        def _on_leave(e, c=card, b=accent_bar, clr=cor):
            c.config(bg=ui_theme.SURFACE)
            for ch in c.winfo_children():
                ui_theme._set_bg_recursive(ch, ui_theme.SURFACE)
            b.config(bg=clr)

        card.bind("<Enter>", _on_enter)
        card.bind("<Leave>", _on_leave)

    ui_theme.status_bar(main)


