import tkinter as tk
from tkinter import filedialog, messagebox
import math
import sys
import os

# Caminho absoluto da raiz do projeto (um nível acima de 'services')
raiz = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if raiz not in sys.path:
    sys.path.insert(0, raiz)



# ── Paleta ──────────────────────────────────────────────────────────────────
BG         = "#0A0A0F"   # Quase preto azulado
SURFACE    = "#12121A"   # Card / superfície
SURFACE2   = "#1C1C28"   # Hover / campo
ACCENT     = "#7C3AED"   # Violeta primário
ACCENT2    = "#A855F7"   # Violeta claro (hover)
ACCENT3    = "#4F46E5"   # Índigo (detalhe)
TEXT       = "#F1F0FF"   # Branco levemente violeta
TEXT_SUB   = "#8884A8"   # Subtítulo / placeholder
BORDER     = "#2A2A3E"   # Borda sutil
SUCCESS    = "#10B981"   # Verde feedback
WARNING    = "#F59E0B"   # Amarelo feedback

FONT_TITLE  = ("Segoe UI", 22, "bold")
FONT_HEAD   = ("Segoe UI", 13, "bold")
FONT_BODY   = ("Segoe UI", 10)
FONT_SMALL  = ("Segoe UI", 9)
FONT_BTN    = ("Segoe UI", 10, "bold")
FONT_MONO   = ("Consolas", 9)

# ── Helpers ─────────────────────────────────────────────────────────────────
def limpar_janela(win):
    for w in win.winfo_children():
        w.destroy()

def pill_button(parent, text, command, width=None, color=ACCENT, hover=ACCENT2, fg=TEXT):
    """Botão estilo pill com hover animado."""
    btn = tk.Button(
        parent, text=text, command=command,
        bg=color, fg=fg,
        activebackground=hover, activeforeground=fg,
        relief="flat", cursor="hand2",
        font=FONT_BTN,
        padx=20, pady=10,
        bd=0
    )
    if width:
        btn.config(width=width)
    btn.bind("<Enter>", lambda e: btn.config(bg=hover))
    btn.bind("<Leave>", lambda e: btn.config(bg=color))
    return btn

def ghost_button(parent, text, command):
    """Botão ghost (sem fundo) para ações secundárias."""
    btn = tk.Button(
        parent, text=f"← {text}", command=command,
        bg=SURFACE, fg=TEXT_SUB,
        activebackground=SURFACE2, activeforeground=TEXT,
        relief="flat", cursor="hand2",
        font=FONT_SMALL,
        padx=14, pady=7,
        bd=0
    )
    btn.bind("<Enter>", lambda e: btn.config(fg=TEXT, bg=SURFACE2))
    btn.bind("<Leave>", lambda e: btn.config(fg=TEXT_SUB, bg=SURFACE))
    return btn

def divider(parent, pady=8):
    tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", padx=24, pady=pady)

def badge(parent, text, color=ACCENT):
    tk.Label(
        parent, text=text,
        bg=color, fg=TEXT,
        font=FONT_SMALL,
        padx=8, pady=2,
        relief="flat"
    ).pack(side="left", padx=(0, 6))

def status_bar(parent, texto="Pronto"):
    bar = tk.Frame(parent, bg=SURFACE, height=28)
    bar.pack(fill="x", side="bottom")
    tk.Label(bar, text="●", fg=SUCCESS, bg=SURFACE, font=("Segoe UI", 9)).pack(side="left", padx=(12, 4))
    label = tk.Label(bar, text=texto, fg=TEXT_SUB, bg=SURFACE, font=FONT_SMALL)
    label.pack(side="left")
    tk.Label(bar, text="v2.0.0", fg=BORDER, bg=SURFACE, font=FONT_SMALL).pack(side="right", padx=12)
    return label

# ── Tela: Menu Principal ─────────────────────────────────────────────────────
def tela_menu_principal(root):
    limpar_janela(root)
    root.geometry("960x660")
    root.configure(bg=BG)

    # ── Sidebar ──
    sidebar = tk.Frame(root, bg=SURFACE, width=220)
    sidebar.pack(side="left", fill="y")
    sidebar.pack_propagate(False)

    tk.Label(sidebar, text="⬡", font=("Segoe UI", 28), bg=SURFACE, fg=ACCENT).pack(pady=(32, 4))
    tk.Label(sidebar, text="AutoPanel", font=("Segoe UI", 13, "bold"), bg=SURFACE, fg=TEXT).pack()
    tk.Label(sidebar, text="Gerenciador de Automação", font=FONT_SMALL, bg=SURFACE, fg=TEXT_SUB).pack(pady=(2, 24))

    divider(sidebar)

    nav_items = [
        ("🗂  Automações",   True),
        ("📋  Outros",   False),
        ("⚙  Configurações", False),
    ]
    
    for label, active in nav_items:
        f_color = ACCENT if active else SURFACE
        t_color = TEXT if active else TEXT_SUB
        item = tk.Frame(sidebar, bg=f_color, cursor="hand2")
        item.pack(fill="x", padx=12, pady=2)
        tk.Label(item, text=label, bg=f_color, fg=t_color,
                 font=FONT_BODY, padx=14, pady=9, anchor="w").pack(fill="x")

    # ── Área principal ──
    main = tk.Frame(root, bg=BG)
    main.pack(side="left", fill="both", expand=True)

    # Cabeçalho
    header = tk.Frame(main, bg=BG)
    header.pack(fill="x", padx=30, pady=(28, 0))
    tk.Label(header, text="Central de Automações", font=FONT_TITLE,
             bg=BG, fg=TEXT).pack(side="left")
    badge_frame = tk.Frame(header, bg=BG)
    badge_frame.pack(side="right", anchor="s", pady=6)
    badge(badge_frame, "4 módulos", ACCENT3)

    tk.Label(main, text="Selecione um módulo para iniciar o processamento",
             font=FONT_BODY, bg=BG, fg=TEXT_SUB).pack(anchor="w", padx=30, pady=(4, 18))

    # ── Grid de cards 2×2 ──
    automacoes = [
        ("📊", "Relatório Despesas",    "Simplificação e montagem de relatório",   ACCENT,  tela_excel_despesa),
        ("📄", "Relatório PDF",      "Geração automática de\nrelatórios em PDF",             ACCENT3, tela_pdf),
        ("🌐", "Scraping Web",       "Coleta de dados de\npáginas e APIs externas",          "#0E7490", tela_web),
        ("🔄", "Sync de Banco",      "Sincronização e migração\nentre bases de dados",       "#065F46", tela_banco),
    ]

    grid = tk.Frame(main, bg=BG)
    grid.pack(fill="both", expand=True, padx=24, pady=0)

    for i, (icon, titulo, desc, cor, tela_fn) in enumerate(automacoes):
        row, col = divmod(i, 2)

        card = tk.Frame(grid, bg=SURFACE, relief="flat", bd=0)
        card.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")
        card.configure(cursor="hand2")
        grid.columnconfigure(col, weight=1)
        grid.rowconfigure(row, weight=1)

        # Faixa colorida no topo do card
        accent_bar = tk.Frame(card, bg=cor, height=4)
        accent_bar.pack(fill="x")

        inner = tk.Frame(card, bg=SURFACE, padx=18, pady=14)
        inner.pack(fill="both", expand=True)

        top_row = tk.Frame(inner, bg=SURFACE)
        top_row.pack(fill="x")
        tk.Label(top_row, text=icon, font=("Segoe UI", 22),
                 bg=SURFACE, fg=cor).pack(side="left")

        tk.Label(inner, text=titulo, font=FONT_HEAD, bg=SURFACE, fg=TEXT, anchor="w").pack(fill="x", pady=(6, 2))
        tk.Label(inner, text=desc, font=FONT_SMALL, bg=SURFACE, fg=TEXT_SUB,
                 justify="left", anchor="w").pack(fill="x")

        fn = tela_fn  # captura correta no loop
        btn = tk.Button(
            inner, text="Abrir  →",
            bg=cor, fg=TEXT,
            activebackground=SURFACE2, activeforeground=TEXT,
            relief="flat", cursor="hand2",
            font=FONT_BTN, padx=14, pady=6,
            command=lambda r=root, f=fn: f(r)
        )
        btn.pack(anchor="e", pady=(12, 0))

        # Hover no card inteiro
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

    status_bar(main)

def _set_bg_recursive(widget, color):
    """Propaga cor de fundo recursivamente (exceto botões e barras coloridas)."""
    try:
        if isinstance(widget, (tk.Button, tk.Entry)):
            return
        widget.config(bg=color)
    except tk.TclError:
        pass
    for child in widget.winfo_children():
        _set_bg_recursive(child, color)

# ── Layout base para telas de automação ──────────────────────────────────────
def _base_tela(root, titulo, icon, cor, conteudo_fn):
    limpar_janela(root)
    root.geometry("860x560")
    root.configure(bg=BG)

    # Topo com faixa colorida
    header_bar = tk.Frame(root, bg=cor, height=5)
    header_bar.pack(fill="x")

    top = tk.Frame(root, bg=BG)
    top.pack(fill="x", padx=32, pady=(22, 0))

    tk.Label(top, text=icon, font=("Segoe UI", 26), bg=BG, fg=cor).pack(side="left", padx=(0, 12))
    title_col = tk.Frame(top, bg=BG)
    title_col.pack(side="left")
    tk.Label(title_col, text=titulo, font=FONT_TITLE, bg=BG, fg=TEXT).pack(anchor="w")
    tk.Label(title_col, text="Configure e execute o módulo abaixo", font=FONT_SMALL, bg=BG, fg=TEXT_SUB).pack(anchor="w")

    ghost_button(top, "Voltar ao menu", lambda: tela_menu_principal(root)).pack(side="right", anchor="center")

    divider(root, pady=14)

    # Área de conteúdo
    content = tk.Frame(root, bg=BG)
    content.pack(fill="both", expand=True, padx=32)

    conteudo_fn(root, content, cor)

    status_bar(root)

# ── Tela Excel ───────────────────────────────────────────────────────────────
def tela_excel_despesa(root):
    def corpo(root, content, cor):
        caminho = tk.StringVar(value="")

        # Card de upload
        card = tk.Frame(content, bg=SURFACE, padx=24, pady=20)
        card.pack(fill="x", pady=(0, 14))

        tk.Label(card, text="Arquivo Excel", font=FONT_HEAD, bg=SURFACE, fg=TEXT).pack(anchor="w")
        tk.Label(card, text="Formatos suportados: .xlsx", font=FONT_SMALL, bg=SURFACE, fg=TEXT_SUB).pack(anchor="w", pady=(2, 12))

        row = tk.Frame(card, bg=SURFACE)
        row.pack(fill="x")

        entry = tk.Entry(row, textvariable=caminho, bg=SURFACE2, fg=TEXT,
                         insertbackground=TEXT, relief="flat",
                         font=FONT_MONO, highlightthickness=1,
                         highlightbackground=BORDER, highlightcolor=cor)
        entry.pack(side="left", fill="x", expand=True, ipady=8, padx=(0, 10))
        entry.insert(0, "Nenhum arquivo selecionado...")

        def selecionar():
            p = filedialog.askopenfilename(
                title="Selecione o arquivo Excel",
                filetypes=[("Excel", "*.xlsx *.xls")]
            )
            if p:
                caminho.set(p)

        pill_button(row, "Navegar", selecionar, color=cor, hover=ACCENT2).pack(side="right")

        # Opções
        card2 = tk.Frame(content, bg=SURFACE, padx=24, pady=16)
        card2.pack(fill="x", pady=(0, 14))
        tk.Label(card2, text="Opções de Processamento", font=FONT_HEAD, bg=SURFACE, fg=TEXT).pack(anchor="w", pady=(0, 10))

        opts = [("Remover duplicatas", True), ("Formatar cabeçalhos", True), ("Exportar relatório", False)]
        vars_ = []
        for label, default in opts:
            v = tk.BooleanVar(value=default)
            vars_.append(v)
            row_opt = tk.Frame(card2, bg=SURFACE)
            row_opt.pack(fill="x", pady=2)
            tk.Checkbutton(row_opt, variable=v, bg=SURFACE, fg=TEXT,
                           selectcolor=cor, activebackground=SURFACE,
                           font=FONT_BODY).pack(side="left")
            tk.Label(row_opt, text=label, bg=SURFACE, fg=TEXT, font=FONT_BODY).pack(side="left")

        def executar():
            arq = caminho.get()
            if not arq or "Nenhum" in arq:
                messagebox.showwarning("Atenção", "Selecione um arquivo antes de executar.")
                return
            
            try:
                
                from despesas.main import iniciar_processamento
                
              
                iniciar_processamento(arq)
                
                messagebox.showinfo("Sucesso", "Processamento concluído!")
            except Exception as e:
                messagebox.showerror("Erro de Módulo", f"Não foi possível iniciar o módulo: {e}")

        pill_button(content, "▶  Executar Automação", executar, color=cor, hover=ACCENT2).pack(fill="x", pady=4, ipady=4)

    _base_tela(root, "Processar Excel", "📊", ACCENT, corpo)

# ── Tela PDF ─────────────────────────────────────────────────────────────────
def tela_pdf(root):
    def corpo(root, content, cor):
        card = tk.Frame(content, bg=SURFACE, padx=24, pady=20)
        card.pack(fill="x", pady=(0, 14))

        tk.Label(card, text="Geração de Relatório PDF", font=FONT_HEAD, bg=SURFACE, fg=TEXT).pack(anchor="w")
        tk.Label(card, text="Configure o template e os dados de origem", font=FONT_SMALL, bg=SURFACE, fg=TEXT_SUB).pack(anchor="w", pady=(2, 14))

        campos = [("Título do relatório", "Ex: Relatório Mensal Q1"), ("Autor / Empresa", "Ex: Acme Corp")]
        for label, ph in campos:
            tk.Label(card, text=label, bg=SURFACE, fg=TEXT_SUB, font=FONT_SMALL).pack(anchor="w", pady=(6, 2))
            e = tk.Entry(card, bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                         relief="flat", font=FONT_BODY,
                         highlightthickness=1, highlightbackground=BORDER, highlightcolor=cor)
            e.insert(0, ph)
            e.config(fg=TEXT_SUB)
            e.bind("<FocusIn>", lambda ev, x=e, p=ph: (x.delete(0, "end"), x.config(fg=TEXT)) if x.get() == p else None)
            e.pack(fill="x", ipady=7, pady=(0, 4))

        def gerar():
            messagebox.showinfo("PDF", "Relatório PDF gerado com sucesso!")

        pill_button(content, "▶  Gerar PDF", gerar, color=cor, hover="#6366F1").pack(fill="x", pady=10, ipady=4)

    _base_tela(root, "Relatório PDF", "📄", ACCENT3, corpo)

# ── Tela Web ──────────────────────────────────────────────────────────────────
def tela_web(root):
    def corpo(root, content, cor):
        card = tk.Frame(content, bg=SURFACE, padx=24, pady=20)
        card.pack(fill="x", pady=(0, 14))

        tk.Label(card, text="Configuração de Scraping", font=FONT_HEAD, bg=SURFACE, fg=TEXT).pack(anchor="w")

        tk.Label(card, text="URL alvo", bg=SURFACE, fg=TEXT_SUB, font=FONT_SMALL).pack(anchor="w", pady=(10, 2))
        url_entry = tk.Entry(card, bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                             relief="flat", font=FONT_MONO,
                             highlightthickness=1, highlightbackground=BORDER, highlightcolor=cor)
        url_entry.insert(0, "https://")
        url_entry.pack(fill="x", ipady=8)

        tk.Label(card, text="Seletor CSS / XPath", bg=SURFACE, fg=TEXT_SUB, font=FONT_SMALL).pack(anchor="w", pady=(10, 2))
        sel_entry = tk.Entry(card, bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                             relief="flat", font=FONT_MONO,
                             highlightthickness=1, highlightbackground=BORDER, highlightcolor=cor)
        sel_entry.insert(0, "div.content > p")
        sel_entry.pack(fill="x", ipady=8)

        card2 = tk.Frame(content, bg=SURFACE, padx=24, pady=16)
        card2.pack(fill="x", pady=(0, 14))
        tk.Label(card2, text="Saída", font=FONT_HEAD, bg=SURFACE, fg=TEXT).pack(anchor="w", pady=(0, 8))
        for opt in ["JSON", "CSV", "Excel"]:
            tk.Radiobutton(card2, text=opt, value=opt, bg=SURFACE, fg=TEXT,
                           selectcolor=cor, activebackground=SURFACE,
                           font=FONT_BODY).pack(side="left", padx=8)

        def iniciar():
            messagebox.showinfo("Web Scraping", "Coleta iniciada!\nAguarde a conclusão.")

        pill_button(content, "▶  Iniciar Coleta", iniciar, color=cor, hover="#0891B2").pack(fill="x", pady=4, ipady=4)

    _base_tela(root, "Scraping Web", "🌐", "#0E7490", corpo)

# ── Tela Banco ────────────────────────────────────────────────────────────────
def tela_banco(root):
    def corpo(root, content, cor):
        card = tk.Frame(content, bg=SURFACE, padx=24, pady=20)
        card.pack(fill="x", pady=(0, 14))

        tk.Label(card, text="Conexão com Banco de Dados", font=FONT_HEAD, bg=SURFACE, fg=TEXT).pack(anchor="w")
        tk.Label(card, text="Configure a origem e o destino da sincronização", font=FONT_SMALL, bg=SURFACE, fg=TEXT_SUB).pack(anchor="w", pady=(2, 14))

        row2 = tk.Frame(card, bg=SURFACE)
        row2.pack(fill="x")
        for label, ph in [("Host de Origem", "localhost:5432"), ("Host de Destino", "prod-server:5432")]:
            col_f = tk.Frame(row2, bg=SURFACE)
            col_f.pack(side="left", fill="x", expand=True, padx=(0, 12))
            tk.Label(col_f, text=label, bg=SURFACE, fg=TEXT_SUB, font=FONT_SMALL).pack(anchor="w", pady=(0, 3))
            e = tk.Entry(col_f, bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                         relief="flat", font=FONT_MONO,
                         highlightthickness=1, highlightbackground=BORDER, highlightcolor=cor)
            e.insert(0, ph)
            e.pack(fill="x", ipady=7)

        card2 = tk.Frame(content, bg=SURFACE, padx=24, pady=14)
        card2.pack(fill="x", pady=(0, 14))
        tk.Label(card2, text="Modo de Sincronização", font=FONT_HEAD, bg=SURFACE, fg=TEXT).pack(anchor="w", pady=(0, 8))
        modo = tk.StringVar(value="incremental")
        for val, lbl in [("incremental", "Incremental (delta)"), ("full", "Completa (full replace)"), ("mirror", "Mirror / Espelho")]:
            tk.Radiobutton(card2, text=lbl, variable=modo, value=val,
                           bg=SURFACE, fg=TEXT, selectcolor=cor,
                           activebackground=SURFACE, font=FONT_BODY).pack(anchor="w", pady=2)

        def sincronizar():
            messagebox.showinfo("Sync", f"Sincronização [{modo.get()}] iniciada!")

        pill_button(content, "▶  Sincronizar Agora", sincronizar, color=cor, hover="#047857").pack(fill="x", pady=4, ipady=4)

    _base_tela(root, "Sync de Banco", "🔄", "#065F46", corpo)

