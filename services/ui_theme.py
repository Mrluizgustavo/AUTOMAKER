import tkinter as tk

# ── Paleta ──────────────────────────────────────────────────────────────────
BG         = "#0A0A0F"   # Quase preto azulado
SURFACE    = "#12121A"   # Card / superfície
SURFACE2   = "#1C1C28"   # Hover / campo
ACCENT     = "#7C3AED"   # Violeta primário
ACCENT2    = "#A855F7"   # Violeta claro (hover)
ACCENT3    = "#4F46E5"   # Índigo (detalhe)
ACCENT4    = "#0891B2"   # Azul Ciano Primário
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

def _set_bg_recursive(widget, color):
    """Propaga cor de fundo recursivamente (exceto botões e campos)."""
    try:
        if isinstance(widget, (tk.Button, tk.Entry)):
            return
        widget.config(bg=color)
    except tk.TclError:
        pass
    for child in widget.winfo_children():
        _set_bg_recursive(child, color)

# ── Layout base para telas de automação ──────────────────────────────────────
def _base_tela(root, titulo, icon, cor, conteudo_fn, roteador=None):
    """
    Monta o shell padrão de uma tela de automação:
    faixa colorida + cabeçalho + botão voltar + área de conteúdo + status bar.
    Chama conteudo_fn(root, content_frame, cor) para preencher o miolo.
    """
    from services.main_interface import tela_menu_principal

    root.configure(bg=BG)

    header_bar = tk.Frame(root, bg=cor, height=5)
    header_bar.pack(fill="x")

    top = tk.Frame(root, bg=BG)
    top.pack(fill="x", padx=32, pady=(22, 0))

    tk.Label(top, text=icon, font=("Segoe UI", 26), bg=BG, fg=cor).pack(side="left", padx=(0, 12))
    title_col = tk.Frame(top, bg=BG)
    title_col.pack(side="left")
    tk.Label(title_col, text=titulo, font=FONT_TITLE, bg=BG, fg=TEXT).pack(anchor="w")
    tk.Label(title_col, text="Configure e execute o módulo abaixo",
             font=FONT_SMALL, bg=BG, fg=TEXT_SUB).pack(anchor="w")

    if roteador:
        voltar_cmd = lambda: roteador(root, lambda p: tela_menu_principal(p, roteador))
    else:
        voltar_cmd = lambda: None
    ghost_button(top, "Voltar ao menu", voltar_cmd).pack(side="right", anchor="center")

    divider(root, pady=14)

    content = tk.Frame(root, bg=BG)
    content.pack(fill="both", expand=True, padx=32)

    conteudo_fn(root, content, cor)

    status_bar(root)


# ── Overlay de loading com blur ───────────────────────────────────────────────
def executar_com_loading(container, tarefa, ao_concluir=None, ao_erro=None,
                         texto="Processando..."):
    """
    Cobre `container` com um overlay desfocado (blur) + ícone de loading animado
    enquanto `tarefa()` roda numa thread. Ao terminar, remove o overlay e chama
    `ao_concluir(resultado)` ou `ao_erro(excecao)` na thread principal.

    - container : frame a ser coberto (ex.: a área do módulo, sem a sidebar).
    - tarefa    : função sem argumentos, executada em background; seu retorno
                  é repassado a ao_concluir.
    """
    import threading

    container.update_idletasks()
    w = max(container.winfo_width(), 1)
    h = max(container.winfo_height(), 1)

    # 1. Captura a área atual e aplica blur (antes de sobrepor o overlay)
    foto = None
    try:
        from PIL import ImageGrab, ImageFilter, ImageEnhance, ImageTk
        x = container.winfo_rootx()
        y = container.winfo_rooty()
        img = ImageGrab.grab(bbox=(x, y, x + w, y + h)).resize((w, h))
        img = img.filter(ImageFilter.GaussianBlur(9))
        img = ImageEnhance.Brightness(img).enhance(0.5)   # escurece o fundo
        foto = ImageTk.PhotoImage(img)
    except Exception:
        foto = None  # fallback: overlay sólido

    # 2. Overlay cobrindo todo o container
    overlay = tk.Frame(container, bg=BG)
    overlay.place(x=0, y=0, relwidth=1, relheight=1)

    if foto is not None:
        bg_label = tk.Label(overlay, image=foto, bd=0, bg=BG)
        bg_label.image = foto  # mantém referência viva
        bg_label.place(x=0, y=0, relwidth=1, relheight=1)

    # 3. Card central com spinner animado
    card = tk.Frame(overlay, bg=SURFACE, padx=44, pady=32,
                    highlightthickness=1, highlightbackground=BORDER)
    card.place(relx=0.5, rely=0.5, anchor="center")

    quadros = ["◐", "◓", "◑", "◒"]
    estado = {"i": 0, "rodando": True}
    spinner = tk.Label(card, text=quadros[0], font=("Segoe UI", 38, "bold"),
                       bg=SURFACE, fg=ACCENT2)
    spinner.pack()
    tk.Label(card, text=texto, font=FONT_BODY, bg=SURFACE, fg=TEXT).pack(pady=(12, 0))

    def animar():
        if not estado["rodando"]:
            return
        estado["i"] = (estado["i"] + 1) % len(quadros)
        spinner.config(text=quadros[estado["i"]])
        overlay.after(110, animar)

    animar()

    # 4. Executa a tarefa em background e monitora a conclusão
    resultado = {}

    def worker():
        try:
            resultado["ok"] = tarefa()
        except Exception as e:
            resultado["erro"] = e

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()

    def checar():
        if thread.is_alive():
            overlay.after(80, checar)
            return
        estado["rodando"] = False
        overlay.destroy()
        if "erro" in resultado:
            if ao_erro:
                ao_erro(resultado["erro"])
        elif ao_concluir:
            ao_concluir(resultado.get("ok"))

    overlay.after(80, checar)
