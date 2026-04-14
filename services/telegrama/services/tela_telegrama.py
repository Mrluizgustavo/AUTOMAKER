import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from services.ui_theme import (
    ACCENT2, ACCENT, BG, BORDER, FONT_BODY, FONT_HEAD, FONT_MONO,
    FONT_SMALL, FONT_TITLE, SURFACE, SURFACE2,
    TEXT, TEXT_SUB, pill_button, _base_tela,
)

ARQ_BASE = r'G:\LUIZ GUSTAVO\PYTHON\AUTOMAKER\services\telegrama\input\Formulário de telegrama - correios.pdf'


def criar_container_scrollable(parent):
    main_frame = tk.Frame(parent, bg=BG)
    main_frame.pack(fill="both", expand=True)

    canvas = tk.Canvas(main_frame, bg=BG, highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)

    scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollable_frame = tk.Frame(canvas, bg=BG)
    canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    def _on_canvas_configure(e):
        canvas.itemconfig(canvas_window, width=e.width)

    def _on_frame_configure(e):
        canvas.configure(scrollregion=canvas.bbox("all"))

    scrollable_frame.bind("<Configure>", _on_frame_configure)
    canvas.bind("<Configure>", _on_canvas_configure)

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # Vincula o scroll apenas ao canvas, não globalmente
    canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
    canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

    return scrollable_frame


def _campo(parent, label, cor, placeholder=""):
    tk.Label(parent, text=label, font=FONT_BODY, bg=SURFACE, fg=TEXT_SUB).pack(anchor="w", pady=(8, 2))

    var = tk.StringVar()
    entry = tk.Entry(
        parent, textvariable=var,
        bg=SURFACE2, fg=TEXT_SUB if placeholder else TEXT,
        insertbackground=TEXT, relief="flat",
        font=FONT_MONO, highlightthickness=1,
        highlightbackground=BORDER, highlightcolor=cor,
    )
    entry.pack(fill="x", ipady=7)

    if placeholder:
        entry.insert(0, placeholder)

        def _in(e):
            if entry.get() == placeholder:
                entry.delete(0, "end")
                entry.config(fg=TEXT)

        def _out(e):
            if not entry.get():
                entry.insert(0, placeholder)
                entry.config(fg=TEXT_SUB)

        entry.bind("<FocusIn>", _in)
        entry.bind("<FocusOut>", _out)

    return var, entry


def tela_telegrama(root):
    COR = ACCENT

    def corpo(root, content, cor):
        # Container scrollável criado aqui dentro, usando o root correto
        content_scroll = criar_container_scrollable(content)

        # ── Card 1: Dados do destinatário ─────────────────────────────────────
        card_dest = tk.Frame(content_scroll, bg=SURFACE, padx=24, pady=18)
        card_dest.pack(fill="x", pady=(10, 10), anchor="n")

        tk.Label(card_dest, text="Dados do Destinatário", font=FONT_HEAD, bg=SURFACE, fg=TEXT).pack(anchor="w")
        tk.Label(card_dest, text="Preencha os campos para montar o telegrama",
                 font=FONT_SMALL, bg=SURFACE, fg=TEXT_SUB).pack(anchor="w", pady=(2, 4))

        v_nome,     _ = _campo(card_dest, "Nome",     cor, "Nome completo do destinatário")
        v_endereco, _ = _campo(card_dest, "Endereço", cor, "Logradouro e número")
        v_cidade,   _ = _campo(card_dest, "Cidade",   cor, "Cidade - UF")

        row_fc = tk.Frame(card_dest, bg=SURFACE)
        row_fc.pack(fill="x", pady=(8, 0))

        col_fone = tk.Frame(row_fc, bg=SURFACE)
        col_fone.pack(side="left", fill="x", expand=True, padx=(0, 12))
        tk.Label(col_fone, text="Fone", font=FONT_BODY, bg=SURFACE, fg=TEXT_SUB).pack(anchor="w", pady=(0, 2))
        v_fone = tk.StringVar()
        ef = tk.Entry(col_fone, textvariable=v_fone, bg=SURFACE2, fg=TEXT_SUB,
                      insertbackground=TEXT, relief="flat", font=FONT_MONO,
                      highlightthickness=1, highlightbackground=BORDER, highlightcolor=cor)
        ef.insert(0, "00 00000-0000")
        ef.pack(fill="x", ipady=7)
        ef.bind("<FocusIn>",  lambda e: (ef.delete(0, "end"), ef.config(fg=TEXT)) if ef.get() == "00 00000-0000" else None)
        ef.bind("<FocusOut>", lambda e: (ef.insert(0, "00 00000-0000"), ef.config(fg=TEXT_SUB)) if not ef.get() else None)

        col_cep = tk.Frame(row_fc, bg=SURFACE)
        col_cep.pack(side="left", fill="x", expand=True)
        tk.Label(col_cep, text="CEP", font=FONT_BODY, bg=SURFACE, fg=TEXT_SUB).pack(anchor="w", pady=(0, 2))
        v_cep = tk.StringVar()
        ec = tk.Entry(col_cep, textvariable=v_cep, bg=SURFACE2, fg=TEXT_SUB,
                      insertbackground=TEXT, relief="flat", font=FONT_MONO,
                      highlightthickness=1, highlightbackground=BORDER, highlightcolor=cor)
        ec.insert(0, "00000-000")
        ec.pack(fill="x", ipady=7)
        ec.bind("<FocusIn>",  lambda e: (ec.delete(0, "end"), ec.config(fg=TEXT)) if ec.get() == "00000-000" else None)
        ec.bind("<FocusOut>", lambda e: (ec.insert(0, "00000-000"), ec.config(fg=TEXT_SUB)) if not ec.get() else None)

        # ── Card 2: Mensagem ──────────────────────────────────────────────────
        card_msg = tk.Frame(content_scroll, bg=SURFACE, padx=24, pady=18)
        card_msg.pack(fill="x", pady=(0, 10))

        tk.Label(card_msg, text="Mensagem", font=FONT_HEAD, bg=SURFACE, fg=TEXT).pack(anchor="w")
        tk.Label(card_msg, text="Texto que será inserido no corpo do telegrama",
                 font=FONT_SMALL, bg=SURFACE, fg=TEXT_SUB).pack(anchor="w", pady=(2, 8))

        frame_txt = tk.Frame(card_msg, bg=SURFACE)
        frame_txt.pack(fill="x")

        txt_msg = tk.Text(
            frame_txt,
            bg=SURFACE2, fg=TEXT,
            insertbackground=TEXT, relief="flat",
            font=FONT_MONO, highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=cor,
            height=8, wrap="word",
            padx=8, pady=6,
        )
        txt_msg.pack(side="left", fill="x", expand=True)

        sb = tk.Scrollbar(frame_txt, command=txt_msg.yview, bg=SURFACE2)
        sb.pack(side="right", fill="y")
        txt_msg.config(yscrollcommand=sb.set)

        # ── Card 3: Onde salvar ───────────────────────────────────────────────
        card_saida = tk.Frame(content_scroll, bg=SURFACE, padx=24, pady=18)
        card_saida.pack(fill="x", pady=(0, 10))

        tk.Label(card_saida, text="Salvar Arquivo Gerado", font=FONT_HEAD, bg=SURFACE, fg=TEXT).pack(anchor="w")
        tk.Label(card_saida, text="Escolha a pasta e o nome do PDF que será gerado",
                 font=FONT_SMALL, bg=SURFACE, fg=TEXT_SUB).pack(anchor="w", pady=(2, 10))

        caminho_saida = tk.StringVar()
        row_saida = tk.Frame(card_saida, bg=SURFACE)
        row_saida.pack(fill="x")

        entry_saida = tk.Entry(
            row_saida, textvariable=caminho_saida,
            bg=SURFACE2, fg=TEXT_SUB,
            insertbackground=TEXT, relief="flat",
            font=FONT_MONO, highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=cor,
        )
        entry_saida.insert(0, "Destino não definido...")
        entry_saida.pack(side="left", fill="x", expand=True, ipady=7, padx=(0, 10))

        def selecionar_saida():
            p = filedialog.asksaveasfilename(
                title="Salvar telegrama como",
                defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")],
                initialfile="Telegrama_gerado.pdf",
            )
            if p:
                caminho_saida.set(p)
                entry_saida.config(fg=TEXT)

        pill_button(row_saida, "Escolher local", selecionar_saida, color=cor, hover=ACCENT2).pack(side="right")

        # ── Botão executar ────────────────────────────────────────────────────
        def executar():
            arq_saida = caminho_saida.get()
            nome      = v_nome.get().strip()
            endereco  = v_endereco.get().strip()
            cidade    = v_cidade.get().strip()
            fone      = v_fone.get().strip()
            cep       = v_cep.get().strip()
            mensagem  = txt_msg.get("1.0", "end").strip()

            PLACEHOLDERS = {
                "Destino não definido...", "Nome completo do destinatário",
                "Logradouro e número", "Cidade - UF",
                "00 00000-0000", "00000-000",
            }

            erros = []
            if not arq_saida or arq_saida in PLACEHOLDERS:
                erros.append("• Defina onde o arquivo será salvo.")
            if not nome or nome in PLACEHOLDERS:
                erros.append("• Preencha o nome do destinatário.")
            if not endereco or endereco in PLACEHOLDERS:
                erros.append("• Preencha o endereço.")
            if not cidade or cidade in PLACEHOLDERS:
                erros.append("• Preencha a cidade.")
            if not mensagem:
                erros.append("• A mensagem não pode estar vazia.")

            if erros:
                messagebox.showwarning("Campos obrigatórios", "\n".join(erros))
                return

            try:
                import services.reporter as rp

                rp.DADOS_POSICOES["destinatario"]["NOME_DESTINATARIO"]     = (40, 625, nome.upper())
                rp.DADOS_POSICOES["destinatario"]["ENDERECO_DESTINATARIO"] = (40, 602, endereco.upper())
                rp.DADOS_POSICOES["destinatario"]["CIDADE_DESTINATARIO"]   = (40, 581, cidade.upper())
                rp.DADOS_POSICOES["destinatario"]["TELEFONE_DESTINATARIO"] = (235, 581, fone)
                rp.DADOS_POSICOES["destinatario"]["CEP"]                   = (445, 581, cep)

                rp.gerar_telegrama(ARQ_BASE, arq_saida, mensagem)

                if messagebox.askyesno("Sucesso", f"Telegrama gerado!\n\n{arq_saida}\n\nDeseja abrir o arquivo?"):
                    os.startfile(arq_saida)

            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível gerar o telegrama:\n{e}")

        pill_button(
            content_scroll, "▶  Gerar Telegrama", executar,
            color=cor, hover=ACCENT2,
        ).pack(fill="x", pady=(4, 0), ipady=4)

    _base_tela(root, "Monte seu Telegrama", "📃", COR, corpo)