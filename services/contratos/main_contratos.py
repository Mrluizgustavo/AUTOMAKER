import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from services.ui_theme import (
    ACCENT2, ACCENT3, BG, BORDER, FONT_HEAD, FONT_MONO,
    FONT_SMALL, SURFACE, SURFACE2, TEXT, TEXT_SUB,
    pill_button, _base_tela, executar_com_loading,
)


def _card_input(content, cor, titulo, subtitulo, placeholder, comando_navegar, label_botao):
    """Monta um card padrão com entry + botão de navegação. Retorna (var, entry)."""
    card = tk.Frame(content, bg=SURFACE, padx=24, pady=20)
    card.pack(fill="x", pady=(0, 14))

    tk.Label(card, text=titulo, font=FONT_HEAD, bg=SURFACE, fg=TEXT).pack(anchor="w")
    tk.Label(card, text=subtitulo, font=FONT_SMALL, bg=SURFACE, fg=TEXT_SUB).pack(anchor="w", pady=(2, 12))

    row = tk.Frame(card, bg=SURFACE)
    row.pack(fill="x")

    var = tk.StringVar(value="")
    entry = tk.Entry(row, textvariable=var, bg=SURFACE2, fg=TEXT,
                     insertbackground=TEXT, relief="flat",
                     font=FONT_MONO, highlightthickness=1,
                     highlightbackground=BORDER, highlightcolor=cor)
    entry.pack(side="left", fill="x", expand=True, ipady=8, padx=(0, 10))
    entry.insert(0, placeholder)

    def navegar():
        p = comando_navegar()
        if p:
            var.set(p)
            entry.config(fg=TEXT)

    pill_button(row, label_botao, navegar, color=cor, hover=ACCENT2).pack(side="right")
    return var, entry


# ── Tela: CONTRATOS ──────────────────────────────────────────────────────────
def tela_contratos(parent_frame, roteador=None):
    COR = ACCENT3

    def corpo(root, content, cor):
        # ── Input 1: PDF de origem ──
        caminho_pdf, _ = _card_input(
            content, cor,
            "Arquivo de Contratos (PDF)",
            "Selecione o PDF com os contratos de trabalho a serem separados",
            "Nenhum arquivo selecionado...",
            lambda: filedialog.askopenfilename(
                title="Selecione o PDF de contratos",
                filetypes=[("PDF", "*.pdf")],
            ),
            "Navegar",
        )

        # ── Input 2: pasta de destino ──
        pasta_saida, _ = _card_input(
            content, cor,
            "Pasta de Destino",
            "Onde os contratos separados serão salvos",
            "Nenhuma pasta selecionada...",
            lambda: filedialog.askdirectory(title="Selecione a pasta de destino"),
            "Escolher pasta",
        )

        # ── Botão executar ──
        def executar():
            arq = caminho_pdf.get()
            pasta = pasta_saida.get()

            if not arq or "Nenhum" in arq:
                messagebox.showwarning("Atenção", "Selecione o PDF de contratos.")
                return
            if not pasta or "Nenhuma" in pasta:
                messagebox.showwarning("Atenção", "Selecione a pasta de destino.")
                return

            from services.contratos.services.separador import separar_contratos

            def tarefa():
                return separar_contratos(arq, pasta)

            def concluir(resultados):
                sem_loja = [r for r in resultados if not r["loja"]]
                resumo = f"{len(resultados)} contrato(s) separado(s) em:\n{pasta}"
                if sem_loja:
                    resumo += (f"\n\n⚠ {len(sem_loja)} sem loja identificada "
                               "(CNPJ não consta na relação).")
                if messagebox.askyesno("Sucesso", resumo + "\n\nDeseja abrir a pasta?"):
                    os.startfile(pasta)

            def erro(e):
                messagebox.showerror("Erro", f"Não foi possível separar os contratos:\n{e}")

            executar_com_loading(root, tarefa, ao_concluir=concluir, ao_erro=erro,
                                 texto="Separando contratos...")

        pill_button(content, "▶  Separar Contratos", executar,
                    color=cor, hover=ACCENT2).pack(fill="x", pady=4, ipady=4)

    _base_tela(parent_frame, "Contratos", "📄", COR, corpo, roteador)
