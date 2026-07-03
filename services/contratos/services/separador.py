"""
Núcleo de separação de contratos de trabalho.

Recebe um PDF contendo vários contratos concatenados e o divide em arquivos
individuais. Cada contrato começa numa página cujo título é
"CONTRATO DE TRABALHO A TÍTULO DE EXPERIÊNCIA".

De cada contrato, apenas da PRIMEIRA página, são extraídos:
  • Nome do funcionário  → texto após "Sr.(a):"
  • CNPJ                  → texto após "CNPJ:"

O CNPJ é traduzido para o Nº da loja via RELACAO_CNPJ_LOJA e cada arquivo é
salvo no padrão: "Nº LOJA - NOME FUNCIONÁRIO.pdf".
"""

import os
import re
import unicodedata

import pdfplumber
from PyPDF2 import PdfReader, PdfWriter


# ── Relação CNPJ → Nº LOJA ────────────────────────────────────────────────────
# Chave = CNPJ da loja | Valor = número da loja.
# O CNPJ pode ser escrito COM ou SEM pontuação — a busca normaliza sozinha.
# Ex.: "12.345.678/0001-90" e "12345678000190" funcionam igual.
RELACAO_CNPJ_LOJA = {
    "01.444.210/0001-31": "1",
    "01.444.210/0010-22": "2",
    "01.444.210/0009-99": "3",
    "01.444.210/0008-08": "4",
    "01.444.210/0007-27": "5",
    "01.444.210/0006-46": "6",
    "01.444.210/0005-65": "7",
    "01.444.210/0004-84": "8",
    "01.444.210/0003-01": "9",
    "01.444.210/0002-12": "10",
    "01.444.210/0011-03": "11",
    "01.444.210/0012-94": "12",
    "01.444.210/0013-75": "13",
    "01.444.210/0014-56": "14",
    "01.444.210/0015-37": "15"

}


TITULO_CONTRATO = "CONTRATO DE TRABALHO A TÍTULO DE EXPERIÊNCIA"

# A página de PRORROGAÇÃO contém o título acima como substring, mas é um anexo
# do MESMO contrato — não deve iniciar um novo bloco.
TITULO_PRORROGACAO = "PRORROGAÇÃO DO CONTRATO DE TRABALHO A TÍTULO DE EXPERIÊNCIA"

# Marcadores que costumam vir logo após o nome, usados para cortar sobras.
_MARCADORES_POS_NOME = (
    " titular", " portadora", " brasileiro", " brasileira",
    " inscrito", " inscrita", " cpf", " rg", " ctps", " residente",
)


# ── Helpers de normalização ───────────────────────────────────────────────────
def _sem_acentos(texto):
    nfkd = unicodedata.normalize("NFKD", texto or "")
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def _so_digitos(texto):
    return re.sub(r"\D", "", texto or "")


# Títulos normalizados (sem acento, maiúsculo) — comparados contra o texto das páginas.
_TITULO_NORM = _sem_acentos(TITULO_CONTRATO).upper()
_PRORROGACAO_NORM = _sem_acentos(TITULO_PRORROGACAO).upper()


def _limpar_nome_arquivo(nome):
    """Remove caracteres inválidos para nome de arquivo no Windows."""
    nome = re.sub(r'[\\/:*?"<>|]', "", nome or "")
    return re.sub(r"\s+", " ", nome).strip()


# ── Extração de campos (sempre da primeira página do contrato) ─────────────────
def _extrair_nome(texto_pagina):
    m = re.search(r"Sr\.?\s*\(a\):\s*([^\n]+)", texto_pagina, re.IGNORECASE)
    if not m:
        return ""
    nome = m.group(1)

    # Corta em vírgula (padrão "Sr.(a): FULANO DE TAL, brasileiro, ...")
    nome = nome.split(",")[0]

    # Rede de segurança: corta em marcadores que sinalizam fim do nome
    nome_low = nome.lower()
    for marc in _MARCADORES_POS_NOME:
        pos = nome_low.find(marc)
        if pos != -1:
            nome = nome[:pos]
            nome_low = nome.lower()

    return nome.strip()


def _e_inicio_contrato(texto_pagina):
    """
    True se a página inicia um NOVO contrato.
    Ignora a página de "PRORROGAÇÃO DO CONTRATO..." — que contém o título base
    como substring, mas é apenas um anexo do contrato anterior.
    """
    t = _sem_acentos(texto_pagina).upper()
    if _TITULO_NORM not in t:
        return False
    # Remove as ocorrências de prorrogação; se ainda sobrar o título base,
    # é um contrato de verdade. Se só existia dentro da prorrogação, não é.
    return _TITULO_NORM in t.replace(_PRORROGACAO_NORM, "")


def _extrair_cnpj(texto_pagina):
    m = re.search(r"CNPJ:\s*([\d.\-/]+)", texto_pagina, re.IGNORECASE)
    return _so_digitos(m.group(1)) if m else ""


def _nome_unico(pasta, nome_arquivo):
    """Evita sobrescrita: acrescenta (2), (3)... se o arquivo já existir."""
    destino = os.path.join(pasta, nome_arquivo)
    if not os.path.exists(destino):
        return destino
    base, ext = os.path.splitext(nome_arquivo)
    i = 2
    while True:
        cand = os.path.join(pasta, f"{base} ({i}){ext}")
        if not os.path.exists(cand):
            return cand
        i += 1


# ── Rotina principal ──────────────────────────────────────────────────────────
def separar_contratos(caminho_pdf, pasta_saida, relacao=None):
    """
    Separa o PDF em contratos individuais e os grava em pasta_saida.
    Retorna uma lista de dicts com o resumo de cada arquivo gerado.
    """
    relacao = relacao if relacao is not None else RELACAO_CNPJ_LOJA
    # Normaliza as chaves para só dígitos: aceita CNPJ com ou sem pontuação.
    relacao = {_so_digitos(k): v for k, v in relacao.items()}
    os.makedirs(pasta_saida, exist_ok=True)

    # 1. Texto de cada página
    with pdfplumber.open(caminho_pdf) as pdf:
        textos = [(pg.extract_text() or "") for pg in pdf.pages]

    total_paginas = len(textos)

    # 2. Páginas que iniciam um novo contrato (a prorrogação NÃO conta)
    inicios = [i for i, txt in enumerate(textos) if _e_inicio_contrato(txt)]

    if not inicios:
        raise ValueError(
            "Nenhum contrato encontrado: o título "
            f'"{TITULO_CONTRATO}" não foi localizado em nenhuma página.'
        )

    # 3. Cada contrato vai da sua página inicial até a página anterior à próxima
    leitor = PdfReader(caminho_pdf)
    resultados = []

    for idx, inicio in enumerate(inicios):
        fim = inicios[idx + 1] if idx + 1 < len(inicios) else total_paginas
        primeira = textos[inicio]

        nome = _extrair_nome(primeira)
        cnpj = _extrair_cnpj(primeira)
        loja = relacao.get(cnpj, "")

        nome_fmt = _limpar_nome_arquivo(nome) or f"CONTRATO_{idx + 1}"
        loja_fmt = _limpar_nome_arquivo(loja) or (cnpj or "SEM_LOJA")
        destino = _nome_unico(pasta_saida, f"{loja_fmt} - {nome_fmt}.pdf")

        escritor = PdfWriter()
        for p in range(inicio, fim):
            escritor.add_page(leitor.pages[p])
        with open(destino, "wb") as f:
            escritor.write(f)

        resultados.append({
            "arquivo": os.path.basename(destino),
            "nome": nome,
            "cnpj": cnpj,
            "loja": loja,
            "paginas": (inicio + 1, fim),
        })

    return resultados
