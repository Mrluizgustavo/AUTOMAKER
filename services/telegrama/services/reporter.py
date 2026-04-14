import io
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Dicionário unificado para facilitar a iteração
DADOS_POSICOES = {
    "padrao": {
        "COPIA_CONFIRMATÓRIA": (321, 759, "X"),
        "PEDIDO_DE_CONFIRMAÇÃO": (321, 742, "X"),
        "NOME_REMETENTE": (40, 696, "COMERCIO DE ALIMENTOS ALVES E FARIAS"),
        "ENDERECO_REMETENTE": (40, 674, "RUA OSVALDO RAMOS, 100 - PARQUE MIKAIL"),
        "CIDADE_REMETENTE": (40, 651, "GUARULHOS - SP"),
        "TELEFONE_REMETENTE": (235, 651, "11 99999-9999"),
        "CEP": (445, 651, "07142-600")
    },
    "destinatario": {
        "NOME_DESTINATARIO": (40, 625, "NOME DO DESTINATARIO EXEMPLO"), 
        "ENDERECO_DESTINATARIO": (40, 602, "ENDERECO DO DESTINATARIO"),
        "CIDADE_DESTINATARIO": (40, 581, "CIDADE - UF"),
        "TELEFONE_DESTINATARIO": (235, 581, "00 00000-0000"),
        "CEP": (445, 581, "00000-000")
    }
}

def gerar_telegrama(caminho_arq_base, caminho_arq_saida, texto):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)

    fonte_nome = "Courier-Bold"
    tamanho_fonte = 12
    c.setFont(fonte_nome, tamanho_fonte)

    for categoria in DADOS_POSICOES:
        for campo, (x, y, valor) in DADOS_POSICOES[categoria].items():
            if campo == "CEP":

                cep_limpo = valor.replace("-", " ")
                for i, char in enumerate(cep_limpo):
                    c.drawString(x + i * 14.4, y, char)
            else:
                c.drawString(x, y, str(valor))

    y = 553
    x_inicial = 22.5
    col_max = 39          # Limite de quadradinhos por linha
    espacamento_x = 14.1  
    espacamento_y = 22.5  

    coluna = 0
    c.setFont(fonte_nome, 11)

    for char in texto:
        # 1. Se for quebra de linha manual (\n)
        if char == "\n":
            y -= espacamento_y
            coluna = 0
            continue

        # 2. Se a coluna atingir o limite (quebra automática)
        if coluna >= col_max:
            y -= espacamento_y
            coluna = 0

        # 3. Desenha o caractere
        c.drawString(x_inicial + (coluna * espacamento_x), y, char.upper())
        
        # 4. Incrementa para o próximo quadradinho
        coluna += 1

    c.save()
    buffer.seek(0)

    try:
        leitor_base = PdfReader(caminho_arq_base)
        leitor_overlay = PdfReader(buffer)
        escritor = PdfWriter()

        
        pagina_base = leitor_base.pages[0]
        pagina_base.merge_page(leitor_overlay.pages[0])
        escritor.add_page(pagina_base)
        
        with open(caminho_arq_saida, "wb") as f:
            escritor.write(f)
            
    except Exception as e:
        print(f"Erro ao processar PDF: {e}")

