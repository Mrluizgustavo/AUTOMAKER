import os
from services.reporter import gerar_telegrama
from textwrap import dedent
import os
import sys

sys.path.append(os.path.dirname(os.path.abspath(__file__)))


def iniciar_telegrama():
    arq = r'G:\LUIZ GUSTAVO\PYTHON\AUTOMAKER\services\telegrama\input\Formulário de telegrama - correios.pdf'
    output = "Telegrama_formatado.pdf"
    txt = """EXTINÇÃO DO CONTRATO DE EXPERIÊNCIA\nÁ (Ao)\nSr.(a): WILLIAN JOSÉ DA ROCHA\nPortador(a) do RG – 42.430.756-X\nVimos pela presente comunicar-lhe que seu contrato de experiência termina na data, 09/03/2014, sendo que a partir de então não necessitamos dos seus trabalhos, devendo, portanto, cessar sua atividade na referida data.\nSolicitamos o seu comparecimento em nossa empresa às 16h00hs do dia seguinte, para a quitação das parcelas a que faz jus de acordo com a Legislação Vigente.\nAtenciosamente,\nSUPERMERCADO PARANÁ JANDAIA LTDA\nGuarulhos, 08 de Março de 2014."""

    txt = dedent(txt)
    gerar_telegrama(arq, output, txt)

    # Retorna o caminho absoluto para facilitar a localizacao do arquivo gerado
    return f"Processamento iniciado: {arq}\nSalvo em: {os.path.abspath(output)}"

if __name__ == "__main__":
    resultado = iniciar_telegrama()
    print(resultado)