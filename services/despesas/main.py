import os
import sys

# Garante que o Python encontre os módulos internos de despesas
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from services.despesas.services import processador
from services.despesas.services import reporter



def iniciar_processamento(caminho_excel: str):
    print("Iniciando o programa...")

    try:
    
        df = processador.buscar_dados(caminho_excel)
        
        dados_totais = processador.processar_dados_totais(df)
        dados_loja = processador.processar_dados_loja(df)

        mes = dados_totais['mes']
        # ano = dados_totais['ano'] # Removido se não estiver em uso

        reporter.gerar_relatorio(dados_totais, mes, aba_nome="TOTAL GERAL")
        
        # 2. Itera sobre o DataFrame de lojas e gera uma aba para cada
        for _, linha in dados_loja.iterrows():

            nome_loja = linha['RATEIO']
            print(f"Gerando aba: {nome_loja}")
            reporter.gerar_relatorio(linha.to_dict(), mes, aba_nome=str(nome_loja))

        print("\nProcesso concluído com sucesso!")
    except Exception as e:
        print(f"Erro no fluxo de despesas: {e}")
        raise e 

if __name__ == "__main__":
    iniciar_processamento()