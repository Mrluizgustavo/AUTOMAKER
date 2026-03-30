import os
import sys

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from services.despesas.services import processador
from services.despesas.services import reporter

CAMINHO_CUSTOS   = r'G:\LUIZ GUSTAVO\CÓPIA - Projeto nova Planilha de custo --2.0.xlsx'
CAMINHO_RESCISAO = r'G:\Despesas\Rescisão\Rescisões.xlsx'
CAMINHO_VT       = r'G:\Despesas\Vale Transporte\DESPESAS VT.xlsx'
CAMINHO_FERIAS   = r'G:\Despesas\Mercado - Planilha de férias (ATUAL).xlsx'
CAMINHO_ALMOXARIFADO   = r'G:\GASTOS\GASTOS.xlsx'




def iniciar_processamento(caminho_excel: str):
    print("Iniciando o programa...")

    try:
        df_custos   = processador.buscar_dados(caminho_excel)
        df_rescisao = processador.buscar_dados(CAMINHO_RESCISAO, aba="Valores rescisões")
        df_VT       = processador.buscar_dados(CAMINHO_VT)
        df_ferias   = processador.buscar_dados_ferias(CAMINHO_FERIAS, caminho_custos=caminho_excel)  # carrega tudo + reatribui ADMs
        df_uniforme  = processador.buscar_dados(CAMINHO_ALMOXARIFADO, aba="Envio uniforme", header=1)
        df_materiais = processador.buscar_dados(CAMINHO_ALMOXARIFADO, aba="Envio de materiais", header=1)
        dados_totais = processador.group_SUM_values(df_custos, df_rescisao, df_VT, df_ferias, df_uniforme, df_materiais)
        
        # --- INÍCIO DEBUG ---
        print("\n[DEBUG MATERIAIS]")
        print(f"Colunas encontradas: {df_materiais.columns.tolist()}")
        if not df_materiais.empty and 'DATA' in df_materiais.columns.astype(str).str.strip().str.upper():
             print(f"Total de linhas lidas: {len(df_materiais)}")
        # --- FIM DEBUG ---

        mes = dados_totais['mes']
        ano = dados_totais['ano']

        dados_loja = processador.group_LOJAS_values(df_custos, df_rescisao, df_VT, df_ferias,df_uniforme, df_materiais, mes, ano)

        reporter.gerar_relatorio(dados_totais, mes, ano, aba_nome="TOTAL GERAL")

        for _, linha in dados_loja.iterrows():
            nome_loja = linha['RATEIO']
            print(f"Gerando aba: {nome_loja}")
            reporter.gerar_relatorio(linha.to_dict(), mes, ano, aba_nome=str(nome_loja))

        print("\nProcesso concluído com sucesso!")

    except Exception as e:
        print(f"Erro no fluxo de despesas: {e}")
        raise e


if __name__ == "__main__":
    iniciar_processamento(CAMINHO_CUSTOS)