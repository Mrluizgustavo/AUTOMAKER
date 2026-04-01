import os
import sys

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from services.despesas.services import processador
from services.despesas.services import reporter

#from services import processador
#from services import reporter

CAMINHO_RESCISAO = r'G:\Despesas\Rescisão\Rescisões.xlsx'
CAMINHO_VT       = r'G:\Despesas\Vale Transporte\DESPESAS VT.xlsx'
CAMINHO_FERIAS   = r'G:\Despesas\Mercado - Planilha de férias (ATUAL).xlsx'
CAMINHO_ALMOXARIFADO = r'G:\GASTOS\GASTOS.xlsx'
CAMINHO_PLANILHA_IMPOSTO = r'G:\PILOTO - CONTROLE DE TRIBUTOS.xlsx'


def iniciar_processamento(caminho_excel: str):
    print("Iniciando o programa...")

    try:

        df_custos    = processador.buscar_dados(caminho_excel)
        df_rescisao  = processador.buscar_dados(CAMINHO_RESCISAO, aba="Valores rescisões")
        df_VT        = processador.buscar_dados_vt(CAMINHO_VT)
        df_ferias    = processador.buscar_dados_ferias(CAMINHO_FERIAS, caminho_custos=caminho_excel) 
        df_uniforme  = processador.buscar_dados(CAMINHO_ALMOXARIFADO, aba="Envio uniforme", header=1)
        df_materiais = processador.buscar_dados(CAMINHO_ALMOXARIFADO, aba="Envio de materiais", header=1)
        dados_totais = processador.group_SUM_values(df_custos, df_rescisao, df_VT, df_ferias, df_uniforme, df_materiais, CAMINHO_PLANILHA_IMPOSTO)
        
        mes = dados_totais['mes']
        ano = dados_totais['ano']
        for cat in ['FGTS', 'FGTS APRENDIZES', 'GPS']:
            if cat not in dados_totais:
                dados_totais[cat] = {}


        reporter.gerar_relatorio(dados_totais, mes, ano, aba_nome="TOTAL GERAL")

        dados_loja = processador.group_LOJAS_values(df_custos, df_rescisao, df_VT, df_ferias,df_uniforme, df_materiais, CAMINHO_PLANILHA_IMPOSTO, mes, ano)


        #RETIRA O VALOR DO ADM DA LOJA 01
        loja_alvo = 1 if 1 in dados_loja.index else '1'

        if loja_alvo in dados_loja.index:
            # Subtraímos os 95.000 do valor que já existe lá
            valor_atual = dados_loja.at[loja_alvo, 'FGTS']
            dados_loja.at[loja_alvo, 'FGTS'] = valor_atual - 95000


        for id_loja, linha in dados_loja.iterrows():
            nome_loja = str(id_loja)
            valores_loja = linha.to_dict()

            if nome_loja.upper() == "ADM":

                valores_loja['FGTS'] = 95000  # Adiciona o valor fixo para ADM
                valores_loja['FGTS APRENDIZES'] = 0 # Adiciona o valor fixo para ADM
                valores_loja['GPS'] = 0 # Adiciona o valor fixo para ADM
            
            print(f"Gerando aba: {nome_loja}")
            reporter.gerar_relatorio(valores_loja, mes, ano, aba_nome=nome_loja)

        print("\nProcesso concluído com sucesso!")

    except Exception as e:
        print(f"Erro no fluxo de despesas: {e}")
        raise e


if __name__ == "__main__":
    iniciar_processamento(r'G:\Backup\2025\Jan2025\Projeto nova Planilha de custo -- versão 2.0.xlsx')