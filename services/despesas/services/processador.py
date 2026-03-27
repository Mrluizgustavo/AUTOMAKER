import pandas as pd
import calendar

def buscar_dados(caminho):   
    df = pd.read_excel(caminho)
    return df


def processar_dados_totais(df):

    total_bruto = df['BRUTO'].sum()

    total_faltas = df['FALTA'].sum() 

    data = df['PERIODO'].mode().iloc[0] 

    mes_num, ano = data.split("/")
    mes_num = int(mes_num)

    mes = calendar.month_abbr[mes_num].upper()
    

    #_VALOR HORAS EXTRAS_60%_E_(100% COM DSR)_
    valor_horas_extras_60 = df['H. EXTRA 60%'].sum()
    valor_horas_extras_100_dsr = (df['H. EXTRA 100%'] + df['DSR']).sum()
    

    #_QUANT. HORAS EXTRAS_60%_E_100%_
    quant_horas_extras_60 = df.loc[df['H. EXTRA 60%'] != 0, 'H. EXTRA 60%'].count()
    quant_horas_extras_100 = df.loc[df['H. EXTRA 100%'] != 0, 'H. EXTRA 100%'].count()

    
    # _BRUTO_
    bruto_real = total_bruto - total_faltas - valor_horas_extras_60 - valor_horas_extras_100_dsr

    # _DESC.FUNC._
    inss_real = df['INSS'].sum()

    #_QTDE. FUNC._
    qtde_func = df['NOME'].count()

    #_VT DESC. FUNC._
    vt_desc_func = df['DESC. V.T.'].sum()

    #_REFEIÇÕES DESC. FUN._
    refeicoes = df['ALMOÇO'].sum()

    convenio = (df['CONVENIO MEDICO'] + df['CO-PARTIC.'] + df['CONVENIO ODONTO']).sum()
    valores = {
        'mes': str(mes),
        'ano' : str(ano),
        'bruto_real': float(bruto_real),
        'inss': float(inss_real),
        'qtde_func': int(qtde_func),
        'vt_desc_func': float(vt_desc_func),
        'refeicoes_desc_func': float(refeicoes),
        'valor_horas_extras_60': float(valor_horas_extras_60),
        'valor_horas_extras_100_com_dsr': float(valor_horas_extras_100_dsr),
        'quant_horas_extras_60': int(quant_horas_extras_60),
        'quant_horas_extras_100': int(quant_horas_extras_100),
        'valor_convenio': float(convenio)
        }
    
    return valores

def processar_dados_loja(df):

    resultado = []

    for loja, grupo in df.groupby('RATEIO'):
        dados = processar_dados_totais(grupo)
        dados['RATEIO'] = loja
        resultado.append(dados)
    
    return pd.DataFrame(resultado)