import pandas as pd

MESES = {
    'JAN': 1, 'FEV': 2, 'MAR': 3, 'ABR': 4,
    'MAI': 5, 'JUN': 6, 'JUL': 7, 'AGO': 8,
    'SET': 9, 'OUT': 10, 'NOV': 11, 'DEZ': 12
}

MESES_PTBR = {v: k for k, v in MESES.items()}  # {1: 'JAN', 2: 'FEV', ...}

MAPA_COLUNAS_FERIAS = {
    'LOJAS': 'LOJA',
    '1098': 'INSS',
    '1169': 'CONV. MEDICO',
    '14':   'CO-PARTIC.',
    '1009': 'CONVENIO ODONTO'
}


def _soma_coluna(df, col):
    """Soma uma coluna numericamente; retorna 0.0 se a coluna não existir."""
    if col not in df.columns:
        return 0.0
    return pd.to_numeric(df[col], errors='coerce').sum()


# ==============================================================================
# LEITURA DE DADOS
# ==============================================================================

def buscar_dados(caminho, aba=0):
    df = pd.read_excel(caminho, sheet_name=aba)
    return df


def buscar_dados_ferias(caminho):
    """
    Lê todas as abas da planilha de férias e consolida em um único DataFrame.
    A filtragem por mês/ano fica a cargo de get_dados_planilha_ferias.
    """
    xls = pd.ExcelFile(caminho)
    frames = []

    for aba in xls.sheet_names:
        try:
            df = pd.read_excel(caminho, sheet_name=aba, header=1)
            df.columns = df.columns.astype(str).str.strip()
            df.rename(columns=MAPA_COLUNAS_FERIAS, inplace=True)

            if 'PAGAMENTO' not in df.columns:
                continue

            df['PAGAMENTO'] = pd.to_datetime(df['PAGAMENTO'], errors='coerce', format='mixed')
            df = df[df['PAGAMENTO'].notna()]

            if df.empty:
                continue

            frames.append(df)

        except Exception as e:
            print(f"  [férias] Aba '{aba}' ignorada por erro: {e}")
            continue

    if not frames:
        print("[férias] Nenhum dado encontrado na planilha.")
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


# ==============================================================================
# CUSTOS
# ==============================================================================

def get_dados_planilha_custos(df_custos):

    total_bruto  = df_custos['BRUTO'].sum()
    total_faltas = df_custos['FALTA'].sum()

    data = df_custos['PERIODO'].mode().iloc[0]
    mes_num, ano = data.split("/")
    mes_num = int(mes_num)

    mes = MESES_PTBR[mes_num]  # PT-BR: 'JAN', 'FEV', etc.

    valor_horas_extras_60      = df_custos['H. EXTRA 60%'].sum()
    valor_horas_extras_100_dsr = (df_custos['H. EXTRA 100%'] + df_custos['DSR']).sum()

    quant_horas_extras_60  = df_custos.loc[df_custos['H. EXTRA 60%']  != 0, 'H. EXTRA 60%'].count()
    quant_horas_extras_100 = df_custos.loc[df_custos['H. EXTRA 100%'] != 0, 'H. EXTRA 100%'].count()

    bruto_real   = total_bruto - total_faltas - valor_horas_extras_60 - valor_horas_extras_100_dsr
    inss_real    = df_custos['INSS'].sum()
    qtde_func    = df_custos['NOME'].count()
    vt_desc_func = df_custos['DESC. V.T.'].sum()
    qtde_func_vt = df_custos.loc[df_custos['DESC. V.T.'] != 0, 'DESC. V.T.'].count()
    refeicoes    = df_custos['ALMOÇO'].sum()
    convenio     = (df_custos['CONVENIO MEDICO'] + df_custos['CO-PARTIC.'] + df_custos['CONVENIO ODONTO']).sum()

    return {
        'mes':                            str(mes),
        'ano':                            str(ano),
        'bruto_real':                     float(bruto_real),
        'inss_planilha_custos':           float(inss_real),
        'qtde_func':                      int(qtde_func),
        'qtde_func_vt':                   int(qtde_func_vt),
        'vt_desc_func':                   float(vt_desc_func),
        'refeicoes_desc_func':            float(refeicoes),
        'valor_horas_extras_60':          float(valor_horas_extras_60),
        'valor_horas_extras_100_com_dsr': float(valor_horas_extras_100_dsr),
        'quant_horas_extras_60':          int(quant_horas_extras_60),
        'quant_horas_extras_100':         int(quant_horas_extras_100),
        'valor_convenio_planilha_custos': float(convenio),
        'rescisoes':                      0.0
    }


def get_dados_planilha_custos_por_loja(df):
    resultado = []
    for loja, grupo in df.groupby('RATEIO'):
        dados = get_dados_planilha_custos(grupo)
        dados['RATEIO'] = loja
        resultado.append(dados)
    return pd.DataFrame(resultado)




def get_dados_planilha_rescisao(df_rescisao, mes, ano):

    num_mes = MESES[mes.upper()]
    num_ano = int(ano)

    df_rescisao['DATA DEMISSÃO'] = pd.to_datetime(df_rescisao['DATA DEMISSÃO'], errors='coerce', format='mixed')

    df_filtrado = df_rescisao[
        (df_rescisao['DATA DEMISSÃO'].dt.month == num_mes) &
        (df_rescisao['DATA DEMISSÃO'].dt.year  == num_ano)
    ]

    total_rescisao = pd.to_numeric(df_filtrado['RESCISÃO'], errors='coerce').sum()
    total_gfd      = pd.to_numeric(df_filtrado['GFD'],      errors='coerce').sum()

    return {
        'rescisao_total': round(float(total_rescisao + total_gfd), 2)
    }


def get_dados_planilha_rescisao_por_loja(df_rescisao, mes, ano):

    lojas_ignoradas = ['CHEGUEI BRASIL', 'ETI']
    df_rescisao = df_rescisao[~df_rescisao['LOJA'].isin(lojas_ignoradas)]

    resultado = []
    for loja, grupo in df_rescisao.groupby('LOJA'):
        dados = get_dados_planilha_rescisao(grupo, mes, ano)
        dados['LOJA'] = loja
        resultado.append(dados)

    return pd.DataFrame(resultado)



def get_dados_planilha_VT(df_vt, mes, ano):

    num_mes = MESES[mes.upper()]
    num_ano = int(ano)

    df_vt['DATA'] = pd.to_datetime(df_vt['DATA'], errors='coerce', format='mixed')

    df_filtrado = df_vt[
        (df_vt['DATA'].dt.month == num_mes) &
        (df_vt['DATA'].dt.year  == num_ano)
    ]

    total = pd.to_numeric(df_filtrado['VALOR'], errors='coerce').sum()

    return {
        'valor_vt': round(float(total), 2)
    }


def get_dados_planilha_VT_por_loja(df_vt, mes, ano):

    lojas_ignoradas = ['CHEGUEI BRASIL', 'ETI']
    df_vt = df_vt[~df_vt['LOJA'].isin(lojas_ignoradas)]

    resultado = []
    for loja, grupo in df_vt.groupby('LOJA'):
        dados = get_dados_planilha_VT(grupo, mes, ano)
        dados['LOJA'] = loja
        resultado.append(dados)

    return pd.DataFrame(resultado)



def get_dados_planilha_ferias(df_ferias, mes, ano):
    """
    Filtra o DataFrame consolidado de férias por mês.
    Em dezembro filtra também por ano, evitando somar registros
    de dezembro do ano anterior que estejam na mesma planilha.
    """
    if df_ferias.empty:
        return {'valor_ferias': 0.0, 'inss_ferias': 0.0, 'convenio_ferias': 0.0}

    num_mes = MESES[mes.upper()]
    num_ano = int(ano)

    if num_mes == 12:
        df_filtrado = df_ferias[
            (df_ferias['PAGAMENTO'].dt.month == num_mes) &
            (df_ferias['PAGAMENTO'].dt.year  == num_ano)
        ]
    else:
        df_filtrado = df_ferias[df_ferias['PAGAMENTO'].dt.month == num_mes]

    total    = pd.to_numeric(df_filtrado['SOMA BRUTO'], errors='coerce').sum()
    inss     = _soma_coluna(df_filtrado, 'INSS')
    convenio = (
        _soma_coluna(df_filtrado, 'CONV. MEDICO')   +
        _soma_coluna(df_filtrado, 'CO-PARTIC.')      +
        _soma_coluna(df_filtrado, 'CONVENIO ODONTO')
    )

    return {
        'valor_ferias':    round(float(total),    2),
        'inss_ferias':     round(float(inss),     2),
        'convenio_ferias': round(float(convenio), 2)
    }


def get_dados_planilha_ferias_por_loja(df_ferias, mes, ano):

    resultado = []
    for loja, grupo in df_ferias.groupby('LOJA'):
        dados = get_dados_planilha_ferias(grupo, mes, ano)
        dados['LOJA'] = loja
        resultado.append(dados)

    return pd.DataFrame(resultado)



def group_SUM_values(planilha_custo, planilha_rescisoes, planilha_VT, df_ferias):

    valores = {}

    custos = get_dados_planilha_custos(planilha_custo)
    mes    = custos['mes']
    ano    = custos['ano']

    rescisao = get_dados_planilha_rescisao(planilha_rescisoes, mes, ano)
    vt       = get_dados_planilha_VT(planilha_VT, mes, ano)
    ferias   = get_dados_planilha_ferias(df_ferias, mes, ano)

    valores.update(custos)
    valores.update(rescisao)
    valores.update(vt)
    valores.update(ferias)

    return valores


def group_LOJAS_values(planilha_custo, planilha_rescisoes, planilha_VT, df_ferias, mes, ano):

    custos   = get_dados_planilha_custos_por_loja(planilha_custo)
    rescisao = get_dados_planilha_rescisao_por_loja(planilha_rescisoes, mes, ano)
    vt       = get_dados_planilha_VT_por_loja(planilha_VT, mes, ano)
    ferias   = get_dados_planilha_ferias_por_loja(df_ferias, mes, ano)

    df = pd.merge(custos,  rescisao, left_on='RATEIO', right_on='LOJA', how='outer')
    df = pd.merge(df,      vt,       left_on='RATEIO', right_on='LOJA', how='outer')
    df = pd.merge(df,      ferias,   left_on='RATEIO', right_on='LOJA', how='outer')

    df.fillna(0, inplace=True)

    return df