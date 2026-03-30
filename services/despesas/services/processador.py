import pandas as pd
import re

MESES = {
    'JAN': 1, 'FEV': 2, 'MAR': 3, 'ABR': 4,
    'MAI': 5, 'JUN': 6, 'JUL': 7, 'AGO': 8,
    'SET': 9, 'OUT': 10, 'NOV': 11, 'DEZ': 12
}

MESES_PTBR = {v: k for k, v in MESES.items()}

MAPA_COLUNAS_FERIAS = {
    'LOJAS': 'LOJA',
    '1098': 'INSS'
}
CODIGOS_CONVENIO = ['1169', '1251', '1204', '1009', '14']
LOJAS_IGNORADAS = {'CHEGUEI BRASIL', 'ETI', 'PADARIA'}


def _normalizar_loja(val):
    """
    Converte 13.0 -> 13 (int), mantém 'ADM' como 'ADM'.
    Evita mismatch entre int e float nas chaves de merge,
    sem descartar strings legítimas como 'ADM'.
    """
    try:
        f = float(val)
        return int(f) if f == int(f) else f
    except (ValueError, TypeError):
        return val


def _normalizar_coluna_loja(series):
    return series.apply(_normalizar_loja)


def _soma_coluna(df, col):
    if col not in df.columns:
        return 0.0
    return pd.to_numeric(df[col], errors='coerce').sum()


# ==============================================================================
# LEITURA DE DADOS
# ==============================================================================

def buscar_dados(caminho, aba=0, header=0):
    return pd.read_excel(caminho, sheet_name=aba, header=header)


def buscar_dados_ferias(caminho, caminho_custos=None):
    nomes_adm = _carregar_nomes_adm(caminho_custos) if caminho_custos else set()

    with pd.ExcelFile(caminho) as xls:
        frames = []
        for aba in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=aba, header=0)
                df.columns = df.columns.astype(str).str.strip().str.upper()

                # CORREÇÃO CRUCIAL: Padroniza o nome da coluna de filial
                if 'LOJAS' in df.columns:
                    df.rename(columns={'LOJAS': 'LOJA'}, inplace=True)
                

                col_data = 'DATA' if 'DATA' in df.columns else 'PAGAMENTO'
                
                if col_data not in df.columns:
                    continue

                df['PAGAMENTO_REF'] = pd.to_datetime(df[col_data], errors='coerce', format='mixed')
                df = df[df['PAGAMENTO_REF'].notna()].copy()

                if df.empty:
                    continue

                if nomes_adm and 'NOME' in df.columns:
                    df = _tratar_adms_ferias(df, nomes_adm)

                frames.append(df)

            except Exception as e:
                print(f"[férias] Aba '{aba}' ignorada: {e}")

    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def _normalizar_nome(nome):
    """Remove espaços nas bordas e espaços duplos internos, converte para maiúsculo."""
    if pd.isna(nome):
        return ''
    return ' '.join(str(nome).strip().upper().split())


def _carregar_nomes_adm(caminho_custos):
    """
    Retorna um set de nomes normalizados da aba ADM da planilha de custos.
    Normalização: strip + upper + espaços duplos internos removidos.
    """
    try:
        df_adm = pd.read_excel(caminho_custos, sheet_name='ADM')
        nomes = set(df_adm['NOME'].dropna().apply(_normalizar_nome))
        nomes.discard('')
        print(f"[ADM] {len(nomes)} nomes carregados da aba ADM.")
        return nomes
    except Exception as e:
        print(f"[ADM] Não foi possível carregar nomes ADM: {e}")
        return set()


def _tratar_adms_ferias(df, nomes_adm):
    if 'LOJA' not in df.columns:
        return df

    nome_normalizado = df['NOME'].apply(_normalizar_nome)
    
    # Previne erro de tipagem forçando a coluna a aceitar strings
    df['LOJA'] = df['LOJA'].astype(object)
    
    mask_is_adm = nome_normalizado.isin(nomes_adm)
    df.loc[mask_is_adm, 'LOJA'] = 'ADM'
    
    loja_normalizada = df['LOJA'].astype(str).str.strip().str.upper()
    mask_adm_invalido = (loja_normalizada == 'ADM') & (~mask_is_adm)
    
    return df[~mask_adm_invalido].copy()


# ==============================================================================
# CUSTOS
# ==============================================================================

def get_dados_planilha_custos(df_custos):
    total_bruto  = df_custos['BRUTO'].sum()
    total_faltas = df_custos['FALTA'].sum()

    data = df_custos['PERIODO'].mode().iloc[0]
    mes_num, ano = data.split("/")
    mes_num = int(mes_num)
    mes = MESES_PTBR[mes_num]

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

    convenio = (
        _soma_coluna(df_custos, 'CONVENIO MEDICO') +
        _soma_coluna(df_custos, 'CO-PARTIC.') +
        _soma_coluna(df_custos, 'CONVENIO ODONTO')
    )

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


# ==============================================================================
# RESCISÃO
# ==============================================================================

def get_dados_planilha_rescisao(df_rescisao, mes, ano):
    num_mes = MESES[mes.upper()]
    num_ano = int(ano)

    df_rescisao = df_rescisao.copy()
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
    df = df_rescisao[~df_rescisao['LOJA'].isin(LOJAS_IGNORADAS)].copy()
    df = df[df['LOJA'].notna()].copy()
    df['LOJA'] = _normalizar_coluna_loja(df['LOJA'])

    resultado = []
    for loja, grupo in df.groupby('LOJA'):
        dados = get_dados_planilha_rescisao(grupo, mes, ano)
        dados['LOJA'] = loja
        resultado.append(dados)

    return pd.DataFrame(resultado)


# ==============================================================================
# VALE TRANSPORTE
# ==============================================================================

def get_dados_planilha_VT(df_vt, mes, ano):
    num_mes = MESES[mes.upper()]
    num_ano = int(ano)

    df_vt = df_vt.copy()
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
    df = df_vt[~df_vt['LOJA'].isin(LOJAS_IGNORADAS)].copy()
    df = df[df['LOJA'].notna()].copy()
    df['LOJA'] = _normalizar_coluna_loja(df['LOJA'])

    resultado = []
    for loja, grupo in df.groupby('LOJA'):
        dados = get_dados_planilha_VT(grupo, mes, ano)
        dados['LOJA'] = loja
        resultado.append(dados)

    return pd.DataFrame(resultado)


# ==============================================================================
# FÉRIAS
# ==============================================================================

def get_dados_planilha_ferias(df_ferias, mes, ano):
    if df_ferias.empty:
        return {'valor_ferias': 0.0, 'inss_ferias': 0.0, 'convenio_ferias': 0.0}

    num_mes = MESES[mes.upper()]
    num_ano = int(ano)

    df_filtrado = df_ferias[
        (df_ferias['PAGAMENTO_REF'].dt.month == num_mes) &
        (df_ferias['PAGAMENTO_REF'].dt.year  == num_ano)
    ].copy()

    # Busca as colunas bases diretamente pelo nome da linha 1
    total = pd.to_numeric(df_filtrado['SOMA BRUTO'], errors='coerce').sum()
    inss  = _soma_coluna(df_filtrado, 'INSS')

    # LÓGICA DINÂMICA: Pega tudo que tiver MÉDICO ou ODONTO no título
    cols_convenio = [col for col in df_filtrado.columns if 'MÉDICO' in col or 'ODONTO' in col]
    convenio = sum(_soma_coluna(df_filtrado, col) for col in cols_convenio)

    return {
        'valor_ferias':    round(float(total),    2),
        'inss_ferias':     round(float(inss),     2),
        'convenio_ferias': round(float(convenio), 2)
    }


def get_dados_planilha_ferias_por_loja(df_ferias, mes, ano):
    if df_ferias.empty:
        return pd.DataFrame()

    df = df_ferias[~df_ferias['LOJA'].isin(LOJAS_IGNORADAS)].copy()
    df = df[df['LOJA'].notna()].copy()
    df['LOJA'] = _normalizar_coluna_loja(df['LOJA'])

    resultado = []
    for loja, grupo in df.groupby('LOJA'):
        dados = get_dados_planilha_ferias(grupo, mes, ano)
        dados['LOJA'] = loja
        resultado.append(dados)

    return pd.DataFrame(resultado)


# =================================================================
# ALMOXARIFADO
# ================================================================= 

def get_dados_gastos_almoxarifado(df_almoxarifado, mes, ano, chave_retorno):
    if df_almoxarifado.empty:
        return {chave_retorno: 0.0}

    df_almoxarifado.columns = df_almoxarifado.columns.astype(str).str.strip().str.upper()

    if 'DATA' not in df_almoxarifado.columns:
        return {chave_retorno: 0.0}

    num_mes = MESES[mes.upper()]
    num_ano = int(ano)

    df_temp = df_almoxarifado.copy()
    df_temp['DATA'] = pd.to_datetime(df_temp['DATA'], errors='coerce', format='mixed')

    df_filtrado = df_temp[
        (df_temp['DATA'].dt.month == num_mes) &
        (df_temp['DATA'].dt.year  == num_ano)
    ]

    total = pd.to_numeric(df_filtrado['VALOR'], errors='coerce').sum()

    return {chave_retorno: round(float(total), 2)}


def get_dados_gastos_almoxarifado_por_loja(df_almoxarifado, mes, ano, chave_retorno):
    if df_almoxarifado.empty:
        return pd.DataFrame()

    df = df_almoxarifado.copy()
    df.columns = df.columns.astype(str).str.strip().str.upper()

    # Padroniza a nomenclatura caso a aba utilize 'DESTINO' em vez de 'LOJA'
    if 'DESTINO' in df.columns:
        df.rename(columns={'DESTINO': 'LOJA'}, inplace=True)

    if 'LOJA' not in df.columns or 'DATA' not in df.columns:
        return pd.DataFrame()

    num_mes = MESES[mes.upper()]
    num_ano = int(ano)

    df['LOJA'] = df['LOJA'].apply(_tratar_loja_almoxarifado)
    df['LOJA'] = _normalizar_coluna_loja(df['LOJA'])
    df['DATA'] = pd.to_datetime(df['DATA'], errors='coerce', format='mixed')

    df_filtrado = df[
        (df['DATA'].dt.month == num_mes) &
        (df['DATA'].dt.year  == num_ano) &
        (~df['LOJA'].isin(LOJAS_IGNORADAS)) &
        (df['LOJA'].notna())
    ]

    resultado = []
    for loja, grupo in df_filtrado.groupby('LOJA'):
        total = pd.to_numeric(grupo['VALOR'], errors='coerce').sum()
        resultado.append({'LOJA': loja, chave_retorno: round(float(total), 2)})

    return pd.DataFrame(resultado)

# ==============================================================================
# AGRUPAMENTOS
# ==============================================================================

def _tratar_loja_almoxarifado(val):
    val_str = str(val).strip().upper()
    
    if 'LOJA' in val_str:
        numero = re.sub(r'\D', '', val_str)
        return int(numero) if numero else 'ADM'
    elif val_str.replace('.', '', 1).isdigit():
        return float(val_str) if '.' in val_str else int(val_str)
    
    return 'ADM'

def group_SUM_values(planilha_custo, planilha_rescisoes, planilha_VT, df_ferias, df_uniforme, df_materiais):
    custos   = get_dados_planilha_custos(planilha_custo)
    mes, ano = custos['mes'], custos['ano']

    rescisao = get_dados_planilha_rescisao(planilha_rescisoes, mes, ano)
    vt       = get_dados_planilha_VT(planilha_VT, mes, ano)
    ferias   = get_dados_planilha_ferias(df_ferias, mes, ano)

    uniformes = get_dados_gastos_almoxarifado(df_uniforme, mes, ano, 'valor_uniforme')
    materiais = get_dados_gastos_almoxarifado(df_materiais, mes, ano, 'valor_materiais')

    valores = {}
    valores.update(custos)
    valores.update(rescisao)
    valores.update(vt)
    valores.update(ferias)
    valores.update(uniformes)
    valores.update(materiais)

    return valores


def group_LOJAS_values(planilha_custo, planilha_rescisoes, planilha_VT, df_ferias, df_uniforme, df_materiais, mes, ano):
    custos    = get_dados_planilha_custos_por_loja(planilha_custo)
    rescisao  = get_dados_planilha_rescisao_por_loja(planilha_rescisoes, mes, ano)
    vt        = get_dados_planilha_VT_por_loja(planilha_VT, mes, ano)
    ferias    = get_dados_planilha_ferias_por_loja(df_ferias, mes, ano)
    uniformes = get_dados_gastos_almoxarifado_por_loja(df_uniforme, mes, ano, 'valor_uniforme')
    materiais = get_dados_gastos_almoxarifado_por_loja(df_materiais, mes, ano, 'valor_materiais')
    
    for df in [custos, rescisao, vt, ferias, uniformes, materiais]:
        col = 'RATEIO' if 'RATEIO' in df.columns else 'LOJA'
        if col in df.columns and not df.empty:
            df[col] = _normalizar_coluna_loja(df[col])

    # LEFT merge: ancora no df de custos e remove a coluna 'LOJA' para evitar colisão de sufixos
    df = pd.merge(custos, rescisao, left_on='RATEIO', right_on='LOJA', how='left')
    if 'LOJA' in df.columns: df.drop(columns=['LOJA'], inplace=True)

    df = pd.merge(df, vt, left_on='RATEIO', right_on='LOJA', how='left')
    if 'LOJA' in df.columns: df.drop(columns=['LOJA'], inplace=True)

    df = pd.merge(df, ferias, left_on='RATEIO', right_on='LOJA', how='left')
    if 'LOJA' in df.columns: df.drop(columns=['LOJA'], inplace=True)
    
    if not uniformes.empty:
        df = pd.merge(df, uniformes, left_on='RATEIO', right_on='LOJA', how='left')
        if 'LOJA' in df.columns: df.drop(columns=['LOJA'], inplace=True)
        
    if not materiais.empty:
        df = pd.merge(df, materiais, left_on='RATEIO', right_on='LOJA', how='left')
        if 'LOJA' in df.columns: df.drop(columns=['LOJA'], inplace=True)

    # Garante estrutura caso DataFrames venham vazios
    if 'valor_uniforme' not in df.columns:
        df['valor_uniforme'] = 0.0
    if 'valor_materiais' not in df.columns:
        df['valor_materiais'] = 0.0

    df.fillna(0, inplace=True)

    return df