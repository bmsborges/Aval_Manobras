import io
import copy
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Sistema de Avaliação de Formandos - ENB",
    page_icon="👨‍🚒",
    layout="wide"
)

st.title("👨‍🚒 Gestão, Registo e Exportação de Avaliações Práticas UFCD")
st.write("Registe os parâmetros no formulário e exporte os ficheiros finais atualizados com todas as cotações.")

# --- FUNÇÕES AUXILIARES DE EXCEL ---

def definir_valor_seguro(sheet, row, col, value):
    """Escreve na célula tratando com segurança células mescladas (MergedCells)."""
    cell = sheet.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for rng in sheet.merged_cells.ranges:
            if cell.coordinate in rng:
                top_left_cell = sheet.cell(row=rng.min_row, column=rng.min_col)
                top_left_cell.value = value
                return
    else:
        cell.value = value

def extrair_nomes_formandos(file_formandos, coluna_letra="K"):
    """Extrai os nomes dos formandos a partir do ficheiro de dados (ex: Coluna K)."""
    col_idx = column_index_from_string(coluna_letra.upper()) - 1

    if file_formandos.name.endswith(".csv"):
        df = pd.read_csv(file_formandos, header=None)
    else:
        df = pd.read_excel(file_formandos, header=None)

    if df.shape[1] <= col_idx:
        raise ValueError(f"O ficheiro tem apenas {df.shape[1]} colunas. A coluna {coluna_letra.upper()} não existe.")

    coluna_dados = df.iloc[:, col_idx]
    palavras_ignorar = ["nome", "nome do formando", "formando", "nome completo", "formandos", "nome:", "cform", "tipo", "data"]

    nomes = []
    for val in coluna_dados:
        if pd.isna(val) or val is None:
            continue
        val_str = str(val).strip()
        if not val_str or any(val_str.lower().startswith(p) for p in palavras_ignorar):
            continue
        nomes.append(val_str)

    return nomes

def exportar_ficheiros_finais(file_modelo, file_pauta, lista_formandos, avaliacoes):
    """Compila os dados guardados em session_state e exporta o ficheiro com as tabs preenchidas e a pauta consolidada."""
    
    # 1. Exportar Grelha de Observação com 1 Tab por Formando
    wb_modelo = openpyxl.load_workbook(file_modelo)
    aba_base = wb_modelo.active

    for idx, nome in enumerate(lista_formandos):
        nome_aba = f"{idx+1}_{nome[:22]}".translate(str.maketrans("", "", r"/*?:[]\\"))
        aba_nova = wb_modelo.copy_worksheet(aba_base)
        aba_nova.title = nome_aba

        # Identificação do formando na tab
        definir_valor_seguro(aba_nova, row=1, col=7, value=f"FORMANDO: {nome}")

        # Se existirem avaliações registadas para este formando, preencher os campos da tab
        if nome in avaliacoes:
            dados = avaliacoes[nome]
            # Mapeamento dos valores das avaliações na tab
            definir_valor_seguro(aba_nova, row=3, col=34, value=dados.get("of1_1", 3))
            definir_valor_seguro(aba_nova, row=4, col=34, value=dados.get("of1_2", 3))
            definir_valor_seguro(aba_nova, row=5, col=34, value=dados.get("of1_3", 3))
            definir_valor_seguro(aba_nova, row=6, col=34, value=dados.get("of1_4", 3))
            definir_valor_seguro(aba_nova, row=7, col=34, value=dados.get("of1_5", 3))
            
            # Observações
            definir_valor_seguro(aba_nova, row=8, col=2, value=f"OBS: {dados.get('obs', '')}")

    # Remover a aba base/template original se foram criadas novas tabs
    if len(wb_modelo.sheetnames) > 1 and aba_base.title in wb_modelo.sheetnames:
        wb_modelo.remove(aba_base)

    out_modelo = io.BytesIO()
    wb_modelo.save(out_modelo)
    out_modelo.seek(0)

    # 2. Exportar Pauta Consolidada (se tiver sido carregada)
    out_pauta = None
    if file_pauta:
        file_pauta.seek(0)
        wb_pauta = openpyxl.load_workbook(file_pauta)
        sheet_pauta = wb_pauta.active

        linha_inicial = 14
        for idx, nome in enumerate(lista_formandos):
            linha_atual = linha_inicial + idx
            definir_valor_seguro(sheet_pauta, row=linha_atual, col=2, value=idx + 1)
            definir_valor_seguro(sheet_pauta, row=linha_atual, col=11, value=nome)

        out_pauta = io.BytesIO()
        wb_pauta.save(out_pauta)
        out_pauta.seek(0)

    return out_modelo, out_pauta

# --- SESSION STATE ---

if "lista_formandos" not in st.session_state:
    st.session_state["lista_formandos"] = []

if "avaliacoes" not in st.session_state:
    st.session_state["avaliacoes"] = {}

# --- PAINEL LATERAL: CARREGAMENTO E CONFIGURAÇÃO ---

st.sidebar.header("📁 Carregamento dos Ficheiros")

file_modelo = st.sidebar.file_uploader("1. Grelha Modelo de Avaliação (.xls / .xlsx)", type=["xlsx", "xls"], key="modelo")
file_formandos = st.sidebar.file_uploader("2. Ficheiro de Formandos DTP (.xlsx, .csv)", type=["xlsx", "xls", "csv"], key="formandos")
file_pauta = st.sidebar.file_uploader("3. Pauta Final Consolidada (.xlsx) [Opcional]", type=["xlsx"], key="pauta")

col_letra = st.sidebar.text_input("Coluna dos Nomes no Ficheiro 2:", value="K", max_chars=2).upper()

if file_modelo and file_formandos:
    if st.sidebar.button("🔄 Carregar Formandos dos Ficheiros"):
        try:
            nomes = extrair_nomes_formandos(file_formandos, col_letra)
            st.session_state["lista_formandos"] = nomes
            st.sidebar.success(f"✅ {len(nomes)} formandos identificados com sucesso!")
        except Exception as e:
            st.sidebar.error(f"Erro ao carregar formandos: {e}")

# --- FORMULÁRIO DE REGISTO DE AVALIAÇÃO ---

st.subheader("📝 Registar Avaliações no Formulário")

if st.session_state["lista_formandos"]:
    formando_selecionado = st.selectbox(
        "Selecione o Formando:",
        options=st.session_state["lista_formandos"],
        key="select_formando_atual"
    )

    dados = st.session_state["avaliacoes"].get(formando_selecionado, {
        "of1_1": 3, "of1_2": 3, "of1_3": 3, "of1_4": 3, "of1_5": 3,
