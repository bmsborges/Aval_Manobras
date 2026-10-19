import io
import zipfile
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Sistema de Avaliação de Formandos - ENB",
    page_icon="👨‍🚒",
    layout="wide",
)

st.title("👨‍🚒 Gestão, Registo e Exportação de Avaliações (.xlsx / .zip)")
st.write(
    "Registe as cotações no formulário e exporte os dois ficheiros convertidos e preenchidos estritamente no formato **.xlsx**."
)


# --- FUNÇÕES DE SUPORTE E CONVERSÃO EXCEL ---


def abrir_workbook_garantido_xlsx(file_obj):
    """Carrega o ficheiro uploaded e converte para um Workbook openpyxl (.xlsx) nativo."""
    file_obj.seek(0)
    nome_ficheiro = getattr(file_obj, "name", "").lower()

    if nome_ficheiro.endswith(".xls") and not nome_ficheiro.endswith(".xlsx"):
        # Se for um formato antigo .xls, converter via Pandas / xlrd
        dict_dfs = pd.read_excel(file_obj, sheet_name=None, header=None)
        wb = openpyxl.Workbook()
        # Remover aba padrão
        wb.remove(wb.active)

        for sheet_name, df in dict_dfs.items():
            ws = wb.create_sheet(title=str(sheet_name)[:31])
            for r_idx, row in df.iterrows():
                for c_idx, val in enumerate(row):
                    if pd.notna(val):
                        ws.cell(
                            row=r_idx + 1, column=c_idx + 1, value=val
                        )
        return wb
    else:
        # Ficheiro .xlsx nativo
        return openpyxl.load_workbook(file_obj)


def definir_valor_seguro(sheet, row, col, value):
    """Escreve o valor na célula tratando com segurança células mescladas (MergedCells)."""
    cell = sheet.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for rng in sheet.merged_cells.ranges:
            if cell.coordinate in rng:
                top_left_cell = sheet.cell(
                    row=rng.min_row, column=rng.min_col
                )
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
        raise ValueError(
            f"O ficheiro tem apenas {df.shape[1]} colunas. A coluna {coluna_letra.upper()} não existe."
        )

    coluna_dados = df.iloc[:, col_idx]
    palavras_ignorar = [
        "nome",
        "nome do formando",
        "formando",
        "nome completo",
        "formandos",
        "nome:",
        "cform",
        "tipo",
        "data",
    ]

    nomes = []
    for val in coluna_dados:
        if pd.isna(val) or val is None:
            continue
        val_str = str(val).strip()
        if not val_str or any(
            val_str.lower().startswith(p) for p in palavras_ignorar
        ):
            continue
        nomes.append(val_str)

    return nomes


def gerar_ficheiros_xlsx_bytes(
    file_modelo, file_pauta, lista_formandos, avaliacoes
):
    """Compila os dois ficheiros garantindo a exportação no formato .xlsx."""

    # 1. Grelha Prática com 1 Tab por Formando (.xlsx)
    wb_modelo = abrir_workbook_garantido_xlsx(file_modelo)
    aba_base = wb_modelo.active

    for idx, nome in enumerate(lista_formandos):
        nome_aba = f"{idx+1}_{nome[:22]}".translate(
            str.maketrans("", "", r"/*?:[]\\")
        )
        aba_nova = wb_modelo.copy_worksheet(aba_base)
        aba_nova.title = nome_aba

        # Identificação do formando na tab
        definir_valor_seguro(
            aba_nova, row=1, col=7, value=f"FORMANDO: {nome}"
        )

        # Inserir cotações registadas no formulário
        if nome in avaliacoes:
            dados = avaliacoes[nome]
            definir_valor_seguro(
                aba_nova, row=3, col=34, value=dados.get("of1_1", 3)
            )
            definir_valor_seguro(
                aba_nova, row=4, col=34, value=dados.get("of1_2", 3)
            )
            definir_valor_seguro(
                aba_nova, row=5, col=34, value=dados.get("of1_3", 3)
            )
            definir_valor_seguro(
                aba_nova, row=6, col=34, value=dados.get("of1_4", 3)
            )
            definir_valor_seguro(
                aba_nova, row=7, col=34, value=dados.get("of1_5", 3)
            )
            definir_valor_seguro(
                aba_nova, row=8, col=2, value=f"OBS: {dados.get('obs', '')}"
            )

    # Remover a aba template original
    if len(wb_modelo.sheetnames) > 1 and aba_base.title in wb_modelo.sheetnames:
        wb_modelo.remove(aba_base)

    bytes_modelo_xlsx = io.BytesIO()
    wb_modelo.save(bytes_modelo_xlsx)
    bytes_modelo_xlsx.seek(0)

    # 2. Pauta Final Consolidada (.xlsx)
    bytes_pauta_xlsx = None
    if file_pauta:
        wb_pauta = abrir_workbook_garantido_xlsx(file_pauta)
        sheet_pauta = wb_pauta.active

        linha_inicial = 14
        for idx, nome in enumerate(lista_formandos):
            linha_atual = linha_inicial + idx
            definir_valor_seguro(
                sheet_pauta, row=linha_atual, col=2, value=idx + 1
            )
            definir_valor_seguro(
                sheet_pauta, row=linha_atual, col=11, value=nome
            )

        bytes_pauta_xlsx = io.BytesIO()
        wb_pauta.save(bytes_pauta_xlsx)
        bytes_pauta_xlsx.seek(0)

    return bytes_modelo_xlsx, bytes_pauta_xlsx


def criar_zip_com_ficheiros_xlsx(bytes_modelo, bytes_pauta):
    """Empacota os ficheiros em formato .xlsx dentro de um arquivo .zip em memória."""
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(
        zip_buffer, "w", zipfile.ZIP_DEFLATED, False
    ) as zip_file:
        zip_file.writestr(
            "Grelha_Observacao_Pratica_Tabs.xlsx", bytes_modelo.getvalue()
        )

        if bytes_pauta:
            zip_file.writestr(
                "Pauta_Final_Avaliacao.xlsx", bytes_pauta.getvalue()
            )

    zip_buffer.seek(0)
    return zip_buffer


# --- SESSION STATE ---

if "lista_formandos" not in st.session_state:
    st.session_state["lista_formandos"] = []

if "avaliacoes" not in st.session_state:
    st.session_state["avaliacoes"] = {}


# --- PAINEL LATERAL: CARREGAMENTO DOS FICHEIROS ---

st.sidebar.header("📁 Carregamento dos Ficheiros")

file_modelo = st.sidebar.file_uploader(
    "1. Grelha Modelo de Avaliação (.xls / .xlsx)",
    type=["xlsx", "xls"],
    key="modelo",
)

file_formandos = st.sidebar.file_uploader(
    "2. Ficheiro de Formandos DTP (.xlsx, .csv)",
    type=["xlsx", "xls", "csv"],
    key="formandos",
)

file_pauta = st.sidebar.file_uploader(
    "3. Pauta Final Consolidada (.xlsx) [Opcional]",
    type=["xlsx"],
    key="pauta",
)

col_letra = st.sidebar.text_input(
    "Coluna dos Nomes no Ficheiro 2:", value="K", max_chars=2
).upper()

if file_modelo and file_formandos:
    if st.sidebar.button("🔄 Carregar Formandos dos Ficheiros"):
        try:
            nomes = extrair_nomes_formandos(file_formandos, col_letra)
            st.session_state["lista_formandos"] = nomes
            st.sidebar.success(
                f"✅ {len(nomes)} formandos identificados com sucesso!"
            )
        except Exception as e:
            st.sidebar.error(f"Erro ao carregar formandos: {e}")


# --- FORMULÁRIO DE REGISTO DE AVALIAÇÃO ---

st.subheader("📝 Registar Avaliações no Formulário")

if st.session_state["lista_formandos"]:
    formando_selecionado = st.selectbox(
        "Selecione o Formando:",
        options=st.session_state["lista_formandos"],
        key="select_formando_atual",
    )

    dados = st.session_state["avaliacoes"].get(
        formando_selecionado,
        {
            "of1_1": 3,
            "of1_2": 3,
            "of1_3": 3,
            "of1_4": 3,
            "of1_5": 3,
            "of2_1": 3,
            "of2_2": 3,
            "of2_3": 3,
            "of2_4": 3,
            "of2_5": 3,
            "as_1": 3,
            "as_2": 3,
            "as_3": 3,
            "as_4": 3,
            "as_5": 3,
            "soc_1": 3,
            "soc_2": 3,
            "soc_3": 3,
            "soc_4": 3,
            "soc_5": 3,
            "obs": "",
        },
    )

    with st.form(key=f"form_avaliacao_{formando_selecionado}"):
        st.markdown(f"### Avaliação de: **{formando_selecionado}**")

        c1, c2 = st.columns(2)

        with c1:
            st.markdown("#### 1. Operador de Ferramentas 1")
            v_of1_1 = st.slider(
                "Estabiliza o(s) veículo(s)",
                0,
                5,
                int(dados["of1_1"]),
                key=f"of1_1_{formando_selecionado}",
            )
            v_of1_2 = st.slider(
                "Procede à abertura de acessos",
                0,
                5,
                int(dados["of1_2"]),
                key=f"of1_2_{formando_selecionado}",
            )
            v_of1_3 = st.slider(
                "Opera corretamente as ferramentas",
                0,
                5,
                int(dados["of1_3"]),
                key=f"of1_3_{formando_selecionado}",
            )
            v_of1_4 = st.slider(
                "Criação de espaço (técnicas)",
                0,
                5,
                int(dados["of1_4"]),
                key=f"of1_4_{formando_selecionado}",
            )
            v_of1_5 = st.slider(
                "Protege vítima(s) e socorrista(s)",
                0,
                5,
                int(dados["of1_5"]),
                key=f"of1_5_{formando_selecionado}",
            )

            st.markdown("#### 2. Operador de Ferramentas 2")
            v_of2_1 = st.slider(
                "Estabiliza o(s) veículo(s) (OF2)",
                0,
                5,
                int(dados["of2_1"]),
                key=f"of2_1_{formando_selecionado}",
            )
            v_of2_2 = st.slider(
                "Abertura de acessos (OF2)",
                0,
                5,
                int(dados["of2_2"]),
                key=f"of2_2_{formando_selecionado}",
            )
            v_of2_3 = st.slider(
                "Operação de ferramentas (OF2)",
                0,
                5,
                int(dados["of2_3"]),
                key=f"of2_3_{formando_selecionado}",
            )
            v_of2_4 = st.slider(
                "Criação de espaço (OF2)",
                0,
                5,
                int(dados["of2_4"]),
                key=f"of2_4_{formando_selecionado}",
            )
            v_of2_5 = st.slider(
                "Proteção vítima/socorrista (OF2)",
                0,
                5,
                int(dados["of2_5"]),
                key=f"of2_5_{formando_selecionado}",
            )

        with c2:
            st.markdown("#### 3. Assistente de Segurança")
            v_as_1 = st.slider(
                "Reconhecimento e avaliação de riscos",
                0,
                5,
                int(dados["as_1"]),
                key=f"as_1_{formando_selecionado}",
            )
            v_as_2 = st.slider(
                "Sinalização e delimitação de zonas",
                0,
                5,
                int(dados["as_2"]),
                key=f"as_2_{formando_selecionado}",
            )
            v_as_3 = st.slider(
                "Estabilização inicial e progressiva",
                0,
                5,
                int(dados["as_3"]),
                key=f"as_3_{formando_selecionado}",
            )
            v_as_4 = st.slider(
                "Limpeza da zona de trabalho",
                0,
                5,
                int(dados["as_4"]),
                key=f"as_4_{formando_selecionado}",
            )
            v_as_5 = st.slider(
                "Proteção nos pontos agressivos",
                0,
                5,
                int(dados["as_5"]),
                key=f"as_5_{formando_selecionado}",
            )

            st.markdown("#### 4. Socorrista")
            v_soc_1 = st.slider(
                "Abordagem da(s) vítima(s)",
                0,
                5,
                int(dados["soc_1"]),
                key=f"soc_1_{formando_selecionado}",
            )
            v_soc_2 = st.slider(
                "Exame primário da(s) vítima(s)",
                0,
                5,
                int(dados["soc_2"]),
                key=f"soc_2_{formando_selecionado}",
            )
            v_soc_3 = st.slider(
                "Definição do estado clínico",
                0,
                5,
                int(dados["soc_3"]),
                key=f"soc_3_{formando_selecionado}",
            )
            v_soc_4 = st.slider(
                "Cuidados de emergência",
                0,
                5,
                int(dados["soc_4"]),
                key=f"soc_4_{formando_selecionado}",
            )
            v_soc_5 = st.slider(
                "Coordenação da extração",
                0,
                5,
                int(dados["soc_5"]),
                key=f"soc_5_{formando_selecionado}",
            )

        st.markdown("#### 5. Observações Gerais")
        v_obs = st.text_area(
            "Observações da Avaliação:",
            value=dados["obs"],
            placeholder="Insira notas do formador...",
            key=f"obs_{formando_selecionado}",
        )

        submetido = st.form_submit_button("💾 Guardar Avaliação do Formando")

        if submetido:
            st.session_state["avaliacoes"][formando_selecionado] = {
                "of1_1": v_of1_1,
                "of1_2": v_of1_2,
                "of1_3": v_of1_3,
                "of1_4": v_of1_4,
                "of1_5": v_of1_5,
                "of2_1": v_of2_1,
                "of2_2": v_of2_2,
                "of2_3": v_of2_3,
                "of2_4": v_of2_4,
                "of2_5": v_of2_5,
                "as_1": v_as_1,
                "as_2": v_as_2,
                "as_3": v_as_3,
                "as_4": v_as_4,
                "as_5": v_as_5,
                "soc_1": v_soc_1,
                "soc_2": v_soc_2,
                "soc_3": v_soc_3,
                "soc_4": v_soc_4,
                "soc_5": v_soc_5,
                "obs": v_obs,
            }
            st.success(
                f"✅ Avaliação do formando **{formando_selecionado}** guardada na sessão!"
            )

    # --- SECÇÃO DE EXPORTAÇÃO COMPACTADA EM ZIP COM FICHEIROS .XLSX ---

    st.markdown("---")
    st.subheader("📦 Exportação de Ficheiros (.xlsx / .zip)")

    if st.button("🚀 Exportar Ficheiros em .xlsx (ZIP)", type="primary"):
        try:
            with st.spinner("A converter e compilar os ficheiros .xlsx..."):
                bytes_mod_xlsx, bytes_pauta_xlsx = (
                    gerar_ficheiros_xlsx_bytes(
                        file_modelo,
                        file_pauta,
                        st.session_state["lista_formandos"],
                        st.session_state["avaliacoes"],
                    )
                )
                buffer_zip = criar_zip_com_ficheiros_xlsx(
                    bytes_mod_xlsx, bytes_pauta_xlsx
                )

            st.success("✅ Ficheiros .xlsx compilados e prontos em arquivo ZIP!")

            st.download_button(
                label="📦 Descarregar Avaliacoes_XLSX.zip",
                data=buffer_zip,
                file_name="Avaliacoes_XLSX.zip",
                mime="application/zip",
            )
        except Exception as e:
            st.error(f"Erro ao gerar a exportação: {e}")
else:
    st.info(
        "Carregue os ficheiros no painel lateral e clique em 'Carregar Formandos dos Ficheiros'."
    )
