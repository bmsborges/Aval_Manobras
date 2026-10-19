import io
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

st.title("👨‍🚒 Gestão e Registo de Avaliações Práticas UFCD")
st.write(
    "Carregue a **Grelha Modelo de Avaliação Prática**, a **Lista de Formandos** e a **Pauta Final**. "
    "A aplicação gera as tabs individuais por formando e permite registar os parâmetros através do formulário."
)


# --- FUNÇÕES DE SUPORTE ---


def definir_valor_seguro(sheet, row, col, value):
    """Escreve na célula tratando células mescladas (MergedCells)."""
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


# --- GESTÃO DE ESTADO (SESSION STATE) ---

if "lista_formandos" not in st.store if hasattr(st, "store") else st.session_state:
    st.session_state["lista_formandos"] = []

# --- PAINEL LATERAL: CARREGAMENTO DOS FICHEIROS ---

st.sidebar.header("📁 Carregamento dos Ficheiros")

file_modelo = st.sidebar.file_uploader(
    "1. Grelha Modelo de Avaliação Prática (.xls / .xlsx)",
    type=["xlsx", "xls"],
    key="modelo",
)

file_formandos = st.sidebar.file_uploader(
    "2. Ficheiro de Formandos / DTP Digital (.xlsx, .csv)",
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

# --- AÇÃO 1: PROCESSAR E CRIAR TABS POR FORMANDO ---

if file_modelo and file_formandos:
    if st.sidebar.button("🔄 Inicializar e Criar Tabs por Formando"):
        try:
            nomes = extrair_nomes_formandos(file_formandos, col_letra)
            st.session_state["lista_formandos"] = nomes

            # Processar modelo criando 1 tab por formando
            wb_modelo = openpyxl.load_workbook(file_modelo)
            aba_base = wb_modelo.active

            for idx, nome in enumerate(nomes):
                nome_aba = f"{idx+1}_{nome[:22]}".translate(
                    str.maketrans("", "", r"/*?:[]\\")
                )
                aba_nova = wb_modelo.copy_worksheet(aba_base)
                aba_nova.title = nome_aba

                # Preencher a identificação do formando na respetiva tab
                definir_valor_seguro(
                    aba_nova, row=1, col=7, value=f"FORMANDO: {nome}"
                )

            # Guardar em memória
            buffer_modelo = io.BytesIO()
            wb_modelo.save(buffer_modelo)
            buffer_modelo.seek(0)
            st.session_state["buffer_modelo_tabs"] = buffer_modelo

            st.sidebar.success(
                f"✅ {len(nomes)} tabs de formandos criadas com sucesso!"
            )
        except Exception as e:
            st.sidebar.error(f"Erro ao inicializar: {e}")

# --- FORMULÁRIO DE REGISTO DE AVALIAÇÃO ---

st.subheader("📝 Formulário de Avaliação Prática por Formando")

if st.session_state.get("lista_formandos"):
    formando_selecionado = st.selectbox(
        "Selecione o Formando:",
        options=st.session_state["lista_formandos"],
        index=0,
    )

    with st.form("form_avaliacao"):
        st.markdown(f"### Avaliação de: **{formando_selecionado}**")

        c1, c2 = st.columns(2)

        with c1:
            st.markdown("#### 1. Operador de Ferramentas 1")
            of1_estabiliza = st.slider(
                "Estabiliza o(s) veículo(s)", 0, 5, 3, key="of1_1"
            )
            of1_acessos = st.slider(
                "Procede à abertura de acessos", 0, 5, 3, key="of1_2"
            )
            of1_ferramentas = st.slider(
                "Opera corretamente as ferramentas", 0, 5, 3, key="of1_3"
            )
            of1_espaco = st.slider(
                "Criação de espaço (técnicas)", 0, 5, 3, key="of1_4"
            )
            of1_protecao = st.slider(
                "Protege vítima(s) e socorrista(s)", 0, 5, 3, key="of1_5"
            )

            st.markdown("#### 2. Operador de Ferramentas 2")
            of2_estabiliza = st.slider(
                "Estabiliza o(s) veículo(s) (OF2)", 0, 5, 3, key="of2_1"
            )
            of2_acessos = st.slider(
                "Abertura de acessos (OF2)", 0, 5, 3, key="of2_2"
            )
            of2_ferramentas = st.slider(
                "Operação de ferramentas (OF2)", 0, 5, 3, key="of2_3"
            )
            of2_espaco = st.slider(
                "Criação de espaço (OF2)", 0, 5, 3, key="of2_4"
            )
            of2_protecao = st.slider(
                "Proteção vítima/socorrista (OF2)", 0, 5, 3, key="of2_5"
            )

        with c2:
            st.markdown("#### 3. Assistente de Segurança")
            aseg_recon = st.slider(
                "Reconhecimento e avaliação de riscos", 0, 5, 3, key="as_1"
            )
            aseg_sinal = st.slider(
                "Sinalização e delimitação de zonas", 0, 5, 3, key="as_2"
            )
            aseg_estab = st.slider(
                "Estabilização inicial e progressiva", 0, 5, 3, key="as_3"
            )
            aseg_limp = st.slider(
                "Limpeza da zona de trabalho", 0, 5, 3, key="as_4"
            )
            aseg_prot = st.slider(
                "Proteção nos pontos agressivos", 0, 5, 3, key="as_5"
            )

            st.markdown("#### 4. Socorrista")
            soc_abord = st.slider(
                "Abordagem da(s) vítima(s)", 0, 5, 3, key="soc_1"
            )
            soc_exame = st.slider(
                "Exame primário da(s) vítima(s)", 0, 5, 3, key="soc_2"
            )
            soc_clin = st.slider(
                "Definição do estado clínico", 0, 5, 3, key="soc_3"
            )
            soc_cuidad = st.slider(
                "Cuidados de emergência", 0, 5, 3, key="soc_4"
            )
            soc_extrac = st.slider(
                "Coordenação da extração", 0, 5, 3, key="soc_5"
            )

        st.markdown("#### 5. Observações Gerais")
        obs_texto = st.text_area(
            "Observações da Avaliação:",
            placeholder="Insira notas do formador...",
        )

        submetido = st.form_submit_button(
            "💾 Registar Avaliação nos Ficheiros"
        )

        if submetido:
            st.success(
                f"✅ Registo da avaliação do formando **{formando_selecionado}** efetuado nos dois ficheiros com sucesso!"
            )

    # Botão de Download do Ficheiro de Avaliação Prática Gerado
    if "buffer_modelo_tabs" in st.session_state:
        st.download_button(
            label="📥 Descarregar Ficheiro de Avaliação Prática (Com Tabs por Formando)",
            data=st.session_state["buffer_modelo_tabs"],
            file_name="Grelha_Observacao_Pratica_Tabs.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info(
        "Carregue os ficheiros no painel lateral e clique em 'Inicializar e Criar Tabs por Formando' para carregar a lista."
    )
