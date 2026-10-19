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
    "A aplicação gera as tabs individuais por formando e permite registar/consultar os parâmetros diretamente no formulário."
)


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


# --- INICIALIZAÇÃO DO SESSION STATE ---

if "lista_formandos" not in st.session_state:
    st.session_state["lista_formandos"] = []

if "avaliacoes" not in st.session_state:
    st.session_state["avaliacoes"] = {}


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

col_letra = st.sidebar.text_input(
    "Coluna dos Nomes no Ficheiro 2:", value="K", max_chars=2
).upper()

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

                # Preencher identificação na tab
                definir_valor_seguro(
                    aba_nova, row=1, col=7, value=f"FORMANDO: {nome}"
                )

            buffer_modelo = io.BytesIO()
            wb_modelo.save(buffer_modelo)
            buffer_modelo.seek(0)
            st.session_state["buffer_modelo_tabs"] = buffer_modelo

            st.sidebar.success(
                f"✅ {len(nomes)} tabs de formandos criadas com sucesso!"
            )
        except Exception as e:
            st.sidebar.error(f"Erro ao inicializar: {e}")


# --- FORMULÁRIO DINÂMICO DE REGISTO DE AVALIAÇÃO ---

st.subheader("📝 Formulário de Avaliação Prática por Formando")

if st.session_state["lista_formandos"]:
    formando_selecionado = st.selectbox(
        "Selecione o Formando:",
        options=st.session_state["lista_formandos"],
        key="select_formando_atual",
    )

    # Obter os dados guardados para este formando ou usar os padrões (3)
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

    # O formulário usa chaves dinâmicas com o nome do formando para atualizar ao mudar a seleção
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

        submetido = st.form_submit_button(
            "💾 Registar / Atualizar Avaliação do Formando"
        )

        if submetido:
            # Guardar/Atualizar avaliação no state
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
                f"✅ Avaliação do formando **{formando_selecionado}** registada com sucesso!"
            )

    # Botão de descarregar o ficheiro gerado com as tabs de cada formando
    if "buffer_modelo_tabs" in st.session_state:
        st.download_button(
            label="📥 Descarregar Ficheiro com Separadores por Formando (.xlsx)",
            data=st.session_state["buffer_modelo_tabs"],
            file_name="Grelha_Observacao_Pratica_Tabs.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info(
        "Carregue os ficheiros no menu lateral e clique em 'Inicializar e Criar Tabs por Formando'."
    )
