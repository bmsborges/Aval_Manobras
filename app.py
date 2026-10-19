import io
import openpyxl
from openpyxl.utils import column_index_from_string
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Importador de Formandos - ENB",
    page_icon="📋",
    layout="centered",
)

st.title("📋 Importador de Formandos para Grelha de Avaliação")
st.write(
    "Carregue a pauta/grelha de avaliação e o ficheiro com os dados dos formandos para preencher automaticamente os nomes."
)


def extrair_nomes_formandos(file_formandos, coluna_letra="K"):
    """Extrai os nomes dos formandos a partir da coluna especificada (ex: 'K')."""
    # Converter a letra da coluna (ex: 'K') para índice 0-based (A=0, B=1, ..., K=10)
    col_idx = column_index_from_string(coluna_letra.upper()) - 1

    # Ler o ficheiro sem assumir que a linha 0 é o cabeçalho
    if file_formandos.name.endswith(".csv"):
        df = pd.read_csv(file_formandos, header=None)
    else:
        df = pd.read_excel(file_formandos, header=None)

    # Verificar se a coluna existe no ficheiro
    if df.shape[1] <= col_idx:
        raise ValueError(
            f"O ficheiro enviado possui apenas {df.shape[1]} colunas, mas a Coluna {coluna_letra.upper()} (coluna {col_idx + 1}) foi selecionada."
        )

    coluna_dados = df.iloc[:, col_idx]

    # Palavras a ignorar (cabeçalhos e metadados comuns)
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
        if not val_str:
            continue

        # Filtrar cabeçalhos e metadados
        if val_str.lower() in palavras_ignorar or any(
            val_str.lower().startswith(p) for p in palavras_ignorar
        ):
            continue

        nomes.append(val_str)

    return nomes


def importar_formandos_stream(
    file_avaliacao, file_formandos, coluna_letra="K", aba_destino="Grelha Avaliacao"
):
    """Lê os ficheiros em memória e preenche a grelha de avaliação Excel."""
    # 1. Extrair a lista de nomes a partir da coluna K (ou selecionada)
    nomes_formandos = extrair_nomes_formandos(file_formandos, coluna_letra)

    if not nomes_formandos:
        raise ValueError(
            f"Nenhum nome válido foi encontrado na Coluna {coluna_letra.upper()} do ficheiro de formandos."
        )

    # 2. Carregar o ficheiro de avaliação preservando formatações e fórmulas
    wb = openpyxl.load_workbook(file_avaliacao)

    if aba_destino in wb.sheetnames:
        sheet = wb[aba_destino]
    else:
        sheet = wb.active

    # 3. Mapeamento de posições na pauta de avaliação
    linha_inicial = 14  # Linha onde começam os formandos na pauta
    coluna_num = 2  # Coluna B (Nº do Formando)
    coluna_nome = 11  # Coluna K (Nome na pauta de avaliação)

    # 4. Inserir os nomes na pauta
    for idx, nome in enumerate(nomes_formandos):
        linha_atual = linha_inicial + idx
        sheet.cell(row=linha_atual, column=coluna_num, value=idx + 1)
        sheet.cell(row=linha_atual, column=coluna_nome, value=nome)

    # 5. Guardar num buffer de memória para download
    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)

    return output_buffer, len(nomes_formandos)


# --- INTERFACE STREAMLIT ---

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Ficheiro de Avaliação")
    file_eval = st.file_uploader(
        "Carregue a Grelha/Pauta de Avaliação (.xlsx)",
        type=["xlsx"],
        key="eval",
    )

with col2:
    st.subheader("2. Ficheiro de Formandos")
    file_list = st.file_uploader(
        "Carregue a Lista de Formandos (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv"],
        key="list",
    )

# Opção de seleção da coluna onde estão os nomes
col_opcao = st.text_input(
    "Letra da Coluna que contém os nomes dos formandos:",
    value="K",
    max_chars=2,
    help="Por padrão é a Coluna K. Pode alterar se estiver numa coluna diferente.",
).upper()

if file_eval and file_list:
    st.success("Ficheiros prontos para processamento!")

    if st.button("🚀 Importar Formandos", type="primary"):
        try:
            with st.spinner("A extrair dados e a preencher a grelha..."):
                excel_gerado, num_total = importar_formandos_stream(
                    file_eval, file_list, coluna_letra=col_opcao
                )

            st.success(
                f"✅ Importação concluída! {num_total} formandos foram inseridos na pauta."
            )

            # Botão de descarregar ficheiro final
            st.download_button(
                label="📥 Descarregar Pauta de Avaliação Preenchida",
                data=excel_gerado,
                file_name="Avaliacao_Formandos_Preenchido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Ocorreu um erro no processamento: {e}")
else:
    st.info(
        "Por favor, carregue ambos os ficheiros acima para realizar a importação."
    )
