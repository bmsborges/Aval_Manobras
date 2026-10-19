import io
import openpyxl
from openpyxl.cell.cell import MergedCell
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


def definir_valor_celula_seguro(sheet, row, col, value):
    """Escreve o valor na célula de forma segura, tratando células mescladas (MergedCells)."""
    cell = sheet.cell(row=row, column=col)

    if isinstance(cell, MergedCell):
        # Se a célula for mesclada, localiza a célula principal (topo esquerdo)
        for rng in sheet.merged_cells.ranges:
            if cell.coordinate in rng:
                top_left_cell = sheet.cell(row=rng.min_row, column=rng.min_col)
                top_left_cell.value = value
                return
    else:
        cell.value = value


def extrair_nomes_formandos(file_formandos, coluna_letra="K"):
    """Extrai os nomes dos formandos a partir da coluna especificada (ex: 'K')."""
    col_idx = column_index_from_string(coluna_letra.upper()) - 1

    if file_formandos.name.endswith(".csv"):
        df = pd.read_csv(file_formandos, header=None)
    else:
        df = pd.read_excel(file_formandos, header=None)

    if df.shape[1] <= col_idx:
        raise ValueError(
            f"O ficheiro enviado possui apenas {df.shape[1]} colunas, mas a Coluna {coluna_letra.upper()} (coluna {col_idx + 1}) foi selecionada."
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
        if not val_str:
            continue

        if val_str.lower() in palavras_ignorar or any(
            val_str.lower().startswith(p) for p in palavras_ignorar
        ):
            continue

        nomes.append(val_str)

    return nomes


def importar_formandos_stream(
    file_avaliacao, file_formandos, coluna_letra="K", aba_destino="Grelha Avaliacao"
):
    """Lê os ficheiros em memória e preenche a grelha de avaliação Excel tratando células mescladas."""
    nomes_formandos = extrair_nomes_formandos(file_formandos, coluna_letra)

    if not nomes_formandos:
        raise ValueError(
            f"Nenhum nome válido foi encontrado na Coluna {coluna_letra.upper()} do ficheiro de formandos."
        )

    wb = openpyxl.load_workbook(file_avaliacao)

    if aba_destino in wb.sheetnames:
        sheet = wb[aba_destino]
    else:
        sheet = wb.active

    # Posições na pauta de avaliação
    linha_inicial = 14  # Linha onde começam os formandos na pauta
    coluna_num = 2  # Coluna B (Nº do Formando)
    coluna_nome = 11  # Coluna K (Nome na pauta de avaliação)

    # Inserir os nomes e números na pauta usando a função segura
    for idx, nome in enumerate(nomes_formandos):
        linha_atual = linha_inicial + idx

        # Número do formando (Coluna B)
        definir_valor_celula_seguro(
            sheet, row=linha_atual, col=coluna_num, value=idx + 1
        )

        # Nome do formando (Coluna K)
        definir_valor_celula_seguro(
            sheet, row=linha_atual, col=coluna_nome, value=nome
        )

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
                f"✅ Importação concluída! {num_total} formandos foram inseridos na pauta com sucesso."
            )

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
