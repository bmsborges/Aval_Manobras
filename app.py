import io
import os
from pathlib import Path
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Importador de Formandos - ENB",
    page_icon="📋",
    layout="centered",
)

st.title("📋 Importador de Formandos para Grelha de Avaliação")
st.write(
    "Carregue a pauta/grelha de avaliação e a lista de formandos (Excel/CSV) para preencher automaticamente os nomes."
)


def importar_formandos_stream(file_avaliacao, file_formandos, aba_destino="Grelha Avaliacao"):
    """Lê os ficheiros enviados via Streamlit em memória e devolve um buffer

    BytesIO do ficheiro preenchido.
    """
    # 1. Ler a lista de formandos
    if file_formandos.name.endswith(".csv"):
        df_formandos = pd.read_csv(file_formandos)
    else:
        df_formandos = pd.read_excel(file_formandos)

    # Identificar a coluna do Nome
    col_nome = next(
        (c for c in df_formandos.columns if "nome" in str(c).lower()), None
    )

    if not col_nome:
        raise ValueError(
            "Não foi encontrada nenhuma coluna com a palavra 'Nome' no ficheiro de formandos."
        )

    # 2. Carregar o modelo Excel preservando formatação e fórmulas
    wb = openpyxl.load_workbook(file_avaliacao)

    if aba_destino in wb.sheetnames:
        sheet = wb[aba_destino]
    else:
        sheet = wb.active

    # 3. Posição inicial na grelha de avaliação
    linha_inicial = 14  # Linha onde começam os formandos
    coluna_num = 2  # Coluna B (Nº)
    coluna_nome = 11  # Coluna K (Nome)

    # 4. Inserir os formandos na grelha
    for idx, row in df_formandos.iterrows():
        linha_atual = linha_inicial + idx
        nome = str(row[col_nome]).strip()

        sheet.cell(row=linha_atual, column=coluna_num, value=idx + 1)
        sheet.cell(row=linha_atual, column=coluna_nome, value=nome)

    # 5. Guardar num buffer de memória
    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)

    return output_buffer, len(df_formandos)


# --- INTERFACE STREAMLIT ---

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Ficheiro de Avaliação")
    file_eval = st.file_uploader(
        "Carregue a Grelha/Ficha (.xlsx)",
        type=["xlsx"],
        key="eval",
    )

with col2:
    st.subheader("2. Lista de Formandos")
    file_list = st.file_uploader(
        "Carregue os Formandos (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv"],
        key="list",
    )

if file_eval and file_list:
    st.success("Ficheiros carregados com sucesso!")

    if st.button("🚀 Importar Formandos", type="primary"):
        try:
            with st.spinner("A processar a importação..."):
                excel_gerado, num_total = importar_formandos_stream(
                    file_eval, file_list
                )

            st.success(
                f"✅ Foram importados {num_total} formandos com sucesso!"
            )

            # Botão de Download do ficheiro final
            st.download_button(
                label="📥 Descarregar Ficheiro de Avaliação Atualizado",
                data=excel_gerado,
                file_name="Avaliacao_Formandos_Preenchido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Ocorreu um erro no processamento: {e}")
else:
    st.info(
        "Por favor, carregue ambos os ficheiros acima para ativar a importação."
    )
