import os
import openpyxl
import pandas as pd


def importar_formandos(
    ficheiro_avaliacao_path, ficheiro_formandos_path, aba_destino="Grelha Avaliacao"
):
    """Importa a lista de formandos para o ficheiro de avaliação Excel.

    :param ficheiro_avaliacao_path: Caminho do ficheiro 'Avaliação
        Formandos.xlsx'
    :param ficheiro_formandos_path: Caminho do ficheiro com a lista de formandos
        (.xlsx ou .csv)
    :param aba_destino: Nome da folha/aba no ficheiro de avaliação
    """
    if not os.path.exists(ficheiro_avaliacao_path):
        raise FileNotFoundError(
            f"Ficheiro de avaliação não encontrado: {ficheiro_avaliacao_path}"
        )

    # 1. Carregar lista de formandos do ficheiro fonte
    if ficheiro_formandos_path.endswith(".csv"):
        df_formandos = pd.read_csv(ficheiro_formandos_path)
    else:
        df_formandos = pd.read_excel(ficheiro_formandos_path)

    print(f"Formandos encontrados no ficheiro de origem: {len(df_formandos)}")

    # Identificar colunas relevantes (ajustar conforme o cabeçalho do ficheiro de origem)
    col_nome = next(
        (c for c in df_formandos.columns if "nome" in str(c).lower()), None
    )
    col_nif = next(
        (
            c
            for c in df_formandos.columns
            if "nif" in str(c).lower() or "doc" in str(c).lower()
        ),
        None,
    )

    if not col_nome:
        raise ValueError(
            "Não foi possível identificar a coluna de 'Nome' no ficheiro de formandos."
        )

    # 2. Abrir o ficheiro Excel de avaliação preservando fórmulas
    wb = openpyxl.load_workbook(ficheiro_avaliacao_path)

    if aba_destino not in wb.sheetnames:
        aba_destino = wb.sheetnames[0]

    sheet = wb[aba_destino]

    # 3. Localizar a linha de início dos formandos (Linha 14 em diante)
    linha_inicial = 14  # Linha standard da grelha de avaliação
    coluna_num = 2  # Coluna B (Nº)
    coluna_nome = 11  # Coluna K (Nome)

    # 4. Preencher a grelha com os formandos
    for idx, row in df_formandos.iterrows():
        linha_atual = linha_inicial + idx

        num_formando = idx + 1
        nome_formando = str(row[col_nome]).strip()

        # Inserir Número e Nome
        sheet.cell(row=linha_atual, column=coluna_num, value=num_formando)
        sheet.cell(row=linha_atual, column=coluna_nome, value=nome_formando)

        print(f"[{num_formando:02d}] Importado: {nome_formando}")

    # 5. Guardar o ficheiro atualizado
    nome_saida = ficheiro_avaliacao_path.replace(".xlsx", "_Atualizado.xlsx")
    wb.save(nome_saida)
    print(
        f"\nImportação concluída com sucesso! Ficheiro gerado: {nome_saida}"
    )


# --- EXEMPLO DE EXECUÇÃO ---
if __name__ == "__main__":
    # Substituir pelos nomes reais dos ficheiros
    FICHEIRO_AVALIACAO = "Avaliação Formandos.xlsx"
    FICHEIRO_LISTA_FORMANDOS = "Lista_Formandos_DTP.xlsx"

    # Criar ficheiro de teste para demonstração se não existir a lista
    if not os.path.exists(FICHEIRO_LISTA_FORMANDOS):
        dados_exemplo = pd.DataFrame({
            "Nº": range(1, 6),
            "Nome": [
                "Angelique Sofia Pais Freitas",
                "António Fernando da Silva Rodrigues",
                "Camila Oliveira Gonçalves",
                "Cláudia Marisa Torres Silva",
                "Fábio José Oliveira Matos",
            ],
            "NIF": [
                "232790370",
                "238595374",
                "252913710",
                "222375167",
                "214940004",
            ],
        })
        dados_exemplo.to_excel(FICHEIRO_LISTA_FORMANDOS, index=False)

    importar_formandos(FICHEIRO_AVALIACAO, FICHEIRO_LISTA_FORMANDOS)
