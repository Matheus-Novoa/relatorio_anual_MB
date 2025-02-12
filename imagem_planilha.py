import pandas as pd
import dataframe_image as dfi
from pathlib import Path



input_dir = Path("extracted_tables")  # Diretório com arquivos Excel"
output_dir = Path("imagens") # Cria o diretório imagens se ele não existir
output_dir.mkdir(exist_ok=True)

for excel_file in input_dir.glob("*.xlsx"):
    # Import just one example excel file from extracted_tables folder. Import the notas sheet
    df = pd.read_excel(
        excel_file,
        sheet_name="notas",
        converters={
            # Garantir que "Nota" seja exibida como string inteira (sem .0)
            'Nota': lambda x: f"{int(x):d}" if pd.notnull(x) else ''
        },
        dtype={'Data': str, 'Cliente': str, 'CPF': str}  # Garantir tipos como string
    )

    # Preencher TODOS os NaNs com string vazia (evitar "nan" na imagem)
    df = df.fillna('')

    # Formatar "Valor_Contabil" como R$ X,XX (vírgula decimal)
    df['Valor Contábil'] = df['Valor Contábil'].apply(lambda x: f"R$ {x:,.2f}".replace(".", "|").replace(",", ".").replace("|", ","))

    # Remover índices do DataFrame
    styled_df = df.style.hide(axis='index')

    # Aplicar bordas e estilos
    styled_df = styled_df.set_properties(**{
        'border': '1px solid black',
        'text-align': 'center'
    }).set_table_styles([
        # Estilo do cabeçalho
        {'selector': 'th', 'props': [('border', '1px solid black')]},
        
        # Remover bordas internas da última linha (exceto bordas externas)
        {
            'selector': 'tr:last-child td',
            'props': [
                ('border-top', '1px solid black'),
                ('border-bottom', '1px solid black'),
                ('border-left', 'none'),
                ('border-right', 'none')
            ]
        },
        # Borda esquerda na primeira célula da última linha
        {
            'selector': 'tr:last-child td:first-child',
            'props': [('border-left', '1px solid black')]
        },
        # Borda direita na última célula da última linha
        {
            'selector': 'tr:last-child td:last-child',
            'props': [('border-right', '1px solid black')]
        }
    ])

    nome_arquivo_imagem = f"{output_dir}/{excel_file.stem}.png"

    dfi.export(styled_df, nome_arquivo_imagem, dpi=300)
    print(f"Imagem salva em: {nome_arquivo_imagem}")
