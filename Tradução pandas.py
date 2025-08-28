import pandas as pd
from deep_translator import GoogleTranslator
import os

def traduzir_workbook_inteiro(caminho_arquivo_entrada, caminho_arquivo_saida, idioma_origem='pt', idioma_destino='en'):
    try:
        with pd.ExcelWriter(caminho_arquivo_saida, engine='openpyxl') as writer:
            xls = pd.ExcelFile(caminho_arquivo_entrada)
            
            print(f"\nArquivo '{os.path.basename(caminho_arquivo_entrada)}' carregado. Abas encontradas: {xls.sheet_names}")

            for nome_aba in xls.sheet_names:
                print(f"\n--- Processando a aba: '{nome_aba}' ---")
                df = pd.read_excel(xls, sheet_name=nome_aba)

                df.columns = df.columns.str.strip()
                print(f"Nomes das colunas da aba '{nome_aba}' (após limpeza):")
                print(list(df.columns))
                print("-" * 40)

                def traduzir_texto(texto):
                    if pd.isna(texto) or str(texto).strip() == '':
                        return ''
                    
                    texto_str = str(texto)
                    if texto_str.isnumeric():
                        return texto_str
                    try:
                        return GoogleTranslator(source=idioma_origem, target=idioma_destino).translate(texto_str)
                    except Exception:
                        return texto_str

                for nome_coluna in df.columns:
                    if df[nome_coluna].dtype == 'object':
                        print(f"Traduzindo coluna: '{nome_coluna}'...")
                        df[nome_coluna] = df[nome_coluna].apply(traduzir_texto)
                
                df.to_excel(writer, sheet_name=nome_aba, index=False)
                print(f"Aba '{nome_aba}' traduzida e adicionada ao arquivo de saída.")

        print(f"\nTradução completa. Arquivo salvo como: '{os.path.basename(caminho_arquivo_saida)}'")

    except Exception as e:
        print(f"Ocorreu um erro ao processar o arquivo {os.path.basename(caminho_arquivo_entrada)}: {e}")

if __name__ == "__main__":
    pasta_de_trabalho = r'C:\Users\User\Desktop\Nova pasta'
    
    print(f"Procurando por arquivos .xlsx na pasta: '{pasta_de_trabalho}'")

    if not os.path.isdir(pasta_de_trabalho):
        print(f"ERRO: A pasta '{pasta_de_trabalho}' não foi encontrada.")
    else:
        arquivos_na_pasta = os.listdir(pasta_de_trabalho)
        
        arquivos_para_traduzir = [f for f in arquivos_na_pasta if f.endswith('.xlsx') and not f.endswith('_TRADUZIDO.xlsx')]

        if not arquivos_para_traduzir:
            print("Nenhum arquivo novo para traduzir foi encontrado.")
        else:
            print(f"Arquivos encontrados para tradução: {arquivos_para_traduzir}")
            for nome_arquivo in arquivos_para_traduzir:
                caminho_entrada = os.path.join(pasta_de_trabalho, nome_arquivo)
                
                nome_base, extensao = os.path.splitext(nome_arquivo)
                nome_saida = f"{nome_base}_TRADUZIDO{extensao}"
                caminho_saida = os.path.join(pasta_de_trabalho, nome_saida)
                
                traduzir_workbook_inteiro(caminho_entrada, caminho_saida)

    print("\nProcesso geral finalizado.")