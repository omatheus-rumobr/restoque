import pandas as pd

def atualizar_estoque_tabela(arquivo_entrada, arquivo_saida):
    """
    Lê um arquivo Excel com duas páginas (Tabela e Estoque) e atualiza a coluna Estoque
    da página Tabela com os valores da coluna quantidade da página Estoque usando EAN como relacionamento.
    
    Args:
        arquivo_entrada (str): Caminho do arquivo Excel de entrada
        arquivo_saida (str): Caminho do arquivo Excel de saída
    """
    
    try:
        # Ler as duas páginas do arquivo Excel
        df_tabela = pd.read_excel(arquivo_entrada, sheet_name='Tabela')
        df_estoque = pd.read_excel(arquivo_entrada, sheet_name='Estoque')
        
        print("Dados da página Tabela:")
        print(df_tabela.head())
        print("\nDados da página Estoque:")
        print(df_estoque.head())
        
        # Verificar se as colunas necessárias existem
        colunas_necessarias_tabela = ['EAN', 'Estoque']
        colunas_necessarias_estoque = ['EAN', 'Estoque Disponivel']
        
        for coluna in colunas_necessarias_tabela:
            if coluna not in df_tabela.columns:
                raise ValueError(f"Coluna '{coluna}' não encontrada na página Tabela")
        
        for coluna in colunas_necessarias_estoque:
            if coluna not in df_estoque.columns:
                raise ValueError(f"Coluna '{coluna}' não encontrada na página Estoque")
        
        # Criar um dicionário de EAN para quantidade do estoque
        estoque_dict = df_estoque.set_index('EAN')['Estoque Disponivel'].to_dict()
        
        # Atualizar a coluna Estoque na tabela
        df_tabela['Estoque'] = df_tabela['EAN'].map(estoque_dict).fillna(0).astype(int)
        
        print("\nTabela atualizada:")
        print(df_tabela.head())
        
        # Salvar o resultado em um novo arquivo Excel
        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            df_tabela.to_excel(writer, sheet_name='Tabela', index=False)
            df_estoque.to_excel(writer, sheet_name='Estoque', index=False)
        
        print(f"\nArquivo salvo com sucesso: {arquivo_saida}")
        
        # Mostrar estatísticas
        eans_encontrados = (df_tabela['Estoque'] > 0).sum()
        eans_nao_encontrados = (df_tabela['Estoque'] == 0).sum()
        
        print(f"\nEstatísticas:")
        print(f"EANs com estoque encontrado: {eans_encontrados}")
        print(f"EANs sem estoque (valor 0): {eans_nao_encontrados}")
        print(f"Total de itens na Tabela: {len(df_tabela)}")
        
    except FileNotFoundError:
        print(f"Erro: Arquivo '{arquivo_entrada}' não encontrado.")
    except ValueError as e:
        print(f"Erro: {e}")
    except Exception as e:
        print(f"Erro inesperado: {e}")

if __name__ == "__main__":
    # Nome do arquivo de entrada e saída
    arquivo_entrada = "tabela.xlsx"
    arquivo_saida = "tabela_atualizada.xlsx"
    
    # Executar a função
    atualizar_estoque_tabela(arquivo_entrada, arquivo_saida)
