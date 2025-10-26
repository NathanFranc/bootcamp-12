import pandas as pd
import os
import duckdb


class LeitorExcel:
    def __init__(self, caminho_arquivo):
        self.caminho_arquivo = caminho_arquivo
        self.dados_completos = None
        
        
    def listar_arquivos_excel(self, pasta): 
        try:
            arquivos = [f for f in os.listdir(pasta) if f.endswith(('.xlsx', '.xls'))]
            print("Arquivos Excel encontrados: ")
            for i, arquivo in enumerate(arquivos, 1):
                print(f"{i}, {arquivo}")
            return arquivos 
        except FileNotFoundError:
             print(f"Pasta '{pasta}'  não encontrada! ")
             return []
     
     
    def encontrar_pasta_downloads(self):
        """Tenta encontrar a pasta Downloads automaticamente"""
        # Tentativas de caminhos comuns para a pasta Downloads
        caminhos_tentativas = [
            os.path.expanduser("~/Downloads"),  # Pasta Downloads do usuário
            "Downloads",
            "./Downloads"
        ]
        
        for caminho in caminhos_tentativas:
            if os.path.exists(caminho):
                print(f"✅ Pasta Downloads encontrada: {caminho}")
                return caminho
        print(f"Não foi possivel encontar a pasta Downloads automaticamente")
        return None
     
     
    def carregar_arquivo(self):  
        #Carrega o arquivo excel para memoria#
        try:
           self.dados_completos = pd.read_excel(self.caminho_arquivo)
           print("Arquivo carregado com sucesso!")
           return True
        except FileNotFoundError:
            print(f"Erro: Arquivo {self.caminho_arquivo} não encontrado!")
            return False
        except Exception as e:
            print(f"Erro ao carregar arquivo: {e}")
            return False
        
               
class AnalisadorDuckDB:
    def __init__(self):
        self.conexao = duckdb.connect()        
        
        print(" conexão DuckDB estabelecida!")
        
    def carregar_dataframe(self, dataframe, nome_tabela='dados_excel'):
        
        self.conexao.register(nome_tabela, dataframe)
        print(f"Dataframe carregado como tabela '{nome_tabela}'")
        return nome_tabela
    
        
    def mostrar_estatisticas(self, nome_tabela="dados_excel"):
        print(f"\n ESTASTICAS DA TABELA '{nome_tabela}':")
              
        total_linhas = self.conexao.execute(f"SELECT COUNT(*) FROM {nome_tabela}").fetchone()[0]
        print(f"Total de linhas: {total_linhas}")
    
        print(f"Estrutura da tabela:")
        estrutura = self.conexao.execute(f"DESCRIBE {nome_tabela}").fetchdf()
        print(estrutura)
    
    def mostrar_primeiras_linhas(self, nome_tabela='dados_excel', limite=10):
    #Mostra as primeiras linhas dos dados
        print(f"\n📋 PRIMEIRAS {limite} LINHAS DOS DADOS:")
        consulta = f"SELECT * FROM {nome_tabela} LIMIT {limite}"
        resultado = self.executar_consulta(consulta)
        if  resultado is not None:
          print(resultado)
    
    def executar_consulta(self, consulta_sql):
    
            try:
                resultado = self.conexao.execute(consulta_sql).fetchdf()
                return resultado
            except Exception as e:
                print(f"❌ Erro na consulta SQL: {e}")
                return None
    
    
    
    
    def __del__(self):
        if hasattr(self, 'conexao'):
            self.conexao.close()
        
        
if __name__ == "__main__":
    leitor = LeitorExcel("")
    
    print("=== PROCURANDO PASTA DOWNLOADS ===")
    pasta_downloads = leitor.encontrar_pasta_downloads()
    
    if pasta_downloads:
        print(f"\n=== PROCURANDO ARQUIVOS EXCEL EM: {pasta_downloads} ===")
        arquivos = leitor.listar_arquivos_excel(pasta_downloads)

        if arquivos:
            print("\nEscolha um arquivo (digite o número):")
            escolha = input("Número do arquivo: ")
            
        try:
                arquivo_escolhido = arquivos[int(escolha) - 1]
                caminho_completo = os.path.join(pasta_downloads, arquivo_escolhido)

                leitor.caminho_arquivo = caminho_completo
                if leitor.carregar_arquivo():
                            print("✅ Arquivo carregado com sucesso!")

                            # 🦆 DUCKDB AQUI!
                            print("\n" + "="*50)
                            print("🦆 INICIANDO ANÁLISE COM DUCKDB")
                            print("="*50)

                            analisador = AnalisadorDuckDB()
                            nome_tabela = analisador.carregar_dataframe(leitor.dados_completos)
                            
                            # ⬇ APENAS UMA CHAMADA ⬇
                            analisador.mostrar_estatisticas(nome_tabela)
                            
                            print("\n" + "="*50)
                            print("📊 ANÁLISE DE VENDAS")
                            print("="*50)
                            
                            # CONSULTA DOS PRODUTOS MAIS VENDIDOS
                            if hasattr(analisador, 'executar_consulta'):  # ✅ VERIFICA SE O MÉTODO EXISTE
                                consulta_produtos = """
                                    SELECT Produto, SUM(Quantidade) as Total_Vendido
                                    FROM dados_excel 
                                    GROUP BY Produto 
                                    ORDER BY Total_Vendido DESC
                                    LIMIT 5
                                """
                                produtos_top = analisador.executar_consulta(consulta_produtos)
                                print("🏆 TOP 5 PRODUTOS MAIS VENDIDOS:")
                                print(produtos_top)
                            else:
                                print("❌ Método executar_consulta não encontrado!")
                            
                
        except (ValueError, IndexError):
                print("❌ Escolha inválida!")
        else:
            print("❌ Nenhum arquivo Excel encontrado na pasta Downloads")
    else:
        print("❌ Pasta Downloads não encontrada. Digite o caminho manualmente:")
        caminho_manual = input("Caminho da pasta: ")
        arquivos = leitor.listar_arquivos_excel(caminho_manual)
        
            


