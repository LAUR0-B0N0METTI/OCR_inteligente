import re
import os
from datetime import datetime

def analisar_log(caminho_do_arquivo):
    """
    Analisa um arquivo de log para contabilizar arquivos classificados e não classificados.
    Calcula a porcentagem de não classificados e, ao final, exibe o relatório
    no console e o salva em um arquivo .txt.

    Args:
        caminho_do_arquivo (str): O caminho completo para o arquivo .log.
    """
    # Verifica se o caminho do arquivo existe
    if not os.path.exists(caminho_do_arquivo):
        print(f"--- ERRO ---")
        print(f"O arquivo não foi encontrado no caminho especificado: {caminho_do_arquivo}")
        print("Por favor, verifique se o caminho está correto e tente novamente.")
        return

    # Contadores e listas para armazenar os resultados
    classificados = 0
    nao_classificados = 0
    arquivos_nao_classificados = []
    
    arquivo_atual = None
    diretorio_raiz = ""

    # Expressões regulares
    padrao_arquivo = re.compile(r"Processando arquivo: (.*?) \s*\(Cliente:")
    padrao_tipo = re.compile(r"-> Tipo: (.*)")
    padrao_diretorio_raiz = re.compile(r"Diretório Raiz para Processamento: (.*)")

    try:
        with open(caminho_do_arquivo, 'r', encoding='utf-8') as f:
            for linha in f:
                match_diretorio = padrao_diretorio_raiz.search(linha)
                if match_diretorio and not diretorio_raiz:
                    diretorio_raiz = match_diretorio.group(1).strip()

                match_arquivo = padrao_arquivo.search(linha)
                if match_arquivo:
                    arquivo_atual = match_arquivo.group(1).strip()
                    continue

                match_tipo = padrao_tipo.search(linha)
                if match_tipo and arquivo_atual:
                    tipo_documento = match_tipo.group(1).strip()
                    
                    if tipo_documento == "Documento Não Classificado":
                        nao_classificados += 1
                        caminho_completo = os.path.join(diretorio_raiz, arquivo_atual)
                        arquivos_nao_classificados.append(caminho_completo)
                    else:
                        classificados += 1
                    
                    arquivo_atual = None

    except Exception as e:
        print(f"Ocorreu um erro inesperado ao processar o arquivo: {e}")
        return

    # --- Cálculo da Porcentagem ---
    total_de_arquivos = classificados + nao_classificados
    # Evita divisão por zero se o log estiver vazio
    if total_de_arquivos > 0:
        porcentagem_nao_classificados = (nao_classificados / total_de_arquivos) * 100
    else:
        porcentagem_nao_classificados = 0

    # --- Montagem do Relatório ---
    linhas_relatorio = []
    
    linhas_relatorio.append("======================================================")
    linhas_relatorio.append("          ANÁLISE DO LOG DE PROCESSAMENTO           ")
    linhas_relatorio.append("======================================================")
    linhas_relatorio.append(f"Arquivo de Log Analisado: {caminho_do_arquivo}")
    linhas_relatorio.append(f"Data da Análise: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    
    # Formata a porcentagem para o padrão brasileiro (vírgula)
    porcentagem_formatada = f"{porcentagem_nao_classificados:.2f}".replace('.', ',')

    linhas_relatorio.append("\nResumo da Classificação:\n")
    linhas_relatorio.append(f"  - Arquivos Classificados com Sucesso: {classificados}")
    linhas_relatorio.append(f"  - Arquivos Não Classificados:         {nao_classificados}")
    if total_de_arquivos > 0:
        linhas_relatorio.append(f"\n  - {porcentagem_formatada}% de arquivos não classificados")

    linhas_relatorio.append("\n------------------------------------------------------\n")

    if arquivos_nao_classificados:
        linhas_relatorio.append("Arquivos com 'Tipo: Documento Não Classificado':\n")
        for i, caminho in enumerate(arquivos_nao_classificados, 1):
            linhas_relatorio.append(f"{i}. {caminho}")
    else:
        linhas_relatorio.append("🎉 Todos os arquivos foram classificados com sucesso!")
        
    linhas_relatorio.append("\n======================================================")

    # --- Apresentação dos Resultados no Console ---
    relatorio_final = "\n".join(linhas_relatorio)
    print(relatorio_final)

    # --- Salvamento do Relatório em Arquivo .txt ---
    nome_arquivo_saida = "relatorio_analise.txt"
    try:
        with open(nome_arquivo_saida, 'w', encoding='utf-8') as f:
            f.write(relatorio_final)
        print(f"\n✅ Relatório também foi salvo com sucesso em: {os.path.abspath(nome_arquivo_saida)}")
    except Exception as e:
        print(f"\n❌ Ocorreu um erro ao salvar o arquivo de relatório: {e}")


# --- INÍCIO DA EXECUÇÃO ---
# Especifique o caminho para o seu arquivo de log aqui.
caminho_do_log = r"C:\Users\laurob\Desktop\processamento_log.log"

# Chama a função principal para iniciar a análise
analisar_log(caminho_do_log)