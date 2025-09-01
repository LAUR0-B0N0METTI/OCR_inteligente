# Sistema OCR Inteligente e Análise de Log

Este repositório contém dois scripts Python interligados que formam um sistema robusto para processamento inteligente de documentos e análise de logs. O `OCR_inteligente.py` é o coração do sistema, responsável por extrair texto de diversos formatos de arquivo (incluindo imagens e PDFs escaneados via OCR), classificar documentos com base em inteligência artificial e organizar os arquivos. O `OCR_inteligente_leitor_log.py` complementa o sistema, fornecendo uma ferramenta para analisar os logs gerados pelo processador OCR, sumarizando a eficiência da classificação.

## Visão Geral do Projeto

O objetivo principal deste sistema é automatizar e otimizar o fluxo de trabalho de gerenciamento de documentos em larga escala. Ele é projetado para lidar com uma vasta gama de tipos de documentos, extrair informações cruciais (como competência e CNPJ) e classificá-los de forma inteligente, reduzindo significativamente a necessidade de intervenção manual. A capacidade de analisar os logs gerados permite monitorar e aprimorar continuamente a performance do sistema de classificação.

### Componentes Principais:

1.  **`OCR_inteligente.py`**: Um processador de documentos multifuncional que integra OCR, extração de dados e um motor de classificação baseado em IA. Ele é capaz de ler, interpretar e categorizar documentos, salvando metadados importantes em formato JSON e organizando os arquivos em uma estrutura de diretórios lógica.

2.  **`OCR_inteligente_leitor_log.py`**: Uma ferramenta de análise de log que processa os arquivos de log gerados pelo `OCR_inteligente.py`. Ele quantifica a quantidade de documentos classificados e não classificados, calcula a porcentagem de falha na classificação e gera um relatório detalhado, auxiliando na identificação de áreas para melhoria.

## `OCR_inteligente.py` - Processador Inteligente de Documentos

### Descrição Detalhada

O `OCR_inteligente.py` é uma solução avançada para a automação do processamento de documentos. Ele combina funcionalidades de extração de texto (OCR), análise de conteúdo e classificação inteligente para gerenciar grandes volumes de arquivos. O script é altamente configurável e projetado para ser resiliente a diferentes formatos e qualidades de documentos.

### Funcionalidades:

*   **Extração de Texto Abrangente**: Suporta a extração de texto de uma ampla variedade de formatos de arquivo:
    *   **Imagens**: JPG, JPEG, PNG, TIFF, TIF, BMP (via Tesseract OCR).
    *   **PDFs**: Leitura direta de texto e OCR para PDFs escaneados ou baseados em imagem (via PyPDF2, PyMuPDF e Tesseract).
    *   **Documentos Office**: DOCX (via `python-docx`), XLSX (via `pandas`).
    *   **Web/Estruturados**: HTML (via BeautifulSoup), XML.
    *   **Texto Plano**: TXT.
*   **Descompactação Automática**: Lida com arquivos compactados (`.zip` e `.rar`), extraindo seu conteúdo para processamento.
*   **Pré-processamento de Imagens**: Aplica técnicas de aprimoramento de imagem (nitidez, escala de cinza, binarização) para otimizar a precisão do OCR.
*   **Extração Inteligente de Competência**: Utiliza um motor de IA com padrões regex e análise contextual para identificar a competência (mês/ano de referência) do documento, mesmo em formatos variados.
*   **Classificação Avançada de Documentos**: Possui um motor de classificação baseado em IA que atribui pontuações de confiança para diferentes tipos de documentos (Nota Fiscal, Extrato Bancário, Boleto, DACTE, SPED Fiscal, Relatório de Faturamento, Fatura de Serviços) com base em indicadores primários, secundários e negativos encontrados no texto.
*   **Extração de CNPJ**: Identifica e valida CNPJs presentes no texto do documento.
*   **Geração de JSON de Metadados**: Para cada documento processado, um arquivo JSON é gerado contendo todos os metadados extraídos e classificados (tipo de documento, subtipo, competência, CNPJ, etc.).
*   **Organização de Arquivos**: Move os documentos processados para uma estrutura de pastas organizada por tipo de documento e competência, facilitando a recuperação.
*   **Registro Detalhado (Logging)**: Gera um arquivo de log (`processamento_log.log`) que registra todas as etapas do processamento, incluindo erros, avisos e resultados da classificação.

### Como Funciona:

1.  **Configuração Inicial**: Define o `BASE_PATH` (diretório raiz para processamento), o caminho para a pasta de saída JSON e os executáveis do Tesseract OCR e WinRAR (para RAR).
2.  **Varredura de Diretórios**: O script percorre recursivamente o `BASE_PATH`, identificando todos os arquivos a serem processados.
3.  **Descompactação**: Se um arquivo compactado (`.zip` ou `.rar`) for encontrado, ele é descompactado em um diretório temporário, e seus conteúdos são adicionados à fila de processamento.
4.  **Extração de Texto e OCR**: Para cada arquivo, o `extract_text_from_file` tenta extrair seu conteúdo textual. Para imagens e PDFs escaneados, ele utiliza o Tesseract OCR, aplicando pré-processamento de imagem para melhorar a qualidade do reconhecimento.
5.  **Análise Inteligente**: O texto extraído é então passado para a classe `IntelligentDocumentAnalyzer`, que contém os motores de IA para:
    *   **Extração de Competência**: O método `extract_competence_with_ai` utiliza padrões regex, análise contextual e análise do nome do arquivo para determinar a competência do documento com um nível de confiança.
    *   **Classificação de Documentos**: O método `classify_document_with_ai` avalia o texto com base em um conjunto de indicadores (palavras-chave, padrões) para determinar o tipo mais provável do documento e um score de confiança.
    *   **Validação de CNPJ**: A classe `CNPJValidator` verifica a validade de CNPJs encontrados.
6.  **Geração de JSON**: Os metadados extraídos (tipo, subtipo, competência, CNPJ, etc.) são compilados em um dicionário e salvos como um arquivo JSON na pasta `01-JSON`.
7.  **Organização de Arquivos**: O arquivo original é movido para uma pasta de destino final, que é determinada pela sua classificação e competência (ex: `BASE_PATH/Nota Fiscal Eletrônica/2024/01-Janeiro/`).
8.  **Registro**: Todas as ações e resultados são registrados no `processamento_log.log`, fornecendo um rastro completo do processamento.

### Dependências:

*   `os`, `json`, `re`, `shutil`, `zipfile`, `rarfile`, `time`, `datetime`, `timedelta`:
    Para operações de sistema de arquivos, manipulação de JSON, expressões regulares, cópia de arquivos, descompactação, e manipulação de tempo.
*   `imgkit`: Para converter HTML em imagem (requer wkhtmltopdf).
*   `reportlab`: Para geração de PDFs.
*   `PIL (Pillow)`: Para processamento de imagens.
*   `pytesseract`: Interface Python para o Tesseract OCR (requer Tesseract instalado e configurado).
*   `docx`: Para trabalhar com arquivos DOCX.
*   `pandas`: Para trabalhar com dados tabulares, especialmente de Excel.
*   `PyPDF2`, `fitz` (PyMuPDF): Para leitura de PDFs.
*   `bs4` (BeautifulSoup): Para parsing de HTML/XML.
*   `logging`: Para geração de logs.
*   `collections.Counter`, `difflib.SequenceMatcher`: Para análise de texto e similaridade.

Para instalar as dependências, execute:

```bash
pip install Pillow pytesseract python-docx pandas PyPDF2 pymupdf beautifulsoup4 rarfile
```

**Observações sobre `rarfile`, `pytesseract` e `imgkit`:**

*   **`rarfile`**: Este módulo requer que o executável `UnRAR.exe` (parte do WinRAR) esteja instalado no seu sistema e que o caminho para ele seja configurado na variável `rarfile.UNRAR_TOOL` no script. Ex: `rarfile.UNRAR_TOOL = r"C:\Program Files\WinRAR\UnRAR.exe"`.
*   **`pytesseract`**: Este módulo requer que o Tesseract OCR esteja instalado no seu sistema. O caminho para o executável `tesseract.exe` deve ser configurado na variável `pytesseract.pytesseract.tesseract_cmd` no script. Ex: `pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"`.
*   **`imgkit`**: Este módulo requer que o `wkhtmltopdf` esteja instalado no seu sistema para converter HTML em imagens. Você pode precisar instalá-lo separadamente e garantir que esteja no PATH do sistema.

### Como Executar:

1.  **Configuração**: Edite o script `OCR_inteligente.py` e ajuste as variáveis `BASE_PATH`, `JSON_OUTPUT_PATH`, `rarfile.UNRAR_TOOL` e `pytesseract.pytesseract.tesseract_cmd` para refletir os caminhos corretos em seu ambiente.
2.  **Execução**: Execute o script Python diretamente:

    ```bash
    python OCR_inteligente.py
    ```

O script iniciará o processamento dos arquivos no `BASE_PATH`, imprimirá o progresso no console, gerará arquivos JSON na pasta `01-JSON` e registrará as atividades em `processamento_log.log`.

## `OCR_inteligente_leitor_log.py` - Analisador de Log de Processamento

### Descrição Detalhada

O `OCR_inteligente_leitor_log.py` é uma ferramenta de linha de comando projetada para analisar o arquivo de log gerado pelo `OCR_inteligente.py` (`processamento_log.log`). Ele extrai informações sobre a classificação dos documentos, quantifica o número de arquivos classificados e não classificados, e calcula a porcentagem de documentos que necessitam de revisão manual. O resultado é um relatório conciso que é exibido no console e salvo em um arquivo de texto.

### Funcionalidades:

*   **Análise de Log**: Lê e interpreta o arquivo `processamento_log.log`.
*   **Contagem de Classificações**: Contabiliza o número de documentos que foram classificados com sucesso e aqueles que foram marcados como "Documento Não Classificado".
*   **Cálculo de Eficiência**: Calcula a porcentagem de documentos não classificados, fornecendo uma métrica da eficácia do sistema de OCR e classificação.
*   **Listagem de Não Classificados**: Lista os caminhos completos dos arquivos que não puderam ser classificados, facilitando a identificação e correção manual.
*   **Geração de Relatório**: Gera um relatório formatado com o resumo da análise, incluindo totais e porcentagens, e o salva em um arquivo (`relatorio_analise.txt`).

### Como Funciona:

1.  **Entrada**: O script recebe o caminho para o arquivo de log (`processamento_log.log`) como entrada.
2.  **Leitura e Parsing**: Ele lê o arquivo de log linha por linha, utilizando expressões regulares para identificar as entradas de "Processando arquivo" e "-> Tipo" para determinar o status de classificação de cada documento.
3.  **Contagem**: Mantém contadores para arquivos classificados e não classificados.
4.  **Geração de Relatório**: Após processar todo o log, ele compila as informações em um relatório textual, incluindo a data da análise, o resumo da classificação e a lista de arquivos não classificados.
5.  **Saída**: O relatório é impresso no console e salvo em um arquivo chamado `relatorio_analise.txt` no mesmo diretório de execução do script.

### Dependências:

*   `re`: Para expressões regulares (parsing do log).
*   `os`: Para operações de sistema de arquivos (verificação de existência de arquivo, manipulação de caminhos).
*   `datetime`: Para incluir a data e hora da análise no relatório.

Para instalar as dependências, execute:

```bash
pip install
```

(Este script não possui dependências externas além das bibliotecas padrão do Python).

### Como Executar:

1.  **Configuração**: Edite o script `OCR_inteligente_leitor_log.py` e ajuste a variável `caminho_do_log` para apontar para o seu arquivo `processamento_log.log`.
2.  **Execução**: Execute o script Python diretamente:

    ```bash
    python OCR_inteligente_leitor_log.py
    ```

O script imprimirá o relatório no console e salvará o arquivo `relatorio_analise.txt` no mesmo diretório.

## Considerações Finais

O sistema OCR Inteligente, em conjunto com seu analisador de log, oferece uma solução completa para o desafio de gerenciar e classificar grandes volumes de documentos. A automação da extração de dados e a inteligência na classificação reduzem a carga de trabalho manual, enquanto a análise de log fornece insights valiosos para a melhoria contínua do sistema. Este conjunto de ferramentas é ideal para empresas que buscam eficiência e precisão no tratamento de seus documentos digitais.

### Autor:
Lauro Bonometti