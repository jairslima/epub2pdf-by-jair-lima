# Especificação Técnica: epub2pdf-by-jair-lima

Este documento define a especificação técnica para o conversor de e-books EPUB para documentos em formato PDF no padrão A5 para impressão física.

## 1. Objetivo
O objetivo principal deste projeto é criar uma ferramenta CLI em Python que converta arquivos EPUB para documentos PDF perfeitamente formatados de acordo com os padrões tipográficos de um livro impresso de alta qualidade: especificamente no tamanho de folha A5.

## 2. Escopo
* **Extração do Pacote:** Descompactação programática do arquivo EPUB.
* **Leitura de Metadados:** Extração automática de dados como título, autor e idioma do e-book.
* **Processamento de HTML:** Parsing de conteúdo XHTML com BeautifulSoup: identificando capítulos, parágrafos, listas, tabelas e formatação em linha.
* **Geração de PDF A5:** Criação de documento PDF estruturado em tamanho A5 com margens espelhadas ajustadas para encadernação de livro físico.
* **Preservação Visual:** Manutenção da cor de destaque vermelho para as palavras de Jesus e outras formatações importantes do livro original.
* **Páginas de Rosto e Sumário:** Inclusão automática de página de rosto e sumário dinâmico paginado.
* **Paginação:** Inserção de cabeçalhos e rodapés com paginação automática.
* **Automação Word COM:** Uso de automação COM do Microsoft Word no Windows para fazer a paginação, a atualização de sumários e a exportação nativa para PDF com qualidade profissional.

## 3. Arquitetura do Projeto
O utilitário é implementado como um script CLI Python autônomo.
O fluxo de execução segue os seguintes passos:
1. **Validação de Argumentos:** A CLI aceita o caminho do arquivo EPUB de entrada, o caminho de saída PDF e argumentos opcionais de configuração (fonte, tamanho da fonte, margens, etc.).
2. **Leitura e Extração:** O EPUB é carregado usando a biblioteca zipfile. O arquivo `content.opf` é lido para mapear os metadados do livro e a ordem lógica de leitura.
3. **Mapeamento de Mídia:** Todas as imagens são extraídas em memória para inserção no documento intermediário.
4. **Montagem do DOCX Temporário:**
   * Configuração das propriedades de página da primeira seção (papel A5, margens espelhadas).
   * Criação e formatação dos estilos básicos utilizando fontes serifadas elegantes (Georgia por padrão).
   * Criação da Página de Rosto e do Sumário Dinâmico (TOC).
   * Varredura dos arquivos XHTML em ordem cronológica: injetando o conteúdo de texto no Word com a formatação adequada.
5. **Automação COM e Exportação:** O script inicia o Microsoft Word de forma oculta via win32com, carrega o arquivo DOCX temporário, atualiza todos os campos de página e o sumário, exporta o documento final em formato PDF e encerra o processo do Word.
6. **Limpeza:** O arquivo DOCX intermediário é excluído do disco.

## 4. Dependências
* Python 3.x
* `python-docx` para geração de documentos Word intermediários.
* `beautifulsoup4` e `lxml` para parsing estruturado de XHTML.
* `pillow` para redimensionamento e validação de imagens.
* `pywin32` para automação do Word via COM no Windows (conversão e exportação para PDF).

## 5. Decisões de Design
* **Conversão via DOCX Intermediário:** O uso da ponte DOCX para PDF via Word COM foi escolhido por ser a única forma de obter um layout A5 perfeitamente paginado com sumário automático atualizado de acordo com o posicionamento físico exato dos capítulos, garantindo 100% de consistência estética com o projeto epub2docx.
* **Margens Espelhadas (Mirror Margins):** O padrão para livro físico exige margens internas maiores. O script configura margens internas de 2.2 cm e externas de 1.8 cm para permitir encadernação sem cortar texto.
* **Tipografia do Livro:** A fonte padrão é Georgia: uma fonte serifada que oferece excelente legibilidade em papel impresso A5. O tamanho padrão é 10.5 pt com espaçamento entre linhas de 1.25.
* **Preservação de Estilos Inline:** O parser de HTML mapeia cores CSS (como `color: red` ou tags com a classe `red`) e aplica a cor vermelha no run correspondente do DOCX antes de convertê-lo.

## 6. Critérios de Aceite
* O arquivo PDF gerado deve ser válido e legível.
* As páginas do documento devem estar configuradas no formato A5.
* As margens do documento devem ser espelhadas.
* O sumário dinâmico no PDF deve conter as referências corretas das páginas reais.
* A paginação automática deve ser exibida no rodapé.
* As palavras de Jesus e marcações em vermelho devem ser perfeitamente reproduzidas no PDF.

## 7. Instruções Operacionais
* Para instalar dependências: `pip install -r requirements.txt`
* Uso básico da ferramenta individual:
   ```powershell
   python epub2pdf.py "caminho/do/livro.epub" "caminho/do/resultado.pdf"
   ```
* Parâmetros avançados suportados:
   * `--font "Nome da Fonte"`: Altera a fonte principal (padrão: Georgia).
   * `--size 11`: Altera o tamanho do texto do corpo (padrão: 10.5).
   * `--margin-inner 2.2`: Margem interna em centímetros.
   * `--margin-outer 1.8`: Margem externa em centímetros.
   * `--title "Título Customizado"`: Substitui o título extraído do EPUB.
   * `--author "Autor Customizado"`: Substitui o autor extraído do EPUB.

---

## 8. Sincronização e Processamento em Lote

O projeto inclui o utilitário `find_and_convert.py` para processar bibliotecas de livros de forma automatizada:

### Objetivo:
Sincronizar a biblioteca de livros em formato EPUB com a biblioteca em formato PDF, convertendo de forma inteligente apenas os arquivos ausentes.

### Funcionamento e Execução:
* O script lê as pastas configuradas (`EPUB/Igreja` e `PDF/Igreja`), compara os títulos e autores extraídos e executa a conversão nos arquivos faltantes.
* Comando para executar a sincronização em lote:
   ```powershell
   python find_and_convert.py
   ```

