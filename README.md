# Automação de Relatórios Odontológicos

Este script Python foi desenvolvido para automatizar a análise e conciliação de demonstrativos de pagamento odontológico (originalmente em PDF, convertidos para formato `.txt`) com relatórios mensais de Guias de Tratamento Odontológico (GTO) digitais (em formato `.csv`). O objetivo principal é extrair informações detalhadas de cada GTO, incluindo valores financeiros, nomes de profissionais e pacientes, e justificativas de glosa, consolidando tudo em um relatório Excel organizado, formatado e ordenado.

## Funcionalidades Principais

* Extrai dados detalhados de GTOs a partir de um arquivo de texto (`relatorio.txt`) gerado a partir do demonstrativo de pagamento PDF.
* Lê uma lista de GTOs de um arquivo CSV mensal (ex: `Relatorio-GTODigital-*.csv`).
* **Identifica automaticamente o arquivo CSV do mês corrente** na pasta do script, baseado no padrão de nome `Relatorio-GTODigital-*.csv`.
* Para cada GTO listada no CSV, busca e associa as seguintes informações extraídas do `relatorio.txt`:
    * Nome do Profissional (Contratado)
    * Nome do Paciente (Nome Civil)
    * Valor Total de Glosa da Guia
    * Valor Total Liberado da Guia
    * Justificativa da Glosa (se houver)
* Trata diferentes formatos de linhas de valores financeiros (com 3 ou 5 colunas de totais).
* Gera um relatório final consolidado em formato **Excel (`.xlsx`)**.
* Formata as colunas de valores financeiros como moeda brasileira (R$).
* **Ordena o relatório final** por Nome do Profissional, depois por Nome do Paciente e, por fim, pelo número da GTO, para facilitar a visualização.
* Calcula e exibe no console (ao final da execução) a **soma total das glosas e dos valores liberados** presentes no relatório.

## Requisitos

* Python 3.6 ou superior
* Bibliotecas Python:
    * `pandas`
    * `XlsxWriter`

## Instalação das Dependências

Para instalar as bibliotecas necessárias, abra seu terminal ou prompt de comando e execute:

```bash
pip install pandas XlsxWriter
```

## Como Usar

Siga estes passos para executar o script:

1.  **Preparação Manual do Arquivo de Texto (Etapa Crucial):**
    * Abra o arquivo PDF do demonstrativo de pagamento no **Adobe Acrobat Reader** (a versão gratuita é suficiente).
    * No menu, vá em `Arquivo` -> `Salvar como outro` -> `Texto...`.
    * Salve o arquivo com o nome exato **`relatorio.txt`** na mesma pasta onde o script Python (`AUTOMACAO.py`) estará localizado.

2.  **Organização dos Arquivos:**
    Certifique-se de que os seguintes arquivos estejam na mesma pasta:
    * `AUTOMACAO.py` (o script Python fornecido)
    * `relatorio.txt` (o arquivo de texto gerado no passo anterior)
    * O arquivo CSV mensal de GTOs Digitais. O script espera um nome no padrão `Relatorio-GTODigital-*.csv` (ex: `Relatorio-GTODigital-01_05_2025-31_05_2025.csv`). Ele buscará automaticamente o arquivo correto se apenas um correspondente ao padrão estiver na pasta.

3.  **Execução do Script:**
    * Abra um terminal ou prompt de comando.
    * Navegue (usando o comando `cd`) até a pasta onde você salvou os arquivos.
    * Execute o script com o comando:
        ```bash
        python AUTOMACAO.py
        ```

## Arquivos de Entrada Detalhados

* **`relatorio.txt`**:
    * **Origem:** Texto extraído do PDF do demonstrativo de pagamento odontológico via Adobe Acrobat Reader.
    * **Estrutura Esperada:** Deve conter seções "Dados do Prestador" (com "7 - Nome do Contratado") e, dentro destas, seções "Dados do Pagamento" (com "12 - Nome Civil", "13 - Número da Guia", "27 - Observação / Justificativa", "32 - Valor Total Liberado Guia (R$)"). A formatação com campos numerados é essencial para o parser. **Este script foi otimizado para o modelo de relatório da OdontoPrev.**
* **Arquivo CSV (`Relatorio-GTODigital-*.csv`)**:
    * **Origem:** Relatório de GTOs digitais do período desejado.
    * **Estrutura Esperada:** Deve conter, no mínimo, uma coluna chamada **"GTO"** com os números das guias que serão pesquisadas no `relatorio.txt`.

## Arquivo de Saída

* **`Resultado_Busca_GTOs_Completo.xlsx`**:
    * Uma planilha Excel gerada pelo script, contendo as seguintes colunas para cada GTO do arquivo CSV:
        * `GTO`
        * `Status` (indicando se foi encontrada no `relatorio.txt`)
        * `Profissional`
        * `Paciente`
        * `Valor Total Glosa (R$)` (formatado como moeda)
        * `Valor Total Liberado (R$)` (formatado como moeda)
        * `Justificativa da Glosa`
    * Os dados no Excel estarão ordenados por Profissional, Paciente e GTO. As colunas de valores financeiros estarão formatadas em Reais (R$), e as larguras das colunas ajustadas para melhor visualização.

## Observações Importantes

### Complexidade de PDFs
A extração de dados diretamente de arquivos PDF é inerentemente complexa devido à grande variedade na estrutura interna desses arquivos. A conversão manual do PDF para um arquivo de texto (`.txt`) simples, utilizando a funcionalidade "Salvar como Texto" do Adobe Acrobat Reader, é um passo fundamental neste fluxo de trabalho. Ela padroniza a entrada para o script Python, aumentando significativamente a precisão e confiabilidade da extração dos dados.

### Compatibilidade do Modelo de Relatório
Este script foi desenvolvido e testado especificamente com base no modelo de demonstrativo de pagamento da **OdontoPrev** (plano de saúde odontológico). Embora o script tente ser flexível, a estrutura de relatórios de outros planos de saúde ou operadoras pode variar significativamente. Portanto, para melhor precisão e funcionamento, **recomenda-se a utilização deste script com relatórios gerados pela OdontoPrev.** A adaptação para outros modelos de relatório provavelmente exigirá modificações nas regras de extração de dados (expressões regulares) contidas no código Python.

---
