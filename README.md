# Automatic PDF Document Generator / Gerador Automático de Documentos PDF

# About the Project

This project automates the generation of PDF documents from an Excel spreadsheet and a DOCX template.

The application reads data from the spreadsheet, automatically fills the template fields, and generates individual PDF files for each record found.

## Features

* Read data from Excel spreadsheets.
* Automatic DOCX template filling.
* Automatic DOCX to PDF conversion using LibreOffice.
* Batch processing of multiple records.
* Automatic generation of customized documents.

## Requirements

* Python 3.10 or later
* LibreOffice
* pandas
* python-docx
* openpyxl

## Installation

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

1. Create or fill the `items.xlsx` spreadsheet.
2. Run the application:

```bash
python main.py
```

3. Enter the applicant's name when prompted.

The PDF documents will be generated automatically.

## Spreadsheet Structure

The `items.xlsx` spreadsheet must contain the following columns:

| Column      | Description         |
| ----------- | ------------------- |
| DESCRIPTION | Item description    |
| CATEGORY    | Item category       |
| PRICE       | Item price          |
| UNIT        | Unit of measurement |
| IMPACT      | Impact level        |
| PRIORITY    | Priority level      |

## Accepted CATEGORY Values

* CATEGORY_A
* CATEGORY_B
* CATEGORY_C
* CATEGORY_D

## Accepted IMPACT Values

* HIGH IMPACT
* MEDIUM IMPACT
* LOW IMPACT

## Accepted PRIORITY Values

* HIGH
* MEDIUM
* LOW

## Example

| DESCRIPTION | CATEGORY   | PRICE   | UNIT | IMPACT        | PRIORITY |
| ----------- | ---------- | ------- | ---- | ------------- | -------- |
| Laptop      | CATEGORY_A | 3500.00 | UN   | HIGH IMPACT   | HIGH     |
| Monitor     | CATEGORY_B | 1200.00 | UN   | MEDIUM IMPACT | MEDIUM   |
| Mouse       | CATEGORY_C | 80.00   | UN   | LOW IMPACT    | LOW      |

## Output

For each row in the spreadsheet:

1. A DOCX document is generated from the template.
2. The document is converted to PDF using LibreOffice.
3. The temporary DOCX file is removed.
4. The final PDF file remains available for use.

# Sobre o Projeto

Este projeto automatiza a criação de documentos PDF a partir de uma planilha Excel e de um modelo DOCX.

O sistema lê os dados da planilha, preenche automaticamente os campos do documento modelo e gera arquivos PDF individuais para cada registro encontrado.

## Funcionalidades:

* Leitura de dados de planilhas Excel.
* Preenchimento automático de modelos DOCX.
* Conversão automática de DOCX para PDF utilizando LibreOffice.
* Processamento em lote de múltiplos registros.
* Geração automática de documentos personalizados.

## Requisitos:

* Python 3.10 ou superior
* LibreOffice
* pandas
* python-docx
* openpyxl

## Instalação:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Como Utilizar:

1. Crie ou preencha a planilha items.xlsx.
2. Execute o programa: 

```bash
python main.py
```

3. Informe o nome do solicitante quando solicitado.

Os documentos PDF serão gerados automaticamente.

## Estrutura da Planilha

A planilha items.xlsx deve conter as seguintes colunas:

| Coluna      | Descrição           |
| ----------- | ------------------- |
| DESCRIPTION | Descrição do item   |
| CATEGORY    | Categoria do item   |
| PRICE       | Valor do item       |
| UNIT        | Unidade de medida   |
| IMPACT      | Nível de impacto    |
| PRIORITY    | Nível de prioridade |

## Valores Aceitos para CATEGORY

* CATEGORY_A
* CATEGORY_B
* CATEGORY_C
* CATEGORY_D

## Valores Aceitos para IMPACT

* HIGH IMPACT
* MEDIUM IMPACT
* LOW IMPACT

## Valores Aceitos para PRIORITY

* HIGH
* MEDIUM
* LOW

## Example

| DESCRIPTION | CATEGORY   | PRICE   | UNIT | IMPACT        | PRIORITY |
| ----------- | ---------- | ------- | ---- | ------------- | -------- |
| Notebook    | CATEGORY_A | 3500.00 | UN   | HIGH IMPACT   | HIGH     |
| Monitor     | CATEGORY_B | 1200.00 | UN   | MEDIUM IMPACT | MEDIUM   |
| Mouse       | CATEGORY_C | 80.00   | UN   | LOW IMPACT    | LOW      |4

## Saída

Para cada linha da planilha:

1. Um documento DOCX é gerado a partir do modelo.
2. O documento é convertido para PDF utilizando LibreOffice.
3. O arquivo DOCX temporário é removido.
4. O PDF final permanece disponível para uso.