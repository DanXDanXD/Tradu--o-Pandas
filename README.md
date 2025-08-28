# Tradutor de Planilhas Excel

Este script em Python automatiza a tradução de arquivos `.xlsx` (Excel) inteiros, incluindo todas as suas abas e colunas de texto, de português para inglês. Ele é ideal para quem trabalha com dados multilíngues e precisa de uma solução rápida para traduzir grandes volumes de texto em planilhas.

---

## Recursos

- **Tradução Automática:** Traduz todo o conteúdo de texto de cada aba de uma planilha.
- **Processamento em Lote:** Percorre uma pasta inteira, identificando e traduzindo todos os arquivos `.xlsx`.
- **Identificação Inteligente:** Detecta e traduz apenas colunas com dados do tipo "objeto" (geralmente texto), ignorando números.
- **Criação de Novo Arquivo:** Salva a versão traduzida em um novo arquivo com o sufixo `_TRADUZIDO.xlsx`, mantendo o arquivo original intacto.

---

## Pré-requisitos

Para rodar este script, você precisará ter o Python instalado. Além disso, as seguintes bibliotecas são necessárias:

- `pandas`
- `deep-translator`
- `openpyxl`

Você pode instalar todas as dependências de uma vez usando `pip`:

```bash
pip install pandas deep-translator openpyxl