## Descrição
Este projeto foi desenvolvido para automação e análise de dados em planilhas Excel da Anvisa, especificamente para classificar e destacar células com base em critérios definidos.
O código modifica as cores das células em colunas específicas e gera um resumo de diagnósticos com base em condições específicas.

## Funcionalidades
- **Alteração de Cores**: Destaca células nas colunas de assunto, CNPJ e NCM com cores diferentes, dependendo de listas de critérios.
- **Resumo de Diagnósticos**: Gera um resumo em uma coluna adicional, identificando situações como "Indeferida" e erros comuns.
- **Manipulação de Planilhas**: Utiliza a biblioteca `openpyxl` para leitura e escrita em arquivos Excel.

## Requisitos
- Python 3.x
- Bibliotecas: `openpyxl`
