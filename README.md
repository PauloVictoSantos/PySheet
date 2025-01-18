# **Sistema de Gestão de Estoque com Relatórios e Gráficos 📊**

Este projeto é um sistema automatizado para gerenciamento de estoque usando **Python** e **OpenPyXL**. Ele gera relatórios detalhados e gráficos que ajudam a visualizar e analisar os dados de estoque de forma eficiente.

## **Funcionalidades 🚀**

- **Cadastro automatizado de produtos**: Geração de nomes, valores, e quantidade.
- **Cálculos automatizados**:
  - Preço de venda com base na lucratividade.
  - Lucro total e valor total de cada produto.
- **Relatórios dinâmicos**: Totais gerais de lucro e valor em estoque.
- **Gráficos interativos**:
  - Gráfico de barras para valores totais por produto.
  - Gráfico de pizza para os 5 produtos com maior lucratividade.
  - Gráfico de linha para quantidades em estoque.

## **Tecnologias Utilizadas 🛠️**

- **Python 3.8+**
- **Poetry ou Pip**: Para gerenciamento de dependências e ambientes virtuais.
- **Biblioteca OpenPyXL**: Para manipulação de planilhas Excel.
- **Microsoft Excel**: Para visualização e edição dos arquivos gerados.

## **Como Utilizar 📋**

### **Pré-requisitos**
1. Certifique-se de ter o Python instalado. [Download do Python](https://www.python.org/downloads/).
2. Instale o Poetry para gerenciamento de dependências:
   - [Instalar o Poetry](https://python-poetry.org/docs/#installation)
   - Ou use o comando:
     ```bash
     pip install poetry
     ```

### **Instalação de Dependências**
Com o **Poetry** instalado, você pode facilmente instalar as dependências do projeto. No terminal, execute os seguintes comandos:

1. Clone este repositório:
    ```bash
    git clone https://github.com/PauloVictoSantos/PySheet.git
    cd PySheet
    ```

2. Instale as dependências do projeto com o Poetry:
    ```bash
    poetry install
    ```

### **Executando o Projeto**
Após instalar as dependências, você pode rodar o script principal com o comando abaixo:

1. Execute o script para gerar o arquivo `estoque.xlsx`:
    ```bash
    poetry run python main.py
    ```

2. Abra o arquivo `estoque.xlsx` gerado na pasta do projeto para visualizar os dados e os gráficos interativos.

