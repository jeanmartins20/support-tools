# 🛠️ Repositório de Ferramentas de Suporte Interno

Este repositório é uma coleção centralizada (monorepo) de ferramentas de automação e suporte (GUIs em Python/Tkinter) desenvolvidas para agilizar tarefas operacionais, consultas de banco de dados e geração de relatórios.

## 🎯 Objetivo

Centralizar, versionar e compartilhar scripts internos de forma segura e profissional, garantindo que as dependências sejam gerenciadas e que as credenciais (`.ini`) *nunca* sejam expostas.

## 🚀 Ferramentas Incluídas

Clique no nome de uma ferramenta abaixo para ver seu README específico, instruções e código-fonte.

1.  ### [📂 sql-query-tool/](./sql-query-tool/)
    * **Descrição:** Uma ferramenta de consulta SQL multi-conexão (V1 e V2) com interface gráfica. Permite consultas `SELECT` seguras e processamento assíncrono de "consultas de campanha" (em lote).
    * **Tecnologias:** `Tkinter`, `pyodbc`, `threading`.

2.  ### [📂 webphone-reporter/](./webphone-reporter/)
    * **Descrição:** Um gerador de relatórios de "Apuração WebPhone". Executa 4 queries de negócios complexas e exporta os resultados consolidados para um **único arquivo Excel (.xlsx)** com múltiplas planilhas formatadas.
    * **Tecnologias:** `Tkinter`, `pyodbc`, `openpyxl`, `threading`.

## ⚙️ Instalação (Para todas as ferramentas)

Recomenda-se fortemente o uso de um ambiente virtual para isolar as dependências.

1.  Clone o repositório:
    ```bash
    git clone [URL_DO_SEU_REPO]
    cd ApuraçãoWebPhoneWhatsApp
    ```

2.  Crie e ative um ambiente virtual:
    ```bash
    # Windows
    python -m venv .venv
    .venv\Scripts\activate
    
    # macOS/Linux
    python3 -m venv .venv
    source .venv/bin/activate
    ```

3.  Instale as dependências:
    ```bash
    pip install -r requirements.txt
    ```

## 🚀 Como Executar

Após instalar as dependências, navegue até a pasta da ferramenta desejada e execute o script Python.

**Exemplo (WA.py):**
```bash
cd sql-query-tool
python WA.py