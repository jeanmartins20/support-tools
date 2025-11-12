# 📈 Gerador de Relatórios de Apuração WebPhone (wp.py)

Uma aplicação desktop (GUI) de uso interno, desenvolvida em Python e Tkinter, para automatizar a apuração de dados de WebPhone. A ferramenta conecta-se a um banco de dados SQL Server (ambiente de leitura V2) e executa quatro queries complexas e pré-definidas para gerar um relatório consolidado.

> [!NOTE]
> *Insira aqui um print-screen (captura de tela) da aba de resultados do aplicativo.*

## 🌟 Recursos Principais

* **Automação de Relatórios:** Executa quatro consultas de negócios essenciais com um único clique:
    1.  Atualização de Plano
    2.  Base de Cliente
    3.  Detalhamento de Faturamento
    4.  Base de Crédito
* **Execução Assíncrona:** Utiliza **`threading`** e **`queue`** para executar todas as quatro consultas em segundo plano, mantendo a interface responsiva e informando o usuário sobre o progresso.
* **Exportação para Excel (Multi-Sheet):** O recurso principal é a exportação de **todos os quatro relatórios** para um **único arquivo `.xlsx`**, onde cada relatório é organizado em sua própria planilha (worksheet) formatada.
* **Persistência de Configuração:** Salva os dados de conexão em um arquivo `.ini` (`apuracao_webphone.ini`).

## 🛠️ Tecnologias Utilizadas

* Python 3
* Tkinter (ttk)
* pyodbc (para conectividade SQL Server)
* openpyxl (para criação e formatação de arquivos `.xlsx`)
* threading / queue (para execução assíncrona)
* configparser (para gerenciamento de `.ini`)

## 🚀 Como Executar

1.  Certifique-se de que está na raiz do repositório (`ApuraçãoWebPhoneWhatsApp`) e que as dependências do `requirements.txt` principal foram instaladas.
2.  Navegue até esta pasta:
    ```bash
    cd webphone-reporter
    ```
3.  Execute o script:
    ```bash
    python wp.py
    ```

**Importante:** Na primeira execução, um arquivo `apuracao_webphone.ini` será criado. Este arquivo contém credenciais e **já está sendo ignorado** pelo `.gitignore` da raiz.
