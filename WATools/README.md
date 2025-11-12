---

### 2. README do SQL Query Tool (Para a pasta: `sql-query-tool/README.md`)

```markdown
# 🧰 Ferramenta de Consulta SQL Multi-Conexão (WA.py)

Uma aplicação desktop (GUI) desenvolvida em Python e Tkinter, projetada para analistas de suporte ou desenvolvedores que precisam gerenciar e consultar simultaneamente dois ambientes de banco de dados SQL Server (por exemplo, V1 e V2) de forma segura e eficiente.

> [!NOTE]
> *Insira aqui um print-screen (captura de tela) da tela principal do aplicativo.*

## 🌟 Recursos Principais

* **Gerenciamento de Conexão Dupla:** Conecte-se e mantenha ativas duas conexões de banco de dados (V1 e V2) de forma independente.
* **Interface Tabulada (TTK):** Navegação limpa usando 6 abas (Conexão, Consulta e Campanha para cada ambiente).
* **Consultas Assíncronas (Campanha):** A funcionalidade "Consulta Campanha" utiliza **`threading`** e **`queue`** para processar listas de IDs em segundo plano. Isso garante que a interface do usuário (UI) **não congele** durante operações longas.
* **Segurança (Read-Only):** O script é estritamente focado em operações `SELECT`, impedindo alterações acidentais nos dados.
* **Exportação de Dados:** Exporte facilmente os resultados das consultas de campanha para arquivos **`.csv`**.
* **Persistência de Configuração:** Salva e carrega informações de conexão (servidor, banco de dados, usuário) no arquivo `sqltool.ini` para agilizar o uso diário.
* **Verificação de Rede:** Inclui uma verificação de `socket` para testar o acesso à porta 1433 antes de tentar a conexão, fornecendo feedback imediato sobre problemas de VPN ou firewall.

## 🛠️ Tecnologias Utilizadas

* Python 3
* Tkinter (ttk)
* pyodbc (para conectividade SQL Server)
* threading / queue (para operações assíncronas)
* configparser (para gerenciamento de `.ini`)

## 🚀 Como Executar

1.  Certifique-se de que está na raiz do repositório (`ApuraçãoWebPhoneWhatsApp`) e que as dependências do `requirements.txt` principal foram instaladas.
2.  Navegue até esta pasta:
    ```bash
    cd sql-query-tool
    ```
3.  Execute o script:
    ```bash
    python WA.py
    ```

**Importante:** Na primeira execução, um arquivo `sqltool.ini` será criado nesta pasta. Este arquivo contém credenciais e **já está sendo ignorado** pelo `.gitignore` da raiz do projeto.
