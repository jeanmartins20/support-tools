import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import pyodbc
from typing import Optional
import socket
import re
import csv
import threading
import queue
import configparser 
import os 

CONFIG_FILE = 'sqltool.ini'

class SQLQueryTool(tk.Tk):
    """
    Ferramenta de UI para consultar SQL Server, agora com duas
    conexões (V1 e V2) e abas de consulta dedicadas.
    """
    def __init__(self):
        super().__init__()
        self.title("Ferramenta de Consulta SQL (V1 & V2)")
        self.geometry("850x650") # Aumentei um pouco para as 6 abas

        # --- Conexão V2 (Antiga Conexão 1) ---
        self.conn_v2: Optional[pyodbc.Connection] = None
        self.cursor_v2: Optional[pyodbc.Cursor] = None
        
        # --- Conexão V1 (Antiga Conexão 2) ---
        self.conn_v1: Optional[pyodbc.Connection] = None
        self.cursor_v1: Optional[pyodbc.Cursor] = None

        # --- Widgets Conexão V2 ---
        self.server_entry_v2 = None
        self.auth_type_combo_v2 = None
        self.db_entry_v2 = None
        self.logon_label_v2 = None
        self.logon_entry_v2 = None
        self.pass_label_v2 = None
        self.pass_entry_v2 = None
        self.status_label_v2 = None
        
        # --- Widgets Conexão V1 ---
        self.server_entry_v1 = None
        self.auth_type_combo_v1 = None
        self.db_entry_v1 = None
        self.logon_label_v1 = None
        self.logon_entry_v1 = None
        self.pass_label_v1 = None
        self.pass_entry_v1 = None
        self.status_label_v1 = None
        
        # --- Widgets Consulta V2 (Geral) ---
        self.query_text_v2 = None
        self.results_tree_v2 = None

        # --- Widgets Consulta V1 (Geral) ---
        self.query_text_v1 = None
        self.results_tree_v1 = None
        self.execute_button_v1 = None
        
        # --- Widgets Campanha V2 ---
        self.campanha_ids_text_v2 = None
        self.campanha_results_tree_v2 = None
        self.campanha_execute_button_v2 = None
        self.campanha_export_button_v2 = None
        self.campanha_cancel_button_v2 = None
        self.campanha_status_label_v2 = None
        self.query_queue_v2 = queue.Queue()
        self.cancel_event_v2 = threading.Event()
        
        # --- Widgets Campanha V1 ---
        self.campanha_ids_text_v1 = None
        self.campanha_results_tree_v1 = None
        self.campanha_execute_button_v1 = None
        self.campanha_export_button_v1 = None
        self.campanha_cancel_button_v1 = None
        self.campanha_status_label_v1 = None
        self.query_queue_v1 = queue.Queue()
        self.cancel_event_v1 = threading.Event()
        
        self.config = configparser.ConfigParser()
        
        self._create_widgets()
        self._on_auth_type_change_v1()
        self._on_auth_type_change_v2()
        
        self._load_config()
        self._monitor_queues()

    def _create_widgets(self):
        """Cria os componentes da interface (abas, campos, botões)."""
        
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(pady=10, padx=10, fill="both", expand=True)

        # Lista de servidores
        server_list = [
            "spotterv1prd.database.windows.net",
            "spotterprdv2-leitura.database.windows.net",
            "spotterprdv2.database.windows.net",
            "queue-observer-prd.database.windows.net",
            "globalserviceprd.database.windows.net",
            "spotterprd-v2-leitura.database.windows.net"
        ]

        # --- ABA 1: Conexão V1 (Antiga conn_frame_2) ---
        self.conn_frame_v1 = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.conn_frame_v1, text="Conexão V1")

        ttk.Label(self.conn_frame_v1, text="Configurações de Conexão V1", 
                  font=("Arial", 14, "bold")).pack(pady=10)
        conn_grid_v1 = ttk.Frame(self.conn_frame_v1)
        conn_grid_v1.pack(pady=5, padx=5)

        ttk.Label(conn_grid_v1, text="Tipo de servidor:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.server_type_combo_v1 = ttk.Combobox(conn_grid_v1, width=38, state="readonly", 
                                              values=["Mecanismo de Banco de Dados"])
        self.server_type_combo_v1.grid(row=0, column=1, padx=5, pady=5)
        self.server_type_combo_v1.current(0)

        ttk.Label(conn_grid_v1, text="Nome do servidor:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.server_entry_v1 = ttk.Combobox(conn_grid_v1, width=38, values=server_list)
        self.server_entry_v1.grid(row=1, column=1, padx=5, pady=5)
        self.server_entry_v1.insert(0, "") # <-- CORRIGIDO

        ttk.Label(conn_grid_v1, text="Autenticação:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.auth_type_combo_v1 = ttk.Combobox(conn_grid_v1, width=38, state="readonly",
                                            values=["Autenticação do SQL Server", "Autenticação do Windows"])
        self.auth_type_combo_v1.grid(row=2, column=1, padx=5, pady=5)
        self.auth_type_combo_v1.current(0) 
        self.auth_type_combo_v1.bind("<<ComboboxSelected>>", self._on_auth_type_change_v1)

        ttk.Label(conn_grid_v1, text="Banco de Dados:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.db_entry_v1 = ttk.Entry(conn_grid_v1, width=40)
        self.db_entry_v1.grid(row=3, column=1, padx=5, pady=5)
        self.db_entry_v1.insert(0, "") # <-- CORRIGIDO

        self.logon_label_v1 = ttk.Label(conn_grid_v1, text="Logon:")
        self.logon_label_v1.grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.logon_entry_v1 = ttk.Entry(conn_grid_v1, width=40)
        self.logon_entry_v1.grid(row=4, column=1, padx=5, pady=5)
        self.logon_entry_v1.insert(0, "") # <-- CORRIGIDO

        self.pass_label_v1 = ttk.Label(conn_grid_v1, text="Senha:")
        self.pass_label_v1.grid(row=5, column=0, sticky="w", padx=5, pady=5)
        self.pass_entry_v1 = ttk.Entry(conn_grid_v1, width=40, show="*")
        self.pass_entry_v1.grid(row=5, column=1, padx=5, pady=5)
        
        self.connect_button_v1 = ttk.Button(self.conn_frame_v1, text="Conectar V1", command=self.connect_db_v1)
        self.connect_button_v1.pack(pady=20)

        self.status_label_v1 = ttk.Label(self.conn_frame_v1, text="Status: Desconectado", 
                                      font=("Arial", 10), foreground="red")
        self.status_label_v1.pack(pady=5)
        
        # --- ABA 2: Conexão V2 (Antiga conn_frame) ---
        self.conn_frame_v2 = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.conn_frame_v2, text="Conexão V2")

        ttk.Label(self.conn_frame_v2, text="Configurações de Conexão V2", 
                  font=("Arial", 14, "bold")).pack(pady=10)
        conn_grid_v2 = ttk.Frame(self.conn_frame_v2)
        conn_grid_v2.pack(pady=5, padx=5)

        ttk.Label(conn_grid_v2, text="Tipo de servidor:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.server_type_combo_v2 = ttk.Combobox(conn_grid_v2, width=38, state="readonly", 
                                              values=["Mecanismo de Banco de Dados"])
        self.server_type_combo_v2.grid(row=0, column=1, padx=5, pady=5)
        self.server_type_combo_v2.current(0)

        ttk.Label(conn_grid_v2, text="Nome do servidor:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.server_entry_v2 = ttk.Combobox(conn_grid_v2, width=38, values=server_list)
        self.server_entry_v2.grid(row=1, column=1, padx=5, pady=5)
        self.server_entry_v2.insert(0, "") # <-- CORRIGIDO

        ttk.Label(conn_grid_v2, text="Autenticação:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.auth_type_combo_v2 = ttk.Combobox(conn_grid_v2, width=38, state="readonly",
                                            values=["Autenticação do SQL Server", "Autenticação do Windows"])
        self.auth_type_combo_v2.grid(row=2, column=1, padx=5, pady=5)
        self.auth_type_combo_v2.current(0) 
        self.auth_type_combo_v2.bind("<<ComboboxSelected>>", self._on_auth_type_change_v2)

        ttk.Label(conn_grid_v2, text="Banco de Dados:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.db_entry_v2 = ttk.Entry(conn_grid_v2, width=40)
        self.db_entry_v2.grid(row=3, column=1, padx=5, pady=5)
        self.db_entry_v2.insert(0, "") # <-- CORRIGIDO

        self.logon_label_v2 = ttk.Label(conn_grid_v2, text="Logon:")
        self.logon_label_v2.grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.logon_entry_v2 = ttk.Entry(conn_grid_v2, width=40)
        self.logon_entry_v2.grid(row=4, column=1, padx=5, pady=5)
        self.logon_entry_v2.insert(0, "") # <-- CORRIGIDO

        self.pass_label_v2 = ttk.Label(conn_grid_v2, text="Senha:")
        self.pass_label_v2.grid(row=5, column=0, sticky="w", padx=5, pady=5)
        self.pass_entry_v2 = ttk.Entry(conn_grid_v2, width=40, show="*")
        self.pass_entry_v2.grid(row=5, column=1, padx=5, pady=5)
        
        self.connect_button_v2 = ttk.Button(self.conn_frame_v2, text="Conectar V2", command=self.connect_db_v2)
        self.connect_button_v2.pack(pady=20)

        self.status_label_v2 = ttk.Label(self.conn_frame_v2, text="Status: Desconectado", 
                                      font=("Arial", 10), foreground="red")
        self.status_label_v2.pack(pady=5)

        # --- ABA 3: Consulta V1 (NOVA) ---
        self.query_frame_v1 = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.query_frame_v1, text="Consulta V1")
        
        ttk.Label(self.query_frame_v1, text="Digite sua consulta SQL (SOMENTE SELECT) - Conexão V1:").pack(anchor="w")
        self.query_text_v1 = scrolledtext.ScrolledText(self.query_frame_v1, height=10, width=80, 
                                                    font=("Courier New", 10))
        self.query_text_v1.pack(pady=5, fill="x", expand=False)
        self.execute_button_v1 = ttk.Button(self.query_frame_v1, text="Executar V1 (F5)", 
                                         command=self.execute_query_v1)
        self.execute_button_v1.pack(pady=5)
        
        ttk.Label(self.query_frame_v1, text="Resultados V1:").pack(anchor="w", pady=(10, 0))
        tree_frame_v1 = ttk.Frame(self.query_frame_v1)
        tree_frame_v1.pack(fill="both", expand=True)
        self.tree_scroll_y_v1 = ttk.Scrollbar(tree_frame_v1, orient="vertical")
        self.tree_scroll_x_v1 = ttk.Scrollbar(tree_frame_v1, orient="horizontal")
        self.results_tree_v1 = ttk.Treeview(tree_frame_v1, 
                                         yscrollcommand=self.tree_scroll_y_v1.set,
                                         xscrollcommand=self.tree_scroll_x_v1.set)
        self.tree_scroll_y_v1.config(command=self.results_tree_v1.yview)
        self.tree_scroll_x_v1.config(command=self.results_tree_v1.xview)
        self.tree_scroll_y_v1.pack(side="right", fill="y")
        self.tree_scroll_x_v1.pack(side="bottom", fill="x")
        self.results_tree_v1.pack(fill="both", expand=True)

        # --- ABA 4: Consulta V2 (Antiga query_frame) ---
        self.query_frame_v2 = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.query_frame_v2, text="Consulta V2")
        
        ttk.Label(self.query_frame_v2, text="Digite sua consulta SQL (SOMENTE SELECT) - Conexão V2:").pack(anchor="w")
        self.query_text_v2 = scrolledtext.ScrolledText(self.query_frame_v2, height=10, width=80, 
                                                    font=("Courier New", 10))
        self.query_text_v2.pack(pady=5, fill="x", expand=False)
        self.execute_button_v2 = ttk.Button(self.query_frame_v2, text="Executar V2 (F6)", 
                                         command=self.execute_query_v2)
        self.execute_button_v2.pack(pady=5)
        
        ttk.Label(self.query_frame_v2, text="Resultados V2:").pack(anchor="w", pady=(10, 0))
        tree_frame_v2 = ttk.Frame(self.query_frame_v2)
        tree_frame_v2.pack(fill="both", expand=True)
        self.tree_scroll_y_v2 = ttk.Scrollbar(tree_frame_v2, orient="vertical")
        self.tree_scroll_x_v2 = ttk.Scrollbar(tree_frame_v2, orient="horizontal")
        self.results_tree_v2 = ttk.Treeview(tree_frame_v2, 
                                         yscrollcommand=self.tree_scroll_y_v2.set,
                                         xscrollcommand=self.tree_scroll_x_v2.set)
        self.tree_scroll_y_v2.config(command=self.results_tree_v2.yview)
        self.tree_scroll_x_v2.config(command=self.results_tree_v2.xview)
        self.tree_scroll_y_v2.pack(side="right", fill="y")
        self.tree_scroll_x_v2.pack(side="bottom", fill="x")
        self.results_tree_v2.pack(fill="both", expand=True)

        # --- ABA 5: Consulta Campanha V1 (Antiga campanha_frame_2) ---
        self.campanha_frame_v1 = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.campanha_frame_v1, text="Campanha V1")

        ttk.Label(self.campanha_frame_v1, text="Cole os IDs (CodigoCampanha) V1 separados por vírgula:").pack(anchor="w")
        self.campanha_ids_text_v1 = scrolledtext.ScrolledText(self.campanha_frame_v1, height=10, width=80,
                                                           font=("Courier New", 10))
        self.campanha_ids_text_v1.pack(pady=5, fill="x", expand=False)
        
        campanha_botoes_frame_v1 = ttk.Frame(self.campanha_frame_v1)
        campanha_botoes_frame_v1.pack(pady=5)

        self.campanha_execute_button_v1 = ttk.Button(campanha_botoes_frame_v1, text="Executar Consulta Campanha V1",
                                                  command=self.start_campanha_query_v1)
        self.campanha_execute_button_v1.pack(side="left", padx=5)

        self.campanha_export_button_v1 = ttk.Button(campanha_botoes_frame_v1, text="Exportar para CSV (Planilha)",
                                                 command=self.export_campanha_to_csv_v1)
        self.campanha_export_button_v1.pack(side="left", padx=5)
        
        self.campanha_cancel_button_v1 = ttk.Button(campanha_botoes_frame_v1, text="Cancelar",
                                                 command=self.cancel_campanha_query_v1, state="disabled")
        self.campanha_cancel_button_v1.pack(side="left", padx=5)
        
        self.campanha_status_label_v1 = ttk.Label(self.campanha_frame_v1, text="", 
                                               font=("Arial", 10), foreground="blue")
        self.campanha_status_label_v1.pack(anchor="w", pady=5)

        ttk.Label(self.campanha_frame_v1, text="Resultados da Campanha V1:").pack(anchor="w", pady=(10, 0))
        campanha_tree_frame_v1 = ttk.Frame(self.campanha_frame_v1)
        campanha_tree_frame_v1.pack(fill="both", expand=True)
        
        self.campanha_tree_scroll_y_v1 = ttk.Scrollbar(campanha_tree_frame_v1, orient="vertical")
        self.campanha_tree_scroll_x_v1 = ttk.Scrollbar(campanha_tree_frame_v1, orient="horizontal")
        self.campanha_results_tree_v1 = ttk.Treeview(campanha_tree_frame_v1,
                                                  yscrollcommand=self.campanha_tree_scroll_y_v1.set,
                                                  xscrollcommand=self.campanha_tree_scroll_x_v1.set)
        self.campanha_tree_scroll_y_v1.config(command=self.campanha_results_tree_v1.yview)
        self.campanha_tree_scroll_x_v1.config(command=self.campanha_results_tree_v1.xview)
        self.campanha_tree_scroll_y_v1.pack(side="right", fill="y")
        self.campanha_tree_scroll_x_v1.pack(side="bottom", fill="x")
        self.campanha_results_tree_v1.pack(fill="both", expand=True)

        # --- ABA 6: Consulta Campanha V2 (Antiga campanha_frame) ---
        self.campanha_frame_v2 = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.campanha_frame_v2, text="Campanha V2")

        ttk.Label(self.campanha_frame_v2, text="Cole os IDs (CodigoCampanha) V2 separados por vírgula:").pack(anchor="w")
        self.campanha_ids_text_v2 = scrolledtext.ScrolledText(self.campanha_frame_v2, height=10, width=80,
                                                           font=("Courier New", 10))
        self.campanha_ids_text_v2.pack(pady=5, fill="x", expand=False)
        
        campanha_botoes_frame_v2 = ttk.Frame(self.campanha_frame_v2)
        campanha_botoes_frame_v2.pack(pady=5)

        self.campanha_execute_button_v2 = ttk.Button(campanha_botoes_frame_v2, text="Executar Consulta Campanha V2",
                                                  command=self.start_campanha_query_v2)
        self.campanha_execute_button_v2.pack(side="left", padx=5)

        self.campanha_export_button_v2 = ttk.Button(campanha_botoes_frame_v2, text="Exportar para CSV (Planilha)",
                                                 command=self.export_campanha_to_csv_v2)
        self.campanha_export_button_v2.pack(side="left", padx=5)
        
        self.campanha_cancel_button_v2 = ttk.Button(campanha_botoes_frame_v2, text="Cancelar",
                                                 command=self.cancel_campanha_query_v2, state="disabled")
        self.campanha_cancel_button_v2.pack(side="left", padx=5)
        
        self.campanha_status_label_v2 = ttk.Label(self.campanha_frame_v2, text="", 
                                               font=("Arial", 10), foreground="blue")
        self.campanha_status_label_v2.pack(anchor="w", pady=5)

        ttk.Label(self.campanha_frame_v2, text="Resultados da Campanha V2:").pack(anchor="w", pady=(10, 0))
        campanha_tree_frame_v2 = ttk.Frame(self.campanha_frame_v2)
        campanha_tree_frame_v2.pack(fill="both", expand=True)
        
        self.campanha_tree_scroll_y_v2 = ttk.Scrollbar(campanha_tree_frame_v2, orient="vertical")
        self.campanha_tree_scroll_x_v2 = ttk.Scrollbar(campanha_tree_frame_v2, orient="horizontal")
        self.campanha_results_tree_v2 = ttk.Treeview(campanha_tree_frame_v2,
                                                  yscrollcommand=self.campanha_tree_scroll_y_v2.set,
                                                  xscrollcommand=self.campanha_tree_scroll_x_v2.set)
        self.campanha_tree_scroll_y_v2.config(command=self.campanha_results_tree_v2.yview)
        self.campanha_tree_scroll_x_v2.config(command=self.campanha_results_tree_v2.xview)
        self.campanha_tree_scroll_y_v2.pack(side="right", fill="y")
        self.campanha_tree_scroll_x_v2.pack(side="bottom", fill="x")
        self.campanha_results_tree_v2.pack(fill="both", expand=True)
        
        
        # Desabilita as abas de consulta por padrão
        # Índices: 0:ConnV1, 1:ConnV2, 2:QueryV1, 3:QueryV2, 4:CampanhaV1, 5:CampanhaV2
        self.notebook.tab(2, state="disabled")
        self.notebook.tab(3, state="disabled") 
        self.notebook.tab(4, state="disabled") 
        self.notebook.tab(5, state="disabled") 

        # Bind de teclas F5 e F6
        self.bind("<F5>", lambda event: self._handle_f5())
        self.bind("<F6>", lambda event: self._handle_f6())


    def _handle_f5(self):
        """Executa a consulta da aba que estiver selecionada (V1 ou V2)."""
        current_tab_index = self.notebook.index(self.notebook.select())
        if current_tab_index == 2: # Aba "Consulta V1"
            self.execute_query_v1()

    # ADICIONE ESTA NOVA FUNÇÃO ABAIXO
    def _handle_f6(self):
        """Executa a consulta da aba V2 se ela estiver selecionada."""
        current_tab_index = self.notebook.index(self.notebook.select())
        if current_tab_index == 3: # Aba "Consulta V2"
            self.execute_query_v2()

    def _on_auth_type_change_v2(self, event=None):
        """Controla os campos de auth da Conexão V2."""
        auth_type = self.auth_type_combo_v2.get()
        state = "disabled" if auth_type == "Autenticação do Windows" else "normal"
        
        self.logon_entry_v2.config(state=state)
        self.pass_entry_v2.config(state=state)
        self.logon_label_v2.config(state=state)
        self.pass_label_v2.config(state=state)
        if state == "disabled":
            self.logon_entry_v2.delete(0, tk.END)
            self.pass_entry_v2.delete(0, tk.END)

    def _on_auth_type_change_v1(self, event=None):
        """Controla os campos de auth da Conexão V1."""
        auth_type = self.auth_type_combo_v1.get()
        state = "disabled" if auth_type == "Autenticação do Windows" else "normal"
        
        self.logon_entry_v1.config(state=state)
        self.pass_entry_v1.config(state=state)
        self.logon_label_v1.config(state=state)
        self.pass_label_v1.config(state=state)
        if state == "disabled":
            self.logon_entry_v1.delete(0, tk.END)
            self.pass_entry_v1.delete(0, tk.END)

    def _check_network_access(self, server: str, port: int, timeout: int = 3) -> bool:
        if server.startswith("tcp:"):
            server = server[4:]
        if ',' in server:
            server = server.split(',')[0]
            
        try:
            ip = socket.gethostbyname(server)
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(timeout)
                s.connect((ip, port))
            return True
        except socket.timeout:
            print(f"DEBUG: Timeout ao tentar conectar a {server}:{port}")
            return False
        except (socket.gaierror, socket.herror):
            print(f"DEBUG: Erro de DNS ao resolver {server}")
            return False
        except Exception as e:
            if "Connection refused" in str(e):
                print(f"DEBUG: Conexão recusada por {server}:{port} (Servidor está online)")
                return True
            print(f"DEBUG: Erro de socket inesperado: {e}")
            return False

    def connect_db_v2(self):
        """Conecta ao banco de dados V2 (Aba 2)."""
        server = self.server_entry_v2.get()
        database = self.db_entry_v2.get()
        auth_type = self.auth_type_combo_v2.get()
        
        if not all([server, database]):
            messagebox.showwarning("Campos Vazios", "Servidor e Banco de Dados são obrigatórios.")
            return

        self.status_label_v2.config(text="Status: Verificando rede (VPN)...", foreground="orange")
        self.update_idletasks() 

        if not self._check_network_access(server, 1433):
            self.status_label_v2.config(text="Status: Desconectado", foreground="red")
            messagebox.showerror("Erro de Rede", f"Não foi possível acessar o servidor: {server}.\n\nVerifique sua conexão VPN.")
            return
        
        self.status_label_v2.config(text="Status: Conectando ao banco de dados...", foreground="orange")
        self.update_idletasks()
        
        self._close_connection_v2()
        
        conn_string = ""
        password = "" # Definido aqui para uso no print de debug
        try:
            server_to_use = server
            if server.endswith(".database.windows.net") and not server.startswith('tcp:'):
                server_to_use = f"tcp:{server}"

            auth_part = ""
            if auth_type == "Autenticação do SQL Server":
                username = self.logon_entry_v2.get()
                password = self.pass_entry_v2.get()
                if not all([username, password]):
                    messagebox.showwarning("Campos Vazios", "Logon e Senha são obrigatórios.")
                    self.status_label_v2.config(text="Status: Desconectado", foreground="red")
                    return
                auth_part = f"UID={username};PWD={password};"
            elif auth_type == "Autenticação do Windows":
                auth_part = "Trusted_Connection=yes;"
            else:
                raise ValueError("Tipo de autenticação inválido.")

            conn_string = (
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={server_to_use};"
                f"DATABASE={database};"
                f"{auth_part}"
                f"Encrypt=yes;"
                f"TrustServerCertificate=yes;" # <--- CORRIGIDO
            )
            
            # --- DEBUGGING V2 ---
            print("--- [DEBUG Conexão V2] ---")
            print_string = conn_string.replace(password, "********") if password else conn_string
            print(f"String: {print_string}")
            print("---------------------------")
            # --- FIM DEBUG ---

            self.conn_v2 = pyodbc.connect(conn_string, timeout=5)
            self.cursor_v2 = self.conn_v2.cursor()

            self.status_label_v2.config(text="Status: Conectado", foreground="green")
            self.notebook.tab(3, state="normal") # Habilita Consulta V2 (Aba 4)
            self.notebook.tab(5, state="normal") # Habilita Campanha V2 (Aba 6)
            self.notebook.select(3) # Seleciona Aba "Consulta V2"
            
            self._save_config()
            messagebox.showinfo("Sucesso", "Conexão V2 estabelecida com sucesso!")

        except pyodbc.Error as e:
            self.conn_v2 = None
            self.cursor_v2 = None
            self.status_label_v2.config(text="Status: Desconectado", foreground="red")
            messagebox.showerror("Erro de Conexão ODBC (V2)", f"O driver ODBC reportou um erro:\n\n{e}")
        except Exception as e:
            self.conn_v2 = None
            self.cursor_v2 = None
            self.status_label_v2.config(text="Status: Desconectado", foreground="red")
            messagebox.showerror("Erro Inesperado (V2)", f"Ocorreu um erro na lógica de conexão:\n{e}")
            
    def connect_db_v1(self):
        """Conecta ao banco de dados V1 (Aba 1)."""
        server = self.server_entry_v1.get()
        database = self.db_entry_v1.get()
        auth_type = self.auth_type_combo_v1.get()
        
        if not all([server, database]):
            messagebox.showwarning("Campos Vazios", "Servidor e Banco de Dados são obrigatórios.")
            return

        self.status_label_v1.config(text="Status: Verificando rede (VPN)...", foreground="orange")
        self.update_idletasks() 

        if not self._check_network_access(server, 1433):
            self.status_label_v1.config(text="Status: Desconectado", foreground="red")
            messagebox.showerror("Erro de Rede", f"Não foi possível acessar o servidor: {server}.\n\nVerifique sua conexão VPN.")
            return
        
        self.status_label_v1.config(text="Status: Conectando ao banco de dados...", foreground="orange")
        self.update_idletasks()
        
        self._close_connection_v1()

        conn_string = ""
        password = "" # Definido aqui para uso no print de debug
        try:
            server_to_use = server
            if server.endswith(".database.windows.net") and not server.startswith('tcp:'):
                server_to_use = f"tcp:{server}"

            auth_part = ""
            if auth_type == "Autenticação do SQL Server":
                username = self.logon_entry_v1.get()
                password = self.pass_entry_v1.get()
                if not all([username, password]):
                    messagebox.showwarning("Campos Vazios", "Logon e Senha são obrigatórios.")
                    self.status_label_v1.config(text="Status: Desconectado", foreground="red")
                    return
                auth_part = f"UID={username};PWD={password};"
            elif auth_type == "Autenticação do Windows":
                auth_part = "Trusted_Connection=yes;"
            else:
                raise ValueError("Tipo de autenticação inválido.")

            conn_string = (
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={server_to_use};"
                f"DATABASE={database};"
                f"{auth_part}"
                f"Encrypt=yes;"
                f"TrustServerCertificate=no;"
            )
            
            # --- DEBUGGING V1 ---
            print("--- [DEBUG Conexão V1] ---")
            print_string = conn_string.replace(password, "********") if password else conn_string
            print(f"String: {print_string}")
            print("---------------------------")
            # --- FIM DEBUG ---

            self.conn_v1 = pyodbc.connect(conn_string, timeout=5)
            self.cursor_v1 = self.conn_v1.cursor()

            self.status_label_v1.config(text="Status: Conectado", foreground="green")
            self.notebook.tab(2, state="normal") # Habilita Consulta V1 (Aba 3)
            self.notebook.tab(4, state="normal") # Habilita Campanha V1 (Aba 5)
            self.notebook.select(2) # Seleciona Aba "Consulta V1"
            
            self._save_config()
            messagebox.showinfo("Sucesso", "Conexão V1 estabelecida com sucesso!")

        except pyodbc.Error as e: # Captura o erro ODBC especificamente
            self.conn_v1 = None
            self.cursor_v1 = None
            self.status_label_v1.config(text="Status: Desconectado", foreground="red")
            # Esta é a mensagem que eu preciso
            messagebox.showerror("Erro de Conexão ODBC (V1)", f"O driver ODBC reportou um erro:\n\n{e}")
        except Exception as e: # Captura outros erros (ex: ValueError do auth_type)
            self.conn_v1 = None
            self.cursor_v1 = None
            self.status_label_v1.config(text="Status: Desconectado", foreground="red")
            messagebox.showerror("Erro Inesperado (V1)", f"Ocorreu um erro na lógica de conexão:\n\n{e}")

    def execute_query_v2(self):
        """Executa a consulta SQL (SOMENTE SELECT) na Conexão V2 (Aba 4)."""
        query = self.query_text_v2.get("1.0", tk.END).strip()

        if not self.conn_v2 or not self.cursor_v2:
            messagebox.showerror("Sem Conexão", "Por favor, conecte-se (Conexão V2) primeiro.")
            return
        if not query:
            messagebox.showwarning("Consulta Vazia", "Por favor, digite uma consulta SQL.")
            return

        self._clear_results_tree(self.results_tree_v2)

        try:
            if not query.lstrip().upper().startswith("SELECT"):
                self.conn_v2.rollback() 
                messagebox.showwarning("Operação Não Permitida", "Permitido apenas consultas (SELECT).")
                return

            self.cursor_v2.execute(query)
            columns = [col[0] for col in self.cursor_v2.description]
            self._setup_tree_columns(self.results_tree_v2, columns)

            rows = self.cursor_v2.fetchall()
            for i, row in enumerate(rows):
                self.results_tree_v2.insert(parent="", index="end", iid=i, 
                                            text="", values=tuple(row))
        except pyodbc.Error as e:
            messagebox.showerror("Erro na Consulta V2", f"Erro:\n{e}")
            try: self.conn_v2.rollback()
            except pyodbc.Error: pass
        except Exception as e:
            messagebox.showerror("Erro Inesperado", f"Ocorreu um erro: {e}")

    def execute_query_v1(self):
        """Executa a consulta SQL (SOMENTE SELECT) na Conexão V1 (Aba 3)."""
        query = self.query_text_v1.get("1.0", tk.END).strip()

        if not self.conn_v1 or not self.cursor_v1:
            messagebox.showerror("Sem Conexão", "Por favor, conecte-se (Conexão V1) primeiro.")
            return
        if not query:
            messagebox.showwarning("Consulta Vazia", "Por favor, digite uma consulta SQL.")
            return

        self._clear_results_tree(self.results_tree_v1)

        try:
            if not query.lstrip().upper().startswith("SELECT"):
                self.conn_v1.rollback() 
                messagebox.showwarning("Operação Não Permitida", "Permitido apenas consultas (SELECT).")
                return

            self.cursor_v1.execute(query)
            columns = [col[0] for col in self.cursor_v1.description]
            self._setup_tree_columns(self.results_tree_v1, columns)

            rows = self.cursor_v1.fetchall()
            for i, row in enumerate(rows):
                self.results_tree_v1.insert(parent="", index="end", iid=i, 
                                            text="", values=tuple(row))
        except pyodbc.Error as e:
            messagebox.showerror("Erro na Consulta V1", f"Erro:\n{e}")
            try: self.conn_v1.rollback()
            except pyodbc.Error: pass
        except Exception as e:
            messagebox.showerror("Erro Inesperado", f"Ocorreu um erro: {e}")

    def start_campanha_query_v2(self):
        """Inicia a consulta de campanha V2 (Aba 6) em uma thread."""
        if not self.conn_v2 or not self.cursor_v2:
            messagebox.showerror("Sem Conexão", "Por favor, conecte-se (Conexão V2) primeiro.")
            return

        id_string_raw = self.campanha_ids_text_v2.get("1.0", tk.END).strip()
        if not id_string_raw:
            messagebox.showwarning("Entrada Vazia", "Por favor, cole os IDs CodigoCampanha.")
            return

        id_list = re.findall(r"'([\d\w]+)'", id_string_raw)
        if not id_list:
            messagebox.showwarning("Formato Inválido", "Formato esperado (ex: 'id1','id2').")
            return
        
        self._clear_results_tree(self.campanha_results_tree_v2)
        self.cancel_event_v2.clear()
        
        self.campanha_execute_button_v2.config(state="disabled")
        self.campanha_export_button_v2.config(state="disabled")
        self.campanha_cancel_button_v2.config(state="normal")
        
        thread = threading.Thread(target=self._run_campanha_thread_v2, args=(id_list,))
        thread.daemon = True
        thread.start()

    def start_campanha_query_v1(self):
        """Inicia a consulta de campanha V1 (Aba 5) em uma thread."""
        if not self.conn_v1 or not self.cursor_v1:
            messagebox.showerror("Sem Conexão", "Por favor, conecte-se (Conexão V1) primeiro.")
            return

        id_string_raw = self.campanha_ids_text_v1.get("1.0", tk.END).strip()
        if not id_string_raw:
            messagebox.showwarning("Entrada Vazia", "Por favor, cole os IDs CodigoCampanha.")
            return

        id_list = re.findall(r"'([\d\w]+)'", id_string_raw)
        if not id_list:
            messagebox.showwarning("Formato Inválido", "Formato esperado (ex: 'id1','id2').")
            return
        
        self._clear_results_tree(self.campanha_results_tree_v1)
        self.cancel_event_v1.clear()
        
        self.campanha_execute_button_v1.config(state="disabled")
        self.campanha_export_button_v1.config(state="disabled")
        self.campanha_cancel_button_v1.config(state="normal")
        
        thread = threading.Thread(target=self._run_campanha_thread_v1, args=(id_list,))
        thread.daemon = True
        thread.start()

    def _run_campanha_thread_v2(self, id_list: list):
        """ Thread para consultar IDs (V2) um a um. """
        
        # Query V2 (Original)
        base_query = """
        select distinct 
        e.id, e.nome, e.id1, 
        CASE 
            WHEN p.funcionalidadeid = 15 THEN Quantidade
        END AS 'Mensagem',
        (select Quantidade
        from funcionalidadeplanoempresa f
        join funcionalidadeplano p on p.id = f.funcionalidadeplanoid
        where f.empresaclienteid = i.EmpresaClienteId
        and f.funcionalidadeid in (16)
        and f.[Situacao] = 1) as 'HSM',
        'Status' = 'Ativo',
        FORMAT(DATEADD(HOUR, -3, f.CreatedAt), 'MM/yyyy') AS 'Data alteração',
        i.CodigoCampanha, 
        FlBloqueio,
        FORMAT(DATEADD(HOUR, -3, e.UpdatedAt), 'MM/yyyy') AS 'Data atualização conta'
        from IntegracaoMensagens i
        join EmpresaCliente e on e.Id = i.EmpresaClienteId
        left join FuncionalidadePlanoEmpresa f on f.EmpresaClienteId = i.EmpresaClienteId and f.FuncionalidadeId = 15 and f.Situacao = 1
        left join funcionalidadeplano p on p.id = f.funcionalidadeplanoid
        where CodigoCampanha = ?
        """
        
        try:
            # --- Cria uma NOVA conexão V2 para esta thread ---
            server = self.server_entry_v2.get()
            database = self.db_entry_v2.get()
            auth_type = self.auth_type_combo_v2.get()
            server_to_use = f"tcp:{server}" if server.endswith(".windows.net") and not server.startswith('tcp:') else server
            
            auth_part = ""
            if auth_type == "Autenticação do SQL Server":
                auth_part = f"UID={self.logon_entry_v2.get()};PWD={self.pass_entry_v2.get()};"
            elif auth_type == "Autenticação do Windows":
                auth_part = "Trusted_Connection=yes;"
            
            conn_string = (f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server_to_use};DATABASE={database};"
                           f"{auth_part}Encrypt=yes;TrustServerCertificate=yes;")
            
            with pyodbc.connect(conn_string, timeout=5) as thread_conn:
                with thread_conn.cursor() as thread_cursor:
                    columns_sent = False
                    for i, id_campanha in enumerate(id_list):
                        if self.cancel_event_v2.is_set():
                            self.query_queue_v2.put(("status", "Consulta V2 cancelada."))
                            break
                        
                        self.query_queue_v2.put(("status", f"Consultando V2 {i+1}/{len(id_list)}: {id_campanha}"))
                        
                        try:
                            thread_cursor.execute(base_query, (id_campanha,))
                            if not columns_sent and thread_cursor.description:
                                columns = [col[0] for col in thread_cursor.description]
                                self.query_queue_v2.put(("columns", columns))
                                columns_sent = True
                            row = thread_cursor.fetchone()
                            if row:
                                self.query_queue_v2.put(("data", tuple(row)))
                        except pyodbc.Error as e:
                            self.query_queue_v2.put(("status", f"Erro V2 {id_campanha}: {e}"))
                    
                    if not self.cancel_event_v2.is_set():
                        self.query_queue_v2.put(("done", "Consulta V2 concluída."))
                        
        except pyodbc.Error as e:
            self.query_queue_v2.put(("error", f"Erro de conexão Thread V2: {e}"))
        except Exception as e:
            self.query_queue_v2.put(("error", f"Erro inesperado Thread V2: {e}"))

    def _run_campanha_thread_v1(self, id_list: list):
        """ Thread para consultar IDs (V1) um a um. """
        
        # Query V1 (Nova query que você forneceu)
        base_query = """
        select distinct 
        e.id, e.nome, e.id1, 
        CASE 
            WHEN p.funcionalidadeid = 15 THEN Quantidade
        END AS 'Mensagem',
        (select Quantidade
        from funcionalidadeplanoempresa f
        join funcionalidadeplano p on p.id = f.funcionalidadeplanoid
        where f.empresaclienteid = i.EmpresaClienteId
        and f.funcionalidadeid in (16)) as 'HSM',
        'Status' = 'Ativo',
        FORMAT(DATEADD(HOUR, -3, f.CreatedAt), 'MM/yyyy') AS 'Data alteração',
        i.CodigoCampanha, 
        FlBloqueio,
        FORMAT(DATEADD(HOUR, -3, e.UpdatedAt), 'MM/yyyy') AS 'Data atualização conta'
        from IntegracaoMensagens i
        join EmpresaCliente e on e.Id = i.EmpresaClienteId
        left join FuncionalidadePlanoEmpresa f on f.EmpresaClienteId = i.EmpresaClienteId and f.FuncionalidadeId = 15
        left join funcionalidadeplano p on p.id = f.funcionalidadeplanoid
        where e.ReadOnly = 0 and CodigoCampanha = ?
        """
        
        try:
            # --- Cria uma NOVA conexão V1 para esta thread ---
            server = self.server_entry_v1.get()
            database = self.db_entry_v1.get()
            auth_type = self.auth_type_combo_v1.get()
            server_to_use = f"tcp:{server}" if server.endswith(".windows.net") and not server.startswith('tcp:') else server

            auth_part = ""
            if auth_type == "Autenticação do SQL Server":
                auth_part = f"UID={self.logon_entry_v1.get()};PWD={self.pass_entry_v1.get()};"
            elif auth_type == "Autenticação do Windows":
                auth_part = "Trusted_Connection=yes;"
            
            conn_string = (f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server_to_use};DATABASE={database};"
                           f"{auth_part}Encrypt=yes;TrustServerCertificate=no;") # <--- CORRIGIDO
            
            with pyodbc.connect(conn_string, timeout=5) as thread_conn:
                with thread_conn.cursor() as thread_cursor:
                    columns_sent = False
                    for i, id_campanha in enumerate(id_list):
                        if self.cancel_event_v1.is_set():
                            self.query_queue_v1.put(("status", "Consulta V1 cancelada."))
                            break
                        
                        self.query_queue_v1.put(("status", f"Consultando V1 {i+1}/{len(id_list)}: {id_campanha}"))
                        
                        try:
                            thread_cursor.execute(base_query, (id_campanha,))
                            if not columns_sent and thread_cursor.description:
                                columns = [col[0] for col in thread_cursor.description]
                                self.query_queue_v1.put(("columns", columns))
                                columns_sent = True
                            row = thread_cursor.fetchone()
                            if row:
                                self.query_queue_v1.put(("data", tuple(row)))
                        except pyodbc.Error as e:
                            self.query_queue_v1.put(("status", f"Erro V1 {id_campanha}: {e}"))
                    
                    if not self.cancel_event_v1.is_set():
                        self.query_queue_v1.put(("done", "Consulta V1 concluída."))
                        
        except pyodbc.Error as e:
            self.query_queue_v1.put(("error", f"Erro de conexão Thread V1: {e}"))
        except Exception as e:
            self.query_queue_v1.put(("error", f"Erro inesperado Thread V1: {e}"))

    def cancel_campanha_query_v2(self):
        self.cancel_event_v2.set()
        self.campanha_status_label_v2.config(text="Cancelando V2...", foreground="orange")

    def cancel_campanha_query_v1(self):
        self.cancel_event_v1.set()
        self.campanha_status_label_v1.config(text="Cancelando V1...", foreground="orange")

    def _reset_campanha_buttons_v2(self):
        self.campanha_execute_button_v2.config(state="normal")
        self.campanha_export_button_v2.config(state="normal")
        self.campanha_cancel_button_v2.config(state="disabled")
        self.cancel_event_v2.clear()

    def _reset_campanha_buttons_v1(self):
        self.campanha_execute_button_v1.config(state="normal")
        self.campanha_export_button_v1.config(state="normal")
        self.campanha_cancel_button_v1.config(state="disabled")
        self.cancel_event_v1.clear()

    def _monitor_queues(self):
        """Verifica ambas as filas por mensagens da thread e atualiza a UI."""
        
        # --- Bloco 1: Processa Fila V2 (Campanha V2) ---
        try:
            while True:
                msg = self.query_queue_v2.get_nowait()
                msg_type, msg_data = msg

                if msg_type == "status":
                    self.campanha_status_label_v2.config(text=msg_data, foreground="blue")
                elif msg_type == "columns":
                    self._setup_tree_columns(self.campanha_results_tree_v2, msg_data)
                elif msg_type == "data":
                    self.campanha_results_tree_v2.insert(parent="", index="end", values=msg_data)
                elif msg_type == "done":
                    self.campanha_status_label_v2.config(text=msg_data, foreground="green")
                    self._reset_campanha_buttons_v2()
                elif msg_type == "error":
                    self.campanha_status_label_v2.config(text=msg_data, foreground="red")
                    messagebox.showerror("Erro na Thread V2", msg_data)
                    self._reset_campanha_buttons_v2()
        except queue.Empty:
            pass
        
        # --- Bloco 2: Processa Fila V1 (Campanha V1) ---
        try:
            while True:
                msg = self.query_queue_v1.get_nowait()
                msg_type, msg_data = msg

                if msg_type == "status":
                    self.campanha_status_label_v1.config(text=msg_data, foreground="blue")
                elif msg_type == "columns":
                    self._setup_tree_columns(self.campanha_results_tree_v1, msg_data)
                elif msg_type == "data":
                    self.campanha_results_tree_v1.insert(parent="", index="end", values=msg_data)
                elif msg_type == "done":
                    self.campanha_status_label_v1.config(text=msg_data, foreground="green")
                    self._reset_campanha_buttons_v1() 
                elif msg_type == "error":
                    self.campanha_status_label_v1.config(text=msg_data, foreground="red")
                    messagebox.showerror("Erro na Thread V1", msg_data)
                    self._reset_campanha_buttons_v1()
        except queue.Empty:
            pass
        
        self.after(100, self._monitor_queues)

    def export_campanha_to_csv_v2(self):
        """Exporta os dados da tabela da Campanha V2 para um arquivo CSV."""
        self._export_tree_to_csv(self.campanha_results_tree_v2, "Salvar resultados da campanha V2")

    def export_campanha_to_csv_v1(self):
        """Exporta os dados da tabela da Campanha V1 para um arquivo CSV."""
        self._export_tree_to_csv(self.campanha_results_tree_v1, "Salvar resultados da campanha V1")

    def _export_tree_to_csv(self, tree_widget: ttk.Treeview, title: str):
        """Função auxiliar para exportar dados de um Treeview para CSV."""
        if not tree_widget.get_children():
            messagebox.showwarning("Nada para Exportar", "Primeiro, execute uma consulta para gerar resultados.")
            return

        try:
            filepath = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")],
                title=title
            )
            if not filepath:
                return 

            with open(filepath, mode='w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';') 
                columns = tree_widget['columns']
                writer.writerow(columns)
                for item_id in tree_widget.get_children():
                    row_values = tree_widget.item(item_id)['values']
                    writer.writerow(row_values)
            messagebox.showinfo("Exportação Concluída", f"Resultados salvos com sucesso em:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Erro na Exportação", f"Não foi possível salvar o arquivo:\n{e}")

    def _clear_results_tree(self, tree_widget: ttk.Treeview):
        """Limpa todos os dados do Treeview de resultados fornecido."""
        for item in tree_widget.get_children():
            tree_widget.delete(item)
        tree_widget["columns"] = ()
        
    def _setup_tree_columns(self, tree_widget: ttk.Treeview, columns: list):
        """Configura as colunas de um Treeview."""
        tree_widget["columns"] = columns
        tree_widget.column("#0", width=0, stretch=tk.NO)
        tree_widget.heading("#0", text="", anchor="center")
        for col in columns:
            tree_widget.column(col, anchor="w", width=120)
            tree_widget.heading(col, text=col, anchor="w")

    def _close_connection_v2(self):
        """Fecha a Conexão V2."""
        if self.cursor_v2: self.cursor_v2.close()
        if self.conn_v2: self.conn_v2.close()
        self.conn_v2, self.cursor_v2 = None, None
        self.status_label_v2.config(text="Status: Desconectado", foreground="red")
        self.notebook.tab(3, state="disabled") # Desabilita Consulta V2
        self.notebook.tab(5, state="disabled") # Desabilita Campanha V2

    def _close_connection_v1(self):
        """Fecha a Conexão V1."""
        if self.cursor_v1: self.cursor_v1.close()
        if self.conn_v1: self.conn_v1.close()
        self.conn_v1, self.cursor_v1 = None, None
        self.status_label_v1.config(text="Status: Desconectado", foreground="red")
        self.notebook.tab(2, state="disabled") # Desabilita Consulta V1
        self.notebook.tab(4, state="disabled") # Desabilita Campanha V1
        
    def _save_config(self):
        """Salva as configurações de ambas as conexões no arquivo .ini"""
        try:
            if not self.config.has_section('ConnectionV1'):
                self.config.add_section('ConnectionV1')
            self.config.set('ConnectionV1', 'server', self.server_entry_v1.get())
            self.config.set('ConnectionV1', 'database', self.db_entry_v1.get())
            self.config.set('ConnectionV1', 'logon', self.logon_entry_v1.get())
            self.config.set('ConnectionV1', 'password', self.pass_entry_v1.get()) 
            
            if not self.config.has_section('ConnectionV2'):
                self.config.add_section('ConnectionV2')
            self.config.set('ConnectionV2', 'server', self.server_entry_v2.get())
            self.config.set('ConnectionV2', 'database', self.db_entry_v2.get())
            self.config.set('ConnectionV2', 'logon', self.logon_entry_v2.get())
            self.config.set('ConnectionV2', 'password', self.pass_entry_v2.get())

            with open(CONFIG_FILE, 'w') as configfile:
                self.config.write(configfile)
        except Exception as e:
            print(f"Erro ao salvar config: {e}")

    def _load_config(self):
        """Carrega as configurações do arquivo .ini ao iniciar"""
        if not os.path.exists(CONFIG_FILE):
            return
            
        try:
            self.config.read(CONFIG_FILE)
            
            if 'ConnectionV1' in self.config:
                cfg = self.config['ConnectionV1']
                self.server_entry_v1.delete(0, tk.END)
                self.server_entry_v1.insert(0, cfg.get('server', '')) # <-- CORRIGIDO
                self.db_entry_v1.delete(0, tk.END)
                self.db_entry_v1.insert(0, cfg.get('database', '')) # <-- CORRIGIDO
                self.logon_entry_v1.delete(0, tk.END)
                self.logon_entry_v1.insert(0, cfg.get('logon', '')) # <-- CORRIGIDO
                self.pass_entry_v1.delete(0, tk.END)
                self.pass_entry_v1.insert(0, cfg.get('password', ''))

            if 'ConnectionV2' in self.config:
                cfg = self.config['ConnectionV2']
                self.server_entry_v2.delete(0, tk.END)
                self.server_entry_v2.insert(0, cfg.get('server', '')) # <-- CORRIGIDO
                self.db_entry_v2.delete(0, tk.END)
                self.db_entry_v2.insert(0, cfg.get('database', '')) # <-- CORRIGIDO
                self.logon_entry_v2.delete(0, tk.END)
                self.logon_entry_v2.insert(0, cfg.get('logon', '')) # <-- CORRIGIDO
                self.pass_entry_v2.delete(0, tk.END)
                self.pass_entry_v2.insert(0, cfg.get('password', ''))
                
        except Exception as e:
            print(f"Erro ao carregar config: {e}")

    def on_closing(self):
        """Chamado quando a janela é fechada."""
        if messagebox.askokcancel("Sair", "Deseja fechar a aplicação?"):
            self.cancel_event_v1.set()
            self.cancel_event_v2.set()
            self._close_connection_v1()
            self._close_connection_v2()
            self.destroy()

if __name__ == "__main__":
    try:
        import pyodbc
    except ImportError:
        print("Erro: A biblioteca 'pyodbc' não está instalada.")
        print("Por favor, instale-a usando o comando: pip install pyodbc")
        exit()

    app = SQLQueryTool()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()
