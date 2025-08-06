import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import threading
import csv
import os
import sys
import json
import io
import shutil
from datetime import datetime
from urllib import request, error
from PIL import Image, ImageTk


# --- FUNÇÕES DE APOIO ---

def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = get_base_path()
    return os.path.join(base_path, relative_path)


def carregar_dados_csv(nome_arquivo):
    external_path = os.path.join(get_base_path(), nome_arquivo)
    internal_path = resource_path(nome_arquivo)
    load_path = external_path if os.path.exists(external_path) else internal_path

    if not os.path.exists(load_path):
        return None, None

    # Tenta ler com as codificações mais comuns
    for encoding in ['utf-8-sig', 'latin-1', 'windows-1252']:
        try:
            with open(load_path, mode='r', encoding=encoding) as infile:
                reader = csv.DictReader(infile, delimiter=';')
                # CORREÇÃO: Padroniza os cabeçalhos para minúsculas e remove espaços
                reader.fieldnames = [header.strip().lower().replace(" ", "-") for header in reader.fieldnames]
                data = [row for row in reader]

            mod_time = os.path.getmtime(load_path)
            last_updated = datetime.fromtimestamp(mod_time).strftime('%d/%m/%Y às %H:%M:%S')
            print(f"Arquivo '{nome_arquivo}' lido com sucesso usando a codificação: {encoding}")
            return data, last_updated
        except UnicodeDecodeError:
            continue  # Tenta a próxima codificação
        except Exception as e:
            messagebox.showerror("Erro de Leitura", f"Ocorreu um erro ao ler o arquivo '{nome_arquivo}':\n{e}")
            return [], None

    # Se todas as tentativas falharem
    messagebox.showerror("Erro de Codificação",
                         f"Não foi possível decodificar o arquivo '{nome_arquivo}'.\nPor favor, abra o arquivo no Excel ou Bloco de Notas e salve-o com a codificação 'UTF-8'.")
    return [], None


def salvar_dados_csv(nome_arquivo, data, headers):
    caminho_salvar = os.path.join(get_base_path(), nome_arquivo)
    try:
        with open(caminho_salvar, mode='w', encoding='utf-8-sig', newline='') as outfile:
            writer = csv.DictWriter(outfile, fieldnames=headers, delimiter=';')
            writer.writeheader()
            writer.writerows(data)
        return True
    except Exception as e:
        messagebox.showerror("Erro ao Salvar",
                             f"Não foi possível salvar as alterações em:\n{caminho_salvar}\n\nErro: {e}")
        return False


def carregar_configuracoes():
    try:
        with open(os.path.join(get_base_path(), 'settings.json'), 'r') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {"ultima_cidade": "Todas", "ultimo_regime": "Simples Nacional", "ultima_organizacao": "Por Nome (A-Z)"}


def salvar_configuracoes(cidade, regime, organizacao):
    config = {"ultima_cidade": cidade, "ultimo_regime": regime, "ultima_organizacao": organizacao}
    with open(os.path.join(get_base_path(), 'settings.json'), 'w') as f:
        json.dump(config, f)


# --- CARREGAMENTO INICIAL ---
EMPRESAS, LAST_UPDATED_EMPRESAS = carregar_dados_csv('empresas.csv')
CIDADES, _ = carregar_dados_csv('cidades.csv')
CONFIGURACOES_INICIAIS = carregar_configuracoes()

# Define os cabeçalhos corretos para salvar, correspondendo aos seus arquivos
HEADERS_EMPRESAS = ["id", "empresa", "cnpj", "login", "senha", "cidade", "regime"]
HEADERS_CIDADES = ["cidade", "url", "cnpj-selector", "login-selector", "senha_seletor"]


class CityForm(ttk.Toplevel):
    """ Janela para Adicionar ou Editar uma Cidade. """

    def __init__(self, parent, cidade_data=None, callback=None):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()

        self.cidade_data = cidade_data
        self.callback = callback
        self.title("Editar Cidade" if cidade_data else "Adicionar Nova Cidade")

        self.entries = {}

        form_frame = ttk.Frame(self, padding="15")
        form_frame.pack(fill=BOTH, expand=YES)

        for i, field in enumerate(HEADERS_CIDADES):
            ttk.Label(form_frame, text=f"{field.replace('_', ' ').title()}:").grid(row=i, column=0, sticky="w", padx=5,
                                                                                   pady=5)
            entry = ttk.Entry(form_frame, width=50)
            if cidade_data:
                entry.insert(0, cidade_data.get(field, ''))
            entry.grid(row=i, column=1, sticky="ew", padx=5, pady=5)
            self.entries[field] = entry

        form_frame.columnconfigure(1, weight=1)

        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=len(HEADERS_CIDADES), column=0, columnspan=2, pady=10)

        save_btn = ttk.Button(button_frame, text="Salvar", command=self.save, bootstyle="success")
        save_btn.pack(side=LEFT, padx=5)
        cancel_btn = ttk.Button(button_frame, text="Cancelar", command=self.destroy, bootstyle="danger")
        cancel_btn.pack(side=LEFT, padx=5)

    def save(self):
        new_data = {field: entry.get() for field, entry in self.entries.items()}

        if not new_data.get("cidade"):
            messagebox.showwarning("Campo Obrigatório", "O campo 'cidade' é obrigatório.", parent=self)
            return

        if self.cidade_data:
            self.cidade_data.update(new_data)
        else:
            CIDADES.append(new_data)

        if self.callback:
            self.callback()

        self.destroy()


class CompanyForm(ttk.Toplevel):
    """ Janela para Adicionar ou Editar uma Empresa. """

    def __init__(self, parent, cidades_lista, empresa_data=None, callback=None):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()

        self.empresa_data = empresa_data
        self.callback = callback
        self.title("Editar Empresa" if empresa_data else "Adicionar Nova Empresa")

        self.entries = {}

        form_frame = ttk.Frame(self, padding="15")
        form_frame.pack(fill=BOTH, expand=YES)

        for i, field in enumerate(HEADERS_EMPRESAS):
            ttk.Label(form_frame, text=f"{field.replace('_', ' ').title()}:").grid(row=i, column=0, sticky="w", padx=5,
                                                                                   pady=5)

            if field == "regime":
                widget = ttk.Combobox(form_frame, values=["Simples Nacional", "LP / LR"], state="readonly")
                widget.set(empresa_data.get(field, 'Simples Nacional') if empresa_data else 'Simples Nacional')
            elif field == "cidade":
                widget = ttk.Combobox(form_frame, values=cidades_lista, state="readonly")
                widget.set(empresa_data.get(field, '') if empresa_data else '')
            else:
                widget = ttk.Entry(form_frame, width=50)
                if empresa_data:
                    widget.insert(0, empresa_data.get(field, ''))

            widget.grid(row=i, column=1, sticky="ew", padx=5, pady=5)
            self.entries[field] = widget

        form_frame.columnconfigure(1, weight=1)

        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=len(HEADERS_EMPRESAS), column=0, columnspan=2, pady=10)

        save_btn = ttk.Button(button_frame, text="Salvar", command=self.save, bootstyle="success")
        save_btn.pack(side=LEFT, padx=5)
        cancel_btn = ttk.Button(button_frame, text="Cancelar", command=self.destroy, bootstyle="danger")
        cancel_btn.pack(side=LEFT, padx=5)

    def save(self):
        new_data = {field: entry.get() for field, entry in self.entries.items()}

        if not new_data.get("id") or not new_data.get("empresa") or not new_data.get("cidade"):
            messagebox.showwarning("Campos Obrigatórios", "Os campos 'id', 'empresa' e 'cidade' são obrigatórios.",
                                   parent=self)
            return

        if self.empresa_data:
            self.empresa_data.update(new_data)
        else:
            EMPRESAS.append(new_data)

        if self.callback:
            self.callback()

        self.destroy()


class CitiesManager(ttk.Toplevel):
    """ Janela para Gerenciar Cidades. """

    def __init__(self, parent, callback):
        super().__init__(parent)
        self.title("Gerenciador de Cidades")
        self.transient(parent)
        self.grab_set()
        self.callback = callback

        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=BOTH, expand=YES)

        self.tree = ttk.Treeview(main_frame, columns=('cidade', 'link'), show='headings')
        self.tree.heading('cidade', text='Cidade')
        self.tree.heading('link', text='Link do Portal')
        self.tree.pack(fill=BOTH, expand=YES)

        self.refresh_list()

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=X, pady=10)
        btn_frame.columnconfigure((0, 1, 2), weight=1)

        add_btn = ttk.Button(btn_frame, text="Adicionar", command=self.add_city, bootstyle="primary")
        add_btn.grid(row=0, column=0, padx=5, sticky="ew")
        edit_btn = ttk.Button(btn_frame, text="Editar", command=self.edit_city, bootstyle="secondary")
        edit_btn.grid(row=0, column=1, padx=5, sticky="ew")
        del_btn = ttk.Button(btn_frame, text="Excluir", command=self.delete_city, bootstyle="danger")
        del_btn.grid(row=0, column=2, padx=5, sticky="ew")

    def refresh_list(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

        CIDADES.sort(key=lambda x: x.get('cidade', ''))
        for cidade in CIDADES:
            self.tree.insert('', tk.END, values=(cidade.get('cidade'), cidade.get('url')), iid=cidade.get('cidade'))

    def get_selected_city(self):
        selection = self.tree.selection()
        if not selection: return None
        selected_id = selection[0]
        return next((c for c in CIDADES if c.get('cidade') == selected_id), None)

    def add_city(self):
        CityForm(self, callback=self.save_and_refresh_cities)

    def edit_city(self):
        cidade = self.get_selected_city()
        if not cidade:
            messagebox.showwarning("Nenhuma Seleção", "Selecione uma cidade para editar.")
            return
        CityForm(self, cidade_data=cidade, callback=self.save_and_refresh_cities)

    def delete_city(self):
        cidade = self.get_selected_city()
        if not cidade:
            messagebox.showwarning("Nenhuma Seleção", "Selecione uma cidade para excluir.")
            return

        confirm = messagebox.askyesno("Confirmar Exclusão",
                                      f"Tem certeza que deseja excluir a cidade '{cidade.get('cidade')}'?")
        if confirm:
            CIDADES.remove(cidade)
            self.save_and_refresh_cities()

    def save_and_refresh_cities(self):
        if salvar_dados_csv('cidades.csv', CIDADES, HEADERS_CIDADES):
            self.refresh_list()
            self.callback()
        else:
            messagebox.showerror("Erro", "Não foi possível salvar as cidades.")


class App:
    def __init__(self, root, last_updated_date):
        self.root = root
        self.root.title("AgroContat - Login Prefeitura")
        self.root.geometry("700x850")
        self.root.minsize(600, 700)

        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=BOTH, expand=YES)
        main_frame.columnconfigure(0, weight=1)

        try:
            logo_path = resource_path("logo.png")
            if os.path.exists(logo_path):
                image = Image.open(logo_path)
                max_width = 300
                ratio = max_width / image.width
                new_height = int(image.height * ratio)
                image = image.resize((max_width, new_height), Image.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(image)
                logo_label = ttk.Label(main_frame, image=self.logo_img)
                logo_label.pack(pady=(0, 25))
            else:
                raise FileNotFoundError
        except Exception as e:
            print(f"Não foi possível carregar o logo local (logo.png): {e}")
            header_label = ttk.Label(main_frame, text="Agrocontat", font=("Helvetica", 22, "bold"), bootstyle="success")
            header_label.pack(pady=(0, 20))

        filters_frame = ttk.LabelFrame(main_frame, text="Filtros", padding=10)
        filters_frame.pack(fill=X, pady=10)
        filters_frame.columnconfigure((1, 3), weight=1)

        ttk.Label(filters_frame, text="Cidade:", font=("Helvetica", 10)).grid(row=0, column=0, padx=(0, 5), sticky="w")
        self.cidades_lista = sorted([c.get('cidade', 'N/A') for c in CIDADES])
        self.cidade_var = tk.StringVar(value=CONFIGURACOES_INICIAIS.get("ultima_cidade"))
        if self.cidade_var.get() not in self.cidades_lista and self.cidades_lista:
            self.cidade_var.set("Todas")
        self.cidade_menu = ttk.Combobox(filters_frame, textvariable=self.cidade_var,
                                        values=["Todas"] + self.cidades_lista, state="readonly")
        self.cidade_menu.grid(row=0, column=1, padx=(0, 10), sticky="ew")
        self.cidade_menu.bind("<<ComboboxSelected>>", self.update_list)

        ttk.Label(filters_frame, text="Regime:", font=("Helvetica", 10)).grid(row=0, column=2, padx=(10, 5), sticky="w")
        self.regimes = sorted(list(set(emp.get('regime', 'N/A') for emp in EMPRESAS)))
        self.regime_var = tk.StringVar(value=CONFIGURACOES_INICIAIS.get("ultimo_regime"))
        self.regime_menu = ttk.Combobox(filters_frame, textvariable=self.regime_var, values=self.regimes,
                                        state="readonly")
        self.regime_menu.grid(row=0, column=3, sticky="ew")
        self.regime_menu.bind("<<ComboboxSelected>>", self.update_list)

        ttk.Label(filters_frame, text="Organizar por:", font=("Helvetica", 10)).grid(row=1, column=0, padx=(0, 5),
                                                                                     pady=(10, 0), sticky="w")
        self.organizacao_var = tk.StringVar(value=CONFIGURACOES_INICIAIS.get("ultima_organizacao"))
        self.organizacao_menu = ttk.Combobox(filters_frame, textvariable=self.organizacao_var,
                                             values=["Por Nome (A-Z)", "Por Código (0-9)"], state="readonly")
        self.organizacao_menu.grid(row=1, column=1, padx=(0, 10), pady=(10, 0), sticky="ew")
        self.organizacao_menu.bind("<<ComboboxSelected>>", self.update_list)

        ttk.Label(filters_frame, text="Buscar:", font=("Helvetica", 10)).grid(row=1, column=2, padx=(10, 5),
                                                                              pady=(10, 0), sticky="w")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(filters_frame, textvariable=self.search_var)
        self.search_entry.grid(row=1, column=3, pady=(10, 0), sticky="ew")
        self.search_var.trace_add("write", self.update_list)

        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=BOTH, expand=YES, pady=10)

        columns = ('empresa', 'cnpj')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', bootstyle="success")
        self.tree.heading('empresa', text='Empresa')
        self.tree.heading('cnpj', text='CNPJ', anchor='e')
        self.tree.column('cnpj', anchor='e', width=180)
        self.tree.column('empresa', anchor='w')

        self.tree.pack(side=LEFT, fill=BOTH, expand=YES)
        self.tree.bind("<<TreeviewSelect>>", self.on_list_select)

        scrollbar = ttk.Scrollbar(list_frame, orient=VERTICAL, command=self.tree.yview, bootstyle="round-success")
        scrollbar.pack(side=RIGHT, fill=Y)
        self.tree.config(yscrollcommand=scrollbar.set)

        mgmt_frame = ttk.LabelFrame(main_frame, text="Gerenciamento", padding=10)
        mgmt_frame.pack(fill=X, pady=(10, 0))
        mgmt_frame.columnconfigure((0, 1, 2, 3), weight=1)

        self.add_button = ttk.Button(mgmt_frame, text="Adicionar Empresa", command=self.add_company,
                                     bootstyle="primary")
        self.add_button.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        self.edit_button = ttk.Button(mgmt_frame, text="Editar Empresa", command=self.edit_company,
                                      bootstyle="secondary", state="disabled")
        self.edit_button.grid(row=0, column=1, sticky="ew", padx=5)
        self.delete_button = ttk.Button(mgmt_frame, text="Excluir Empresa", command=self.delete_company,
                                        bootstyle="danger", state="disabled")
        self.delete_button.grid(row=0, column=2, sticky="ew", padx=(5, 0))
        self.manage_cities_button = ttk.Button(mgmt_frame, text="Gerenciar Cidades", command=self.open_cities_manager,
                                               bootstyle="outline-info")
        self.manage_cities_button.grid(row=0, column=3, sticky="ew", padx=(10, 0))

        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=X, pady=(10, 0))
        action_frame.columnconfigure((0, 1), weight=1)

        self.view_button = ttk.Button(action_frame, text="Ver Credenciais", command=self.show_credentials_window,
                                      bootstyle="info", state="disabled")
        self.view_button.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        self.login_button = ttk.Button(action_frame, text="Fazer Login", command=self.start_login_thread,
                                       bootstyle="success")
        self.login_button.grid(row=0, column=1, sticky="ew", padx=(5, 0))

        footer_frame = ttk.LabelFrame(main_frame, text="Dados", padding=10)
        footer_frame.pack(fill=X, pady=(10, 0))
        footer_frame.columnconfigure(1, weight=1)

        credits_label = ttk.Label(footer_frame, text="Feito por Lucas Marques \u00A9", font=("Helvetica", 8))
        credits_label.grid(row=0, column=0, sticky="w")

        self.last_updated_label = ttk.Label(footer_frame, text=f"Empresas atualizadas em: {last_updated_date}",
                                            font=("Helvetica", 8))
        self.last_updated_label.grid(row=0, column=1, sticky="e")

        self.update_list()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        salvar_configuracoes(self.cidade_var.get(), self.regime_var.get(), self.organizacao_var.get())
        self.root.destroy()

    def on_list_select(self, event=None):
        is_selected = bool(self.tree.selection())
        state = "normal" if is_selected else "disabled"
        self.view_button.config(state=state)
        self.edit_button.config(state=state)
        self.delete_button.config(state=state)

    def update_list(self, *args):
        for i in self.tree.get_children():
            self.tree.delete(i)

        self.on_list_select()

        cidade_selecionada = self.cidade_var.get()
        regime_selecionado = self.regime_var.get()
        search_term = self.search_var.get().lower()
        organizacao = self.organizacao_var.get()

        lista_filtrada = EMPRESAS
        if cidade_selecionada != "Todas":
            lista_filtrada = [emp for emp in lista_filtrada if emp.get('cidade') == cidade_selecionada]

        lista_filtrada = [emp for emp in lista_filtrada if emp.get('regime') == regime_selecionado]

        self.filtered_list = [
            emp for emp in lista_filtrada
            if search_term in emp.get('empresa', '').lower() or search_term in emp.get('cnpj', '').lower()
        ]

        if organizacao == "Por Nome (A-Z)":
            self.filtered_list.sort(key=lambda x: x.get('empresa', ''))
        elif organizacao == "Por Código (0-9)":
            self.filtered_list.sort(key=lambda x: int(x.get('id', 0) or 0))

        for emp in self.filtered_list:
            empresa_text = f" {emp.get('id')} - {emp.get('empresa', '')}"
            cnpj_text = emp.get('cnpj', '')
            self.tree.insert('', tk.END, values=(empresa_text, cnpj_text), iid=emp.get('id'))

    def get_selected_company(self):
        selection = self.tree.selection()
        if not selection: return None
        selected_id = selection[0]
        return next((emp for emp in self.filtered_list if emp.get('id') == selected_id), None)

    def save_and_refresh_empresas(self):
        if salvar_dados_csv('empresas.csv', EMPRESAS, HEADERS_EMPRESAS):
            self.reload_all_data()
        else:
            messagebox.showerror("Erro", "Não foi possível salvar as empresas.")

    def add_company(self):
        CompanyForm(self.root, self.cidades_lista, callback=self.save_and_refresh_empresas)

    def edit_company(self):
        empresa = self.get_selected_company()
        if not empresa: return
        CompanyForm(self.root, self.cidades_lista, empresa_data=empresa, callback=self.save_and_refresh_empresas)

    def delete_company(self):
        empresa = self.get_selected_company()
        if not empresa: return

        confirm = messagebox.askyesno("Confirmar Exclusão",
                                      f"Tem certeza que deseja excluir a empresa:\n\n{empresa.get('empresa')}?",
                                      parent=self.root)
        if confirm:
            EMPRESAS.remove(empresa)
            self.save_and_refresh_empresas()

    def show_credentials_window(self):
        empresa = self.get_selected_company()
        if not empresa: return

        popup = ttk.Toplevel(self.root)
        popup.title(f"Credenciais - {empresa.get('empresa')}")
        popup.geometry("400x150")
        popup.transient(self.root)
        popup.grab_set()

        popup_frame = ttk.Frame(popup, padding="15")
        popup_frame.pack(fill=BOTH, expand=YES)

        def copy_to_clipboard(text, button):
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            button.config(text="Copiado!", bootstyle="success")
            self.root.after(1500, lambda: button.config(text="Copiar", bootstyle="primary"))

        ttk.Label(popup_frame, text="Login:").grid(row=0, column=0, sticky="w")
        login_entry = ttk.Entry(popup_frame)
        login_entry.insert(0, empresa.get('login', ''))
        login_entry.config(state="readonly")
        login_entry.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        login_btn = ttk.Button(popup_frame, text="Copiar", bootstyle="primary",
                               command=lambda: copy_to_clipboard(empresa.get('login', ''), login_btn))
        login_btn.grid(row=1, column=1, padx=(5, 0), pady=(0, 10))

        ttk.Label(popup_frame, text="Senha:").grid(row=2, column=0, sticky="w")
        senha_entry = ttk.Entry(popup_frame)
        senha_entry.insert(0, empresa.get('senha', ''))
        senha_entry.config(state="readonly")
        senha_entry.grid(row=3, column=0, sticky="ew")

        senha_btn = ttk.Button(popup_frame, text="Copiar", bootstyle="primary",
                               command=lambda: copy_to_clipboard(empresa.get('senha', ''), senha_btn))
        senha_btn.grid(row=3, column=1, padx=(5, 0))
        popup_frame.columnconfigure(0, weight=1)

    def start_login_thread(self):
        empresa = self.get_selected_company()
        if not empresa:
            messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione uma empresa na lista.")
            return

        cidade_nome = empresa.get('cidade')
        cidade_info = next((c for c in CIDADES if c.get('cidade') == cidade_nome), None)

        if not cidade_info:
            messagebox.showerror("Erro de Dados",
                                 f"As informações de login para a cidade '{cidade_nome}' não foram encontradas no arquivo cidades.csv.")
            return

        self.login_button.config(state="disabled", text="Iniciando...")
        self.view_button.config(state="disabled")

        thread = threading.Thread(target=self.login_empresa_automation, args=(empresa, cidade_info))
        thread.daemon = True
        thread.start()

    def login_empresa_automation(self, empresa, cidade):
        try:
            options = webdriver.ChromeOptions()
            profile_path = os.path.join(get_base_path(), "chrome_profile")
            options.add_argument(f"--user-data-dir={profile_path}")
            options.add_argument("--start-maximized")
            options.add_experimental_option("prefs",
                                            {"download.prompt_for_download": True, "safebrowsing.enabled": True})
            options.add_experimental_option("detach", True)

            driver_path = resource_path("chromedriver.exe")
            service = webdriver.ChromeService(executable_path=driver_path)

            driver = webdriver.Chrome(service=service, options=options)
            driver.get(cidade.get('url'))

            self._preencher_campo(driver, cidade.get('cnpj-selector'), empresa.get('cnpj'))
            self._preencher_campo(driver, cidade.get('login-selector'), empresa.get('login'))
            self._preencher_campo(driver, cidade.get('senha_seletor'), empresa.get('senha'))

        except Exception as e:
            messagebox.showerror("Erro de Automação", f"Ocorreu um erro:\n{e}")
        finally:
            self.root.after(0, self.enable_buttons_after_automation)

    def enable_buttons_after_automation(self):
        self.login_button.config(state="normal", text="Fazer Login")
        self.on_list_select()

    def _preencher_campo(self, driver, seletor, valor):
        if not seletor or (isinstance(valor, str) and not valor.strip()):
            return
        try:
            wait = WebDriverWait(driver, 10)
            elemento = wait.until(
                EC.presence_of_element_located((By.XPATH, f"//*[@id='{seletor}' or @name='{seletor}']")))
            elemento.clear()
            elemento.send_keys(valor)
        except (TimeoutException, NoSuchElementException):
            print(f"AVISO: Campo com seletor '{seletor}' não foi encontrado.")

    def open_cities_manager(self):
        CitiesManager(self.root, callback=self.reload_all_data)

    def reload_all_data(self):
        """ Recarrega todos os dados dos CSVs e atualiza a interface. """
        global EMPRESAS, CIDADES, LAST_UPDATED_EMPRESAS
        EMPRESAS, LAST_UPDATED_EMPRESAS = carregar_dados_csv('empresas.csv')
        CIDADES, _ = carregar_dados_csv('cidades.csv')

        self.last_updated_label.config(text=f"Empresas atualizadas em: {LAST_UPDATED_EMPRESAS}")

        self.cidades_lista = sorted([c.get('cidade', 'N/A') for c in CIDADES])
        self.cidade_menu['values'] = ["Todas"] + self.cidades_lista
        if self.cidade_var.get() not in self.cidades_lista:
            self.cidade_var.set("Todas")

        self.regimes = sorted(list(set(emp.get('regime', 'N/A') for emp in EMPRESAS)))
        self.regime_menu['values'] = self.regimes
        if self.regime_var.get() not in self.regimes and self.regimes:
            self.regime_var.set(self.regimes[0])

        self.update_list()


if __name__ == "__main__":
    if EMPRESAS is None or CIDADES is None:
        messagebox.showerror("Erro Crítico",
                             "Não foi possível encontrar 'empresas.csv' e/ou 'cidades.csv'. A aplicação será fechada.")
    else:
        root = ttk.Window(themename="litera")

        try:
            icon_path = resource_path("logo.ico")
            if os.path.exists(icon_path):
                root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Não foi possível definir o ícone da janela: {e}")

        style = root.style
        style.colors.primary = "#005a41"
        style.colors.secondary = "#6c757d"
        style.colors.success = "#00835f"
        style.colors.info = "#54b4d3"
        style.colors.selectbg = "#a6d8c9"
        style.colors.selectfg = "#000000"

        app = App(root, LAST_UPDATED_EMPRESAS)
        root.mainloop()

