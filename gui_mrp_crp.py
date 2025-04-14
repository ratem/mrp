import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
from datetime import datetime
from mrp import MRP
from crp import CRP


class MRP_GUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema MRP")
        self.mrp = None
        self.pasta_trabalho = None
        self.stdout_original = sys.stdout

        # Configurar o encerramento adequado
        self.root.protocol("WM_DELETE_WINDOW", self.sair)

        self.create_widgets()

    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Botões principais (lado esquerdo)
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=0, column=0, sticky=(tk.W, tk.N, tk.S))

        # Seção MRP
        mrp_label = ttk.Label(buttons_frame, text="Operações MRP", font=("Helvetica", 12, "bold"))
        mrp_label.grid(row=0, column=0, pady=10, padx=5, sticky=tk.W)

        ttk.Button(buttons_frame, text="Definir Pasta de Trabalho", command=self.definir_pasta_trabalho).grid(row=1,
                                                                                                              column=0,
                                                                                                              pady=5,
                                                                                                              padx=5,
                                                                                                              sticky=tk.W)
        ttk.Button(buttons_frame, text="Inicializar MRP", command=self.inicializar_mrp).grid(row=2, column=0, pady=5,
                                                                                             padx=5, sticky=tk.W)
        ttk.Button(buttons_frame, text="Planejar Produção", command=self.planejar_producao).grid(row=3, column=0,
                                                                                                 pady=5, padx=5,
                                                                                                 sticky=tk.W)
        ttk.Button(buttons_frame, text="Executar Controle", command=self.executar_controle).grid(row=4, column=0,
                                                                                                 pady=5, padx=5,
                                                                                                 sticky=tk.W)
        ttk.Button(buttons_frame, text="Exportar Resultados", command=self.exportar_resultados).grid(row=5, column=0,
                                                                                                     pady=5, padx=5,
                                                                                                     sticky=tk.W)
        ttk.Button(buttons_frame, text="Listar Ordens", command=self.listar_ordens).grid(row=6, column=0, pady=5,
                                                                                         padx=5, sticky=tk.W)
        ttk.Button(buttons_frame, text="Modificar Ordens", command=self.modificar_ordens).grid(row=7, column=0, pady=5,
                                                                                               padx=5, sticky=tk.W)
        ttk.Button(buttons_frame, text="Atualizar Custos e Leadtimes", command=self.atualizar_custos_leadtimes).grid(
            row=8, column=0, pady=5, padx=5, sticky=tk.W)

        # Seção CRP
        crp_label = ttk.Label(buttons_frame, text="Operações CRP", font=("Helvetica", 12, "bold"))
        crp_label.grid(row=9, column=0, pady=10, padx=5, sticky=tk.W)

        ttk.Button(buttons_frame, text="Inicializar CRP", command=self.inicializar_crp).grid(row=10, column=0, pady=5,
                                                                                             padx=5, sticky=tk.W)
        ttk.Button(buttons_frame, text="Carregar Demanda por Recursos", command=self.carregar_demanda_recursos).grid(
            row=11, column=0, pady=5, padx=5, sticky=tk.W)
        ttk.Button(buttons_frame, text="Carregar Capacidade de Recursos",
                   command=self.carregar_capacidade_recursos).grid(row=12, column=0, pady=5, padx=5, sticky=tk.W)
        ttk.Button(buttons_frame, text="Carregar Exceções de Capacidade",
                   command=self.carregar_excecoes_capacidade).grid(row=13, column=0, pady=5, padx=5, sticky=tk.W)
        ttk.Button(buttons_frame, text="Calcular Demanda por Operação",
                   command=self.calcular_demanda_por_operacao).grid(row=14, column=0, pady=5, padx=5, sticky=tk.W)
        ttk.Button(buttons_frame, text="Criar Planilha CRP", command=self.criar_planilha_crp_gui).grid(row=15, column=0,
                                                                                                   pady=5, padx=5,
                                                                                                   sticky=tk.W)

        # Botão de saída
        ttk.Separator(buttons_frame, orient="horizontal").grid(row=16, column=0, pady=10, sticky=(tk.W, tk.E))
        ttk.Button(buttons_frame, text="Sair", command=self.sair).grid(row=17, column=0, pady=5, padx=5, sticky=tk.W)

        # Painel de saída (lado direito)
        output_frame = ttk.LabelFrame(main_frame, text="Saída do Sistema")
        output_frame.grid(row=0, column=1, padx=10, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configurar o painel de saída para expandir
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=1)

        # Área de texto com barra de rolagem
        scrollbar = ttk.Scrollbar(output_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.output_text = tk.Text(output_frame, wrap=tk.WORD, width=50, height=20,
                                   yscrollcommand=scrollbar.set)
        self.output_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.output_text.yview)

        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Bem-vindo ao Sistema MRP-CRP")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))

        # Configurar redirecionamento de saída
        self.configurar_redirecionamento_saida()

        # Configurar o redimensionamento da janela
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

    def configurar_redirecionamento_saida(self):
        """
        Configura o redirecionamento da saída padrão para o painel de texto.
        """

        class TextRedirector:
            def __init__(self, text_widget):
                self.text_widget = text_widget
                self.buffer = ""

            def write(self, string):
                self.buffer += string
                self.text_widget.insert(tk.END, string)
                self.text_widget.see(tk.END)  # Auto-scroll

            def flush(self):
                pass

        sys.stdout = TextRedirector(self.output_text)

    def restaurar_saida_original(self):
        """
        Restaura a saída padrão original.
        """
        sys.stdout = self.stdout_original

    def definir_pasta_trabalho(self):
        pasta_default = f"MRP_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        pasta = filedialog.askdirectory(title="Selecione a pasta de trabalho ou crie uma nova")

        if not pasta:
            # Se o usuário cancelou, usar a pasta padrão dentro do diretório atual
            pasta = os.path.join(os.getcwd(), pasta_default)

        if not os.path.exists(pasta):
            os.makedirs(pasta)

        self.pasta_trabalho = pasta

        # Inicializar MRP e CRP com a mesma pasta de trabalho
        self.mrp = MRP(self.pasta_trabalho)
        if hasattr(self, 'crp'):
            self.crp = CRP(self.pasta_trabalho)

        self.status_var.set(f"Pasta de trabalho definida: {self.pasta_trabalho}")
        messagebox.showinfo("Pasta de Trabalho", f"Pasta de trabalho definida: {self.pasta_trabalho}")

    def inicializar_mrp(self):
        if not self.mrp:
            messagebox.showerror("Erro", "Pasta de trabalho não definida!")
            return
        try:
            self.mrp.inicializar_dados()
            self.status_var.set("Dados do MRP inicializados com sucesso!")
            messagebox.showinfo("Inicialização", "Dados do MRP inicializados com sucesso!")
        except Exception as e:
            self.status_var.set(f"Erro ao inicializar dados: {e}")
            messagebox.showerror("Erro", f"Erro ao inicializar dados: {e}")

    def planejar_producao(self):
        if not self.mrp:
            messagebox.showerror("Erro", "MRP não inicializado!")
            return
        try:
            demanda = self.input_demanda()
            if demanda:
                self.mrp.planejar_producao(demanda)
                self.status_var.set("Planejamento de produção concluído com sucesso!")
                messagebox.showinfo("Planejamento", "Planejamento de produção concluído com sucesso!")
        except Exception as e:
            self.status_var.set(f"Erro ao planejar produção: {e}")
            messagebox.showerror("Erro", f"Erro ao planejar produção: {e}")

    def executar_controle(self):
        if not self.mrp:
            messagebox.showerror("Erro", "MRP não inicializado!")
            return
        try:
            self.mrp.iniciar_execucao()
            self.status_var.set("Execução iniciada com sucesso!")
            messagebox.showinfo("Execução", "Execução iniciada com sucesso!")
        except Exception as e:
            self.status_var.set(f"Erro ao iniciar execução: {e}")
            messagebox.showerror("Erro", f"Erro ao iniciar execução: {e}")

    def exportar_resultados(self):
        if not self.mrp:
            messagebox.showerror("Erro", "MRP não inicializado!")
            return
        try:
            # Usar a pasta de trabalho definida
            self.mrp.exportar_quadro_planejamento("planejamento_mrp.xlsx")
            self.mrp.exportar_ordens_producao("ordens_producao.xlsx")
            self.mrp.exportar_custos_materiais("custos_materiais.xlsx")
            self.status_var.set("Resultados exportados com sucesso!")
            messagebox.showinfo("Exportação", "Resultados exportados com sucesso!")
        except Exception as e:
            self.status_var.set(f"Erro ao exportar resultados: {e}")
            messagebox.showerror("Erro", f"Erro ao exportar resultados: {e}")

    def listar_ordens(self):
        if not self.mrp:
            messagebox.showerror("Erro", "MRP não inicializado!")
            return
        try:
            # Chama o método original para mostrar no console (que agora é redirecionado para a GUI)
            self.mrp.listar_ordens_controle()

            # Obtém as ordens diretamente para exibir na interface gráfica
            ordens = self.obter_ordens_controle()
            if not ordens:
                messagebox.showinfo("Informação", "Não há ordens de controle para listar.")
            else:
                self.mostrar_ordens(ordens)
        except Exception as e:
            self.status_var.set(f"Erro ao listar ordens: {e}")
            messagebox.showerror("Erro", f"Erro ao listar ordens: {e}")

    def obter_ordens_controle(self):
        """
        Obtém as ordens de controle diretamente do objeto MRP.

        Returns:
            dict: Dicionário de ordens de controle ou None se não disponível.
        """
        if not self.mrp or not hasattr(self.mrp, 'ordens_controle'):
            return None
        return self.mrp.ordens_controle

    def atualizar_custos_leadtimes(self):
        if not self.mrp:
            messagebox.showerror("Erro", "MRP não inicializado!")
            return
        try:
            arquivo_cotacoes = filedialog.askopenfilename(title="Selecione o arquivo de cotações",
                                                          filetypes=[("Excel files", "*.xlsx")])
            if arquivo_cotacoes:
                sucesso, alertas = self.mrp.atualizar_custos_leadtimes(arquivo_cotacoes)
                if sucesso:
                    self.status_var.set("Custos e leadtimes atualizados com sucesso!")
                    messagebox.showinfo("Atualização", "Custos e leadtimes atualizados com sucesso!")
                    if alertas:
                        self.mostrar_alertas(alertas)
                else:
                    self.status_var.set("Falha ao atualizar custos e leadtimes.")
                    messagebox.showerror("Erro", "Falha ao atualizar custos e leadtimes.")
        except Exception as e:
            self.status_var.set(f"Erro ao atualizar custos e leadtimes: {e}")
            messagebox.showerror("Erro", f"Erro ao atualizar custos e leadtimes: {e}")

    def input_demanda(self):
        if not self.mrp:
            messagebox.showerror("Erro", "MRP não inicializado!")
            return None

        # Obter a lista de produtos finais disponíveis
        produtos_finais = []
        for arquivo in os.listdir(self.pasta_trabalho):
            if arquivo.endswith("_BOM.xlsx"):
                produto = arquivo.replace("_BOM.xlsx", "")
                produtos_finais.append(produto)

        if not produtos_finais:
            messagebox.showerror("Erro", "Nenhum produto final encontrado. Verifique os arquivos de BOM.")
            return None

        demanda_window = tk.Toplevel(self.root)
        demanda_window.title("Entrada de Demanda")
        demanda_window.transient(self.root)  # Torna a janela dependente da janela principal
        demanda_window.grab_set()  # Foca a janela

        # Criar um frame com scroll caso haja muitos produtos
        canvas = tk.Canvas(demanda_window)
        scrollbar = ttk.Scrollbar(demanda_window, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)

        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Dicionário para armazenar as entradas
        entries = {}

        # Criar campos para cada produto
        for i, produto in enumerate(produtos_finais):
            ttk.Label(scroll_frame, text=f"{produto}:").grid(row=i, column=0, pady=5, padx=5, sticky=tk.W)
            entry = ttk.Entry(scroll_frame)
            entry.grid(row=i, column=1, pady=5, padx=5, sticky=(tk.W, tk.E))
            entry.insert(0, "0")  # Valor padrão
            entries[produto] = entry

        demanda = {}

        def confirmar():
            try:
                for produto, entry in entries.items():
                    valor = int(entry.get())
                    if valor > 0:  # Só incluir produtos com demanda positiva
                        demanda[produto] = valor
                demanda_window.destroy()
            except ValueError:
                messagebox.showerror("Erro", "Por favor, insira valores numéricos válidos.")

        ttk.Button(demanda_window, text="Confirmar", command=confirmar).pack(pady=10)

        # Centralizar a janela
        demanda_window.update_idletasks()
        width = demanda_window.winfo_width()
        height = demanda_window.winfo_height()
        x = (demanda_window.winfo_screenwidth() // 2) - (width // 2)
        y = (demanda_window.winfo_screenheight() // 2) - (height // 2)
        demanda_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))

        self.root.wait_window(demanda_window)
        return demanda if demanda else None

    def mostrar_ordens(self, ordens):
        ordens_window = tk.Toplevel(self.root)
        ordens_window.title("Ordens de Controle")
        ordens_window.transient(self.root)  # Torna a janela dependente da janela principal

        # Frame com barra de rolagem
        frame = ttk.Frame(ordens_window)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Adicionar barra de rolagem
        scrollbar_y = ttk.Scrollbar(frame, orient="vertical")
        scrollbar_y.pack(side="right", fill="y")

        scrollbar_x = ttk.Scrollbar(frame, orient="horizontal")
        scrollbar_x.pack(side="bottom", fill="x")

        tree = ttk.Treeview(frame, columns=("Material", "Estoque Atual", "Retirada", "Produção", "Aquisição", "Status"),
                            show="headings", yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)

        tree.heading("Material", text="Material")
        tree.heading("Estoque Atual", text="Estoque Atual")
        tree.heading("Retirada", text="Retirada")
        tree.heading("Produção", text="Produção")
        tree.heading("Aquisição", text="Aquisição")
        tree.heading("Status", text="Status")

        # Definir largura das colunas
        tree.column("Material", width=150)
        tree.column("Estoque Atual", width=100)
        tree.column("Retirada", width=100)
        tree.column("Produção", width=100)
        tree.column("Aquisição", width=100)
        tree.column("Status", width=100)

        for material, dados in ordens.items():
            tree.insert("", "end", values=(
                material,
                dados.get("Estoque Atual", ""),
                dados.get("Retirada de Estoque", ""),
                dados.get("Produção", ""),
                dados.get("Aquisição", ""),
                dados.get("Status", "")
            ))

        tree.pack(side="left", fill="both", expand=True)

        # Botão de fechar
        ttk.Button(ordens_window, text="Fechar", command=ordens_window.destroy).pack(pady=10)

        # Definir tamanho e posição da janela
        ordens_window.geometry("800x400")
        ordens_window.update_idletasks()
        width = ordens_window.winfo_width()
        height = ordens_window.winfo_height()
        x = (ordens_window.winfo_screenwidth() // 2) - (width // 2)
        y = (ordens_window.winfo_screenheight() // 2) - (height // 2)
        ordens_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))

    def modificar_ordens(self):
        """
        Permite ao usuário modificar ordens durante a fase de execução e controle.
        """
        if not self.mrp:
            messagebox.showerror("Erro", "MRP não inicializado!")
            return

        try:
            # Obter as ordens de controle existentes
            ordens = self.obter_ordens_controle()
            if not ordens:
                messagebox.showinfo("Informação", "Não há ordens de controle para modificar.")
                return

            # Criar janela para modificar ordens
            janela = tk.Toplevel(self.root)
            janela.title("Modificar Ordens")
            janela.transient(self.root)
            janela.grab_set()

            # Frame principal com barra de rolagem
            main_frame = ttk.Frame(janela, padding="10")
            main_frame.pack(fill="both", expand=True)

            # Adicionar barras de rolagem
            scrollbar_y = ttk.Scrollbar(main_frame, orient="vertical")
            scrollbar_y.pack(side="right", fill="y")

            scrollbar_x = ttk.Scrollbar(main_frame, orient="horizontal")
            scrollbar_x.pack(side="bottom", fill="x")

            # Criar tabela para exibir e modificar ordens
            tree = ttk.Treeview(main_frame,
                                columns=("Material", "Estoque Atual", "Retirada", "Produção", "Aquisição", "Status"),
                                show="headings", yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

            scrollbar_y.config(command=tree.yview)
            scrollbar_x.config(command=tree.xview)

            tree.heading("Material", text="Material")
            tree.heading("Estoque Atual", text="Estoque Atual")
            tree.heading("Retirada", text="Retirada de Estoque")
            tree.heading("Produção", text="Produção")
            tree.heading("Aquisição", text="Aquisição")
            tree.heading("Status", text="Status")

            # Definir largura das colunas
            tree.column("Material", width=150)
            tree.column("Estoque Atual", width=100)
            tree.column("Retirada", width=120)
            tree.column("Produção", width=100)
            tree.column("Aquisição", width=100)
            tree.column("Status", width=120)

            # Preencher a tabela com as ordens existentes
            for material, dados in ordens.items():
                tree.insert("", "end", values=(
                    material,
                    dados.get("Estoque Atual", ""),
                    dados.get("Retirada de Estoque", ""),
                    dados.get("Produção", ""),
                    dados.get("Aquisição", ""),
                    dados.get("Status", "")
                ), tags=(material,))

            tree.pack(fill="both", expand=True)

            # Frame para edição de ordem selecionada
            edit_frame = ttk.LabelFrame(janela, text="Editar Ordem", padding="10")
            edit_frame.pack(fill="x", padx=10, pady=10)

            # Campos de edição
            ttk.Label(edit_frame, text="Material:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
            material_var = tk.StringVar()
            material_entry = ttk.Entry(edit_frame, textvariable=material_var, state="readonly")
            material_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

            ttk.Label(edit_frame, text="Status:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
            status_var = tk.StringVar()
            status_combo = ttk.Combobox(edit_frame, textvariable=status_var)
            status_combo['values'] = ('Planejado', 'Em Execução', 'Concluído', 'Cancelado')
            status_combo.grid(row=1, column=1, sticky="ew", padx=5, pady=5)

            # Função para carregar os dados da ordem selecionada
            def item_selecionado(event):
                selected_items = tree.selection()
                if selected_items:
                    item = selected_items[0]
                    valores = tree.item(item, "values")
                    material_var.set(valores[0])
                    status_var.set(valores[5])

            tree.bind('<<TreeviewSelect>>', item_selecionado)

            # Função para atualizar o status da ordem
            def atualizar_status():
                material = material_var.get()
                novo_status = status_var.get()

                if not material or not novo_status:
                    messagebox.showwarning("Aviso", "Selecione um material e um status.")
                    return

                try:
                    # Atualizar o status da ordem no MRP
                    self.mrp.atualizar_status_ordem(material, novo_status)

                    # Atualizar a visualização na tabela
                    for item in tree.selection():
                        valores = list(tree.item(item, "values"))
                        valores[5] = novo_status
                        tree.item(item, values=valores)

                    messagebox.showinfo("Sucesso", f"Status da ordem para {material} atualizado para {novo_status}.")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao atualizar status: {e}")

            # Botões de ação
            button_frame = ttk.Frame(janela)
            button_frame.pack(fill="x", padx=10, pady=10)

            ttk.Button(button_frame, text="Atualizar Status", command=atualizar_status).pack(side="left", padx=5)
            ttk.Button(button_frame, text="Fechar", command=janela.destroy).pack(side="right", padx=5)

            # Centralizar a janela
            janela.update_idletasks()
            width = 800
            height = 600
            x = (janela.winfo_screenwidth() // 2) - (width // 2)
            y = (janela.winfo_screenheight() // 2) - (height // 2)
            janela.geometry(f'{width}x{height}+{x}+{y}')

        except Exception as e:
            self.status_var.set(f"Erro ao modificar ordens: {e}")
            messagebox.showerror("Erro", f"Erro ao modificar ordens: {e}")


    def mostrar_alertas(self, alertas):
        alertas_window = tk.Toplevel(self.root)
        alertas_window.title("Alertas")
        alertas_window.transient(self.root)  # Torna a janela dependente da janela principal

        # Frame com barra de rolagem
        frame = ttk.Frame(alertas_window)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Adicionar barra de rolagem
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side="right", fill="y")

        # Área de texto
        text_area = tk.Text(frame, wrap="word", yscrollcommand=scrollbar.set)
        text_area.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=text_area.yview)

        # Inserir os alertas
        for alerta in alertas:
            text_area.insert(tk.END, f"• {alerta}\n")

        text_area.config(state="disabled")  # Tornar somente leitura

        # Botão de fechar
        ttk.Button(alertas_window, text="Fechar", command=alertas_window.destroy).pack(pady=10)

        # Definir tamanho e posição da janela
        alertas_window.geometry("500x300")
        alertas_window.update_idletasks()
        width = alertas_window.winfo_width()
        height = alertas_window.winfo_height()
        x = (alertas_window.winfo_screenwidth() // 2) - (width // 2)
        y = (alertas_window.winfo_screenheight() // 2) - (height // 2)
        alertas_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))

    def fluxo_crp(self):
        """
        Exibe a interface para o fluxo CRP.
        """
        if not self.mrp:
            messagebox.showerror("Erro", "MRP não inicializado!")
            return

        # Criar janela para o fluxo CRP
        crp_window = tk.Toplevel(self.root)
        crp_window.title("Fluxo CRP")
        crp_window.transient(self.root)
        crp_window.grab_set()

        # Frame principal
        main_frame = ttk.Frame(crp_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Botões para as operações do CRP
        ttk.Button(main_frame, text="Inicializar CRP", command=self.inicializar_crp).grid(row=0, column=0, pady=5,
                                                                                          padx=5, sticky=tk.W)
        ttk.Button(main_frame, text="Carregar Demanda por Recursos", command=self.carregar_demanda_recursos).grid(row=1,
                                                                                                                  column=0,
                                                                                                                  pady=5,
                                                                                                                  padx=5,
                                                                                                                  sticky=tk.W)
        ttk.Button(main_frame, text="Carregar Capacidade de Recursos", command=self.carregar_capacidade_recursos).grid(
            row=2, column=0, pady=5, padx=5, sticky=tk.W)
        ttk.Button(main_frame, text="Carregar Exceções de Capacidade", command=self.carregar_excecoes_capacidade).grid(
            row=3, column=0, pady=5, padx=5, sticky=tk.W)
        ttk.Button(main_frame, text="Calcular Demanda por Operação", command=self.calcular_demanda_por_operacao).grid(
            row=4, column=0, pady=5, padx=5, sticky=tk.W)
        ttk.Button(main_frame, text="Criar Planilha CRP", command=self.criar_planilha_crp_gui).grid(row=5, column=0, pady=5,
                                                                                                padx=5, sticky=tk.W)
        ttk.Button(main_frame, text="Fechar", command=crp_window.destroy).grid(row=6, column=0, pady=5, padx=5,
                                                                               sticky=tk.W)

        # Centralizar a janela
        crp_window.update_idletasks()
        width = crp_window.winfo_width()
        height = crp_window.winfo_height()
        x = (crp_window.winfo_screenwidth() // 2) - (width // 2)
        y = (crp_window.winfo_screenheight() // 2) - (height // 2)
        crp_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))

    def inicializar_crp(self):
        """
        Inicializa o CRP carregando o planejamento do MRP.
        """
        if not hasattr(self, 'pasta_trabalho'):
            messagebox.showerror("Erro", "Pasta de trabalho não definida!")
            return

        if not hasattr(self, 'crp'):
            self.crp = CRP(self.pasta_trabalho)

        try:
            # Verificar se existe um quadro de planejamento exportado
            arquivo_planejamento = "planejamento_mrp.xlsx"
            caminho_planejamento = os.path.join(self.pasta_trabalho, arquivo_planejamento)

            if not os.path.exists(caminho_planejamento):
                messagebox.showerror("Erro",
                                     "Arquivo de planejamento MRP não encontrado. Exporte o quadro de planejamento primeiro.")
                return

            # Carregar o planejamento do MRP
            resultado = self.crp.carregar_planejamento_mrp(arquivo_planejamento)

            if resultado:
                self.status_var.set("CRP inicializado com sucesso!")
                messagebox.showinfo("Inicialização CRP", "CRP inicializado com sucesso!")
            else:
                self.status_var.set("Erro ao inicializar o CRP.")
                messagebox.showerror("Erro", "Erro ao inicializar o CRP.")
        except Exception as e:
            self.status_var.set(f"Erro ao inicializar o CRP: {e}")
            messagebox.showerror("Erro", f"Erro ao inicializar o CRP: {e}")

    def carregar_demanda_recursos(self):
        """
        Carrega a planilha de demanda por recursos.
        """
        if not hasattr(self, 'crp'):
            messagebox.showerror("Erro", "CRP não inicializado!")
            return

        try:
            arquivo = filedialog.askopenfilename(
                title="Selecione o arquivo de demanda por recursos",
                filetypes=[("Excel files", "*.xlsx")],
                initialdir=self.pasta_trabalho
            )

            if arquivo:
                nome_arquivo = os.path.basename(arquivo)
                # Copiar o arquivo para a pasta de trabalho se não estiver lá
                if os.path.dirname(arquivo) != self.pasta_trabalho:
                    import shutil
                    shutil.copy(arquivo, os.path.join(self.pasta_trabalho, nome_arquivo))

                resultado = self.crp.carregar_demanda_recursos(nome_arquivo)

                if resultado:
                    self.status_var.set("Demanda por recursos carregada com sucesso!")
                    messagebox.showinfo("Demanda por Recursos", "Demanda por recursos carregada com sucesso!")
                else:
                    self.status_var.set("Erro ao carregar demanda por recursos.")
                    messagebox.showerror("Erro", "Erro ao carregar demanda por recursos.")
        except Exception as e:
            self.status_var.set(f"Erro ao carregar demanda por recursos: {e}")
            messagebox.showerror("Erro", f"Erro ao carregar demanda por recursos: {e}")

    def carregar_capacidade_recursos(self):
        """
        Carrega a planilha de capacidade de recursos.
        """
        if not hasattr(self, 'crp'):
            messagebox.showerror("Erro", "CRP não inicializado!")
            return

        try:
            arquivo = filedialog.askopenfilename(
                title="Selecione o arquivo de capacidade de recursos",
                filetypes=[("Excel files", "*.xlsx")],
                initialdir=self.pasta_trabalho
            )

            if arquivo:
                nome_arquivo = os.path.basename(arquivo)
                # Copiar o arquivo para a pasta de trabalho se não estiver lá
                if os.path.dirname(arquivo) != self.pasta_trabalho:
                    import shutil
                    shutil.copy(arquivo, os.path.join(self.pasta_trabalho, nome_arquivo))

                resultado = self.crp.carregar_capacidade_recursos(nome_arquivo)

                if resultado:
                    self.status_var.set("Capacidade de recursos carregada com sucesso!")
                    messagebox.showinfo("Capacidade de Recursos", "Capacidade de recursos carregada com sucesso!")
                else:
                    self.status_var.set("Erro ao carregar capacidade de recursos.")
                    messagebox.showerror("Erro", "Erro ao carregar capacidade de recursos.")
        except Exception as e:
            self.status_var.set(f"Erro ao carregar capacidade de recursos: {e}")
            messagebox.showerror("Erro", f"Erro ao carregar capacidade de recursos: {e}")

    def carregar_excecoes_capacidade(self):
        """
        Carrega a planilha de exceções de capacidade.
        """
        if not hasattr(self, 'crp'):
            messagebox.showerror("Erro", "CRP não inicializado!")
            return

        try:
            arquivo = filedialog.askopenfilename(
                title="Selecione o arquivo de exceções de capacidade",
                filetypes=[("Excel files", "*.xlsx")],
                initialdir=self.pasta_trabalho
            )

            if arquivo:
                nome_arquivo = os.path.basename(arquivo)
                # Copiar o arquivo para a pasta de trabalho se não estiver lá
                if os.path.dirname(arquivo) != self.pasta_trabalho:
                    import shutil
                    shutil.copy(arquivo, os.path.join(self.pasta_trabalho, nome_arquivo))

                resultado = self.crp.carregar_excecoes_capacidade(nome_arquivo)

                if resultado:
                    self.status_var.set("Exceções de capacidade carregadas com sucesso!")
                    messagebox.showinfo("Exceções de Capacidade", "Exceções de capacidade carregadas com sucesso!")
                else:
                    self.status_var.set("Erro ao carregar exceções de capacidade.")
                    messagebox.showerror("Erro", "Erro ao carregar exceções de capacidade.")
        except Exception as e:
            self.status_var.set(f"Erro ao carregar exceções de capacidade: {e}")
            messagebox.showerror("Erro", f"Erro ao carregar exceções de capacidade: {e}")

    def calcular_demanda_por_operacao(self):
        """
        Calcula a demanda por operação e exibe os resultados.
        """
        if not hasattr(self, 'crp'):
            messagebox.showerror("Erro", "CRP não inicializado!")
            return

        try:
            demanda_por_operacao = self.crp.calcular_demanda_por_operacao()

            if demanda_por_operacao:
                # Exibir os resultados em uma janela de texto
                self.mostrar_demanda_por_operacao(demanda_por_operacao)
                self.status_var.set("Demanda por operação calculada com sucesso!")
            else:
                self.status_var.set("Erro ao calcular demanda por operação.")
                messagebox.showerror("Erro", "Erro ao calcular demanda por operação.")
        except Exception as e:
            self.status_var.set(f"Erro ao calcular demanda por operação: {e}")
            messagebox.showerror("Erro", f"Erro ao calcular demanda por operação: {e}")

    def mostrar_demanda_por_operacao(self, demanda_por_operacao):
        """
        Exibe a demanda por operação em uma janela de texto.
        """
        janela = tk.Toplevel(self.root)
        janela.title("Demanda por Operação")
        janela.transient(self.root)

        # Frame com barra de rolagem
        frame = ttk.Frame(janela)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Adicionar barra de rolagem
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side="right", fill="y")

        # Área de texto
        text_area = tk.Text(frame, wrap="word", yscrollcommand=scrollbar.set)
        text_area.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=text_area.yview)

        # Inserir os resultados
        for operacao, produtos in demanda_por_operacao.items():
            text_area.insert(tk.END, f"Operação: {operacao}\n")
            for produto, minutos in produtos.items():
                text_area.insert(tk.END, f"  {produto}: {minutos} minutos\n")
            text_area.insert(tk.END, "\n")

        text_area.config(state="disabled")  # Tornar somente leitura

        # Botão de fechar
        ttk.Button(janela, text="Fechar", command=janela.destroy).pack(pady=10)

        # Definir tamanho e posição da janela
        janela.geometry("500x400")
        janela.update_idletasks()
        width = janela.winfo_width()
        height = janela.winfo_height()
        x = (janela.winfo_screenwidth() // 2) - (width // 2)
        y = (janela.winfo_screenheight() // 2) - (height // 2)
        janela.geometry('{}x{}+{}+{}'.format(width, height, x, y))

    def criar_planilha_crp(self, nome_arquivo, data_planejamento, numero_dias):
        """
        Cria uma planilha para o CRP com base nos dados de capacidade e demanda calculados anteriormente.

        Args:
            nome_arquivo (str): Nome do arquivo Excel a ser criado.
            data_planejamento (str): Data inicial do planejamento (formato YYYY-MM-DD).
            numero_dias (int): Número de dias para o planejamento.

        Returns:
            str: Caminho completo do arquivo Excel criado.
        """
        import openpyxl
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        from openpyxl.utils import get_column_letter
        from openpyxl.worksheet.datavalidation import DataValidation
        from datetime import datetime, timedelta

        # Verificar se todos os dados necessários foram carregados
        if not hasattr(self, 'capacidade_recursos') or not hasattr(self, 'demanda_recursos'):
            print("Erro: Capacidade de recursos ou demanda por recursos não foram carregados.")
            return None

        # Identificar produtos finais (aqueles que estão na demanda_recursos)
        produtos_finais = list(self.demanda_recursos.keys())

        # Criar um novo workbook
        wb = openpyxl.Workbook()

        # Criar aba de demanda total
        ws_demanda = wb.active
        ws_demanda.title = "Demanda Total"

        # Cabeçalhos da aba de demanda
        cabecalhos_demanda = ["Produto", "Demanda Total", "Alocado", "Pendente"]

        for col_num, cabecalho in enumerate(cabecalhos_demanda, 1):
            cell = ws_demanda.cell(row=1, column=col_num, value=cabecalho)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")

        # Preencher dados de produtos finais e demanda total
        linha_atual = 2
        for produto in produtos_finais:
            if produto in self.planejamento_mrp:
                planejamento = self.planejamento_mrp[produto]
                # Calcular demanda total
                demanda_total = sum(quantidade for data, quantidade in planejamento.items()
                                    if data != "Estoque Atual")

                if demanda_total > 0:
                    ws_demanda.cell(row=linha_atual, column=1, value=produto)
                    ws_demanda.cell(row=linha_atual, column=2, value=demanda_total)
                    # Fórmula para calcular o total alocado (será preenchida depois)
                    ws_demanda.cell(row=linha_atual, column=3, value=0)
                    # Fórmula para calcular o pendente
                    ws_demanda.cell(row=linha_atual, column=4, value=f"=B{linha_atual}-C{linha_atual}")
                    linha_atual += 1

        # Ajustar largura das colunas
        for col_num in range(1, len(cabecalhos_demanda) + 1):
            ws_demanda.column_dimensions[get_column_letter(col_num)].width = 20

        # Adicionar bordas
        borda_fina = Border(left=Side(style='thin'), right=Side(style='thin'),
                            top=Side(style='thin'), bottom=Side(style='thin'))
        for row in ws_demanda[f"A1:D{linha_atual - 1}"]:
            for cell in row:
                cell.border = borda_fina

        # Criar planilhas de alocação para cada dia
        for dia in range(numero_dias):
            data_atual = datetime.strptime(data_planejamento, '%Y-%m-%d') + timedelta(days=dia)
            data_str = data_atual.strftime('%Y-%m-%d')
            ws = wb.create_sheet(title=f"Alocação {data_str}")

            # Cabeçalhos da planilha de alocação
            cabecalhos = ["Produto", "Demanda Pendente", "Quantidade Alocada"]

            for col_num, cabecalho in enumerate(cabecalhos, 1):
                cell = ws.cell(row=1, column=col_num, value=cabecalho)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")

            # Preencher dados de produtos finais e demanda
            linha_atual = 2
            for i, produto in enumerate(produtos_finais):
                if produto in self.planejamento_mrp:
                    produto_idx = i + 2  # +2 porque a linha 1 é cabeçalho
                    ws.cell(row=linha_atual, column=1, value=produto)
                    ws.cell(row=linha_atual, column=2, value=f"='Demanda Total'!D{produto_idx}")
                    ws.cell(row=linha_atual, column=3, value=0)
                    linha_atual += 1

            # Adicionar validação de dados para "Quantidade Alocada"
            dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1=0)
            dv.error = "A quantidade alocada deve ser um número inteiro não negativo."
            dv.errorTitle = "Entrada inválida"
            ws.add_data_validation(dv)
            dv.add(f"C2:C{linha_atual - 1}")

            # Ajustar largura das colunas
            for col_num in range(1, len(cabecalhos) + 1):
                ws.column_dimensions[get_column_letter(col_num)].width = 20

            # Adicionar bordas
            for row in ws[f"A1:C{linha_atual - 1}"]:
                for cell in row:
                    cell.border = borda_fina

            # Adicionar tabela de alocação de recursos para este dia
            linha_recursos = linha_atual + 2
            ws.cell(row=linha_recursos, column=1, value="Alocação de Recursos")
            ws.cell(row=linha_recursos, column=1).font = Font(bold=True)

            # Cabeçalhos da tabela de recursos
            cabecalhos_recursos = ["Recurso", "Operação", "Capacidade Nominal", "Exceção",
                                   "Capacidade Disponível", "Consumo", "% Utilização"]

            for col_num, cabecalho in enumerate(cabecalhos_recursos, 1):
                cell = ws.cell(row=linha_recursos + 1, column=col_num, value=cabecalho)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")

            # Preencher dados de recursos
            linha_atual_recursos = linha_recursos + 2
            for recurso, operacoes in self.capacidade_recursos.items():
                for operacao, capacidade_nominal in operacoes.items():
                    if capacidade_nominal > 0:  # Só incluir operações que o recurso pode realizar
                        # Verificar se há exceção para este recurso/operação nesta data
                        excecao = 0
                        if hasattr(self, 'excecoes_capacidade') and recurso in self.excecoes_capacidade:
                            if operacao in self.excecoes_capacidade[recurso]:
                                if data_str in self.excecoes_capacidade[recurso][operacao]:
                                    excecao = self.excecoes_capacidade[recurso][operacao][data_str]

                        # Calcular capacidade disponível -> valores negativos de exceção reduzem a capacidade nominal
                        capacidade_disponivel = max(0, capacidade_nominal + excecao)

                        # Fórmula para calcular o consumo baseado nas quantidades alocadas
                        formula_consumo = "="
                        primeiro = True

                        for i, produto in enumerate(produtos_finais):
                            if produto in self.demanda_recursos and operacao in self.demanda_recursos[produto]:
                                minutos_por_unidade = self.demanda_recursos[produto][operacao]
                                # Encontrar a linha do produto na planilha de alocação
                                produto_linha = i + 2  # +2 porque a linha 1 é cabeçalho
                                if not primeiro:
                                    formula_consumo += "+"
                                formula_consumo += f"C{produto_linha}*{minutos_por_unidade}"
                                primeiro = False

                        if formula_consumo == "=":
                            formula_consumo = "=0"

                        ws.cell(row=linha_atual_recursos, column=1, value=recurso)
                        ws.cell(row=linha_atual_recursos, column=2, value=operacao)
                        ws.cell(row=linha_atual_recursos, column=3, value=capacidade_nominal)
                        ws.cell(row=linha_atual_recursos, column=4, value=excecao)
                        ws.cell(row=linha_atual_recursos, column=5, value=capacidade_disponivel)
                        ws.cell(row=linha_atual_recursos, column=6, value=formula_consumo)
                        ws.cell(row=linha_atual_recursos, column=7,
                                value=f"=F{linha_atual_recursos}/E{linha_atual_recursos}")
                        ws.cell(row=linha_atual_recursos, column=7).number_format = "0.00%"

                        linha_atual_recursos += 1

            # Ajustar largura das colunas
            for col_num in range(1, len(cabecalhos_recursos) + 1):
                ws.column_dimensions[get_column_letter(col_num)].width = 20

            # Adicionar bordas
            for row in ws[f"A{linha_recursos + 1}:G{linha_atual_recursos - 1}"]:
                for cell in row:
                    cell.border = borda_fina

            # Adicionar formatação condicional para % Utilização
            from openpyxl.formatting.rule import ColorScaleRule
            color_scale = ColorScaleRule(start_type='num', start_value=0, start_color='00FF00',
                                         mid_type='num', mid_value=0.7, mid_color='FFFF00',
                                         end_type='num', end_value=1, end_color='FF0000')
            ws.conditional_formatting.add(f"G{linha_recursos + 2}:G{linha_atual_recursos - 1}", color_scale)

        # Agora que todas as abas foram criadas, adicionar fórmulas para atualizar a aba de Demanda Total
        for i, produto in enumerate(produtos_finais):
            if produto in self.planejamento_mrp:
                produto_idx = i + 2  # +2 porque a linha 1 é cabeçalho

                # Construir fórmula para somar todas as quantidades alocadas nas abas diárias
                formula_alocado = "="
                primeiro = True

                for dia in range(numero_dias):
                    data_atual = datetime.strptime(data_planejamento, '%Y-%m-%d') + timedelta(days=dia)
                    data_str = data_atual.strftime('%Y-%m-%d')

                    if not primeiro:
                        formula_alocado += "+"
                    formula_alocado += f"'Alocação {data_str}'!C{produto_idx}"
                    primeiro = False

                # Atualizar a célula com a fórmula
                ws_demanda.cell(row=produto_idx, column=3, value=formula_alocado)

        # Remover a planilha padrão criada pelo openpyxl se existir
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

        # Definir o caminho completo do arquivo
        caminho_arquivo = os.path.join(self.pasta_arquivos, nome_arquivo)

        # Salvar o arquivo
        wb.save(caminho_arquivo)

        print(f"Planilha CRP criada com sucesso: {caminho_arquivo}")
        return caminho_arquivo

    def criar_planilha_crp_gui(self):
        if not hasattr(self, 'crp'):
            messagebox.showerror("Erro", "CRP não inicializado!")
            return

        try:
            # Criar janela para entrada de dados
            janela = tk.Toplevel(self.root)
            janela.title("Criar Planilha CRP")
            janela.transient(self.root)
            janela.grab_set()

            frame = ttk.Frame(janela, padding="10")
            frame.pack(fill="both", expand=True)

            # Campos de entrada
            ttk.Label(frame, text="Nome do Arquivo:").grid(row=0, column=0, pady=5, padx=5, sticky=tk.W)
            nome_arquivo_entry = ttk.Entry(frame)
            nome_arquivo_entry.grid(row=0, column=1, pady=5, padx=5, sticky=(tk.W, tk.E))
            nome_arquivo_entry.insert(0, "crp_planejamento.xlsx")

            ttk.Label(frame, text="Data de Início (YYYY-MM-DD):").grid(row=1, column=0, pady=5, padx=5, sticky=tk.W)
            data_entry = ttk.Entry(frame)
            data_entry.grid(row=1, column=1, pady=5, padx=5, sticky=(tk.W, tk.E))
            data_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))

            ttk.Label(frame, text="Número de Dias:").grid(row=2, column=0, pady=5, padx=5, sticky=tk.W)
            dias_entry = ttk.Entry(frame)
            dias_entry.grid(row=2, column=1, pady=5, padx=5, sticky=(tk.W, tk.E))
            dias_entry.insert(0, "5")

            # Função para criar a planilha
            def confirmar():
                try:
                    nome_arquivo = nome_arquivo_entry.get()
                    data = data_entry.get()
                    dias = int(dias_entry.get())

                    caminho_crp = self.crp.criar_planilha_crp(nome_arquivo, data, dias)

                    if caminho_crp:
                        self.status_var.set(f"Planilha CRP criada com sucesso: {nome_arquivo}")
                        messagebox.showinfo("Planilha CRP", f"Planilha CRP criada com sucesso: {nome_arquivo}")
                        janela.destroy()
                    else:
                        self.status_var.set("Erro ao criar planilha CRP.")
                        messagebox.showerror("Erro", "Erro ao criar planilha CRP.")
                except ValueError:
                    messagebox.showerror("Erro", "Por favor, insira valores válidos.")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao criar planilha CRP: {e}")

            # Botões
            ttk.Button(frame, text="Confirmar", command=confirmar).grid(row=3, column=0, pady=10, padx=5)
            ttk.Button(frame, text="Cancelar", command=janela.destroy).grid(row=3, column=1, pady=10, padx=5)

        except Exception as e:
            self.status_var.set(f"Erro ao criar planilha CRP: {e}")
            messagebox.showerror("Erro", f"Erro ao criar planilha CRP: {e}")


    def abrir_arquivo(self, caminho_arquivo):
        """
        Abre um arquivo com o aplicativo padrão do sistema.
        """
        import platform
        import subprocess

        try:
            if platform.system() == 'Windows':
                os.startfile(caminho_arquivo)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.call(('open', caminho_arquivo))
            else:  # Linux e outros
                subprocess.call(('xdg-open', caminho_arquivo))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir o arquivo: {e}")

    def sair(self):
        """
        Restaura a saída original e fecha a aplicação.
        """
        self.restaurar_saida_original()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = MRP_GUI(root)
    root.mainloop()
