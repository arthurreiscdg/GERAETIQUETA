import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from controller.etiqueta_controller import EtiquetaController
import os
from datetime import datetime
import threading

class EtiquetaView:
    def __init__(self):
        """Inicializa a interface gráfica"""
        self.controller = EtiquetaController()
        self.root = tk.Tk()

        # Dados
        self.current_data = []
        self.filtered_data = []

        # Paginação
    # Paginação removida
        
        # Modo de visualização
        self.grouped_view = True  # True = agrupado por OP, False = lista simples

        # Loading state
        self.is_loading = False
        self.loading_window = None

        # Estado dos cards
        self.card_widgets = {}  # mapeia op -> widget do card
        self.selected_op = None

        # Configuração da janela principal
        self.setup_main_window()

        # Criação dos widgets
        self.create_widgets()

        # Atualiza a lista inicial
        self.refresh_data()
    
    def setup_main_window(self):
        """Configura a janela principal"""
        self.root.title("Sistema de Gestão de Etiquetas")
        self.root.geometry("1200x700")
        self.root.minsize(1000, 600)
        
        # Centraliza a janela
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (1200 // 2)
        y = (self.root.winfo_screenheight() // 2) - (700 // 2)
        self.root.geometry(f"1200x700+{x}+{y}")
        
        # Configura estilo
        style = ttk.Style()
        style.theme_use('clam')
    
    def create_widgets(self):
        """Cria todos os widgets da interface"""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Título
        title_label = ttk.Label(main_frame, text="Sistema de Gestão de Etiquetas", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Frame de botões (lado esquerdo)
        self.create_buttons_frame(main_frame)
        
        # Frame de dados (lado direito)
        self.create_data_frame(main_frame)
        
        # Frame de status (parte inferior)
        self.create_status_frame(main_frame)
    
    def create_buttons_frame(self, parent):
        """Cria o frame com os botões de ação"""
        buttons_frame = ttk.LabelFrame(parent, text="Ações", padding="10")
        buttons_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N), padx=(0, 10))
        
        # Botão Importar Excel
        import_btn = ttk.Button(buttons_frame, text="📁 Importar Excel", 
                               command=self.import_excel, width=20)
        import_btn.grid(row=0, column=0, pady=3, sticky=tk.W)
        
        # Separador
        ttk.Separator(buttons_frame, orient='horizontal').grid(row=1, column=0, sticky=(tk.W, tk.E), pady=8)
        
        # Frame de pesquisa
        search_frame = ttk.LabelFrame(buttons_frame, text="Pesquisar", padding="5")
        search_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)

        # Campo de pesquisa (apenas OP)
        ttk.Label(search_frame, text="Campo:").grid(row=0, column=0, sticky=tk.W)
        self.search_field = ttk.Combobox(search_frame, values=["op"], 
                        state="readonly", width=10)
        self.search_field.set("op")
        self.search_field.grid(row=0, column=1, padx=5)

        ttk.Label(search_frame, text="Valor:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        self.search_value = ttk.Entry(search_frame, width=15)
        self.search_value.grid(row=1, column=1, padx=5, pady=(5, 0))
        self.search_value.bind('<Return>', lambda event: self.search_data())

        search_btn = ttk.Button(search_frame, text="🔍 Buscar", command=self.search_data, width=20)
        search_btn.grid(row=2, column=0, columnspan=2, pady=3)

        clear_search_btn = ttk.Button(search_frame, text="🗙 Limpar", command=self.clear_search, width=20)
        clear_search_btn.grid(row=3, column=0, columnspan=2, pady=3)
        
        # Separador
        ttk.Separator(buttons_frame, orient='horizontal').grid(row=3, column=0, sticky=(tk.W, tk.E), pady=8)
        
        # Botões de PDF
        pdf_frame = ttk.LabelFrame(buttons_frame, text="Gerar PDF", padding="5")
        pdf_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=5)
        
        labels_btn = ttk.Button(pdf_frame, text="🏷️ Etiquetas", 
                               command=self.generate_labels_pdf, width=20)
        labels_btn.grid(row=0, column=0, pady=3)
        
        list_btn = ttk.Button(pdf_frame, text="📋 Relatório", 
                             command=self.generate_list_pdf, width=20)
        list_btn.grid(row=1, column=0, pady=3)
        
        # Separador
        ttk.Separator(buttons_frame, orient='horizontal').grid(row=5, column=0, sticky=(tk.W, tk.E), pady=8)
        
        # Botões de gerenciamento
        mgmt_frame = ttk.LabelFrame(buttons_frame, text="Gerenciar", padding="5")
        mgmt_frame.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=5)
        
        refresh_btn = ttk.Button(mgmt_frame, text="🔄 Atualizar", 
                                command=self.refresh_data, width=20)
        refresh_btn.grid(row=0, column=0, pady=3)
        
        delete_btn = ttk.Button(mgmt_frame, text="🗑️ Excluir Selecionado", 
                               command=self.delete_selected, width=20)
        delete_btn.grid(row=1, column=0, pady=3)
        
        # Botão para alternar visualização
        self.view_btn = ttk.Button(mgmt_frame, text="📋 Visualização Simples", 
                                  command=self.toggle_view, width=20)
        self.view_btn.grid(row=2, column=0, pady=3)
        
    # Botão de limpar tudo removido por segurança
        
        # Informações do sistema
        info_frame = ttk.LabelFrame(buttons_frame, text="Informações", padding="5")
        info_frame.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=5)
        
        info_btn = ttk.Button(info_frame, text="ℹ️ Sobre", 
                             command=self.show_info, width=20)
        info_btn.grid(row=0, column=0, pady=3)

    # paginação removida
    
    def create_data_frame(self, parent):
        """Cria o frame com a visualização dos dados"""
        data_frame = ttk.LabelFrame(parent, text="Registros", padding="10")
        data_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        data_frame.columnconfigure(0, weight=1)
        data_frame.rowconfigure(0, weight=1)  # A tabela expande
        # data_frame.rowconfigure(2) não tem weight, então paginação fica fixa
        
        # Area para mostrar os dados: pode ser Treeview (lista) ou cards (agrupado)
        # Treeview (lista)
        columns = ("ID", "OP", "Unidade", "Arquivo", "Qtde", "Nome")
        self.tree = ttk.Treeview(data_frame, columns=columns, show="headings", height=20)

        # Configurar colunas
        self.tree.heading("ID", text="ID")
        self.tree.heading("OP", text="OP")
        self.tree.heading("Unidade", text="Unidade")
        self.tree.heading("Arquivo", text="Arquivo")
        self.tree.heading("Qtde", text="Qtde")
        self.tree.heading("Nome", text="Nome")

        self.tree.column("ID", width=50, anchor=tk.CENTER)
        self.tree.column("OP", width=100, anchor=tk.CENTER)
        self.tree.column("Unidade", width=150, anchor=tk.W)
        self.tree.column("Arquivo", width=250, anchor=tk.W)  # Reduzida para dar espaço ao Nome
        self.tree.column("Qtde", width=80, anchor=tk.CENTER)
        self.tree.column("Nome", width=150, anchor=tk.W)

        # Scrollbars
        v_scrollbar = ttk.Scrollbar(data_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(data_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # Grid
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # Canvas para cards (visível apenas em grouped_view)
        self.cards_canvas = tk.Canvas(data_frame)
        self.cards_frame = ttk.Frame(self.cards_canvas)
        self.cards_scroll = ttk.Scrollbar(data_frame, orient=tk.VERTICAL, command=self.cards_canvas.yview)
        self.cards_canvas.configure(yscrollcommand=self.cards_scroll.set)
        self.cards_window = self.cards_canvas.create_window((0,0), window=self.cards_frame, anchor='nw')

        # Bind para ajustar tamanho do frame interno
        def _on_frame_config(event):
            self.cards_canvas.configure(scrollregion=self.cards_canvas.bbox('all'))
        self.cards_frame.bind('<Configure>', _on_frame_config)

        # Inicialmente oculta (mostramos dependendo do modo)
        self.cards_canvas.grid_forget()
        self.cards_scroll.grid_forget()
        
    # Paginação removida (lista completa ou agrupada por OP)
        
        # Menu de contexto
        self.create_context_menu()
    
    def create_context_menu(self):
        """Cria menu de contexto para o treeview"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Excluir", command=self.delete_selected)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Gerar Etiquetas", command=self.generate_labels_pdf)
        self.context_menu.add_command(label="Gerar Relatório", command=self.generate_list_pdf)
        
        # Bind do menu de contexto
        self.tree.bind("<Button-3>", self.show_context_menu)
    
    def create_status_frame(self, parent):
        """Cria o frame de status na parte inferior"""
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        status_frame.columnconfigure(1, weight=1)
        
        # Label de status
        self.status_label = ttk.Label(status_frame, text="Pronto")
        self.status_label.grid(row=0, column=0, sticky=tk.W)
        
        # Label de estatísticas
        self.stats_label = ttk.Label(status_frame, text="")
        self.stats_label.grid(row=0, column=1, sticky=tk.E)
        
        # Atualiza estatísticas
        self.update_stats()
    
    def import_excel(self):
        """Abre diálogo para importar arquivo Excel"""
        file_path = filedialog.askopenfilename(
            title="Selecionar arquivo Excel",
            filetypes=[
                ("Arquivos Excel", "*.xlsx *.xls"),
                ("Todos os arquivos", "*.*")
            ]
        )
        
        if file_path:
            def import_worker():
                try:
                    self.show_loading("Validando arquivo Excel...")
                    self.root.after(100, lambda: self.update_loading_message("Importando dados..."))
                    
                    success = self.controller.import_excel_file(file_path)
                    
                    self.root.after(0, lambda: self._finish_import(success))
                except Exception as e:
                    self.root.after(0, lambda: self._finish_import(False, str(e)))
            
            # Executa importação em thread separada
            thread = threading.Thread(target=import_worker, daemon=True)
            thread.start()
    
    def _finish_import(self, success, error_msg=None):
        """Finaliza a importação do Excel"""
        self.hide_loading()
        
        if success:
            self.refresh_data()
            self.status_label.config(text="Excel importado com sucesso")
            messagebox.showinfo("Sucesso", "Arquivo Excel importado com sucesso!")
        else:
            error_text = f"Falha ao importar Excel: {error_msg}" if error_msg else "Falha ao importar Excel"
            self.status_label.config(text=error_text)
            messagebox.showerror("Erro", error_text)
    
    def search_data(self):
        """Executa a pesquisa com loading"""
        def search_worker():
            try:
                # Força pesquisa apenas por OP
                campo = 'op'
                valor = self.search_value.get().strip()
                
                self.root.after(0, lambda: self.show_loading("Pesquisando registros..."))
                
                if not valor:
                    # Sem paginação: carrega todos os registros
                    self.current_data = self.controller.get_all_registros()
                    self.filtered_data = self.current_data
                else:
                    # Executa pesquisa por campo
                    self.filtered_data = self.controller.search_registros(campo, valor)
                
                self.root.after(0, lambda: self._finish_search())
            except Exception as e:
                self.root.after(0, lambda: self._finish_search_error(str(e)))
        
        # Se já está carregando, não faz nada
        if self.is_loading:
            return
            
        thread = threading.Thread(target=search_worker, daemon=True)
        thread.start()
    
    def _finish_search(self):
        """Finaliza a pesquisa"""
        self.hide_loading()
        self.update_tree_data(self.filtered_data)
    # Sem paginação
        self.status_label.config(text=f"Encontrados {len(self.filtered_data)} registros")
    
    def _finish_search_error(self, error_msg):
        """Finaliza pesquisa com erro"""
        self.hide_loading()
        self.status_label.config(text=f"Erro na pesquisa: {error_msg}")
        messagebox.showerror("Erro", f"Erro na pesquisa: {error_msg}")
    
    def clear_search(self):
        """Limpa a pesquisa com loading"""
        def clear_worker():
            try:
                self.root.after(0, lambda: self.show_loading("Limpando pesquisa..."))
                
                # Retorna para dados paginados
                # Agora carregamos todos os registros
                self.current_data = self.controller.get_all_registros()
                self.filtered_data = self.current_data
                
                self.root.after(0, lambda: self._finish_clear_search())
            except Exception as e:
                self.root.after(0, lambda: self._finish_clear_search_error(str(e)))
        
        # Limpa o campo de pesquisa
        self.search_value.delete(0, tk.END)
        
        # Se já está carregando, não faz nada
        if self.is_loading:
            return
            
        thread = threading.Thread(target=clear_worker, daemon=True)
        thread.start()
    
    def _finish_clear_search(self):
        """Finaliza a limpeza da pesquisa"""
        self.hide_loading()
        self.update_tree_data(self.filtered_data)
    # Sem paginação
        self.status_label.config(text="Pesquisa limpa")
    
    def _finish_clear_search_error(self, error_msg):
        """Finaliza limpeza com erro"""
        self.hide_loading()
        self.status_label.config(text=f"Erro ao limpar pesquisa: {error_msg}")
    
    def refresh_data(self):
        """Atualiza os dados da tela"""
        def refresh_worker():
            try:
                self.root.after(0, lambda: self.show_loading("Carregando dados do banco..."))
                # Se há filtro (pesquisa), mantemos comportamento atual (busca completa)
                # Carrega todos os registros (remoção de paginação)
                self.current_data = self.controller.get_all_registros()
                self.filtered_data = self.current_data
                
                self.root.after(0, lambda: self._finish_refresh())
            except Exception as e:
                self.root.after(0, lambda: self._finish_refresh_error(str(e)))
        
        # Se já está carregando, não faz nada
        if self.is_loading:
            return
            
        thread = threading.Thread(target=refresh_worker, daemon=True)
        thread.start()
    
    def _finish_refresh(self):
        """Finaliza o refresh dos dados"""
        self.hide_loading()
        self.update_tree_data(self.current_data)
        self.update_stats()
        self.status_label.config(text=f"Dados atualizados - {len(self.current_data)} registros")
    
    def _finish_refresh_error(self, error_msg):
        """Finaliza refresh com erro"""
        self.hide_loading()
        self.status_label.config(text=f"Erro ao carregar dados: {error_msg}")
        messagebox.showerror("Erro", f"Erro ao carregar dados: {error_msg}")
    
    def toggle_view(self):
        """Alterna entre visualização agrupada e simples"""
        self.grouped_view = not self.grouped_view
        
        # Atualiza o texto do botão
        if self.grouped_view:
            self.view_btn.config(text="📋 Visualização Simples")
        else:
            self.view_btn.config(text="🗂️ Agrupar por OP")
        
        # Atualiza a visualização
        self.update_tree_data(self.filtered_data)

    def update_tree_data(self, data):
        """Atualiza os dados do treeview com opção de agrupamento por OP"""
        # Limpa dados existentes
        # Se modo agrupado: mostrar cards com resumo de OPs
        if self.grouped_view:
            # Oculta treeview e mostra canvas de cards
            try:
                self.tree.grid_remove()
                v = self.tree.master.nametowidget(self.tree.winfo_parent()).grid_slaves(row=0, column=1)
            except Exception:
                pass
            self.cards_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            self.cards_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))

            # Limpa cards existentes
            for child in self.cards_frame.winfo_children():
                child.destroy()

            # Pega resumo de grupos (mais leve que trazer todos os registros)
            # Se foi fornecido 'data' (por ex. resultado de busca), calculamos resumo a partir dele.
            groups = []
            try:
                if data:
                    ops = {}
                    for registro in data:
                        op = registro[1]
                        if op not in ops:
                            ops[op] = {'total_itens': 0, 'total_qtde': 0}
                        ops[op]['total_itens'] += 1
                        try:
                            ops[op]['total_qtde'] += int(registro[4])
                        except Exception:
                            pass
                    # transforma em lista de tuplas ordenada
                    groups = [(op, v['total_itens'], v['total_qtde']) for op, v in ops.items()]
                    groups.sort(reverse=True)
                else:
                    groups = self.controller.get_groups_summary()
            except Exception:
                groups = []

            if not groups:
                lbl = ttk.Label(self.cards_frame, text="Nenhum grupo disponível", padding=10)
                lbl.grid()
                return

            # Renderiza cada grupo como um card
            # Limpa mapeamento
            self.card_widgets = {}

            for idx, grp in enumerate(groups):
                op, total_itens, total_qtde = grp

                # Usar tk.Frame para permitir alteração de background
                card = tk.Frame(self.cards_frame, relief=tk.RIDGE, bd=1, padx=12, pady=10, bg="#FFFFFF", highlightthickness=1, highlightbackground="#E0E0E0")

                # Armazena widget
                self.card_widgets[op] = card

                # Configura colunas internas para permitir expansão
                card.columnconfigure(0, weight=1)
                card.columnconfigure(1, weight=0)

                # Conteúdo do card (lado esquerdo)
                # Garantir que o nome da OP não seja cortado: wrap e alinhamento
                lbl_op = ttk.Label(card, text=f"OP: {op}", font=("Arial", 11, "bold"), anchor='w', justify='left')
                lbl_itens = ttk.Label(card, text=f"Itens: {total_itens}")
                lbl_qtde = ttk.Label(card, text=f"Qtde total: {total_qtde}")

                # Campo de status no card: calculado a partir de total_itens/total_qtde
                if total_itens == 0:
                    status_text = 'Vazio'
                    status_bg = '#D3D3D3'  # cinza
                    status_fg = '#333333'
                elif total_qtde > 100:
                    status_text = 'Alto'
                    status_bg = '#FFA500'  # laranja
                    status_fg = '#000000'
                else:
                    status_text = 'OK'
                    status_bg = '#4CAF50'  # verde
                    status_fg = '#FFFFFF'

                # Badge de status (label com background colorido)
                lbl_status = tk.Label(card, text=status_text, bg=status_bg, fg=status_fg, padx=6, pady=2, font=("Arial", 9, "bold"))

                lbl_op.grid(row=0, column=0, sticky=(tk.W, tk.E))
                lbl_itens.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(6,0))
                lbl_qtde.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(2,0))
                lbl_status.grid(row=3, column=0, sticky=(tk.W, tk.W), pady=(6,0))

                # Armazena referencias de labels de texto para alternar cor quando o card for selecionado
                try:
                    # lista de widgets que devem ficar brancos quando selecionados
                    card._text_widgets = [lbl_op, lbl_itens, lbl_qtde, lbl_status]
                    # guarda cor padrão de cada widget
                    card._text_default = [w.cget('foreground') if hasattr(w, 'cget') else '#000000' for w in card._text_widgets]
                except Exception:
                    card._text_widgets = []
                    card._text_default = []

                # Frame de botões (lado direito) com botões padronizados
                btns_frame = ttk.Frame(card)
                # Reservar largura para os botões para não diminuir espaço do nome da OP
                try:
                    btns_frame.configure(width=120)
                except Exception:
                    pass
                btns_frame.grid(row=0, column=1, rowspan=3, padx=(12,0), sticky=(tk.N, tk.E))

                btn_w = 16
                btn_details = ttk.Button(btns_frame, text="Ver detalhes", command=lambda op=op: self.on_card_click(op), width=btn_w)
                btn_download = ttk.Button(btns_frame, text="Etiquetas", command=lambda op=op: self.download_op_labels(op), width=btn_w)
                btn_report = ttk.Button(btns_frame, text="Relatório", command=lambda op=op: self.download_op_report(op), width=btn_w)

                # Empilha os botões verticalmente com espaçamento uniforme
                btn_details.pack(fill=tk.X, pady=(0,4))
                btn_download.pack(fill=tk.X, pady=(0,4))
                btn_report.pack(fill=tk.X)

                # Definir wraplength dinâmico razoável para o nome (não cortar)
                try:
                    lbl_op.configure(wraplength=180)
                except Exception:
                    pass

                # Efeito de hover para melhor visual
                def on_enter(e, w=card):
                    try:
                        if getattr(self, 'selected_op', None) != op:
                            w.configure(bg="#F5FBFF")
                    except Exception:
                        pass

                def on_leave(e, w=card):
                    try:
                        if getattr(self, 'selected_op', None) != op:
                            w.configure(bg="#FFFFFF")
                    except Exception:
                        pass

                card.bind("<Enter>", on_enter)
                card.bind("<Leave>", on_leave)
                card.bind("<Button-1>", lambda e, op=op: self.select_card(op))
                lbl_op.bind("<Button-1>", lambda e, op=op: self.select_card(op))
                lbl_itens.bind("<Button-1>", lambda e, op=op: self.select_card(op))
                lbl_qtde.bind("<Button-1>", lambda e, op=op: self.select_card(op))
                lbl_status.bind("<Button-1>", lambda e, op=op: self.select_card(op))

                # Adiciona ao frame (layout será organizado por layout_cards)
                card.grid(row=0, column=idx, padx=8, pady=8, sticky=(tk.N, tk.S, tk.E, tk.W))
                # Permite que o conteúdo do card expanda horizontalmente
                card.update_idletasks()

            # Forçar layout responsivo
            self.cards_frame.update_idletasks()
            self.layout_cards()

            # Vincula redimensionamento para recomputar layout
            self.cards_canvas.bind('<Configure>', lambda e: self.layout_cards())

            return

        # Caso contrário, mostra a lista (treeview)
        # Limpa tree
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Exibe registros na treeview
        for registro in data:
            self.tree.insert("", tk.END, values=registro)
    
    def update_stats(self):
        """Atualiza as estatísticas"""
        stats = self.controller.get_statistics()
        stats_text = (f"Registros: {stats['total_registros']} | "
                     f"OPs: {stats['total_ops']} | "
                     f"Unidades: {stats['total_unidades']} | "
                     f"Qtde Total: {stats['total_quantidade']}")
        self.stats_label.config(text=stats_text)

    def on_card_click(self, op: str):
        """Abre modal e carrega os registros da OP sob demanda"""
        modal = tk.Toplevel(self.root)
        modal.title(f"Detalhes da OP {op}")
        # Tamanho do modal
        modal_w, modal_h = 800, 400
        # Centraliza modal em relação à janela principal
        self.root.update_idletasks()
        rx = self.root.winfo_x()
        ry = self.root.winfo_y()
        rw = self.root.winfo_width()
        rh = self.root.winfo_height()
        mx = rx + (rw // 2) - (modal_w // 2)
        my = ry + (rh // 2) - (modal_h // 2)
        modal.geometry(f"{modal_w}x{modal_h}+{mx}+{my}")
        modal.transient(self.root)
        modal.grab_set()

        # Frame e progress
        frame = ttk.Frame(modal, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        loading_lbl = ttk.Label(frame, text="Carregando detalhes...")
        loading_lbl.pack(pady=10)

        records_tree = None

        def load_details():
            nonlocal records_tree
            try:
                # Busca apenas os registros para essa OP
                registros = self.controller.get_registros_by_op(op)

                # Remove label e cria tree
                loading_lbl.pack_forget()

                cols = ("ID", "OP", "Unidade", "Arquivo", "Qtde", "Nome")
                records_tree = ttk.Treeview(frame, columns=cols, show='headings')
                for c in cols:
                    records_tree.heading(c, text=c)
                    records_tree.column(c, width=100, anchor=tk.W)

                vs = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=records_tree.yview)
                records_tree.configure(yscrollcommand=vs.set)
                records_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
                vs.pack(fill=tk.Y, side=tk.RIGHT)

                for r in registros:
                    records_tree.insert('', tk.END, values=r)

            except Exception as e:
                loading_lbl.config(text=f"Erro ao carregar detalhes: {e}")

        # Carrega em thread para não bloquear UI
        thread = threading.Thread(target=load_details, daemon=True)
        thread.start()

    def layout_cards(self):
        """Organiza os cards para ocupar o espaço disponível de forma responsiva"""
        try:
            width = self.cards_canvas.winfo_width() or self.cards_canvas.winfo_reqwidth()
            if width <= 0:
                return

            # força o frame interno a acompanhar a largura do canvas
            try:
                self.cards_canvas.itemconfig(self.cards_window, width=width)
            except Exception:
                pass

            # calcula número de colunas baseado na largura média do card (~220px)
            card_min_w = 220
            # Por padrão queremos 2 OPs por linha. Se a largura não comportar 2, cai para 1.
            desired_cols = 2
            if width < card_min_w * desired_cols:
                cols = max(1, width // card_min_w)
            else:
                cols = desired_cols

            children = self.cards_frame.winfo_children()
            # Não ter mais colunas do que cards
            cols = max(1, min(cols, len(children)))

            for idx, child in enumerate(children):
                r = idx // cols
                c = idx % cols
                child.grid_configure(row=r, column=c, sticky=(tk.N, tk.S, tk.E, tk.W))

            # Ajusta colunas para expandir igualmente
            for c in range(cols):
                self.cards_frame.columnconfigure(c, weight=1)
        except Exception:
            pass

    def select_card(self, op: str):
        """Marca visualmente o card selecionado e guarda estado"""
        # Desmarca anterior
        if self.selected_op and self.selected_op in self.card_widgets:
            prev = self.card_widgets[self.selected_op]
            try:
                prev.configure(bg="#FFFFFF", bd=1, highlightbackground="#E0E0E0")
                # restaurar cor dos textos
                for w, col in zip(getattr(prev, '_text_widgets', []), getattr(prev, '_text_default', [])):
                    try:
                        w.configure(foreground=col)
                    except Exception:
                        pass
            except Exception:
                prev.configure(bg="#FFFFFF")

        # Marca novo
        widget = self.card_widgets.get(op)
        if widget:
            try:
                widget.configure(bg="#D9EFFF", bd=2, highlightbackground="#4FA3FF")
                # pintar textos como branco
                for w in getattr(widget, '_text_widgets', []):
                    try:
                        w.configure(foreground='#FFFFFF')
                    except Exception:
                        pass
            except Exception:
                widget.configure(bg="#D9EFFF")
            self.selected_op = op

    def download_op_labels(self, op: str):
        """Gera/baixa etiquetas apenas para uma OP específica"""
        try:
            registros = self.controller.get_registros_by_op(op)
            if not registros:
                messagebox.showwarning("Aviso", f"Nenhum registro encontrado para OP {op}")
                return

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_name = f"etiquetas_{op}_{timestamp}.pdf"
            file_path = filedialog.asksaveasfilename(
                title=f"Salvar PDF de Etiquetas - {op}",
                defaultextension=".pdf",
                initialfile=default_name,
                filetypes=[("Arquivos PDF", "*.pdf")]
            )

            if not file_path:
                return

            def worker():
                try:
                    self.root.after(0, lambda: self.show_loading("Gerando etiquetas para OP..."))
                    self.root.after(100, lambda: self.update_loading_message("Gerando PDF de etiquetas..."))
                    success = self.controller.generate_labels_pdf(registros, file_path)
                    self.root.after(0, lambda: self._finish_generate_labels(success, file_path))
                except Exception as e:
                    self.root.after(0, lambda: self._finish_generate_labels(False, file_path, str(e)))

            thread = threading.Thread(target=worker, daemon=True)
            thread.start()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar etiquetas para OP {op}: {e}")

    def download_op_report(self, op: str):
        """Gera/baixa relatório (lista) apenas para uma OP específica"""
        try:
            registros = self.controller.get_registros_by_op(op)
            if not registros:
                messagebox.showwarning("Aviso", f"Nenhum registro encontrado para OP {op}")
                return

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_name = f"relatorio_{op}_{timestamp}.pdf"
            file_path = filedialog.asksaveasfilename(
                title=f"Salvar Relatório PDF - {op}",
                defaultextension=".pdf",
                initialfile=default_name,
                filetypes=[("Arquivos PDF", "*.pdf")]
            )

            if not file_path:
                return

            def worker():
                try:
                    self.root.after(0, lambda: self.show_loading("Gerando relatório para OP..."))
                    self.root.after(100, lambda: self.update_loading_message("Gerando PDF do relatório..."))
                    success = self.controller.generate_list_pdf(registros, file_path)
                    self.root.after(0, lambda: self._finish_generate_report(success, file_path))
                except Exception as e:
                    self.root.after(0, lambda: self._finish_generate_report(False, file_path, str(e)))

            thread = threading.Thread(target=worker, daemon=True)
            thread.start()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar relatório para OP {op}: {e}")
    
    def show_loading(self, message="Carregando..."):
        """Mostra janela de loading"""
        if self.is_loading:
            return
            
        self.is_loading = True
        
        # Cria janela de loading
        self.loading_window = tk.Toplevel(self.root)
        self.loading_window.title("Aguarde")
        self.loading_window.geometry("300x120")
        self.loading_window.resizable(False, False)
        self.loading_window.transient(self.root)
        self.loading_window.grab_set()
        
        # Centraliza a janela de loading
        self.loading_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 150
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 60
        self.loading_window.geometry(f"300x120+{x}+{y}")
        
        # Frame principal
        frame = ttk.Frame(self.loading_window, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Progressbar animada
        self.progress = ttk.Progressbar(frame, mode='indeterminate')
        self.progress.pack(pady=(10, 5))
        self.progress.start(10)
        
        # Label de mensagem
        self.loading_label = ttk.Label(frame, text=message, font=("Arial", 10))
        self.loading_label.pack(pady=(5, 10))
        
        # Desabilita botões principais
        self.disable_buttons(True)
        
        self.root.update()
    
    def hide_loading(self):
        """Esconde janela de loading"""
        if not self.is_loading:
            return
            
        self.is_loading = False
        
        if self.loading_window:
            self.progress.stop()
            self.loading_window.grab_release()
            self.loading_window.destroy()
            self.loading_window = None
        
        # Reabilita botões principais
        self.disable_buttons(False)
        
        self.root.update()
    
    def update_loading_message(self, message):
        """Atualiza mensagem do loading"""
        if self.is_loading and self.loading_label:
            self.loading_label.config(text=message)
            self.root.update()
    
    def disable_buttons(self, disabled=True):
        """Desabilita/habilita botões durante loading"""
        state = 'disabled' if disabled else 'normal'
        
        # Encontra todos os botões na interface
        for widget in self.root.winfo_children():
            self._disable_widget_recursive(widget, state)
    
    def _disable_widget_recursive(self, widget, state):
        """Recursivamente desabilita widgets"""
        try:
            if isinstance(widget, (ttk.Button, tk.Button)):
                widget.config(state=state)
            elif hasattr(widget, 'winfo_children'):
                for child in widget.winfo_children():
                    self._disable_widget_recursive(child, state)
        except:
            pass
    
    def get_selected_records(self):
        """Retorna os registros selecionados (apenas registros filhos, não nós de OP)"""
        selection = self.tree.selection()
        if not selection:
            return []
        
        selected_records = []
        for item in selection:
            # Verifica se é um nó pai (OP) ou filho (registro)
            tags = self.tree.item(item, "tags")
            
            if "op_parent" in tags:
                # Se selecionou uma OP, inclui todos os filhos dela
                children = self.tree.get_children(item)
                for child in children:
                    values = self.tree.item(child, "values")
                    if values[0]:  # Se tem ID (não é linha de total)
                        try:
                            nome_val = values[5] if len(values) > 5 else ""
                            record = (int(values[0]), values[1], values[2], values[3], int(values[4]), nome_val)
                            selected_records.append(record)
                        except (ValueError, IndexError):
                            continue
            elif "op_child" in tags:
                # Se selecionou um registro específico
                values = self.tree.item(item, "values")
                if values[0]:  # Se tem ID
                    try:
                        nome_val = values[5] if len(values) > 5 else ""
                        record = (int(values[0]), values[1], values[2], values[3], int(values[4]), nome_val)
                        selected_records.append(record)
                    except (ValueError, IndexError):
                        continue
        
        return selected_records
    
    def generate_labels_pdf(self):
        """Gera PDF com etiquetas dos registros selecionados"""
        selected = self.get_selected_records()
        
        if not selected:
            # Se nada selecionado, usar todos os dados filtrados
            if not self.filtered_data:
                messagebox.showwarning("Aviso", "Nenhum registro disponível!")
                return
            
            resposta = messagebox.askyesno(
                "Gerar Etiquetas",
                f"Nenhum registro selecionado.\n\n"
                f"Deseja gerar etiquetas para todos os {len(self.filtered_data)} registros visíveis?"
            )
            
            if not resposta:
                return
            
            selected = self.filtered_data
        
        # Diálogo para salvar arquivo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"etiquetas_{timestamp}.pdf"
        
        file_path = filedialog.asksaveasfilename(
            title="Salvar PDF de Etiquetas",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        
        if file_path:
            def generate_worker():
                try:
                    self.root.after(0, lambda: self.show_loading("Preparando etiquetas..."))
                    self.root.after(100, lambda: self.update_loading_message("Gerando PDF de etiquetas..."))
                    
                    success = self.controller.generate_labels_pdf(selected, file_path)
                    
                    self.root.after(0, lambda: self._finish_generate_labels(success, file_path))
                except Exception as e:
                    self.root.after(0, lambda: self._finish_generate_labels(False, file_path, str(e)))
            
            thread = threading.Thread(target=generate_worker, daemon=True)
            thread.start()
    
    def _finish_generate_labels(self, success, file_path, error_msg=None):
        """Finaliza a geração de etiquetas"""
        self.hide_loading()
        
        if success:
            self.status_label.config(text="PDF de etiquetas gerado com sucesso")
            result = messagebox.askyesno("Sucesso", 
                f"PDF de etiquetas gerado com sucesso!\n\nDeseja abrir o arquivo?\n{file_path}")
            if result:
                os.startfile(file_path)
        else:
            error_text = f"Falha ao gerar PDF: {error_msg}" if error_msg else "Falha ao gerar PDF de etiquetas"
            self.status_label.config(text=error_text)
            messagebox.showerror("Erro", error_text)
    
    def generate_list_pdf(self):
        """Gera PDF com relatório dos registros selecionados"""
        selected = self.get_selected_records()
        
        if not selected:
            # Se nada selecionado, usar todos os dados filtrados
            if not self.filtered_data:
                messagebox.showwarning("Aviso", "Nenhum registro disponível!")
                return
            selected = self.filtered_data
        
        # Diálogo para salvar arquivo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"relatorio_{timestamp}.pdf"
        
        file_path = filedialog.asksaveasfilename(
            title="Salvar Relatório PDF",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        
        if file_path:
            def generate_worker():
                try:
                    self.root.after(0, lambda: self.show_loading("Preparando relatório..."))
                    self.root.after(100, lambda: self.update_loading_message("Gerando PDF do relatório..."))
                    
                    success = self.controller.generate_list_pdf(selected, file_path)
                    
                    self.root.after(0, lambda: self._finish_generate_report(success, file_path))
                except Exception as e:
                    self.root.after(0, lambda: self._finish_generate_report(False, file_path, str(e)))
            
            thread = threading.Thread(target=generate_worker, daemon=True)
            thread.start()
    
    def _finish_generate_report(self, success, file_path, error_msg=None):
        """Finaliza a geração do relatório"""
        self.hide_loading()
        
        if success:
            self.status_label.config(text="Relatório PDF gerado com sucesso")
            result = messagebox.askyesno("Sucesso", 
                f"Relatório PDF gerado com sucesso!\n\nDeseja abrir o arquivo?\n{file_path}")
            if result:
                os.startfile(file_path)
        else:
            error_text = f"Falha ao gerar relatório: {error_msg}" if error_msg else "Falha ao gerar relatório PDF"
            self.status_label.config(text=error_text)
            messagebox.showerror("Erro", error_text)
    
    def delete_selected(self):
        """Exclui os registros selecionados"""
        selected = self.get_selected_records()
        
        if not selected:
            messagebox.showwarning("Aviso", "Nenhum registro selecionado!")
            return
        
        resposta = messagebox.askyesno(
            "Confirmar Exclusão",
            f"Tem certeza que deseja excluir {len(selected)} registro(s) selecionado(s)?"
        )
        
        if resposta:
            sucessos = 0
            for record in selected:
                if self.controller.delete_registro(record[0]):  # record[0] é o ID
                    sucessos += 1
            
            if sucessos > 0:
                self.refresh_data()
                self.status_label.config(text=f"{sucessos} registro(s) excluído(s)")
            else:
                self.status_label.config(text="Falha ao excluir registros")
    
    
    def show_context_menu(self, event):
        """Mostra o menu de contexto"""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
    
    def show_info(self):
        """Mostra informações sobre o sistema"""
        pdf_info = self.controller.get_pdf_info()
        
        info_text = f"""Sistema de Gestão de Etiquetas
        
Versão: 1.0
Desenvolvido para gestão de ordens de produção e geração de etiquetas.

Configurações de Etiquetas:
• Tamanho: {pdf_info['label_width_mm']:.0f}mm x {pdf_info['label_height_mm']:.0f}mm
• Etiquetas por linha: {pdf_info['labels_per_row']}
• Etiquetas por coluna: {pdf_info['labels_per_col']}
• Etiquetas por página: {pdf_info['labels_per_page']}

Estrutura do Excel:
• A1 = OP (ordem de produção)
• A2+ = arquivos dessa OP
• B1 = unidade (nome da unidade)
• B2+ = quantidade

Funcionalidades:
• Importação de dados do Excel
• Busca e filtros
• Geração de etiquetas em PDF
• Relatórios em PDF
• Gerenciamento de registros"""
        
        messagebox.showinfo("Sobre o Sistema", info_text)
    
    def run(self):
        """Inicia a aplicação"""
        self.root.mainloop()

def main():
    """Função principal"""
    app = EtiquetaView()
    app.run()

if __name__ == "__main__":
    main()
