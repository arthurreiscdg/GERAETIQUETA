import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from controller.etiqueta_controller import EtiquetaController
import os
from datetime import datetime

class EtiquetaView:
    def __init__(self):
        """Inicializa a interface gráfica"""
        self.controller = EtiquetaController()
        self.root = tk.Tk()
        self.current_data = []
        self.filtered_data = []
        
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
        import_btn.grid(row=0, column=0, pady=5, sticky=tk.W)
        
        # Separador
        ttk.Separator(buttons_frame, orient='horizontal').grid(row=1, column=0, sticky=(tk.W, tk.E), pady=10)
        
        # Frame de pesquisa
        search_frame = ttk.LabelFrame(buttons_frame, text="Pesquisar", padding="5")
        search_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # Campo de pesquisa
        ttk.Label(search_frame, text="Campo:").grid(row=0, column=0, sticky=tk.W)
        self.search_field = ttk.Combobox(search_frame, values=["op", "unidade", "arquivos"], 
                                        state="readonly", width=10)
        self.search_field.set("op")
        self.search_field.grid(row=0, column=1, padx=5)
        
        ttk.Label(search_frame, text="Valor:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        self.search_value = ttk.Entry(search_frame, width=15)
        self.search_value.grid(row=1, column=1, padx=5, pady=(5, 0))
        self.search_value.bind('<KeyRelease>', self.on_search_change)
        
        search_btn = ttk.Button(search_frame, text="🔍 Buscar", command=self.search_data)
        search_btn.grid(row=2, column=0, columnspan=2, pady=5)
        
        clear_search_btn = ttk.Button(search_frame, text="🗙 Limpar", command=self.clear_search)
        clear_search_btn.grid(row=3, column=0, columnspan=2, pady=(0, 5))
        
        # Separador
        ttk.Separator(buttons_frame, orient='horizontal').grid(row=3, column=0, sticky=(tk.W, tk.E), pady=10)
        
        # Botões de PDF
        pdf_frame = ttk.LabelFrame(buttons_frame, text="Gerar PDF", padding="5")
        pdf_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=5)
        
        labels_btn = ttk.Button(pdf_frame, text="🏷️ Etiquetas", 
                               command=self.generate_labels_pdf, width=18)
        labels_btn.grid(row=0, column=0, pady=2)
        
        list_btn = ttk.Button(pdf_frame, text="📋 Relatório", 
                             command=self.generate_list_pdf, width=18)
        list_btn.grid(row=1, column=0, pady=2)
        
        # Separador
        ttk.Separator(buttons_frame, orient='horizontal').grid(row=5, column=0, sticky=(tk.W, tk.E), pady=10)
        
        # Botões de gerenciamento
        mgmt_frame = ttk.LabelFrame(buttons_frame, text="Gerenciar", padding="5")
        mgmt_frame.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=5)
        
        refresh_btn = ttk.Button(mgmt_frame, text="🔄 Atualizar", 
                                command=self.refresh_data, width=18)
        refresh_btn.grid(row=0, column=0, pady=2)
        
        delete_btn = ttk.Button(mgmt_frame, text="🗑️ Excluir Selecionado", 
                               command=self.delete_selected, width=18)
        delete_btn.grid(row=1, column=0, pady=2)
        
        clear_all_btn = ttk.Button(mgmt_frame, text="⚠️ Limpar Tudo", 
                                  command=self.clear_all_data, width=18)
        clear_all_btn.grid(row=2, column=0, pady=2)
        
        # Informações do sistema
        info_frame = ttk.LabelFrame(buttons_frame, text="Informações", padding="5")
        info_frame.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=5)
        
        info_btn = ttk.Button(info_frame, text="ℹ️ Sobre", 
                             command=self.show_info, width=18)
        info_btn.grid(row=0, column=0, pady=2)
    
    def create_data_frame(self, parent):
        """Cria o frame com a visualização dos dados"""
        data_frame = ttk.LabelFrame(parent, text="Registros", padding="10")
        data_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        data_frame.columnconfigure(0, weight=1)
        data_frame.rowconfigure(0, weight=1)
        
        # Treeview para mostrar os dados
        columns = ("ID", "OP", "Unidade", "Arquivo", "Qtde")
        self.tree = ttk.Treeview(data_frame, columns=columns, show="headings", height=20)
        
        # Configurar colunas
        self.tree.heading("ID", text="ID")
        self.tree.heading("OP", text="OP")
        self.tree.heading("Unidade", text="Unidade")
        self.tree.heading("Arquivo", text="Arquivo")
        self.tree.heading("Qtde", text="Qtde")
        
        self.tree.column("ID", width=50, anchor=tk.CENTER)
        self.tree.column("OP", width=100, anchor=tk.CENTER)
        self.tree.column("Unidade", width=150, anchor=tk.W)
        self.tree.column("Arquivo", width=300, anchor=tk.W)
        self.tree.column("Qtde", width=80, anchor=tk.CENTER)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(data_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(data_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
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
            self.status_label.config(text="Importando Excel...")
            self.root.update()
            
            success = self.controller.import_excel_file(file_path)
            
            if success:
                self.refresh_data()
                self.status_label.config(text="Excel importado com sucesso")
            else:
                self.status_label.config(text="Falha ao importar Excel")
    
    def search_data(self):
        """Executa a pesquisa"""
        campo = self.search_field.get()
        valor = self.search_value.get().strip()
        
        if not valor:
            self.filtered_data = self.current_data
        else:
            self.filtered_data = self.controller.search_registros(campo, valor)
        
        self.update_tree_data(self.filtered_data)
        self.status_label.config(text=f"Encontrados {len(self.filtered_data)} registros")
    
    def on_search_change(self, event):
        """Executa pesquisa em tempo real"""
        self.search_data()
    
    def clear_search(self):
        """Limpa a pesquisa"""
        self.search_value.delete(0, tk.END)
        self.filtered_data = self.current_data
        self.update_tree_data(self.filtered_data)
        self.status_label.config(text="Pesquisa limpa")
    
    def refresh_data(self):
        """Atualiza os dados da tela"""
        self.status_label.config(text="Carregando dados...")
        self.root.update()
        
        self.current_data = self.controller.get_all_registros()
        self.filtered_data = self.current_data
        self.update_tree_data(self.current_data)
        self.update_stats()
        
        self.status_label.config(text=f"Dados atualizados - {len(self.current_data)} registros")
    
    def update_tree_data(self, data):
        """Atualiza os dados do treeview"""
        # Limpa dados existentes
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Adiciona novos dados
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
    
    def get_selected_records(self):
        """Retorna os registros selecionados"""
        selection = self.tree.selection()
        if not selection:
            return []
        
        selected_records = []
        for item in selection:
            values = self.tree.item(item, "values")
            # Converte para tupla com tipos corretos
            record = (int(values[0]), values[1], values[2], values[3], int(values[4]))
            selected_records.append(record)
        
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
            self.status_label.config(text="Gerando PDF de etiquetas...")
            self.root.update()
            
            success = self.controller.generate_labels_pdf(selected, file_path)
            
            if success:
                self.status_label.config(text="PDF de etiquetas gerado com sucesso")
            else:
                self.status_label.config(text="Falha ao gerar PDF de etiquetas")
    
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
            self.status_label.config(text="Gerando relatório PDF...")
            self.root.update()
            
            success = self.controller.generate_list_pdf(selected, file_path)
            
            if success:
                self.status_label.config(text="Relatório PDF gerado com sucesso")
            else:
                self.status_label.config(text="Falha ao gerar relatório PDF")
    
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
    
    def clear_all_data(self):
        """Limpa todos os dados"""
        if self.controller.clear_all_data():
            self.refresh_data()
    
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
