from model.database import Database
from service.excel_service import ExcelService
from service.pdf_service import PDFService
from typing import List, Tuple, Optional
from tkinter import messagebox
import os
import logging

# Configurar logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class EtiquetaController:
    def __init__(self):
        """Inicializa o controller com os serviços necessários"""
        # Permite configurar o banco via variável de ambiente SQLITE_DB_URL
        db_url = os.environ.get('SQLITE_DB_URL', 'sqlitecloud://cv0idhxxhk.g2.sqlite.cloud:8860/auth.sqlitecloud?apikey=4gtJpnQlCzrAfmGgn9QOdDrFDvalmk3APBcawzNvssc')
        self.database = Database(db_url)
        self.excel_service = ExcelService()
        self.pdf_service = PDFService()
    
    def import_excel_file(self, file_path: str) -> bool:
        """
        Importa dados de um arquivo Excel para o banco de dados
        
        Args:
            file_path (str): Caminho para o arquivo Excel
            
        Returns:
            bool: True se importado com sucesso, False caso contrário
        """
        try:
            # Lê os dados do Excel
            registros = self.excel_service.read_excel_data(file_path)
            
            if registros is None:
                return False
            
            if not registros:
                messagebox.showwarning("Aviso", "Nenhum registro válido encontrado no arquivo!")
                return False
            
            # Valida qualidade dos dados
            qualidade = self.excel_service.validate_data_quality(registros)
            
            if qualidade['registros_com_problemas'] > 0:
                problemas_text = "\n".join(qualidade['problemas'][:5])
                if qualidade['registros_com_problemas'] > 5:
                    problemas_text += f"\n... e mais {qualidade['registros_com_problemas'] - 5} problemas"
                
                resposta = messagebox.askyesno(
                    "Dados com Problemas",
                    f"Encontrados {qualidade['registros_com_problemas']} registros com problemas:\n\n" +
                    problemas_text + "\n\n" +
                    f"Registros válidos: {qualidade['registros_validos']}\n\n" +
                    "Deseja continuar mesmo assim?"
                )
                
                if not resposta:
                    return False
            
            # Verifica duplicatas em vez de perguntar sobre limpar dados
            verificacao_duplicatas = self.database.check_duplicates(registros)
            
            if verificacao_duplicatas['total_duplicatas'] > 0:
                duplicatas_info = []
                for dup in verificacao_duplicatas['duplicatas'][:5]:  # Mostra até 5 exemplos
                    novo = dup['novo']
                    existente = dup['existente']
                    status_qtde = "✓ mesma qtde" if dup['mesmo_qtde'] else f"⚠️ qtde diferente ({existente[4]} → {novo[3]})"
                    duplicatas_info.append(f"• OP: {novo[0]} | Unidade: {novo[1]} | Arquivo: {novo[2]} ({status_qtde})")
                
                duplicatas_text = "\n".join(duplicatas_info)
                if verificacao_duplicatas['total_duplicatas'] > 5:
                    duplicatas_text += f"\n... e mais {verificacao_duplicatas['total_duplicatas'] - 5} duplicatas"
                
                messagebox.showerror(
                    "Dados Duplicados Encontrados",
                    f"Encontradas {verificacao_duplicatas['total_duplicatas']} duplicatas que não serão importadas:\n\n" +
                    duplicatas_text + "\n\n" +
                    f"Registros novos que serão importados: {verificacao_duplicatas['total_novos']}\n\n" +
                    "Para evitar duplicatas, apenas registros únicos serão adicionados."
                )
                
                # Se não há registros novos, cancela a importação
                if verificacao_duplicatas['total_novos'] == 0:
                    messagebox.showwarning(
                        "Nenhum Registro Novo",
                        "Todos os registros já existem no banco de dados.\nNenhum dado foi importado."
                    )
                    return False
                
                # Usa apenas os registros novos
                registros = verificacao_duplicatas['novos']
            
            # Insere apenas os registros novos (sem duplicatas)
            success = self.database.insert_multiple_registros(registros)
            
            if success:
                messagebox.showinfo(
                    "Sucesso",
                    f"Importação concluída!\n\n" +
                    f"Registros importados: {len(registros)}\n" +
                    f"Total de registros no banco: {self.get_total_registros()}"
                )
                return True
            else:
                messagebox.showerror("Erro", "Falha ao salvar os dados no banco!")
                return False
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro durante importação:\n{str(e)}")
            return False
    
    def get_all_registros(self) -> List[Tuple]:
        """
        Retorna todos os registros do banco
        
        Returns:
            List[Tuple]: Lista com todos os registros
        """
        return self.database.get_all_registros()
    
    def search_registros(self, campo: str, valor: str) -> List[Tuple]:
        """
        Busca registros por um campo específico
        
        Args:
            campo (str): Campo para busca
            valor (str): Valor para buscar
            
        Returns:
            List[Tuple]: Lista com registros encontrados
        """
        if not valor.strip():
            return self.get_all_registros()
        
        return self.database.search_registros(campo, valor)
    
    def delete_registro(self, registro_id: int) -> bool:
        """
        Deleta um registro específico
        
        Args:
            registro_id (int): ID do registro
            
        Returns:
            bool: True se deletado com sucesso
        """
        return self.database.delete_registro(registro_id)
    
    def generate_labels_pdf(self, registros: List[Tuple], output_path: str) -> bool:
        """
        Gera PDF com etiquetas dos registros selecionados
        
        Args:
            registros (List[Tuple]): Lista de registros
            output_path (str): Caminho para salvar o PDF
            
        Returns:
            bool: True se gerado com sucesso
        """
        if not registros:
            messagebox.showwarning("Aviso", "Nenhum registro selecionado para gerar etiquetas!")
            return False
        
        try:
            # Agora geramos uma etiqueta por registro selecionado. A quantidade
            # (qtde) será exibida em cada etiqueta.
            total_etiquetas = len(registros)

            # Confirma a geração
            resposta = messagebox.askyesno(
                "Confirmar Geração",
                f"Serão geradas {total_etiquetas} etiquetas (1 por registro) para {len(registros)} registros.\n\n" +
                "Deseja continuar?"
            )
            
            if not resposta:
                return False
            
            # Gera o PDF — para Zebra 10x5 cm (100x50 mm) imprimimos 1 etiqueta por página
            # Usuário já solicitou etiquetas Zebra 10x5: usamos label_size_mm=(100,50) e single_per_page=True
            success = self.pdf_service.generate_labels_pdf(registros, output_path, label_size_mm=(100, 50), single_per_page=True)
            
            if success:
                messagebox.showinfo(
                    "Sucesso",
                    f"PDF gerado com sucesso!\n\n" +
                    f"Arquivo: {output_path}\n" +
                    f"Etiquetas geradas: {total_etiquetas}"
                )
                return True
            else:
                messagebox.showerror("Erro", "Falha ao gerar o PDF!")
                return False
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar PDF:\n{str(e)}")
            return False
    
    def generate_list_pdf(self, registros: List[Tuple], output_path: str) -> bool:
        """
        Gera PDF com lista simples dos registros
        
        Args:
            registros (List[Tuple]): Lista de registros
            output_path (str): Caminho para salvar o PDF
            
        Returns:
            bool: True se gerado com sucesso
        """
        if not registros:
            messagebox.showwarning("Aviso", "Nenhum registro selecionado!")
            return False
        
        try:
            success = self.pdf_service.generate_simple_list_pdf(registros, output_path)
            
            if success:
                messagebox.showinfo(
                    "Sucesso",
                    f"Relatório gerado com sucesso!\n\n" +
                    f"Arquivo: {output_path}\n" +
                    f"Registros: {len(registros)}"
                )
                return True
            else:
                messagebox.showerror("Erro", "Falha ao gerar o relatório!")
                return False
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar relatório:\n{str(e)}")
            return False
    
    def get_excel_preview(self, file_path: str) -> Optional[dict]:
        """
        Retorna prévia dos dados do Excel
        
        Args:
            file_path (str): Caminho para o arquivo Excel
            
        Returns:
            Optional[dict]: Dados da prévia ou None se erro
        """
        return self.excel_service.get_excel_preview(file_path)
    
    def get_statistics(self) -> dict:
        """
        Retorna estatísticas dos dados
        
        Returns:
            dict: Estatísticas dos dados
        """
        return self.database.get_statistics()
    
    def get_total_registros(self) -> int:
        """
        Retorna o total de registros no banco
        
        Returns:
            int: Total de registros
        """
        return len(self.database.get_all_registros())
    
    def clear_all_data(self) -> bool:
        """
        Limpa todos os dados do banco
        
        Returns:
            bool: True se limpo com sucesso
        """
        try:
            resposta = messagebox.askyesno(
                "Confirmar Exclusão",
                "Tem certeza que deseja excluir TODOS os registros?\n\n" +
                "Esta ação não pode ser desfeita!"
            )
            
            if not resposta:
                return False
            
            success = self.database.clear_all_registros()
            
            if success:
                messagebox.showinfo("Sucesso", "Todos os registros foram excluídos!")
                return True
            else:
                messagebox.showerror("Erro", "Falha ao excluir os registros!")
                return False
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao excluir registros:\n{str(e)}")
            return False
    
    def get_pdf_info(self) -> dict:
        """
        Retorna informações sobre o layout das etiquetas PDF
        
        Returns:
            dict: Informações do layout
        """
        return self.pdf_service.get_label_dimensions_info()
