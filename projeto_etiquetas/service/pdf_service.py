from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.units import mm, cm
from reportlab.pdfgen import canvas
from reportlab.lib.colors import black, white
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from typing import List, Tuple
import os
from datetime import datetime

class PDFService:
    def __init__(self):
        """Inicializa o serviço de geração de PDF"""
        self.page_width, self.page_height = A4
        # Configurações da etiqueta (em mm convertido para pontos)
        self.label_width = 80 * mm
        self.label_height = 50 * mm
        self.margin = 10 * mm
        
        # Calcula quantas etiquetas cabem por página
        self.labels_per_row = int((self.page_width - 2 * self.margin) // self.label_width)
        self.labels_per_col = int((self.page_height - 2 * self.margin) // self.label_height)
        self.labels_per_page = self.labels_per_row * self.labels_per_col
    
    def generate_labels_pdf(self, registros: List[Tuple], output_path: str) -> bool:
        """
        Gera PDF com etiquetas baseado nos registros
        
        Args:
            registros (List[Tuple]): Lista de registros (id, op, unidade, arquivos, qtde)
            output_path (str): Caminho para salvar o PDF
            
        Returns:
            bool: True se gerado com sucesso, False caso contrário
        """
        try:
            # Cria o canvas
            c = canvas.Canvas(output_path, pagesize=A4)
            
            # Prepara as etiquetas (expandindo por quantidade)
            etiquetas = self._prepare_labels_data(registros)
            
            if not etiquetas:
                return False
            
            # Gera as etiquetas
            total_labels = len(etiquetas)
            current_label = 0
            
            while current_label < total_labels:
                # Desenha as etiquetas da página atual
                labels_on_page = min(self.labels_per_page, total_labels - current_label)
                
                for i in range(labels_on_page):
                    row = i // self.labels_per_row
                    col = i % self.labels_per_row
                    
                    # Calcula posição da etiqueta
                    x = self.margin + col * self.label_width
                    y = self.page_height - self.margin - (row + 1) * self.label_height
                    
                    # Desenha a etiqueta
                    self._draw_single_label(c, etiquetas[current_label + i], x, y)
                
                current_label += labels_on_page
                
                # Se ainda há etiquetas, cria nova página
                if current_label < total_labels:
                    c.showPage()
            
            # Salva o PDF
            c.save()
            return True
            
        except Exception as e:
            print(f"Erro ao gerar PDF: {e}")
            return False
    
    def _prepare_labels_data(self, registros: List[Tuple]) -> List[dict]:
        """
        Prepara os dados das etiquetas expandindo por quantidade
        
        Args:
            registros (List[Tuple]): Lista de registros do banco
            
        Returns:
            List[dict]: Lista de etiquetas individuais
        """
        # Agora cada registro gera UMA etiqueta. A quantidade (qtde) será
        # exibida na própria etiqueta ao invés de criar múltiplas cópias.
        etiquetas = []

        for registro in registros:
            # registro = (id, op, unidade, arquivos, qtde)
            if len(registro) >= 5:
                _, op, unidade, arquivos, qtde = registro
                etiqueta = {
                    'op': str(op),
                    'unidade': str(unidade),
                    'arquivo': str(arquivos),
                    'qtde': int(qtde) if qtde is not None else 0
                }
                etiquetas.append(etiqueta)

        return etiquetas
    
    def _draw_single_label(self, c: canvas.Canvas, etiqueta: dict, x: float, y: float):
        """
        Desenha uma única etiqueta no PDF
        
        Args:
            c (canvas.Canvas): Canvas do ReportLab
            etiqueta (dict): Dados da etiqueta
            x (float): Posição X
            y (float): Posição Y
        """
        # Desenha borda da etiqueta
        c.setStrokeColor(black)
        c.setLineWidth(1)
        c.rect(x, y, self.label_width, self.label_height)

        # Configurações de fonte
        title_font_size = 10
        text_font_size = 8
        small_font_size = 6

        # Margens internas
        padding = 3 * mm

        # Posições dentro da etiqueta
        text_x = x + padding
        current_y = y + self.label_height - padding - 12

        # Título principal (OP)
        c.setFont("Helvetica-Bold", title_font_size)
        c.drawString(text_x, current_y, f"OP: {etiqueta.get('op', '')}")
        current_y -= 15

        # Unidade
        c.setFont("Helvetica-Bold", text_font_size)
        c.drawString(text_x, current_y, f"Unidade: {etiqueta.get('unidade', '')}")
        current_y -= 12

        # Arquivo
        c.setFont("Helvetica", text_font_size)
        arquivo_text = etiqueta.get('arquivo', '')
        # Trunca se muito longo
        if len(arquivo_text) > 25:
            arquivo_text = arquivo_text[:22] + "..."
        c.drawString(text_x, current_y, f"Arquivo: {arquivo_text}")
        current_y -= 12

        # Quantidade (exibida no canto inferior direito)
        c.setFont("Helvetica-Bold", small_font_size)
        qtde_text = f"Qtde: {etiqueta.get('qtde', 0)}"

        text_width = c.stringWidth(qtde_text, "Helvetica-Bold", small_font_size)
        qtde_x = x + self.label_width - padding - text_width
        c.drawString(qtde_x, y + padding, qtde_text)

        # Data/hora de geração (canto inferior esquerdo)
        c.setFont("Helvetica", small_font_size)
        data_geracao = datetime.now().strftime("%d/%m/%Y %H:%M")
        c.drawString(text_x, y + padding, data_geracao)
    
    def generate_simple_list_pdf(self, registros: List[Tuple], output_path: str) -> bool:
        """
        Gera PDF com lista simples dos registros (sem etiquetas)
        
        Args:
            registros (List[Tuple]): Lista de registros
            output_path (str): Caminho para salvar o PDF
            
        Returns:
            bool: True se gerado com sucesso, False caso contrário
        """
        try:
            c = canvas.Canvas(output_path, pagesize=A4)
            
            # Configurações
            margin = 20 * mm
            line_height = 15
            font_size = 10
            
            # Cabeçalho
            c.setFont("Helvetica-Bold", 14)
            c.drawString(margin, self.page_height - margin, "Relatório de Registros")
            
            c.setFont("Helvetica", 8)
            data_geracao = datetime.now().strftime("%d/%m/%Y às %H:%M")
            c.drawString(margin, self.page_height - margin - 20, f"Gerado em: {data_geracao}")
            
            # Cabeçalho da tabela
            y_position = self.page_height - margin - 50
            c.setFont("Helvetica-Bold", font_size)
            
            headers = ["ID", "OP", "Unidade", "Arquivo", "Qtde"]
            col_widths = [30, 80, 120, 200, 50]
            x_positions = [margin]
            
            for width in col_widths[:-1]:
                x_positions.append(x_positions[-1] + width)
            
            # Desenha cabeçalhos
            for i, header in enumerate(headers):
                c.drawString(x_positions[i], y_position, header)
            
            y_position -= 20
            
            # Linha de separação
            c.line(margin, y_position, margin + sum(col_widths), y_position)
            y_position -= 10
            
            # Dados
            c.setFont("Helvetica", font_size - 1)
            
            for registro in registros:
                if y_position < margin + 30:  # Nova página se necessário
                    c.showPage()
                    y_position = self.page_height - margin
                
                # registro = (id, op, unidade, arquivos, qtde)
                row_data = [
                    str(registro[0]),  # ID
                    str(registro[1]),  # OP
                    str(registro[2])[:15] + "..." if len(str(registro[2])) > 15 else str(registro[2]),  # Unidade
                    str(registro[3])[:25] + "..." if len(str(registro[3])) > 25 else str(registro[3]),  # Arquivo
                    str(registro[4])   # Qtde
                ]
                
                for i, data in enumerate(row_data):
                    c.drawString(x_positions[i], y_position, data)
                
                y_position -= line_height
            
            # Total de registros
            y_position -= 20
            c.setFont("Helvetica-Bold", font_size)
            c.drawString(margin, y_position, f"Total de registros: {len(registros)}")
            
            c.save()
            return True
            
        except Exception as e:
            print(f"Erro ao gerar PDF da lista: {e}")
            return False
    
    def get_label_dimensions_info(self) -> dict:
        """
        Retorna informações sobre as dimensões das etiquetas
        
        Returns:
            dict: Informações sobre layout das etiquetas
        """
        return {
            'label_width_mm': self.label_width / mm,
            'label_height_mm': self.label_height / mm,
            'labels_per_row': self.labels_per_row,
            'labels_per_col': self.labels_per_col,
            'labels_per_page': self.labels_per_page,
            'page_width_mm': self.page_width / mm,
            'page_height_mm': self.page_height / mm
        }
