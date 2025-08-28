from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.units import mm, cm
from reportlab.pdfgen import canvas
from reportlab.lib.colors import black, white, blue
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.graphics import renderPDF
from reportlab.graphics.shapes import Drawing, Rect, String
from typing import List, Tuple
import os
from datetime import datetime
try:
    from reportlab.graphics import renderPDF
    from svglib.svglib import svg2rlg
    SVG_AVAILABLE = True
except ImportError:
    SVG_AVAILABLE = False

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
    
    def generate_labels_pdf(self, registros: List[Tuple], output_path: str, label_size_mm: Tuple[float, float] = None, single_per_page: bool = False) -> bool:
        """
        Gera PDF com etiquetas baseado nos registros
        
        Args:
            registros (List[Tuple]): Lista de registros (id, op, unidade, arquivos, qtde)
            output_path (str): Caminho para salvar o PDF
            
        Returns:
            bool: True se gerado com sucesso, False caso contrário
        """
        try:
            # Decide dimensões da etiqueta (em pontos)
            if label_size_mm:
                lw = label_size_mm[0] * mm
                lh = label_size_mm[1] * mm
            else:
                lw = self.label_width
                lh = self.label_height

            # Se for uma etiqueta por página (Zebra), vamos criar um PDF com o tamanho da etiqueta
            if single_per_page:
                page_size = (lw, lh)
                margin = 3 * mm
            else:
                page_size = A4
                margin = self.margin

            # Cria o canvas
            c = canvas.Canvas(output_path, pagesize=page_size)

            # Prepara as etiquetas
            etiquetas = self._prepare_labels_data(registros)

            if not etiquetas:
                return False

            # Determina como as etiquetas serão dispostas na página atual
            page_width, page_height = page_size
            labels_per_row = int((page_width - 2 * margin) // lw) if not single_per_page else 1
            labels_per_col = int((page_height - 2 * margin) // lh) if not single_per_page else 1
            labels_per_page = max(1, labels_per_row * labels_per_col)

            # Gera as etiquetas
            total_labels = len(etiquetas)
            current_label = 0

            while current_label < total_labels:
                # Desenha as etiquetas da página atual
                labels_on_page = min(labels_per_page, total_labels - current_label)

                for i in range(labels_on_page):
                    row = i // labels_per_row
                    col = i % labels_per_row

                    # Calcula posição da etiqueta
                    if single_per_page:
                        # Para 1 etiqueta por página, desenhamos a etiqueta ocupando toda a página
                        x = 0
                        y = 0
                        effective_lw = page_width
                        effective_lh = page_height
                        effective_margin = 0
                    else:
                        x = margin + col * lw
                        y = page_height - margin - (row + 1) * lh
                        effective_lw = lw
                        effective_lh = lh
                        effective_margin = margin

                    # Temporariamente substitui dimensões internas para desenhar corretamente
                    old_lw, old_lh, old_margin = self.label_width, self.label_height, self.margin
                    self.label_width, self.label_height, self.margin = effective_lw, effective_lh, effective_margin
                    try:
                        self._draw_single_label(c, etiquetas[current_label + i], x, y)
                    finally:
                        self.label_width, self.label_height, self.margin = old_lw, old_lh, old_margin

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
    
    def _wrap_text(self, text: str, font_name: str, font_size: int, max_width: float, canvas_obj) -> List[str]:
        """
        Quebra texto em múltiplas linhas para caber na largura especificada
        
        Args:
            text (str): Texto para quebrar
            font_name (str): Nome da fonte
            font_size (int): Tamanho da fonte
            max_width (float): Largura máxima em pontos
            canvas_obj: Objeto canvas para calcular largura do texto
            
        Returns:
            List[str]: Lista de linhas quebradas
        """
        if not text:
            return [""]
            
        words = text.split(' ')
        lines = []
        current_line = ""
        
        for word in words:
            test_line = current_line + (" " if current_line else "") + word
            text_width = canvas_obj.stringWidth(test_line, font_name, font_size)
            
            if text_width <= max_width:
                current_line = test_line
            else:
                if current_line:
                    lines.append(current_line)
                    current_line = word
                else:
                    # Palavra muito longa, força quebra
                    lines.append(word)
        
        if current_line:
            lines.append(current_line)
            
        return lines if lines else [""]

    def _draw_logo(self, c: canvas.Canvas, x: float, y: float, width: float, height: float):
        """
        Desenha a logo da empresa CDG usando o arquivo SVG
        
        Args:
            c (canvas.Canvas): Canvas do ReportLab
            x (float): Posição X
            y (float): Posição Y
            width (float): Largura da logo
            height (float): Altura da logo
        """
        try:
            assets_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'assets')
            jpg_path = os.path.join(assets_dir, 'cdg_logo.jpg')
            png_path = os.path.join(assets_dir, 'cdg_logo.png')
            svg_path = os.path.join(assets_dir, 'cdg_logo.svg')

            # 1) Prefer JPG/PNG raster image (mais simples para impressão Zebra)
            for img_path in (jpg_path, png_path):
                try:
                    if os.path.exists(img_path):
                        # Calcula escala mantendo proporção
                        from PIL import Image
                        with Image.open(img_path) as im:
                            img_w, img_h = im.size
                        if img_w > 0 and img_h > 0:
                            scale_x = width / img_w
                            scale_y = height / img_h
                            scale = min(scale_x, scale_y)
                            scaled_width = img_w * scale
                            scaled_height = img_h * scale
                            offset_x = (width - scaled_width) / 2
                            offset_y = (height - scaled_height) / 2

                            # drawImage expects coords with origin at bottom-left
                            c.drawImage(img_path, x + offset_x, y + offset_y, width=scaled_width, height=scaled_height, preserveAspectRatio=True, mask='auto')
                            return
                except Exception:
                    # continua para o próximo formato
                    pass

            # 2) Tenta usar o SVG se disponível
            if SVG_AVAILABLE and os.path.exists(svg_path):
                try:
                    drawing = svg2rlg(svg_path)
                    if drawing:
                        svg_width = drawing.width
                        svg_height = drawing.height
                        if svg_width > 0 and svg_height > 0:
                            scale_x = width / svg_width
                            scale_y = height / svg_height
                            scale = min(scale_x, scale_y)  # Mantém proporção

                            # Centraliza a logo redimensionada
                            scaled_width = svg_width * scale
                            scaled_height = svg_height * scale
                            offset_x = (width - scaled_width) / 2
                            offset_y = (height - scaled_height) / 2

                            # Salva estado do canvas
                            c.saveState()

                            # Aplica transformação
                            c.translate(x + offset_x, y + offset_y)
                            c.scale(scale, scale)

                            # Desenha o SVG
                            renderPDF.draw(drawing, c, 0, 0)

                            # Restaura estado
                            c.restoreState()
                            return
                except Exception as e:
                    # se falhar, cai para fallback
                    print(f"SVG draw failed: {e}")
            
            # Fallback: desenha logo simples se SVG não funcionar
            self._draw_fallback_logo(c, x, y, width, height)
            
        except Exception as e:
            print(f"Erro ao desenhar logo SVG: {e}")
            # Fallback em caso de erro
            self._draw_fallback_logo(c, x, y, width, height)
    
    def _draw_fallback_logo(self, c: canvas.Canvas, x: float, y: float, width: float, height: float):
        """
        Desenha uma logo simples como fallback
        """
        # Salva o estado atual das cores
        c.saveState()
        
        # Fundo com gradiente simulado (usando tons de azul)
        c.setFillColor(blue)
        c.setStrokeColor(blue)
        c.rect(x, y, width, height, fill=1, stroke=1)
        
        # Adiciona uma borda mais escura
        from reportlab.lib.colors import Color
        dark_blue = Color(0, 0, 0.7)
        c.setStrokeColor(dark_blue)
        c.setLineWidth(0.5)
        c.rect(x, y, width, height, fill=0, stroke=1)
        
        # Texto "CDG" em branco com fonte maior
        c.setFillColor(white)
        font_size = min(int(height * 0.6), 10)  # Ajusta o tamanho da fonte baseado na altura
        c.setFont("Helvetica-Bold", font_size)
        
        # Centraliza o texto na logo
        text_width = c.stringWidth("CDG", "Helvetica-Bold", font_size)
        text_x = x + (width - text_width) / 2
        text_y = y + (height - font_size) / 2
        
        c.drawString(text_x, text_y, "CDG")
        
        # Adiciona um pequeno sublinhado decorativo
        line_y = text_y - 2
        line_start_x = text_x
        line_end_x = text_x + text_width
        c.setLineWidth(0.8)
        c.setStrokeColor(white)
        c.line(line_start_x, line_y, line_end_x, line_y)
        
        # Restaura o estado das cores
        c.restoreState()

    def _draw_single_label(self, c: canvas.Canvas, etiqueta: dict, x: float, y: float):
        """
        Desenha uma única etiqueta no PDF com quebra automática de linhas e logo
        
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
        
        # Dimensões da logo
        logo_width = 15 * mm
        logo_height = 8 * mm
        
        # Largura disponível para texto (descontando padding e logo)
        available_width = self.label_width - (2 * padding) - logo_width - (2 * mm)

        # Desenha a logo no canto superior direito
        logo_x = x + self.label_width - padding - logo_width
        logo_y = y + self.label_height - padding - logo_height
        self._draw_logo(c, logo_x, logo_y, logo_width, logo_height)

        # Posições dentro da etiqueta
        text_x = x + padding
        current_y = y + self.label_height - padding - 12

        # Título principal (OP)
        c.setFillColor(black)  # Garante que o texto seja preto
        c.setFont("Helvetica-Bold", title_font_size)
        op_text = f"OP: {etiqueta.get('op', '')}"
        op_lines = self._wrap_text(op_text, "Helvetica-Bold", title_font_size, available_width, c)
        for line in op_lines:
            c.drawString(text_x, current_y, line)
            current_y -= 12

        current_y -= 3  # Espaço extra após OP

        # Unidade
        c.setFont("Helvetica-Bold", text_font_size)
        unidade_text = f"Unidade: {etiqueta.get('unidade', '')}"
        unidade_lines = self._wrap_text(unidade_text, "Helvetica-Bold", text_font_size, available_width, c)
        for line in unidade_lines:
            c.drawString(text_x, current_y, line)
            current_y -= 10

        current_y -= 2  # Espaço extra após Unidade

        # Arquivo - quebra automática de linha
        c.setFont("Helvetica", text_font_size)
        arquivo_text = f"Arquivo: {etiqueta.get('arquivo', '')}"
        arquivo_lines = self._wrap_text(arquivo_text, "Helvetica", text_font_size, available_width, c)
        
        # Limita a 3 linhas para arquivo para não ocupar muito espaço
        max_arquivo_lines = 3
        for i, line in enumerate(arquivo_lines[:max_arquivo_lines]):
            if i == max_arquivo_lines - 1 and len(arquivo_lines) > max_arquivo_lines:
                # Adiciona "..." se há mais linhas
                if len(line) > 40:
                    line = f"{line[:40]}..."
                else:
                    line += "..."
            c.drawString(text_x, current_y, line)
            current_y -= 10

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
