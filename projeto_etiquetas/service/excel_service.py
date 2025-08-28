import pandas as pd
import os
from typing import List, Tuple, Optional
from tkinter import messagebox

class ExcelService:
    def __init__(self):
        """Inicializa o serviço de Excel"""
        pass
    
    def validate_excel_structure(self, file_path: str) -> bool:
        """
        Valida se o arquivo Excel tem a estrutura esperada
        
        Args:
            file_path (str): Caminho para o arquivo Excel
            
        Returns:
            bool: True se a estrutura estiver correta, False caso contrário
        """
        try:
            # Lê o arquivo Excel
            df = pd.read_excel(file_path, header=None)
            
            # Verifica se tem pelo menos 2 linhas e 2 colunas
            if df.shape[0] < 2 or df.shape[1] < 2:
                return False
            
            # Verifica se A1 contém algo (OP) e B1 contém algo (unidade)
            if pd.isna(df.iloc[0, 0]) or pd.isna(df.iloc[0, 1]):
                return False
            
            return True
            
        except Exception as e:
            print(f"Erro ao validar estrutura do Excel: {e}")
            return False
    
    def read_excel_data(self, file_path: str) -> Optional[List[Tuple[str, str, str, int]]]:
        """
        Lê os dados do Excel e retorna uma lista de registros
        
        Args:
            file_path (str): Caminho para o arquivo Excel
            
        Returns:
            Optional[List[Tuple]]: Lista de tuplas (op, unidade, arquivos, qtde) ou None se erro
        """
        try:
            if not os.path.exists(file_path):
                messagebox.showerror("Erro", "Arquivo não encontrado!")
                return None
            
            # Valida a estrutura
            if not self.validate_excel_structure(file_path):
                messagebox.showerror("Erro", 
                    "Estrutura do Excel inválida!\n\n" +
                    "Estrutura esperada:\n" +
                    "A1 = OP (identificador da ordem de produção)\n" +
                    "A2 em diante = arquivos\n" +
                    "B1 = unidade (nome da unidade)\n" +
                    "B2 em diante = quantidade")
                return None
            
            # Lê o arquivo Excel sem cabeçalho
            df = pd.read_excel(file_path, header=None)
            
            # Extrai OP (A1) e unidade (B1)
            op = str(df.iloc[0, 0]).strip()
            unidade = str(df.iloc[0, 1]).strip()
            
            registros = []
            
            # Percorre as linhas a partir da linha 2 (índice 1)
            for i in range(1, len(df)):
                # Arquivo está na coluna A (índice 0)
                arquivo = df.iloc[i, 0]
                
                # Quantidade está na coluna B (índice 1)
                qtde = df.iloc[i, 1]
                
                # Pula linhas vazias
                if pd.isna(arquivo) and pd.isna(qtde):
                    continue
                
                # Verifica se arquivo não está vazio
                if pd.isna(arquivo) or str(arquivo).strip() == "":
                    continue
                
                # Converte quantidade para inteiro, se não conseguir, usa 0
                try:
                    qtde = int(float(qtde)) if not pd.isna(qtde) else 0
                except (ValueError, TypeError):
                    qtde = 0
                
                # Se quantidade for 0 ou negativa, pula
                if qtde <= 0:
                    continue
                
                arquivo_str = str(arquivo).strip()
                registros.append((op, unidade, arquivo_str, qtde))
            
            if not registros:
                messagebox.showwarning("Aviso", "Nenhum registro válido encontrado no arquivo!")
                return None
            
            return registros
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler arquivo Excel:\n{str(e)}")
            return None
    
    def get_excel_preview(self, file_path: str, max_rows: int = 10) -> Optional[dict]:
        """
        Retorna uma prévia dos dados do Excel para validação
        
        Args:
            file_path (str): Caminho para o arquivo Excel
            max_rows (int): Número máximo de linhas para prévia
            
        Returns:
            Optional[dict]: Dicionário com informações da prévia ou None se erro
        """
        try:
            if not os.path.exists(file_path):
                return None
            
            # Lê o arquivo Excel
            df = pd.read_excel(file_path, header=None)
            
            # Informações básicas
            info = {
                'total_rows': len(df),
                'total_cols': len(df.columns) if not df.empty else 0,
                'op': str(df.iloc[0, 0]).strip() if len(df) > 0 and not pd.isna(df.iloc[0, 0]) else "N/A",
                'unidade': str(df.iloc[0, 1]).strip() if len(df) > 0 and len(df.columns) > 1 and not pd.isna(df.iloc[0, 1]) else "N/A",
                'preview_data': []
            }
            
            # Dados de prévia (máximo max_rows linhas)
            end_row = min(len(df), max_rows)
            for i in range(end_row):
                row_data = []
                for j in range(min(len(df.columns), 5)):  # Máximo 5 colunas
                    cell_value = df.iloc[i, j]
                    if pd.isna(cell_value):
                        row_data.append("")
                    else:
                        row_data.append(str(cell_value))
                info['preview_data'].append(row_data)
            
            return info
            
        except Exception as e:
            print(f"Erro ao obter prévia do Excel: {e}")
            return None
    
    def validate_data_quality(self, registros: List[Tuple[str, str, str, int]]) -> dict:
        """
        Valida a qualidade dos dados extraídos
        
        Args:
            registros (List[Tuple]): Lista de registros
            
        Returns:
            dict: Relatório de qualidade dos dados
        """
        if not registros:
            return {
                'total_registros': 0,
                'registros_validos': 0,
                'registros_com_problemas': 0,
                'problemas': []
            }
        
        registros_validos = 0
        problemas = []
        
        for i, (op, unidade, arquivo, qtde) in enumerate(registros, 1):
            linha_ok = True
            
            # Verifica OP
            if not op or op.strip() == "":
                problemas.append(f"Linha {i}: OP vazia")
                linha_ok = False
            
            # Verifica unidade
            if not unidade or unidade.strip() == "":
                problemas.append(f"Linha {i}: Unidade vazia")
                linha_ok = False
            
            # Verifica arquivo
            if not arquivo or arquivo.strip() == "":
                problemas.append(f"Linha {i}: Arquivo vazio")
                linha_ok = False
            
            # Verifica quantidade
            if qtde <= 0:
                problemas.append(f"Linha {i}: Quantidade inválida ({qtde})")
                linha_ok = False
            
            if linha_ok:
                registros_validos += 1
        
        return {
            'total_registros': len(registros),
            'registros_validos': registros_validos,
            'registros_com_problemas': len(registros) - registros_validos,
            'problemas': problemas[:10]  # Máximo 10 problemas na lista
        }
