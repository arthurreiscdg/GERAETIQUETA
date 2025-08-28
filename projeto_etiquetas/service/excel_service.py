import pandas as pd
import os
import logging
from typing import List, Tuple, Optional
from tkinter import messagebox

# Configurar logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

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
        Lê os dados de todas as planilhas do Excel e retorna uma lista de registros
        
        Args:
            file_path (str): Caminho para o arquivo Excel
            
        Returns:
            Optional[List[Tuple]]: Lista de tuplas (op, unidade, arquivos, qtde) ou None se erro
        """
        try:
            if not os.path.exists(file_path):
                messagebox.showerror("Erro", "Arquivo não encontrado!")
                return None
            
            # Lê todas as planilhas do arquivo Excel
            try:
                excel_file = pd.ExcelFile(file_path)
                sheet_names = excel_file.sheet_names
                print(f"Planilhas encontradas: {sheet_names}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao acessar arquivo Excel:\n{str(e)}")
                return None
            
            todos_registros = []
            planilhas_processadas = 0
            
            # Processa cada planilha
            for sheet_name in sheet_names:
                print(f"Processando planilha: {sheet_name}")
                
                try:
                    # Lê a planilha atual
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                    
                    # Verifica se tem pelo menos 2 linhas e 2 colunas
                    if df.shape[0] < 2 or df.shape[1] < 2:
                        print(f"Planilha '{sheet_name}' ignorada: estrutura insuficiente")
                        continue
                    
                    # Verifica se A1 e B1 têm dados
                    if pd.isna(df.iloc[0, 0]) or pd.isna(df.iloc[0, 1]):
                        print(f"Planilha '{sheet_name}' ignorada: A1 ou B1 vazios")
                        continue
                    
                    # Extrai OP (A1) e unidade (B1)
                    op = str(df.iloc[0, 0]).strip()
                    unidade = str(df.iloc[0, 1]).strip()
                    
                    registros_planilha = []
                    
                    # Primeiro, vamos encontrar o nome na última linha válida da coluna B
                    nome_planilha = ""
                    # Percorre de trás para frente procurando um nome na coluna B
                    for i in range(len(df) - 1, 0, -1):  # De baixo para cima, excluindo linha 0
                        cell_b = df.iloc[i, 1]
                        cell_a = df.iloc[i, 0]
                        
                        # Pula linha "Quantidade total" 
                        if not pd.isna(cell_a):
                            cell_a_str = str(cell_a).strip().lower()
                            if "quantidade total" in cell_a_str or "qtde total" in cell_a_str:
                                continue
                        
                        # Se encontrou algo na coluna B que não é número
                        if not pd.isna(cell_b):
                            cell_b_str = str(cell_b).strip()
                            if cell_b_str:
                                # Testa se não é número
                                try:
                                    float(cell_b_str)
                                    # É número, continua procurando
                                except (ValueError, TypeError):
                                    # Não é número, pode ser o nome
                                    if cell_b_str.lower() not in ["quantidade", "qtde", "total"]:
                                        nome_planilha = cell_b_str
                                        logger.debug(f"Nome da planilha encontrado na linha {i+1}: '{nome_planilha}'")
                                        break
                    
                    if not nome_planilha:
                        logger.debug("Nenhum nome encontrado na planilha")
                    
                    # Percorre as linhas a partir da linha 2 (índice 1)
                    i = 1
                    while i < len(df):
                        # Arquivo está na coluna A (índice 0)
                        arquivo = df.iloc[i, 0]
                        
                        # Quantidade está na coluna B (índice 1)
                        qtde = df.iloc[i, 1]
                        
                        # Pula linhas vazias
                        if pd.isna(arquivo) and pd.isna(qtde):
                            i += 1
                            continue
                        
                        # Verifica se arquivo não está vazio
                        if pd.isna(arquivo) or str(arquivo).strip() == "":
                            i += 1
                            continue
                        
                        # Ignora linhas onde A contém "Quantidade total" ou similar
                        arquivo_str = str(arquivo).strip().lower() if not pd.isna(arquivo) else ""
                        if "quantidade total" in arquivo_str or "qtde total" in arquivo_str or arquivo_str == "total":
                            print(f"Ignorando linha {i+1}: '{arquivo_str}' na coluna A")
                            i += 1
                            continue
                        
                        # Converte quantidade para inteiro, se não conseguir, usa 0
                        try:
                            qtde = int(float(qtde)) if not pd.isna(qtde) else 0
                        except (ValueError, TypeError):
                            qtde = 0
                        
                        # Se quantidade for 0 ou negativa, pula
                        if qtde <= 0:
                            i += 1
                            continue
                        
                        logger.debug(f"Linha {i+1}: arquivo='{str(arquivo).strip()}', qtde={qtde}")
                        
                        arquivo_str = str(arquivo).strip()
                        # Usa o nome encontrado na planilha para todos os registros
                        registro = (op, unidade, arquivo_str, qtde, nome_planilha)
                        logger.debug(f"Criando registro: {registro}")
                        registros_planilha.append(registro)
                        
                        i += 1
                    
                    if registros_planilha:
                        todos_registros.extend(registros_planilha)
                        planilhas_processadas += 1
                        print(f"Planilha '{sheet_name}': {len(registros_planilha)} registros")
                    else:
                        print(f"Planilha '{sheet_name}': nenhum registro válido")
                        
                except Exception as e:
                    print(f"Erro ao processar planilha '{sheet_name}': {e}")
                    continue
            
            if not todos_registros:
                messagebox.showwarning("Aviso", 
                    f"Nenhum registro válido encontrado em nenhuma das {len(sheet_names)} planilhas!\n\n" +
                    "Estrutura esperada por planilha:\n" +
                    "A1 = OP (identificador da ordem de produção)\n" +
                    "A2 em diante = arquivos\n" +
                    "B1 = unidade (nome da unidade)\n" +
                    "B2 em diante = quantidade")
                return None
            
            messagebox.showinfo("Importação", 
                f"Processamento concluído!\n\n" +
                f"Planilhas processadas: {planilhas_processadas}/{len(sheet_names)}\n" +
                f"Total de registros encontrados: {len(todos_registros)}")
            
            return todos_registros
            
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
        
        for i, registro in enumerate(registros, 1):
            # Desempacota os dados - pode ter 4 ou 5 elementos
            if len(registro) == 4:
                op, unidade, arquivo, qtde = registro
                nome = ""
            elif len(registro) >= 5:
                op, unidade, arquivo, qtde, nome = registro[:5]
            else:
                problemas.append(f"Linha {i}: Registro inválido")
                continue
                
            linha_ok = True
            logger.debug(f"Validando linha {i}: op={op}, unidade={unidade}, arquivo={arquivo}, qtde={qtde}, nome={nome}")
            
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
