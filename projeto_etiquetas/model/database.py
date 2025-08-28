import sqlite3
import os
from typing import List, Tuple, Optional

class Database:
    def __init__(self, db_path: str = "etiquetas.db"):
        """
        Inicializa a conexão com o banco de dados SQLite
        
        Args:
            db_path (str): Caminho para o arquivo do banco de dados
        """
        self.db_path = db_path
        self.init_database()
    
    def init_database(self):
        """Cria a tabela se ela não existir"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS etiquetas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    op TEXT NOT NULL,
                    unidade TEXT NOT NULL,
                    arquivos TEXT NOT NULL,
                    qtde INTEGER NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            conn.commit()
    
    def insert_registro(self, op: str, unidade: str, arquivos: str, qtde: int) -> bool:
        """
        Insere um novo registro na tabela
        
        Args:
            op (str): Ordem de produção
            unidade (str): Nome da unidade
            arquivos (str): Nome do arquivo
            qtde (int): Quantidade
            
        Returns:
            bool: True se inserido com sucesso, False caso contrário
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO etiquetas (op, unidade, arquivos, qtde)
                    VALUES (?, ?, ?, ?)
                ''', (op, unidade, arquivos, qtde))
                conn.commit()
                return True
        except sqlite3.Error as e:
            print(f"Erro ao inserir registro: {e}")
            return False
    
    def insert_multiple_registros(self, registros: List[Tuple[str, str, str, int]]) -> bool:
        """
        Insere múltiplos registros de uma vez
        
        Args:
            registros (List[Tuple]): Lista de tuplas (op, unidade, arquivos, qtde)
            
        Returns:
            bool: True se todos foram inseridos com sucesso, False caso contrário
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.executemany('''
                    INSERT INTO etiquetas (op, unidade, arquivos, qtde)
                    VALUES (?, ?, ?, ?)
                ''', registros)
                conn.commit()
                return True
        except sqlite3.Error as e:
            print(f"Erro ao inserir múltiplos registros: {e}")
            return False
    
    def get_all_registros(self) -> List[Tuple]:
        """
        Retorna todos os registros da tabela
        
        Returns:
            List[Tuple]: Lista com todos os registros
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('SELECT id, op, unidade, arquivos, qtde FROM etiquetas ORDER BY id DESC')
                return cursor.fetchall()
        except sqlite3.Error as e:
            print(f"Erro ao buscar registros: {e}")
            return []
    
    def search_registros(self, campo: str, valor: str) -> List[Tuple]:
        """
        Busca registros por um campo específico
        
        Args:
            campo (str): Campo para busca (op, unidade, arquivos)
            valor (str): Valor para buscar
            
        Returns:
            List[Tuple]: Lista com os registros encontrados
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                if campo in ['op', 'unidade', 'arquivos']:
                    query = f'SELECT id, op, unidade, arquivos, qtde FROM etiquetas WHERE {campo} LIKE ? ORDER BY id DESC'
                    cursor.execute(query, (f'%{valor}%',))
                    return cursor.fetchall()
                else:
                    return []
        except sqlite3.Error as e:
            print(f"Erro ao buscar registros: {e}")
            return []
    
    def delete_registro(self, registro_id: int) -> bool:
        """
        Deleta um registro específico
        
        Args:
            registro_id (int): ID do registro a ser deletado
            
        Returns:
            bool: True se deletado com sucesso, False caso contrário
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('DELETE FROM etiquetas WHERE id = ?', (registro_id,))
                conn.commit()
                return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Erro ao deletar registro: {e}")
            return False
    
    def clear_all_registros(self) -> bool:
        """
        Limpa todos os registros da tabela
        
        Returns:
            bool: True se limpo com sucesso, False caso contrário
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('DELETE FROM etiquetas')
                conn.commit()
                return True
        except sqlite3.Error as e:
            print(f"Erro ao limpar registros: {e}")
            return False
    
    def get_statistics(self) -> dict:
        """
        Retorna estatísticas dos dados
        
        Returns:
            dict: Dicionário com estatísticas
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                # Total de registros
                cursor.execute('SELECT COUNT(*) FROM etiquetas')
                total_registros = cursor.fetchone()[0]
                
                # Total de OPs únicas
                cursor.execute('SELECT COUNT(DISTINCT op) FROM etiquetas')
                total_ops = cursor.fetchone()[0]
                
                # Total de unidades únicas
                cursor.execute('SELECT COUNT(DISTINCT unidade) FROM etiquetas')
                total_unidades = cursor.fetchone()[0]
                
                # Soma total de quantidade
                cursor.execute('SELECT SUM(qtde) FROM etiquetas')
                total_qtde = cursor.fetchone()[0] or 0
                
                return {
                    'total_registros': total_registros,
                    'total_ops': total_ops,
                    'total_unidades': total_unidades,
                    'total_quantidade': total_qtde
                }
        except sqlite3.Error as e:
            print(f"Erro ao obter estatísticas: {e}")
            return {
                'total_registros': 0,
                'total_ops': 0,
                'total_unidades': 0,
                'total_quantidade': 0
            }
