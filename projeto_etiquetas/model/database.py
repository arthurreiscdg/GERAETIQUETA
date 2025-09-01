import psycopg2
import psycopg2.extras
import os
import logging
from typing import List, Tuple

# Configurar logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class Database:
    def __init__(self, db_path: str = None):
        """
        Inicializa conexão com PostgreSQL no Supabase.
        
        Usa as credenciais fornecidas para conectar ao banco PostgreSQL.
        """
        # Configurações do banco PostgreSQL Supabase
        self.db_config = {
            'user': 'postgres.hftlofgdapnbsobjugla',
            'password': 'z1g0GZt53164fVDI',
            'host': 'aws-1-sa-east-1.pooler.supabase.com',
            'port': 6543,
            'dbname': 'postgres'
        }

        # Verifica disponibilidade do pacote psycopg2
        try:
            import psycopg2
            self._psycopg2_available = True
        except ImportError:
            raise RuntimeError("Pacote 'psycopg2-binary' não está instalado. Instale com: pip install psycopg2-binary")

        # Inicializa esquema (create table se necessário)
        self.init_database()
        
        # Migra a tabela para adicionar coluna nome se necessário
        self._migrate_add_nome_column()
        
        # Migra a tabela para adicionar coluna status se necessário
        self._migrate_add_status_column()

    def init_database(self):
        """Cria a tabela se ela não existir no PostgreSQL"""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS etiquetas (
                    id SERIAL PRIMARY KEY,
                    op TEXT NOT NULL,
                    unidade TEXT NOT NULL,
                    arquivos TEXT NOT NULL,
                    qtde INTEGER NOT NULL,
                    nome TEXT DEFAULT '',
                    status TEXT DEFAULT 'Pendente',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            conn.commit()
        except Exception as e:
            print(f"Erro ao inicializar banco de dados: {e}")
        finally:
            if conn:
                conn.close()

    def _get_connection(self):
        """Retorna uma nova conexão via psycopg2."""
        return psycopg2.connect(**self.db_config)

    def _migrate_add_nome_column(self):
        """Adiciona a coluna 'nome' à tabela se ela não existir"""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            # Verifica se a coluna 'nome' já existe no PostgreSQL
            cursor.execute("""
                SELECT column_name FROM information_schema.columns 
                WHERE table_name='etiquetas' AND column_name='nome'
            """)
            column_exists = cursor.fetchone()
            
            if not column_exists:
                print("Adicionando coluna 'nome' à tabela etiquetas...")
                cursor.execute("ALTER TABLE etiquetas ADD COLUMN nome TEXT DEFAULT ''")
                conn.commit()
                print("Coluna 'nome' adicionada com sucesso!")
            else:
                print("Coluna 'nome' já existe na tabela.")
                
        except Exception as e:
            print(f"Erro ao migrar tabela: {e}")
        finally:
            if conn:
                conn.close()

    def _migrate_add_status_column(self):
        """Adiciona a coluna 'status' à tabela se ela não existir"""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            # Verifica se a coluna 'status' já existe no PostgreSQL
            cursor.execute("""
                SELECT column_name FROM information_schema.columns 
                WHERE table_name='etiquetas' AND column_name='status'
            """)
            column_exists = cursor.fetchone()
            
            if not column_exists:
                print("Adicionando coluna 'status' à tabela etiquetas...")
                cursor.execute("ALTER TABLE etiquetas ADD COLUMN status TEXT DEFAULT 'Pendente'")
                conn.commit()
                print("Coluna 'status' adicionada com sucesso!")
            else:
                print("Coluna 'status' já existe na tabela.")
                
        except Exception as e:
            print(f"Erro ao migrar tabela (status): {e}")
        finally:
            if conn:
                conn.close()

    def insert_registro(self, op: str, unidade: str, arquivos: str, qtde: int, nome: str = "") -> bool:
        """Insere um novo registro na tabela."""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO etiquetas (op, unidade, arquivos, qtde, nome)
                VALUES (%s, %s, %s, %s, %s)
            ''', (op, unidade, arquivos, qtde, nome))
            conn.commit()
            return True
        except Exception as e:
            print(f"Erro ao inserir registro: {e}")
            return False
        finally:
            if conn:
                conn.close()

    def insert_multiple_registros(self, registros: List[Tuple]) -> bool:
        """Insere múltiplos registros de uma vez. Aceita tuplas com 4 ou 5 elementos."""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            logger.debug(f"Inserindo {len(registros)} registros")
            
            # Processa cada registro para garantir que tenha nome
            registros_processados = []
            for i, registro in enumerate(registros):
                logger.debug(f"Registro {i+1}: {registro} (len={len(registro)})")
                
                if len(registro) == 4:
                    # Formato antigo: (op, unidade, arquivos, qtde)
                    op, unidade, arquivos, qtde = registro
                    nome = ""
                    logger.debug(f"Formato antigo detectado: op={op}, unidade={unidade}, arquivo={arquivos}, qtde={qtde}")
                elif len(registro) >= 5:
                    # Formato novo: (op, unidade, arquivos, qtde, nome)
                    op, unidade, arquivos, qtde, nome = registro[:5]
                    logger.debug(f"Formato novo detectado: op={op}, unidade={unidade}, arquivo={arquivos}, qtde={qtde}, nome={nome}")
                else:
                    logger.warning(f"Registro inválido ignorado: {registro}")
                    continue  # Pula registros inválidos
                
                registros_processados.append((op, unidade, arquivos, qtde, nome))
            
            logger.debug(f"Processados {len(registros_processados)} registros")
            
            cursor.executemany('''
                INSERT INTO etiquetas (op, unidade, arquivos, qtde, nome)
                VALUES (%s, %s, %s, %s, %s)
            ''', registros_processados)
            conn.commit()
            return True
        except Exception as e:
            logger.error(f"Erro ao inserir múltiplos registros: {e}")
            print(f"Erro ao inserir múltiplos registros: {e}")
            return False
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def get_all_registros(self) -> List[Tuple]:
        """Retorna todos os registros da tabela."""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT id, op, unidade, arquivos, qtde, nome, status FROM etiquetas ORDER BY id DESC')
            return cursor.fetchall()
        except Exception as e:
            print(f"Erro ao buscar registros: {e}")
            return []
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def get_registros_paginated(self, page: int = 1, page_size: int = 50) -> tuple:
        """
        Retorna registros paginados junto com o total de registros.

        Args:
            page (int): página (1-based)
            page_size (int): número de registros por página

        Returns:
            tuple: (lista_de_registros, total_registros)
        """
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()

            # Total
            cursor.execute('SELECT COUNT(*) FROM etiquetas')
            total = cursor.fetchone()[0] or 0

            # Calcula offset
            if page < 1:
                page = 1
            offset = (page - 1) * page_size

            cursor.execute(
                'SELECT id, op, unidade, arquivos, qtde, nome, status FROM etiquetas ORDER BY id DESC LIMIT %s OFFSET %s',
                (page_size, offset)
            )
            rows = cursor.fetchall()
            return rows, total
        except Exception as e:
            print(f"Erro ao buscar registros paginados: {e}")
            return [], 0
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def search_registros(self, campo: str, valor: str) -> List[Tuple]:
        """Busca registros por um campo específico."""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            if campo in ("op", "unidade", "arquivos", "nome", "status"):
                query = f'SELECT id, op, unidade, arquivos, qtde, nome, status FROM etiquetas WHERE {campo} LIKE %s ORDER BY id DESC'
                cursor.execute(query, (f'%{valor}%',))
                return cursor.fetchall()
            return []
        except Exception as e:
            print(f"Erro ao buscar registros: {e}")
            return []
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def delete_registro(self, registro_id: int) -> bool:
        """Deleta um registro específico."""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM etiquetas WHERE id = %s', (registro_id,))
            conn.commit()
            return cursor.rowcount > 0
        except Exception as e:
            print(f"Erro ao deletar registro: {e}")
            return False
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def update_status_by_op(self, op: str, status: str) -> bool:
        """Atualiza o status de todos os registros de uma OP específica"""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('UPDATE etiquetas SET status = %s WHERE op = %s', (status, op))
            conn.commit()
            return cursor.rowcount > 0
        except Exception as e:
            print(f"Erro ao atualizar status da OP {op}: {e}")
            return False
        finally:
            if conn:
                conn.close()

    def update_status_by_ids(self, ids: List[int], status: str) -> bool:
        """Atualiza o status de registros específicos por IDs"""
        if not ids:
            return False
            
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            placeholders = ','.join(['%s'] * len(ids))
            query = f'UPDATE etiquetas SET status = %s WHERE id IN ({placeholders})'
            cursor.execute(query, [status] + ids)
            conn.commit()
            return cursor.rowcount > 0
        except Exception as e:
            print(f"Erro ao atualizar status dos registros {ids}: {e}")
            return False
        finally:
            if conn:
                conn.close()

    def clear_all_registros(self) -> bool:
        """Limpa todos os registros da tabela."""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM etiquetas')
            conn.commit()
            return True
        except Exception as e:
            print(f"Erro ao limpar registros: {e}")
            return False
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def check_duplicates(self, registros: List[Tuple[str, str, str, int]]) -> dict:
        """
        Verifica se existem registros duplicados que seriam inseridos.
        
        Args:
            registros: Lista de tuplas (op, unidade, arquivos, qtde) para verificar
            
        Returns:
            dict: Resultado da verificação contendo duplicatas encontradas
        """
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            duplicatas = []
            novos = []
            
            for registro in registros:
                # Suporta tanto formato antigo (4 elementos) quanto novo (5 elementos)
                if len(registro) == 4:
                    op, unidade, arquivos, qtde = registro
                    nome = ""
                else:
                    op, unidade, arquivos, qtde, nome = registro[:5]
                
                # Verifica se já existe registro com mesma OP, unidade e arquivo
                cursor.execute('''
                    SELECT id, op, unidade, arquivos, qtde, nome, status 
                    FROM etiquetas 
                    WHERE op = ? AND unidade = ? AND arquivos = ?
                ''', (op, unidade, arquivos))
                
                existente = cursor.fetchone()
                
                if existente:
                    duplicatas.append({
                        'novo': (op, unidade, arquivos, qtde, nome),
                        'existente': existente,
                        'mesmo_qtde': existente[4] == qtde
                    })
                else:
                    novos.append(registro)
            
            return {
                'duplicatas': duplicatas,
                'novos': novos,
                'total_duplicatas': len(duplicatas),
                'total_novos': len(novos)
            }
            
        except Exception as e:
            print(f"Erro ao verificar duplicatas: {e}")
            return {
                'duplicatas': [],
                'novos': registros,
                'total_duplicatas': 0,
                'total_novos': len(registros)
            }
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def get_statistics(self) -> dict:
        """Retorna estatísticas dos dados."""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT COUNT(*) FROM etiquetas')
            total_registros = cursor.fetchone()[0]
            cursor.execute('SELECT COUNT(DISTINCT op) FROM etiquetas')
            total_ops = cursor.fetchone()[0]
            cursor.execute('SELECT COUNT(DISTINCT unidade) FROM etiquetas')
            total_unidades = cursor.fetchone()[0]
            cursor.execute('SELECT SUM(qtde) FROM etiquetas')
            total_qtde = cursor.fetchone()[0] or 0
            return {
                'total_registros': total_registros,
                'total_ops': total_ops,
                'total_unidades': total_unidades,
                'total_quantidade': total_qtde
            }
        except Exception as e:
            print(f"Erro ao obter estatísticas: {e}")
            return {
                'total_registros': 0,
                'total_ops': 0,
                'total_unidades': 0,
                'total_quantidade': 0
            }

    def get_groups_summary(self) -> List[Tuple]:
        """
        Retorna um resumo por grupo (op) com contagem de itens e soma de quantidade.

        Returns:
            List[Tuple]: Lista de tuplas (op, total_itens, total_qtde)
        """
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                SELECT op, COUNT(*) as total_itens, SUM(qtde) as total_qtde
                FROM etiquetas
                GROUP BY op
                ORDER BY op DESC
            ''')
            return cursor.fetchall()
        except Exception as e:
            print(f"Erro ao obter resumo de grupos: {e}")
            return []
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass
