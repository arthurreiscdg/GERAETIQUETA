import sqlite3
import os
import logging
from typing import List, Tuple

# Configurar logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class Database:
    def __init__(self, db_path: str = None):
        """
        Inicializa conexão EXCLUSIVA com sqlitecloud.
        
        Usa a variável de ambiente SQLITE_DB_URL ou string hardcoded para o projeto.
        """
        # Prioriza variável de ambiente, senão usa string de conexão padrão do projeto
        env_db = os.environ.get("SQLITE_DB_URL")
        if env_db:
            self.db_path = env_db
        else:
            # String de conexão hardcoded para este projeto
            self.db_path = "sqlitecloud://cv0idhxxhk.g2.sqlite.cloud:8860/auth.sqlitecloud?apikey=4gtJpnQlCzrAfmGgn9QOdDrFDvalmk3APBcawzNvssc"

        # Valida formato da URI
        if not isinstance(self.db_path, str) or not self.db_path.startswith("sqlitecloud://"):
            raise RuntimeError(f"URI deve começar com 'sqlitecloud://'. Valor atual: {self.db_path}")

        self.use_cloud = True

        # Verifica disponibilidade do pacote sqlitecloud
        try:
            import sqlitecloud  # type: ignore
            self._cloud_available = True
        except ImportError:
            raise RuntimeError("Pacote 'sqlitecloud' não está instalado. Instale com: pip install sqlitecloud")

        # Inicializa esquema (create table se necessário)
        self.init_database()
        
        # Migra a tabela para adicionar coluna nome se necessário
        self._migrate_add_nome_column()

    def init_database(self):
        """Cria a tabela se ela não existir no sqlitecloud"""
        conn = None
        try:
            import sqlitecloud  # type: ignore
            conn = sqlitecloud.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS etiquetas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    op TEXT NOT NULL,
                    unidade TEXT NOT NULL,
                    arquivos TEXT NOT NULL,
                    qtde INTEGER NOT NULL,
                    nome TEXT DEFAULT '',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            try:
                conn.commit()
            except Exception:
                # alguns conectores na nuvem podem auto-commit ou não implementar commit
                pass
        except Exception as e:
            print(f"Erro ao inicializar banco de dados: {e}")
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def _get_connection(self):
        """Retorna uma nova conexão via sqlitecloud."""
        import sqlitecloud  # type: ignore
        return sqlitecloud.connect(self.db_path)

    def _migrate_add_nome_column(self):
        """Adiciona a coluna 'nome' à tabela se ela não existir"""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            # Verifica se a coluna 'nome' já existe
            cursor.execute("PRAGMA table_info(etiquetas)")
            columns = [column[1] for column in cursor.fetchall()]
            
            if 'nome' not in columns:
                print("Adicionando coluna 'nome' à tabela etiquetas...")
                cursor.execute("ALTER TABLE etiquetas ADD COLUMN nome TEXT DEFAULT ''")
                try:
                    conn.commit()
                    print("Coluna 'nome' adicionada com sucesso!")
                except Exception:
                    pass
            else:
                print("Coluna 'nome' já existe na tabela.")
                
        except Exception as e:
            print(f"Erro ao migrar tabela: {e}")
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def insert_registro(self, op: str, unidade: str, arquivos: str, qtde: int, nome: str = "") -> bool:
        """Insere um novo registro na tabela."""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO etiquetas (op, unidade, arquivos, qtde, nome)
                VALUES (?, ?, ?, ?, ?)
            ''', (op, unidade, arquivos, qtde, nome))
            try:
                conn.commit()
            except Exception:
                pass
            return True
        except Exception as e:
            print(f"Erro ao inserir registro: {e}")
            return False
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

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
                VALUES (?, ?, ?, ?, ?)
            ''', registros_processados)
            try:
                conn.commit()
            except Exception:
                pass
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
            cursor.execute('SELECT id, op, unidade, arquivos, qtde, nome FROM etiquetas ORDER BY id DESC')
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

    def search_registros(self, campo: str, valor: str) -> List[Tuple]:
        """Busca registros por um campo específico."""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            if campo in ("op", "unidade", "arquivos", "nome"):
                query = f'SELECT id, op, unidade, arquivos, qtde, nome FROM etiquetas WHERE {campo} LIKE ? ORDER BY id DESC'
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
            cursor.execute('DELETE FROM etiquetas WHERE id = ?', (registro_id,))
            try:
                conn.commit()
            except Exception:
                pass
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

    def clear_all_registros(self) -> bool:
        """Limpa todos os registros da tabela."""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM etiquetas')
            try:
                conn.commit()
            except Exception:
                pass
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
                    SELECT id, op, unidade, arquivos, qtde, nome 
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
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass
