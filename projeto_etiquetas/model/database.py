import sqlite3
import os
from typing import List, Tuple


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

    def insert_registro(self, op: str, unidade: str, arquivos: str, qtde: int) -> bool:
        """Insere um novo registro na tabela."""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO etiquetas (op, unidade, arquivos, qtde)
                VALUES (?, ?, ?, ?)
            ''', (op, unidade, arquivos, qtde))
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

    def insert_multiple_registros(self, registros: List[Tuple[str, str, str, int]]) -> bool:
        """Insere múltiplos registros de uma vez."""
        conn = None
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.executemany('''
                INSERT INTO etiquetas (op, unidade, arquivos, qtde)
                VALUES (?, ?, ?, ?)
            ''', registros)
            try:
                conn.commit()
            except Exception:
                pass
            return True
        except Exception as e:
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
            cursor.execute('SELECT id, op, unidade, arquivos, qtde FROM etiquetas ORDER BY id DESC')
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
            if campo in ("op", "unidade", "arquivos"):
                query = f'SELECT id, op, unidade, arquivos, qtde FROM etiquetas WHERE {campo} LIKE ? ORDER BY id DESC'
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
