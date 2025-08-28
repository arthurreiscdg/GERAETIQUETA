import sqlite3
import os
from typing import List, Tuple


class Database:
    def __init__(self, db_path: str = "etiquetas.db"):
        """
        Inicializa a conexão com o banco de dados SQLite.

        Se a variável de ambiente SQLITE_DB_URL estiver definida, ela sobrescreve o
        argumento db_path (útil para conexões como sqlitecloud://...).
        """
        # Permite sobrescrever pelo ambiente: SQLITE_DB_URL (útil para sqlitecloud)
        env_db = os.environ.get("SQLITE_DB_URL")
        self.db_path = env_db or db_path

        # Detecta se deve usar sqlitecloud (URI começando com sqlitecloud://)
        self.use_cloud = isinstance(self.db_path, str) and self.db_path.startswith("sqlitecloud://")

        # Verifica disponibilidade do pacote sqlitecloud se for necessário
        self._cloud_available = False
        if self.use_cloud:
            try:
                import sqlitecloud  # type: ignore
                self._cloud_available = True
            except ImportError:
                print("Pacote 'sqlitecloud' não está instalado. Instale com: pip install sqlitecloud")

        # Inicializa esquema (create table se necessário)
        # se a conexão em nuvem for requerida mas o pacote estiver ausente,
        # init_database será chamada mesmo assim - métodos posteriores vão
        # gerar um erro claro ao tentar abrir a conexão.
        self.init_database()

    def init_database(self):
        """Cria a tabela se ela não existir"""
        conn = None
        try:
            if self.use_cloud:
                if not getattr(self, "_cloud_available", False):
                    # Não tentamos criar esquema na nuvem se o pacote estiver ausente
                    raise ImportError("sqlitecloud não disponível")
                import sqlitecloud  # type: ignore
                conn = sqlitecloud.connect(self.db_path)
            else:
                conn = sqlite3.connect(self.db_path)

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
        except ImportError:
            print("Pacote 'sqlitecloud' não está instalado. Instale com: pip install sqlitecloud")
        except Exception as e:
            print(f"Erro ao inicializar banco de dados: {e}")
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def _get_connection(self):
        """Retorna uma nova conexão, seja local (sqlite3) ou via sqlitecloud."""
        if self.use_cloud:
            if not getattr(self, "_cloud_available", False):
                raise RuntimeError(
                    "Conector 'sqlitecloud' não está instalado. Instale com: pip install sqlitecloud "
                    "ou ajuste a variável de ambiente SQLITE_DB_URL para um caminho local."
                )
            import sqlitecloud  # type: ignore

            return sqlitecloud.connect(self.db_path)
        return sqlite3.connect(self.db_path)

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
