#!/usr/bin/env python3
"""
Sistema de Gestão de Etiquetas
==============================

Sistema para importação de dados do Excel, armazenamento em banco SQLite
e geração de etiquetas em PDF.

Estrutura do Excel esperada:
- A1 = OP (ordem de produção)
- A2+ = arquivos dessa OP
- B1 = unidade (nome da unidade)  
- B2+ = quantidade

Funcionalidades:
- Importar dados do Excel
- Visualizar e pesquisar registros
- Gerar etiquetas em PDF
- Gerar relatórios em PDF
- Gerenciar registros (excluir, limpar)

Autor: Sistema Automático
Data: 2025
"""

import sys
import os
import tkinter as tk
from tkinter import messagebox

# Adiciona o diretório do projeto ao path para imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def check_dependencies():
    """
    Verifica se todas as dependências estão instaladas
    
    Returns:
        bool: True se todas as dependências estão disponíveis
    """
    missing_packages = []
    
    try:
        import pandas
    except ImportError:
        missing_packages.append("pandas")
    
    try:
        import openpyxl
    except ImportError:
        missing_packages.append("openpyxl")
    
    try:
        import reportlab
    except ImportError:
        missing_packages.append("reportlab")
    
    try:
        from PIL import Image
    except ImportError:
        missing_packages.append("Pillow")
    
    if missing_packages:
        error_msg = f"""Dependências não encontradas: {', '.join(missing_packages)}

Para instalar as dependências, execute:
pip install {' '.join(missing_packages)}

Ou instale todas as dependências do projeto:
pip install -r requirements.txt"""
        
        print(error_msg)
        
        # Tenta mostrar messagebox se tkinter estiver disponível
        try:
            root = tk.Tk()
            root.withdraw()  # Esconde a janela principal
            messagebox.showerror("Dependências não encontradas", error_msg)
            root.destroy()
        except:
            pass
        
        return False
    
    return True

def create_sample_excel():
    """
    Cria um arquivo Excel de exemplo para teste
    """
    try:
        import pandas as pd
        
        # Dados de exemplo
        data = {
            'A': ['OP001', 'arquivo1.txt', 'arquivo2.pdf', 'arquivo3.docx'],
            'B': ['UNIDADE_TESTE', 5, 3, 2]
        }
        
        df = pd.DataFrame(data)
        
        sample_file = "exemplo_excel.xlsx"
        df.to_excel(sample_file, index=False, header=False)
        
        print(f"Arquivo de exemplo criado: {sample_file}")
        return sample_file
        
    except Exception as e:
        print(f"Erro ao criar arquivo de exemplo: {e}")
        return None

def main():
    """
    Função principal da aplicação
    """
    print("=" * 50)
    print("Sistema de Gestão de Etiquetas")
    print("=" * 50)
    
    # Verifica dependências
    if not check_dependencies():
        print("\nErro: Dependências não encontradas!")
        print("Instale as dependências e tente novamente.")
        input("\nPressione Enter para sair...")
        return False
    
    print("Dependências verificadas: OK")
    
    # Importa e executa a aplicação
    try:
        from view.etiqueta_view import EtiquetaView
        
        print("Iniciando aplicação...")
        
        app = EtiquetaView()
        app.run()
        
        print("Aplicação encerrada.")
        return True
        
    except ImportError as e:
        error_msg = f"Erro ao importar módulos da aplicação: {e}"
        print(error_msg)
        messagebox.showerror("Erro de Importação", error_msg)
        return False
        
    except Exception as e:
        error_msg = f"Erro inesperado: {e}"
        print(error_msg)
        messagebox.showerror("Erro", error_msg)
        return False

def print_help():
    """
    Imprime informações de ajuda
    """
    help_text = """
Sistema de Gestão de Etiquetas - Ajuda
=====================================

USO:
    python main.py                 - Inicia a aplicação
    python main.py --help          - Mostra esta ajuda
    python main.py --sample        - Cria arquivo Excel de exemplo

ESTRUTURA DO EXCEL:
    A1: OP (ordem de produção)     - Ex: "OP001"
    A2+: arquivos desta OP         - Ex: "arquivo1.txt", "arquivo2.pdf"
    B1: unidade                    - Ex: "UNIDADE_TESTE"
    B2+: quantidade               - Ex: 5, 3, 2

DEPENDÊNCIAS:
    - pandas>=2.0.3
    - openpyxl>=3.1.2
    - reportlab>=4.0.4
    - Pillow>=10.0.0
    - tkinter (incluído no Python)

INSTALAÇÃO DAS DEPENDÊNCIAS:
    pip install -r requirements.txt

FUNCIONALIDADES:
    - Importar dados do Excel
    - Visualizar e pesquisar registros
    - Gerar etiquetas em PDF
    - Gerar relatórios em PDF
    - Gerenciar registros

ARQUIVOS GERADOS:
    - etiquetas.db: Banco de dados SQLite
    - etiquetas_*.pdf: Etiquetas geradas
    - relatorio_*.pdf: Relatórios gerados
"""
    print(help_text)

if __name__ == "__main__":
    # Verifica argumentos da linha de comando
    if len(sys.argv) > 1:
        if sys.argv[1] in ["--help", "-h", "help"]:
            print_help()
            sys.exit(0)
        elif sys.argv[1] in ["--sample", "-s", "sample"]:
            print("Criando arquivo Excel de exemplo...")
            sample_file = create_sample_excel()
            if sample_file:
                print(f"Arquivo criado: {sample_file}")
                print("\nUse este arquivo para testar a importação.")
            sys.exit(0)
        else:
            print(f"Argumento desconhecido: {sys.argv[1]}")
            print("Use --help para ver as opções disponíveis.")
            sys.exit(1)
    
    # Executa a aplicação
    try:
        success = main()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\nAplicação interrompida pelo usuário.")
        sys.exit(0)
    except Exception as e:
        print(f"\nErro fatal: {e}")
        sys.exit(1)
