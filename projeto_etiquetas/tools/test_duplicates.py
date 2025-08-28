"""
Script de teste para verificar a funcionalidade de duplicatas
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from model.database import Database

def test_duplicates():
    """Testa a verificação de duplicatas"""
    print("Testando sistema de verificação de duplicatas...")
    
    try:
        # Cria instância do banco
        db = Database()
        
        # Dados de teste
        registros_teste = [
            ("OP001", "Unidade A", "arquivo1.pdf", 10),
            ("OP002", "Unidade B", "arquivo2.pdf", 5),
            ("OP003", "Unidade A", "arquivo3.pdf", 15),
        ]
        
        print("\n1. Inserindo registros de teste...")
        success = db.insert_multiple_registros(registros_teste)
        if success:
            print("✓ Registros de teste inseridos com sucesso")
        else:
            print("✗ Falha ao inserir registros de teste")
            return False
        
        # Testa verificação de duplicatas
        print("\n2. Testando verificação de duplicatas...")
        registros_com_duplicatas = [
            ("OP001", "Unidade A", "arquivo1.pdf", 10),  # Duplicata exata
            ("OP002", "Unidade B", "arquivo2.pdf", 8),   # Duplicata com qtde diferente
            ("OP004", "Unidade C", "arquivo4.pdf", 20),  # Registro novo
        ]
        
        resultado = db.check_duplicates(registros_com_duplicatas)
        
        print(f"Total de duplicatas encontradas: {resultado['total_duplicatas']}")
        print(f"Total de registros novos: {resultado['total_novos']}")
        
        if resultado['total_duplicatas'] > 0:
            print("\nDuplicatas encontradas:")
            for dup in resultado['duplicatas']:
                novo = dup['novo']
                existente = dup['existente']
                print(f"  - OP: {novo[0]}, Unidade: {novo[1]}, Arquivo: {novo[2]}")
                print(f"    Qtde nova: {novo[3]}, Qtde existente: {existente[4]}")
                print(f"    Mesma quantidade: {dup['mesmo_qtde']}")
        
        if resultado['total_novos'] > 0:
            print("\nRegistros novos:")
            for novo in resultado['novos']:
                print(f"  - OP: {novo[0]}, Unidade: {novo[1]}, Arquivo: {novo[2]}, Qtde: {novo[3]}")
        
        print("\n✓ Teste de duplicatas concluído com sucesso!")
        return True
        
    except Exception as e:
        print(f"✗ Erro durante teste: {e}")
        return False

if __name__ == "__main__":
    success = test_duplicates()
    print(f"\nResultado do teste: {'SUCESSO' if success else 'FALHA'}")
    sys.exit(0 if success else 1)
