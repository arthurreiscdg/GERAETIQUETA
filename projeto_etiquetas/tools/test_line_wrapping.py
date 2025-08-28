"""
Script de teste para verificar quebra de linhas nas etiquetas
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from service.pdf_service import PDFService

def test_line_wrapping():
    """Testa a quebra de linhas em etiquetas com texto longo"""
    print("Testando quebra de linhas em etiquetas...")
    
    try:
        # Cria instância do serviço PDF
        pdf_service = PDFService()
        
        # Dados de teste com textos longos
        registros_teste = [
            (1, "OP: Simulados EEAr + EPCAr 9° a 2° Militar (21/07)", "Itaguaí", "SIMULADO EPCAr -1° Mil... Este é um nome de arquivo muito longo que deveria quebrar automaticamente", 25),
            (2, "OP: GoA Jr + N1 + N2 + N3 (27/08)", "Visão Campinas", "Quantidade Total - Arquivo com nome extremamente longo para testar quebra", 15),
            (3, "OP123", "Unidade Teste", "arquivo_simples.pdf", 5),
        ]
        
        # Gera PDF de teste
        output_path = "teste_quebra_linhas.pdf"
        
        print(f"Gerando PDF de teste: {output_path}")
        success = pdf_service.generate_labels_pdf(registros_teste, output_path)
        
        if success:
            print("✓ PDF gerado com sucesso!")
            print(f"✓ Arquivo salvo como: {output_path}")
            print("✓ Verifique o PDF para confirmar que as linhas longas foram quebradas corretamente")
            
            # Verifica se o arquivo foi criado
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                print(f"✓ Tamanho do arquivo: {file_size} bytes")
            else:
                print("✗ Arquivo PDF não foi encontrado após geração")
                return False
        else:
            print("✗ Falha ao gerar PDF")
            return False
        
        return True
        
    except Exception as e:
        print(f"✗ Erro durante teste: {e}")
        return False

if __name__ == "__main__":
    success = test_line_wrapping()
    print(f"\nResultado do teste: {'SUCESSO' if success else 'FALHA'}")
    if success:
        print("\nAbra o arquivo 'teste_quebra_linhas.pdf' para verificar se as linhas longas foram quebradas corretamente.")
    sys.exit(0 if success else 1)
