"""
Script de teste para verificar quebra de linhas e logo nas etiquetas
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from service.pdf_service import PDFService

def test_labels_with_logo():
    """Testa a quebra de linhas e logo em etiquetas"""
    print("Testando etiquetas com quebra de linhas e logo CDG...")
    
    try:
        # Cria instância do serviço PDF
        pdf_service = PDFService()
        
        # Dados de teste com textos longos (baseados na imagem fornecida)
        registros_teste = [
            (1, "Simulados EEAr + EPCAr 9° a 2° Militar (21/07)", "Itaguaí", "SIMULADO EPCAr -1° Mil... Este é um nome de arquivo muito longo que deveria quebrar automaticamente em múltiplas linhas", 25),
            (2, "GoA Jr + N1 + N2 + N3 (27/08)", "Visão Campinas", "Quantidade Total - Arquivo com nome extremamente longo para testar quebra de linhas e posicionamento da logo", 15),
            (3, "OP Teste Curta", "Unidade Teste", "arquivo_simples.pdf", 5),
            (4, "OP: Preparatório para Concursos Militares - Curso Completo EsPCEx + AMAN + AFA + EPCAr + EEAr + CFS + CHO + CFO", "Rio de Janeiro", "Material_Didatico_Completo_Matematica_Fisica_Quimica_Portugues_Historia_Geografia_2024_2025.pdf", 50),
        ]
        
        # Gera PDF de teste
        output_path = "teste_etiquetas_com_logo.pdf"
        
        print(f"Gerando PDF de teste: {output_path}")
        success = pdf_service.generate_labels_pdf(registros_teste, output_path)
        
        if success:
            print("✓ PDF gerado com sucesso!")
            print(f"✓ Arquivo salvo como: {output_path}")
            print("✓ Verifique o PDF para confirmar:")
            print("  - Linhas longas foram quebradas corretamente")
            print("  - Logo CDG aparece no canto superior direito")
            print("  - Texto não sobrepõe a logo")
            
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
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_labels_with_logo()
    print(f"\nResultado do teste: {'SUCESSO' if success else 'FALHA'}")
    if success:
        print("\nAbra o arquivo 'teste_etiquetas_com_logo.pdf' para verificar:")
        print("- Quebra de linhas automática")
        print("- Logo CDG no canto superior direito")
        print("- Layout otimizado")
    sys.exit(0 if success else 1)
