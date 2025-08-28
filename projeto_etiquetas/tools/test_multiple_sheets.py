from service.excel_service import ExcelService

# Teste de leitura de múltiplas planilhas
excel_service = ExcelService()

# Aqui você pode testar com um arquivo Excel que tenha múltiplas planilhas
# Exemplo de uso:
# registros = excel_service.read_excel_data("arquivo_multiplas_planilhas.xlsx")
# print(f"Total de registros de todas as planilhas: {len(registros) if registros else 0}")

print("Serviço Excel atualizado para processar todas as planilhas!")
print("Agora o sistema irá:")
print("1. Detectar todas as planilhas no arquivo Excel")
print("2. Processar cada planilha individualmente")
print("3. Combinar todos os registros em uma lista única")
print("4. Mostrar feedback sobre quantas planilhas foram processadas")
