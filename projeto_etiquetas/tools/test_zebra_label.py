from service.pdf_service import PDFService

pdf = PDFService()
registros = [(1, 'OP12345', 'Unidade X', 'Arquivo muito longo que precisa quebrar em múltiplas linhas para caber na etiqueta', 2)]
output = 'zebra_test.pdf'

success = pdf.generate_labels_pdf(registros, output, label_size_mm=(100,50), single_per_page=True)
print('Generated:', success, '->', output)
