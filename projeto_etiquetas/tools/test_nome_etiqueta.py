from service.pdf_service import PDFService

# Teste com dados que incluem nome
pdf = PDFService()

# Simula registros com nome: (id, op, unidade, arquivo, qtde, nome)
registros_com_nome = [
    (1, 'OP12345', 'Unidade X', 'arquivo1.pdf', 5, 'João Silva'),
    (2, 'OP12346', 'Unidade Y', 'arquivo2.pdf', 3, 'Maria Santos'),
    (3, 'OP12347', 'Unidade Z', 'arquivo3.pdf', 7, '')  # sem nome
]

output = 'teste_com_nome.pdf'
success = pdf.generate_labels_pdf(registros_com_nome, output, label_size_mm=(100,50), single_per_page=True)

print(f'PDF gerado: {success} -> {output}')
print('As etiquetas agora devem incluir o campo Nome!')
print('Registros testados:')
for reg in registros_com_nome:
    print(f'  OP: {reg[1]}, Nome: {reg[5] if reg[5] else "(vazio)"}')
