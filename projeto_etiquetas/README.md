# Sistema de Gestão de Etiquetas

Sistema completo em Python para importação de dados do Excel, armazenamento em banco SQLite e geração de etiquetas em PDF.

## 📋 Funcionalidades

- **Importação do Excel**: Lê dados de planilhas com estrutura fixa
- **Banco de dados**: Armazena dados em SQLite (4 colunas: OP, unidade, arquivos, qtde)
- **Interface gráfica**: Tkinter com botões intuitivos
- **Pesquisa**: Filtros por OP, unidade ou arquivo
- **Geração de etiquetas**: PDF com layout profissional
- **Relatórios**: PDF com lista de registros
- **Gerenciamento**: Visualizar, excluir e limpar registros

## 🗂️ Estrutura do Projeto

```
projeto_etiquetas/
├── main.py                     # Arquivo principal para rodar o sistema
├── controller/
│   └── etiqueta_controller.py  # Lógica de negócio
├── model/
│   └── database.py             # Gerenciamento do banco SQLite
├── view/
│   └── etiqueta_view.py        # Interface gráfica Tkinter
├── service/
│   ├── excel_service.py        # Leitura e importação do Excel
│   └── pdf_service.py          # Geração de etiquetas em PDF
└── requirements.txt            # Dependências do projeto
```

## 📊 Estrutura do Excel

O sistema espera uma estrutura fixa no arquivo Excel:

| A | B |
|---|---|
| **OP001** | **UNIDADE_TESTE** |
| arquivo1.txt | 5 |
| arquivo2.pdf | 3 |
| arquivo3.docx | 2 |

- **A1**: OP (identificador da ordem de produção)
- **A2+**: Arquivos dessa OP
- **B1**: Unidade (nome da unidade)
- **B2+**: Quantidade de cada arquivo

## 🚀 Instalação e Uso

### 1. Instalar Dependências

```bash
pip install -r requirements.txt
```

Ou instalar individualmente:
```bash
pip install pandas openpyxl reportlab Pillow
```

### 2. Executar o Sistema

```bash
python main.py
```

### 3. Criar Arquivo de Exemplo

```bash
python main.py --sample
```

### 4. Ver Ajuda

```bash
python main.py --help
```

## 🖥️ Interface do Sistema

### Painel de Ações (Esquerda)
- **📁 Importar Excel**: Seleciona e importa arquivo Excel
- **🔍 Pesquisar**: Busca por OP, unidade ou arquivo
- **🏷️ Etiquetas**: Gera PDF com etiquetas dos registros selecionados
- **📋 Relatório**: Gera PDF com lista dos registros
- **🔄 Atualizar**: Recarrega dados do banco
- **🗑️ Excluir**: Remove registros selecionados
- **⚠️ Limpar Tudo**: Remove todos os registros
- **ℹ️ Sobre**: Informações do sistema

### Painel de Dados (Direita)
- **Tabela**: Visualização de todos os registros
- **Colunas**: ID, OP, Unidade, Arquivo, Quantidade
- **Seleção múltipla**: Use Ctrl+clique para selecionar vários registros
- **Menu de contexto**: Clique direito para ações rápidas

### Barra de Status (Inferior)
- **Status**: Mostra o status atual da operação
- **Estatísticas**: Total de registros, OPs, unidades e quantidade

## 🏷️ Layout das Etiquetas

Cada etiqueta contém:
- **OP**: Ordem de produção
- **Unidade**: Nome da unidade
- **Arquivo**: Nome do arquivo
- **Contador**: Número atual/total (ex: 1/5, 2/5...)
- **Data/hora**: Timestamp da geração

### Configurações das Etiquetas
- **Tamanho**: 80mm x 50mm
- **Etiquetas por página**: 6 (2 colunas x 3 linhas)
- **Formato**: PDF A4

## 🗄️ Banco de Dados

O sistema cria automaticamente um arquivo `etiquetas.db` (SQLite) com a tabela:

```sql
CREATE TABLE etiquetas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    op TEXT NOT NULL,
    unidade TEXT NOT NULL, 
    arquivos TEXT NOT NULL,
    qtde INTEGER NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

## 📁 Arquivos Gerados

- **etiquetas.db**: Banco de dados SQLite
- **etiquetas_YYYYMMDD_HHMMSS.pdf**: Etiquetas geradas
- **relatorio_YYYYMMDD_HHMMSS.pdf**: Relatórios gerados
- **exemplo_excel.xlsx**: Arquivo de exemplo (com --sample)

## 🔧 Validações

### Excel
- ✅ Estrutura fixa (A1=OP, B1=unidade)
- ✅ Dados obrigatórios (OP, unidade, arquivo, quantidade)
- ✅ Quantidade deve ser número positivo
- ✅ Relatório de qualidade dos dados

### Interface
- ✅ Confirmações para ações destrutivas
- ✅ Validação de seleções
- ✅ Feedback visual de operações
- ✅ Tratamento de erros

## 🎯 Casos de Uso

### 1. Importar Novos Dados
1. Prepare Excel com estrutura correta
2. Clique "📁 Importar Excel"
3. Selecione o arquivo
4. Confirme a importação

### 2. Gerar Etiquetas
1. Selecione registros (ou deixe vazio para todos)
2. Clique "🏷️ Etiquetas"
3. Escolha local para salvar PDF
4. Etiquetas são geradas expandindo por quantidade

### 3. Pesquisar Registros
1. Escolha campo (OP, unidade, arquivos)
2. Digite valor para buscar
3. Resultados são filtrados em tempo real

### 4. Gerar Relatório
1. Selecione registros (ou deixe vazio para todos)
2. Clique "📋 Relatório"
3. PDF com lista é gerado

## ⚠️ Observações Importantes

- **Backup**: Faça backup do arquivo `etiquetas.db` regularmente
- **Excel**: Use apenas arquivos `.xlsx` ou `.xls`
- **Quantidade**: Cada unidade de quantidade gera uma etiqueta
- **Seleção**: Se nada estiver selecionado, usa todos os registros visíveis
- **Filtros**: Pesquisas são mantidas até serem limpas

## 🐛 Solução de Problemas

### Erro de Dependências
```bash
pip install --upgrade pandas openpyxl reportlab Pillow
```

### Erro de Importação do Excel
- Verifique se A1 contém a OP e B1 contém a unidade
- Certifique-se que há dados nas linhas A2+ e B2+
- Verifique se as quantidades são números válidos

### Erro de Geração de PDF
- Verifique permissões de escrita na pasta
- Certifique-se que o arquivo não está aberto em outro programa
- Tente um caminho mais simples (sem caracteres especiais)

### Interface Não Abre
- Verifique se Python tem suporte ao Tkinter
- No Ubuntu/Debian: `sudo apt-get install python3-tk`

## 📞 Suporte

Para dúvidas ou problemas:
1. Verifique este README
2. Execute `python main.py --help`
3. Teste com arquivo de exemplo: `python main.py --sample`

---

**Versão**: 1.0  
**Desenvolvido**: 2025  
**Compatibilidade**: Python 3.7+
