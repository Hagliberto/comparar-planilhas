# Comparador de Verbas - Sistema Modernizado

## 📋 Visão Geral

Este projeto é uma versão modernizada do sistema de comparação de verbas, desenvolvido com uma arquitetura robusta e interface moderna usando React e Flask.

## 🏗️ Arquitetura

### Frontend
- **Tecnologia**: React 18 com Material-UI
- **Porta**: 3000
- **Características**:
  - Interface responsiva e moderna
  - Ícones do Google Material Design
  - Stepper visual para guiar o usuário
  - Sistema de notificações integrado
  - Validação de arquivos em tempo real

### Backend
- **Tecnologia**: Flask com Python
- **Porta**: 5000
- **Características**:
  - API RESTful
  - Processamento robusto de dados
  - Geração de Excel com formatação
  - CORS habilitado
  - Tratamento de erros

## 🚀 Como Executar

### Pré-requisitos
- Python 3.11+
- Node.js 20+
- npm

### Backend (Flask)
```bash
cd backend
pip install -r requirements.txt
python app.py
```

### Frontend (React)
```bash
cd frontend
npm install
npm start
```

## 📁 Estrutura do Projeto

```
projeto/
├── backend/
│   ├── app.py              # Servidor Flask principal
│   ├── comparador.py       # Lógica de comparação
│   └── requirements.txt    # Dependências Python
├── frontend/
│   ├── src/
│   │   ├── App.js         # Componente principal React
│   │   └── index.js       # Ponto de entrada
│   ├── public/
│   │   └── index.html     # Template HTML
│   └── package.json       # Dependências Node.js
└── README.md              # Este arquivo
```

## 🔧 Funcionalidades

### Upload de Arquivos
- Suporte apenas para arquivos .xlsx
- Validação automática de formato
- Feedback visual do status

### Processamento
- Leitura das 3 primeiras colunas: Matrícula, Verba, Valor
- Normalização de dados (remoção de zeros à esquerda)
- Conversão automática de vírgulas para pontos
- Tolerância de 0.01 para comparações

### Resultado
- Arquivo Excel com 3 abas:
  - **Diferenças**: Apenas registros com divergências
  - **Estatísticas**: Resumo dos dados
  - **Lado a Lado**: Comparação completa
- Formatação colorida para destacar diferenças
- Download automático

## 🎨 Interface

### Características da UI
- Design Material Design
- Cores profissionais (azul primário)
- Stepper para guiar o processo
- Cards organizados
- Instruções integradas
- Status em tempo real

### Fluxo do Usuário
1. **Carregar Planilhas**: Upload dos dois arquivos .xlsx
2. **Configurar Opções**: Definir se há cabeçalho
3. **Comparar Dados**: Processamento automático
4. **Baixar Resultado**: Download do Excel gerado

## 🔍 Validações

### Arquivos
- Formato .xlsx obrigatório
- Mínimo de 3 colunas
- Dados válidos na coluna Valor

### Dados
- Matrícula: Texto (zeros à esquerda removidos)
- Verba: Texto
- Valor: Numérico (vírgula convertida para ponto)

## 🚨 Tratamento de Erros

- Validação de formato de arquivo
- Verificação de estrutura das planilhas
- Tratamento de valores inválidos
- Mensagens de erro claras
- Logs no backend

## 📊 Comparação

### Tipos de Diferenças
- **Apenas na Planilha 1**: Registro existe só na primeira
- **Apenas na Planilha 2**: Registro existe só na segunda
- **Valor Divergente**: Diferença > 0.01 entre valores

### Formatação Excel
- Verde: Valores maiores na Planilha 1
- Vermelho: Valores menores na Planilha 1
- Amarelo: Dados ausentes em uma planilha

## 🔧 Configuração de Desenvolvimento

### Variáveis de Ambiente
- Backend roda em `0.0.0.0:5000`
- Frontend roda em `localhost:3000`
- CORS configurado para desenvolvimento

### Dependências Principais

#### Backend
- Flask 3.1.1
- Flask-CORS 6.0.1
- Pandas 2.3.0
- OpenPyXL 3.1.5
- XlsxWriter 3.2.3

#### Frontend
- React 18.3.1
- Material-UI 5.15.19
- Material Icons

## 🐛 Solução de Problemas

### Backend não inicia
- Verificar se todas as dependências estão instaladas
- Confirmar que a porta 5000 está livre

### Frontend não carrega
- Executar `npm install` novamente
- Verificar se a porta 3000 está livre
- Confirmar que o Node.js está atualizado

### Erro de CORS
- Verificar se o backend está rodando
- Confirmar configuração de CORS no Flask

### Upload falha
- Verificar formato do arquivo (.xlsx)
- Confirmar estrutura das colunas
- Verificar tamanho do arquivo

## 📈 Melhorias Implementadas

### Em relação ao sistema original:
1. **Interface moderna** com Material-UI
2. **Arquitetura separada** (frontend/backend)
3. **Validações robustas** de entrada
4. **Feedback visual** em tempo real
5. **Tratamento de erros** aprimorado
6. **Responsividade** para diferentes telas
7. **Instruções integradas** na interface
8. **Sistema de notificações** com Snackbar

## 🎯 Próximos Passos

Para produção, considere:
- Configurar servidor WSGI (Gunicorn)
- Implementar autenticação
- Adicionar logs estruturados
- Configurar monitoramento
- Implementar testes automatizados
- Otimizar performance para arquivos grandes

