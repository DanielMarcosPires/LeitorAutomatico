# Leitor Automático de Planilhas Excel

Este projeto é uma aplicação desktop para gerenciamento automatizado de planilhas Excel, com interface gráfica moderna usando CustomTkinter.

## Funcionalidades

### Interface Principal
- **Login Seguro**: Formulário de autenticação para acesso ao sistema
- **Dashboard Intuitivo**: Interface com dois painéis (ações e lista de planilhas)

### Operações com Planilhas
- **Criar Nova Planilha**: Gera planilhas Excel formatadas automaticamente
- **Ler Planilhas Selecionadas**: Processa múltiplas planilhas simultaneamente
- **Visualizar Planilha**: Interface gráfica para visualizar dados com cores dinâmicas
- **Gerar Pastas**: Cria estrutura de diretórios organizada por cliente
- **Relatório por Cliente**: Gera relatórios financeiros consolidados

### Recursos Técnicos
- Detecção automática de status (Pago/Pendente) com cores visuais
- Processamento em lote de arquivos Excel
- Interface responsiva com scroll para grandes volumes de dados
- Sistema de logout seguro

## Estrutura do Projeto

```
LeitorAutomatico/
├── main.py                 # Ponto de entrada alternativo
├── Interface/
│   ├── app.py             # Aplicação principal
│   ├── form.py            # Formulário de login
│   ├── dashboard.py       # Dashboard principal
│   ├── Classes/
│   │   ├── excel.py       # Utilitários para Excel
│   │   ├── folders.py     # Gerenciamento de pastas
│   │   └── bgColors.py    # Definições de cores
│   └── Fonts/
│       └── fonts.py       # Configurações de fonte
├── Planilhas/             # Diretório para planilhas geradas
└── Clientes/              # Diretório para relatórios por cliente
```

## Tecnologias Utilizadas

- **Python 3.13**
- **CustomTkinter**: Interface gráfica moderna
- **OpenPyXL**: Manipulação de arquivos Excel
- **Glob**: Busca recursiva de arquivos
- **OS**: Operações do sistema de arquivos

## Como Executar

1. Instale o venv nesse projeto utilizando o comando:
```bash
   python -m venv venv
```
2. Após a instalação ative o Ambiente virtual usando o comando: 
```bash
   ./venv/Scripts/Activate
```
3. Com o venv instalado antes de executar o projeto precisamos instalar as dependencias do requirements.txt
com o seguinte comando:
```bash
   pip install -r requirements.txt
```
4. E por fim execute o projeto utilizando o comando:
```bash 
   python ./main.py
```

## Funcionalidades Detalhadas

### Visualização de Planilhas
- Exibe dados em grid responsivo
- Cores dinâmicas baseadas no conteúdo:
  - Verde: Status positivo (Pago, Sim)
  - Laranja: Status de atenção (Pendente, Não)
  - Branco: Texto padrão

### Geração de Relatórios
- Cria pastas organizadas por cliente
- Consolida dados financeiros automaticamente
- Estrutura hierárquica de diretórios

### Segurança
- Sistema de login obrigatório
- Logout seguro com retorno à tela inicial

## Desenvolvimento

O projeto segue boas práticas de organização de código:
- Separação clara entre interface e lógica de negócio
- Classes especializadas para cada funcionalidade
- Tratamento adequado de erros e exceções
- Interface responsiva e intuitiva

## Suporte

Para dúvidas ou problemas, verifique:
1. Compatibilidade das dependências
2. Permissões de escrita nos diretórios
3. Formato dos arquivos Excel de entrada