🧰 Controle de Equipamentos
Sistema desktop desenvolvido em Python com interface gráfica (PyQt6) integrado ao Google Sheets , permitindo o registro, visualização e geração de relatórios de chamados técnicos de equipamentos.

📌 Sumário
Visão Geral
Funcionalidades
Tecnologias Utilizadas
Pré-Requisitos
Como Executar
Estrutura do Projeto
Detalhes Técnicos
Manual do Usuário
Recomendações Futuras
Licença
Contribuição
👀 Visão Geral
Este é um sistema de controle de equipamentos projetado para facilitar a gestão de chamados técnicos. Ele permite:

Registro manual de chamados
Envio automático dos dados para uma planilha do Google Sheets
Visualização em tabela com destaque por status
Exportação de relatório PDF com gráfico e análise filtrada
Ordenação alfabética da tabela
✅ Funcionalidades
Registro de Chamados
Preencha os campos necessários e envie diretamente para a planilha.
Carregar Dados
Exibe todos os chamados já registrados na planilha.
Destaque Visual
Linhas coloridas conforme o status: verde (concluído), vermelho (crítico) e amarelo (em andamento).
Filtragem por Período
Selecione um intervalo de datas para exibir apenas chamados relevantes.
Gerar Relatório PDF
Exporte os dados filtrados com gráficos e tabelas detalhadas.
Ordenação
Botões para ordenar os registros alfabeticamente pela coluna "Tipo".

⚙️ Tecnologias Utilizadas
Python 3
PyQt6 : Interface gráfica
gspread + OAuth2Client : Integração com Google Sheets
FPDF + matplotlib : Criação de relatórios PDF com gráficos
datetime : Manipulação de datas
unicodedata : Normalização de texto
🔧 Pré-Requisitos
Instale as dependências usando pip:

bash


1
pip install PyQt6 gspread oauth2client fpdf matplotlib
📁 Configurações Necessárias
Ative a API do Google Sheets no Google Cloud Console
Crie uma conta de serviço (Service Account) e baixe o arquivo .json
Compartilhe sua planilha com o email gerado no arquivo JSON
Substitua as seguintes variáveis no código:
python


1
2
CREDENTIALS_FILE = r"seu/caminho/do/arquivo.json"
SHEET_ID = "ID_da_sua_planilha_aqui"
▶️ Como Executar
Clone o repositório:
bash


1
git clone https://github.com/seuusuario/controle-equipamentos.git
Instale as dependências:
bash


1
pip install -r requirements.txt
Execute o programa:
bash


1
python main.py
📁 Estrutura do Projeto


1
2
3
4
5
controle-equipamentos/
│
├── main.py               → Código principal com interface e lógica
├── credentials.json      → Arquivo de credenciais do Google Cloud Project (você deve adicionar)
└── requirements.txt      → Lista das dependências Python
💡 Detalhes Técnicos
O projeto é dividido em:

📦 Funções Auxiliares
normalizar(texto) → Remove acentos e converte texto para minúsculas
calcular_dias_corridos(data_str) → Dias entre data atual e data do chamado
calcular_dias_uteis(data_str) → Conta dias úteis (segunda a sexta-feira)
🖥️ Classe Principal: EquipamentoApp
enviar_dados() → Envia novos dados para o Google Sheets
carregar_dados_tabela() → Busca dados existentes e preenche a tabela GUI
ordenar_tabela(coluna, ordem) → Organiza a tabela por tipo
aplicar_cores() → Aplica cores às linhas conforme status
gerar_pdf() → Exporta dados filtrados em formato PDF com gráficos
📋 Manual do Usuário
Registrar Chamado
Preencha todos os campos obrigatórios:

Tipo Equipamento
Tombamento
Origem
Endereço
Nº O.S
Ticket TCORP
Situação
Data Chamado
A Data Finalização é opcional. 

Enviar para Planilha
Clique no botão “Enviar para Planilha” para salvar os dados na planilha do Google Sheets.

Visualizar Dados
Os dados já existentes aparecerão automaticamente na tabela. As cores indicam:

✅ Verde: Concluído
⚠️ Vermelho: Crítico (mais de 4 dias úteis sem fechamento)
🔶 Amarelo: Em andamento
Filtro por Período e Relatório PDF
Selecione um período nas caixas de data e clique em “Gerar Relatório PDF” para criar um relatório com:

Gráfico de pizza com estatísticas
Listagens agrupadas por tipo e origem
Informações resumidas
Ordenar Tabela
Use os botões “Ordenar A-Z (Tipo)” ou “Ordenar Z-A (Tipo)” para organizar a tabela.

📈 Análise de Dados
O relatório mostra métricas como:

Total de chamados
Chamados concluídos
Chamados críticos
Chamados em andamento
Inclui também gráficos de pizza e listas agrupadas por:

Tipo de equipamento
Origem do chamado
