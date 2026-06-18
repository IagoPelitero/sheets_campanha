# 🚀 Dashboard de Gestão de Campanhas - Copa

Este projeto é uma aplicação Web profissional do tipo Single Page Application (SPA), desenvolvida para a gestão de campanhas de incentivo, auditoria de atendimentos e monitoramento de desempenho de equipes. O sistema foi desenhado para ser integrado de forma nativa com o **Google Apps Script (GAS)**, utilizando o **Google Sheets** como banco de dados em tempo real.

## ✨ Destaques e Funcionalidades Principais

O sistema é dividido em 4 painéis (abas) principais de acesso rápido:

### 📊 1. Placar Geral (Dashboard)
- **Pódio Dinâmico:** Visualização automática de 1º, 2º e 3º lugares com base na pontuação total acumulada.
- **Cartão do Time Líder:** Informações de destaque para a equipe que lidera o ranking.
- **Filtro Avançado por Período:** Modal com integração nativa ao back-end para filtrar o ranking entre datas específicas e recalcular top performances.
- **Painéis e Tabelas de Pontuação:** Tabelas atualizadas em tempo real contendo o ranking de equipes e individual.
- **Auditoria de Atendimentos:** Tabela listando todos os atendimentos auditados, com filtros combinados (texto, equipe e status de retenção) e cálculo dinâmico de pontuação e valores retidos no rodapé.

### 🎧 2. Dados de Atendimento
- **Formulário de Entrada:** Formulário validado para lançar novos atendimentos com cálculo automatizado de pontos no front-end e no back-end.
- **Regras de Negócio Inseridas:** 
  - Contatos Potenciais (+10 pontos) / Não Potenciais (-5 pontos).
  - Retenção de cliente (+25 pontos).
  - Bônus para prêmios retidos acima de R$3.000,00 (+10 pontos).
- **Blindagem e Máscaras:** Inputs de CPF, Data e Valores Monetários protegidos contra caracteres inválidos. Bloqueio automático de duplo clique (Debounce).

### 🎯 3. Registro de Pontos (Manuais)
- **Lançamento Direto:** Possibilidade de buscar rapidamente um colaborador, aplicar penalidades ou bônus operacionais com inserção de justificativas.
- **Filtros e Visualização:** Motor de busca e filtros combinados com cálculo dinâmico dos pontos filtrados na tela.

### ⚙️ 4. Configurações e Administração
- **Gestão de Equipes e Colaboradores:** Formulários em modais modernos para criar, editar ou inativar colaboradores e equipes. As alterações propagam-se para todas as métricas históricas no painel.
- **Customização de Interface (UI):** Permite renomear o título principal da campanha e alternar dinamicamente entre **3 temas visuais** (Padrão/Brasil, Corporativo Azul e Modo Escuro).

## 🛠 Tecnologias e Arquitetura
- **Front-end:** HTML5 semântico, CSS3 (variáveis, flexbox, grid, responsividade) e JavaScript (Vanilla) focado em eventos, manipulação de DOM e debounce protection.
- **Back-end e Database:** Hospedagem direta no **Google Apps Script** utilizando `google.script.run` de forma assíncrona. O **Google Sheets** atua como banco de dados NoSQL por meio de manipulação otimizada de vetores bidimensionais.

## 📱 Responsividade
O layout é 100% responsivo, adaptando-se do Desktop a dispositivos móveis e tablets de forma inteligente e ergonômica, garantindo fácil acesso aos gestores de campanhas.

## 🚀 Como Executar e Configurar

Para rodar este projeto e conectá-lo ao seu banco de dados:

1. **Crie a Base de Dados no Sheets:**
   - Crie uma nova planilha no Google Sheets.
   - Guarde o ID da Planilha (presente na URL entre `/d/` e `/edit`).

2. **Configuração do Google Apps Script:**
   - Na sua planilha, vá em `Extensões > Apps Script`.
   - Substitua o código do arquivo `Código.gs` gerado automaticamente pelo conteúdo do arquivo `code.gs` deste repositório.
   - *Importante:* Altere a constante `SPREADSHEET_ID` no `code.gs` com o ID da planilha que você criou no passo 1.
   - Crie um arquivo HTML chamado `Index.html` e copie todo o conteúdo do `index.html` para ele.

3. **Deploy (Publicação):**
   - Clique em "Implantar" > "Nova implantação".
   - Selecione o tipo "Aplicativo da Web".
   - Defina "Executar como: Eu" e "Quem pode acessar: Qualquer pessoa" (ou restrinja aos usuários do seu domínio do Google Workspace).
   - A inicialização estrutural das abas (Atendimento, Pontuação, Equipes, Colaboradores) no Google Sheets será gerada automaticamente pelo sistema no primeiro carregamento.
   
## 🔮 Melhorias Futuras (Preparadas no Código)
A base de código já contém as estruturas e comentários mapeados para ativação simples das seguintes melhorias futuras:
- **Analistas e Tipos de Transferência:** Colunas de controle sobre quem registrou o atendimento (Analista) e se a transferência para o time foi Devida ou Indevida. *(Verifique os comentários HTML e JS para fácil ativação).*

---
*Desenvolvido como parte do meu portfólio profissional de automação e desenvolvimento web.*