# 🚀 Dashboard de Gestão de Campanhas - Copa

Este projeto é uma aplicação Web profissional desenvolvida para a gestão de campanhas de incentivo, auditoria de atendimentos e monitorização de desempenho de equipas. O sistema foi desenhado para ser integrado com o **Google Apps Script**, utilizando o **Google Sheets** como banco de dados em tempo real.

## 🛠 Tecnologias Utilizadas
- **HTML5**: Estrutura semântica e organizada.
- **CSS3**: Layout responsivo, design moderno, temas customizáveis (Light/Dark Mode) e variáveis CSS para manutenção facilitada.
- **JavaScript (Vanilla)**: Lógica de front-end para manipulação de tabelas, validação de formulários, navegação entre abas (SPA) e gestão de estados de sistema.

## 🎯 Funcionalidades Principais

### 📊 Placar Geral (Dashboard)
- Visualização de Pódio dinâmico (1º, 2º e 3º lugares).
- Cartões informativos de alta visibilidade para o time líder.
- Tabelas de pontuação de times e colaboradores.

### 🎧 Auditoria de Atendimento (Tab 2)
- Formulário inteligente com validação de dados em tempo real.
- Máscaras para inputs (CPF, Datas) e campos obrigatórios para garantir a integridade dos dados antes da auditoria.

### 🎯 Registro de Pontos (Tab 3)
- Motor de busca e filtragem em tempo real (filtros por nome ou time).
- Interface de botões de ação rápida para pontuação (+25, +10, -5).

### ⚙️ Configurações e Administração (Tab 4)
- **Gestão de Entidades**: Cadastros de equipes e colaboradores via modais modernos.
- **Personalização de Sistema**: Opção de alterar o título da campanha e alternar entre **3 temas visuais** (Padrão, Corporativo Azul e Dark Mode).
- **Feedback ao Usuário**: Modais de sucesso e erro para confirmação de ações, garantindo uma experiência de usuário (UX) profissional.

## 📱 Responsividade
O projeto é 100% responsivo, adaptando-se automaticamente a dispositivos móveis e tablets, garantindo que o gestor possa auditar atendimentos ou verificar o placar em qualquer lugar.

## 🚀 Como Executar
1. Clone este repositório.
2. Abra o ficheiro `index.html` no seu navegador.
3. (Opcional) Para integração com Google Sheets:
   - Crie um projeto no Google Apps Script.
   - Utilize a função `doGet()` para servir este ficheiro HTML.
   - Conecte o `google.script.run` às funções `doPost()` para persistência dos dados na sua folha de cálculo.

---
*Desenvolvido como parte do meu portfólio profissional de automação e desenvolvimento web.*