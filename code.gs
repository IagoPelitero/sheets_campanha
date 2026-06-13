/**
 * Sistema de Campanhas Integrado - Backend Google Apps Script
 */

const SPREADSHEET_ID = "sheets_id";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Sistema de Campanhas - Copa')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obterPlanilha() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function obterAbaSegura(ss, nomeDesejado) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toLowerCase() === nomeDesejado.toLowerCase()) {
      return sheets[i];
    }
  }
  return null;
}

function buscarDadosIniciais() {
  const cache = CacheService.getScriptCache();
  const CACHE_KEY = 'dados_iniciais_cache';
  
  // Tenta buscar os dados do cache primeiro
  const cachedData = cache.get(CACHE_KEY);
  if (cachedData) {
    Logger.log("Dados carregados do cache.");
    return JSON.parse(cachedData);
  }

  Logger.log("Cache vazio. Buscando dados da planilha.");
  verificarESetupPlanilhas();

  const ss = obterPlanilha();
  const userEmail = Session.getActiveUser().getEmail() || "Acesso Interno";

  const sheetEquipes = obterAbaSegura(ss, "equipes");
  const sheetColaboradores = obterAbaSegura(ss, "colaboradores");
  const sheetAtendimento = obterAbaSegura(ss, "atendimento");

  const equipesRaw = sheetEquipes ? sheetEquipes.getDataRange().getValues() : [];
  const colaboradoresRaw = sheetColaboradores ? sheetColaboradores.getDataRange().getValues() : [];
  const atendimentosRaw = sheetAtendimento ? sheetAtendimento.getDataRange().getValues() : [];

  // 1. Processar Equipes
  const equipes = [];
  if (equipesRaw.length > 1) {
    for (let e = 1; e < equipesRaw.length; e++) {
      if (equipesRaw[e][0]) equipes.push(equipesRaw[e][0].toString().trim());
    }
  }

  // 2. Processar Colaboradores
  const colaboradores = [];
  if (colaboradoresRaw.length > 1) {
    for (let i = 1; i < colaboradoresRaw.length; i++) {
      if (colaboradoresRaw[i][0]) {
        const nomeC = colaboradoresRaw[i][0].toString().trim();
        const equipeC = colaboradoresRaw[i][2] ? colaboradoresRaw[i][2].toString().trim() : "";
        colaboradores.push({
          nome: nomeC,
          cargo: colaboradoresRaw[i][1] ? colaboradoresRaw[i][1].toString().trim() : "",
          equipe: equipeC
        });
      }
    }
  }

  // 3. Processar Atendimentos para exibição (📋 Atendimentos Auditados)
  const atendimentos = [];
  if (atendimentosRaw.length > 1) {
    const limite = Math.max(1, atendimentosRaw.length - 50);
    for (let j = atendimentosRaw.length - 1; j >= limite; j--) {
      if (atendimentosRaw[j][0]) {
        const rawDate = atendimentosRaw[j][3];
        let fDate = rawDate;
        if (rawDate instanceof Date) {
          fDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }

        atendimentos.push({
          linha: j + 1,
          nome: atendimentosRaw[j][0] ? atendimentosRaw[j][0].toString() : "",
          cargo: atendimentosRaw[j][1] ? atendimentosRaw[j][1].toString() : "",
          cpf: atendimentosRaw[j][2] ? atendimentosRaw[j][2].toString() : "",
          dataContato: fDate,
          apolice: atendimentosRaw[j][4] ? atendimentosRaw[j][4].toString() : "-",
          premio: atendimentosRaw[j][5] ? atendimentosRaw[j][5].toString() : "0,00",
          tipoContato: atendimentosRaw[j][6] ? atendimentosRaw[j][6].toString() : "-",
          retido: atendimentosRaw[j][7] ? atendimentosRaw[j][7].toString() : "",
          pontos: parseFloat(atendimentosRaw[j][8]) || 0
        });
      }
    }
  }

  // 4. Processar Rankings (Lógica unificada)
  const { scoreIndividualMap, scoreTeamMap } = _calcularPontuacoes({});

  const rankingIndividual = Object.keys(scoreIndividualMap).map(function(nome) {
    return { nome: nome, equipe: scoreIndividualMap[nome].equipe, pontos: scoreIndividualMap[nome].pontos };
  }).sort(function(a, b) { return b.pontos - a.pontos; });

  const rankingTimes = Object.keys(scoreTeamMap).map(function(nome) {
    return { nome: nome, pontos: scoreTeamMap[nome] };
  }).sort(function(a, b) { return b.pontos - a.pontos; });

  const dadosParaRetorno = {
    usuario: userEmail,
    equipes: equipes,
    colaboradores: colaboradores,
    atendimentos: atendimentos,
    rankingIndividual: rankingIndividual,
    rankingTimes: rankingTimes
  };

  // Armazena os dados recém-buscados no cache por 3 horas (10800 segundos)
  // O tempo máximo é de 6 horas (21600 segundos)
  cache.put(CACHE_KEY, JSON.stringify(dadosParaRetorno), 10800);
  Logger.log("Dados armazenados no cache.");

  return dadosParaRetorno;
}

// ========================
// FUNÇÕES DE ESCRITA E DELEÇÃO
// ========================

function incluirEquipe(nome) {
  const sheet = obterAbaSegura(obterPlanilha(), "equipes");
  if (sheet) { 
    sheet.appendRow([nome]);
    limparCache();
    return true; 
  }
  return false;
}

function incluirColaborador(nome, cargo, equipe) {
  const sheet = obterAbaSegura(obterPlanilha(), "colaboradores");
  if (sheet) { 
    sheet.appendRow([nome, cargo, equipe]);
    limparCache();
    return true; 
  }
  return false;
}

/**
 * Salva ou edita um atendimento incluindo o novo campo de pontuação direta.
 */
function salvarAtendimento(dados, linhaEdit) {
  try {
    const sheet = obterAbaSegura(obterPlanilha(), "atendimento");
    if (sheet) {
      
      // Força o cálculo automatizado direto no servidor baseado nas regras informadas
      const pontosCalculados = calcularPontosAutomatizados(dados.tipoContato, dados.retido, dados.premio);
      
      const rowData = [
        dados.nome, 
        dados.cargo, 
        dados.cpf, 
        dados.dataContato,
        dados.apolice, 
        dados.premio, 
        dados.tipoContato, 
        dados.retido, 
        pontosCalculados, // Grava a pontuação exata gerada pela inteligência do script
        new Date()
      ];
      
      if (linhaEdit) {
        sheet.getRange(linhaEdit, 1, 1, rowData.length).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
      }
      limparCache();
      return true;
    }
    return false;
  } catch(e) { 
    Logger.log("Erro ao salvar atendimento: " + e.toString());
    return false; 
  }
}

function lancarPontuacao(nome, equipe, pontos, motivo) {
  const sheet = obterAbaSegura(obterPlanilha(), "pontuacao");
  if (sheet) { 
    sheet.appendRow([nome, equipe, pontos, motivo, new Date()]);
    limparCache();
    return true; 
  }
  return false;
}

function excluirEquipeSheets(nomeEquipe) {
  const ss = obterPlanilha();
  const sheetEquipes = obterAbaSegura(ss, "equipes");
  let sucesso = false;
  const equipeRef = nomeEquipe.toString().trim().toLowerCase();
  
  if (sheetEquipes) {
    const data = sheetEquipes.getDataRange().getValues();
    // Exclui de baixo para cima para não quebrar os índices
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][0] && data[i][0].toString().trim().toLowerCase() === equipeRef) { 
        sheetEquipes.deleteRow(i + 1); 
        sucesso = true;
      }
    }
  }
  
  if (sucesso) {
    const sheetColaboradores = obterAbaSegura(ss, "colaboradores");
    if (sheetColaboradores) {
      const dataColab = sheetColaboradores.getDataRange().getValues();
      for (let j = dataColab.length - 1; j >= 0; j--) {
        if (dataColab[j][2] && dataColab[j][2].toString().trim().toLowerCase() === equipeRef) {
          sheetColaboradores.deleteRow(j + 1);
        }
      }
    }
  }
  
  if (sucesso) limparCache();
  return sucesso;
}

function excluirColaboradorSheets(nomeColab) {
  const sheet = obterAbaSegura(obterPlanilha(), "colaboradores");
  if (!sheet) return false;
  
  const data = sheet.getDataRange().getValues();
  const nomeRef = nomeColab.toString().trim().toLowerCase();
  
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === nomeRef) { 
      sheet.deleteRow(i + 1); 
      limparCache();
      return true; 
    }
  }
  return false;
}

function excluirAtendimentoSheets(linha) {
  try {
    const sheet = obterAbaSegura(obterPlanilha(), "atendimento");
    if (sheet) {
      sheet.deleteRow(linha);
      limparCache();
      return true;
    }
    return false;
  } catch(e) { return false; }
}

/**
 * Alinha e garante a estrutura correta de colunas de todas as abas.
 * Nova coluna inserida: "Pontuação Atendimento" na aba de atendimentos.
 */
function verificarESetupPlanilhas() {
  const ss = obterPlanilha();
  const estruturas = {
    "atendimento": ["Nome Colaborador", "Cargo", "CPF", "Data do Contato", "Nº da Apólice", "Valor do Prêmio", "Tipo de Contato", "Foi Retido?", "Pontuação Atendimento", "Data Registro"],
    "pontuacao": ["Nome Colaborador", "Equipe", "Pontos Atribuídos", "Motivo", "Data do Lançamento"],
    "colaboradores": ["Nome Completo", "Cargo", "Equipe Associada"],
    "equipes": ["Nome da Equipe"]
  };
  
  Object.keys(estruturas).forEach(function(nomeAba) {
    let sheet = obterAbaSegura(ss, nomeAba);
    if (!sheet) {
      sheet = ss.insertSheet(nomeAba);
      sheet.appendRow(estruturas[nomeAba]);
      sheet.getRange(1, 1, 1, estruturas[nomeAba].length).setFontWeight("bold").setBackground("#002776").setFontColor("#ffffff");
    }
  });
}

/**
 * =========================================================================
 * FUNÇÕES DE EDIÇÃO E ATUALIZAÇÃO (CONFIGURAÇÕES)
 * =========================================================================
 */

function atualizarEquipeNoSheets(nomeAntigo, novoNome) {
  try {
    const ss = obterPlanilha();
    const nomeAntigoRef = nomeAntigo.toString().trim().toLowerCase();
    
    const sheetEquipes = obterAbaSegura(ss, "equipes"); 
    if (sheetEquipes) {
      const dadosEq = sheetEquipes.getDataRange().getValues();
      for (let i = 0; i < dadosEq.length; i++) {
        if (dadosEq[i][0] && dadosEq[i][0].toString().trim().toLowerCase() === nomeAntigoRef) {
          sheetEquipes.getRange(i + 1, 1).setValue(novoNome.trim());
          break;
        }
      }
    }
    
    const sheetColaboradores = obterAbaSegura(ss, "colaboradores");
    if (sheetColaboradores) {
      const dadosColab = sheetColaboradores.getDataRange().getValues();
      for (let j = 0; j < dadosColab.length; j++) {
        if (dadosColab[j][2] && dadosColab[j][2].toString().trim().toLowerCase() === nomeAntigoRef) {
          sheetColaboradores.getRange(j + 1, 3).setValue(novoNome.trim());
        }
      }
    }
    
    limparCache();
    return { sucesso: true, mensagem: "Equipe e vínculos atualizados com sucesso!" };
  } catch (erro) {
    Logger.log("Erro ao atualizar equipe: " + erro.toString());
    throw new Error("Erro no servidor: " + erro.message);
  }
}

function atualizarColaboradorNoSheets(nomeAntigo, nomeNovo, cargoNovo, equipeNova) {
  try {
    const ss = obterPlanilha();
    const sheetColaboradores = obterAbaSegura(ss, "colaboradores");
    
    if (!sheetColaboradores) {
      throw new Error("Aba 'Colaboradores' não foi encontrada na planilha.");
    }
    
    const dadosColab = sheetColaboradores.getDataRange().getValues();
    let localizado = false;
    const nomeAntigoRef = nomeAntigo.toString().trim().toLowerCase();
    
    for (let i = 0; i < dadosColab.length; i++) {
      if (dadosColab[i][0] && dadosColab[i][0].toString().trim().toLowerCase() === nomeAntigoRef) {
        sheetColaboradores.getRange(i + 1, 1).setValue(nomeNovo.trim());
        sheetColaboradores.getRange(i + 1, 2).setValue(cargoNovo.trim());
        sheetColaboradores.getRange(i + 1, 3).setValue(equipeNova.trim());
        localizado = true;
        break;
      }
    }
    
    if (!localizado) {
      throw new Error("O colaborador '" + nomeAntigo + "' não foi encontrado para edição.");
    }
    
    limparCache();
    return { sucesso: true, mensagem: "Colaborador atualizado com sucesso!" };
  } catch (erro) {
    Logger.log("Erro ao atualizar colaborador: " + erro.toString());
    throw new Error("Erro no servidor: " + erro.message);
  }
}

/**
 * Realiza o filtro dinâmico por datas integrando a pontuação unificada de Atendimento + Manual
 * Retorna apenas o agrupamento dos 3 colaboradores de maior pontuação.
 */
function obterDadosFiltradosAba1(dataInicioStr, dataFimStr) {
  try {
    const dataInicio = new Date(dataInicioStr + "T00:00:00");
    const dataFim = new Date(dataFimStr + "T23:59:59");

    // Utiliza a função centralizada para calcular as pontuações com filtro de data
    const { scoreIndividualMap } = _calcularPontuacoes({ dataInicio, dataFim });

    // Converte o mapa de pontuação em um array para ordenação
    const arrayAtendentes = Object.keys(scoreIndividualMap).map(nome => ({
      nome: nome,
      equipe: scoreIndividualMap[nome].equipe || "Sem Equipe",
      pontos: scoreIndividualMap[nome].pontos
    }));

    // Filtra atendentes com pontuação zero e ordena o ranking
    arrayAtendentes.sort((a, b) => b.pontos - a.pontos);
    const rankingFiltrado = arrayAtendentes.filter(a => a.pontos !== 0);

    return {
      sucesso: true,
      topAtendentes: rankingFiltrado.slice(0, 3) // Retorna exclusivamente os 3 primeiros
    };

  } catch (erro) {
    Logger.log("Erro no processamento do ranking simplificado: " + erro.toString());
    return { erro: erro.toString() };
  }
}

/**
 * Função centralizada para calcular pontuações de indivíduos e equipes.
 * Pode filtrar por um intervalo de datas.
 * @param {object} opcoes Opções de filtro.
 * @param {Date} [opcoes.dataInicio] Data de início para o filtro.
 * @param {Date} [opcoes.dataFim] Data de fim para o filtro.
 * @returns {{scoreIndividualMap: object, scoreTeamMap: object}} Mapas de pontuação.
 * @private
 */
function _calcularPontuacoes(opcoes) {
  opcoes = opcoes || {};
  const ss = obterPlanilha();
  const sheetAtendimentos = obterAbaSegura(ss, "atendimento");
  const sheetPontuacao = obterAbaSegura(ss, "pontuacao");
  const sheetColaboradores = obterAbaSegura(ss, "colaboradores");

  const atendimentosRaw = sheetAtendimentos ? sheetAtendimentos.getDataRange().getValues() : [];
  const pontuacoesRaw = sheetPontuacao ? sheetPontuacao.getDataRange().getValues() : [];
  const colaboradoresRaw = sheetColaboradores ? sheetColaboradores.getDataRange().getValues() : [];

  // 1. Mapear colaboradores e equipes
  const colabEquipeMap = {};
  const scoreIndividualMap = {};
  if (colaboradoresRaw.length > 1) {
    for (let i = 1; i < colaboradoresRaw.length; i++) {
      if (colaboradoresRaw[i][0]) {
        const nomeC = colaboradoresRaw[i][0].toString().trim();
        const equipeC = colaboradoresRaw[i][2] ? colaboradoresRaw[i][2].toString().trim() : "Sem Equipe";
        colabEquipeMap[nomeC] = equipeC;
        scoreIndividualMap[nomeC] = { pontos: 0, equipe: equipeC };
      }
    }
  }

  // 2. Mapear equipes
  const scoreTeamMap = {};
  const sheetEquipes = obterAbaSegura(ss, "equipes");
  if (sheetEquipes) {
    const equipesRaw = sheetEquipes.getDataRange().getValues();
    if (equipesRaw.length > 1) {
      for (let e = 1; e < equipesRaw.length; e++) {
        if (equipesRaw[e][0]) {
          scoreTeamMap[equipesRaw[e][0].toString().trim()] = 0;
        }
      }
    }
  }

  // 3. Somar pontos da aba PONTUACAO (lançamentos manuais)
  if (pontuacoesRaw.length > 1) {
    for (let k = 1; k < pontuacoesRaw.length; k++) {
      const linha = pontuacoesRaw[k];
      const dataLancamento = linha[4] ? new Date(linha[4]) : null;

      if (opcoes.dataInicio && (!dataLancamento || dataLancamento < opcoes.dataInicio)) continue;
      if (opcoes.dataFim && (!dataLancamento || dataLancamento > opcoes.dataFim)) continue;

      const pNome = linha[0] ? linha[0].toString().trim() : null;
      if (!pNome) continue;

      const pValor = parseFloat(linha[2]) || 0;
      const pEquipe = linha[1] ? linha[1].toString().trim() : colabEquipeMap[pNome] || "Sem Equipe";

      if (scoreIndividualMap[pNome]) {
        scoreIndividualMap[pNome].pontos += pValor;
      } else {
        scoreIndividualMap[pNome] = { pontos: pValor, equipe: pEquipe };
      }

      if (scoreTeamMap.hasOwnProperty(pEquipe)) {
        scoreTeamMap[pEquipe] += pValor;
      }
    }
  }

  // 4. Somar pontos da aba ATENDIMENTO
  if (atendimentosRaw.length > 1) {
    for (let m = 1; m < atendimentosRaw.length; m++) {
      const linha = atendimentosRaw[m];
      const dataContato = linha[3] ? new Date(linha[3]) : null;

      if (opcoes.dataInicio && (!dataContato || dataContato < opcoes.dataInicio)) continue;
      if (opcoes.dataFim && (!dataContato || dataContato > opcoes.dataFim)) continue;

      const aNome = linha[0] ? linha[0].toString().trim() : null;
      if (!aNome) continue;

      const aPts = parseFloat(linha[8]) || 0;
      const aEquipe = colabEquipeMap[aNome] || "Sem Equipe";

      if (scoreIndividualMap[aNome]) {
        scoreIndividualMap[aNome].pontos += aPts;
      } else {
        scoreIndividualMap[aNome] = { pontos: aPts, equipe: aEquipe };
      }

      if (scoreTeamMap.hasOwnProperty(aEquipe)) {
        scoreTeamMap[aEquipe] += aPts;
      }
    }
  }
  
  // Limpeza de possíveis nomes de cabeçalho que possam ter sido lidos
  delete scoreIndividualMap["Nome Completo"];
  delete scoreIndividualMap["Nome Colaborador"];
  delete scoreTeamMap["Nome da Equipe"];
  delete scoreTeamMap["Equipe"];

  return { scoreIndividualMap, scoreTeamMap };
}

// LÓGICA DE SOMA AUTOMÁTICA REQUISITADA
function calcularPontosAutomatizados(tipoContato, retido, premioStr) {
  let totalPontos = 0;
  
  // Regra 1: Potencial (+10) ou Não Potencial (-5)
  if (tipoContato === "Potencial") {
    totalPontos += 10;
  } else if (tipoContato === "Não potencial") {
    totalPontos -= 5;
  }
  
  // Regra 2: Retido (+25) ou Não Retido (0)
  if (retido === "Sim") {
    totalPontos += 25;
  }
  
  // Regra 3: Prêmio acima de R$3.000 (+10)
  if (premioStr) {
    // Remove "R$", pontos de milhar e converte a vírgula decimal para ponto
    const limpo = premioStr.toString().replace("R$", "").replace(/\./g, "").replace(",", ".").trim();
    const valorPremio = parseFloat(limpo) || 0;
    if (valorPremio > 3000 && retido === "Sim" && tipoContato === "Potencial") {
      totalPontos += 10;
    }
  }
  
  return totalPontos;
}

/**
 * Remove os dados do cache. Deve ser chamada sempre que houver uma escrita
 * na planilha para garantir que os dados sejam recarregados.
 */
function limparCache() {
  CacheService.getScriptCache().remove('dados_iniciais_cache');
  Logger.log("Cache limpo.");
}