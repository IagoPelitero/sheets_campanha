/**
 * =========================================================================
 * Sistema de Campanhas Integrado — Backend Google Apps Script
 * =========================================================================
 *
 * Responsabilidades:
 *   - Servir a interface web (doGet)
 *   - Ler e processar dados das abas do Sheets (buscarDadosIniciais)
 *   - Salvar, editar e excluir registros (salvarAtendimento, excluir...)
 *   - Calcular pontuações automaticamente (calcularPontosAutomatizados)
 *   - Garantir a estrutura correta das abas (verificarESetupPlanilhas)
 *
 * Estrutura de colunas da aba "atendimento":
 *   [0]  Nome Colaborador      [1]  Cargo
 *   [2]  CPF                   [3]  Data do Contato
 *   [4]  Nº da Apólice         [5]  Valor do Prêmio
 *   [6]  Tipo de Contato       [7]  Foi Retido?
 *   [8]  Nome do Analista      [9]  Transferência
 *   [10] Pontuação Atendimento [11] Data Registro
 * =========================================================================
 */

/** ID fixo da planilha. Altere aqui se mover para outra planilha. */
var SPREADSHEET_ID = "1kS-cCzgiUD5XS0WYEkYnpIQzS8qC7pt5hclCvoN7E70";

/**
 * Ponto de entrada da aplicação web.
 * Serve o arquivo HTML 'Index' como interface para o usuário.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Sistema de Campanhas - Copa')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Abre a planilha pelo ID fixo definido em SPREADSHEET_ID.
 * Usar openById() é mais seguro que getActiveSpreadsheet() em web apps,
 * pois garante que o script sempre encontre a planilha correta.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function obterPlanilha() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * Busca uma aba (sheet) pelo nome de forma insensível a maiúsculas/minúsculas.
 * Útil para evitar erros caso o usuário renomeie a aba com capitalização diferente.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - A planilha.
 * @param {string} nomeDesejado - Nome da aba a localizar.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} A aba encontrada ou null.
 */
function obterAbaSegura(ss, nomeDesejado) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toLowerCase() === nomeDesejado.toLowerCase()) {
      return sheets[i];
    }
  }
  return null; // Aba não encontrada
}

/**
 * Função principal chamada pelo front-end ao inicializar a página.
 * Lê todas as abas, processa os dados e retorna um objeto consolidado
 * com equipes, colaboradores, atendimentos e rankings calculados.
 * @returns {Object} { usuario, equipes, colaboradores, atendimentos, rankingIndividual, rankingTimes }
 */
function buscarDadosIniciais() {
  // Garante que todas as abas necessárias existam antes de tentar ler
  verificarESetupPlanilhas();
  
  var ss = obterPlanilha();
  var userEmail = Session.getActiveUser().getEmail() || "Acesso Interno";
  
  var sheetEquipes = obterAbaSegura(ss, "equipes");
  var sheetColaboradores = obterAbaSegura(ss, "colaboradores");
  var sheetAtendimento = obterAbaSegura(ss, "atendimento");
  var sheetPontuacao = obterAbaSegura(ss, "pontuacao");

  var equipesRaw = sheetEquipes ? sheetEquipes.getDataRange().getValues() : [];
  var colaboradoresRaw = sheetColaboradores ? sheetColaboradores.getDataRange().getValues() : [];
  var atendimentosRaw = sheetAtendimento ? sheetAtendimento.getDataRange().getValues() : [];
  var pontuacoesRaw = sheetPontuacao ? sheetPontuacao.getDataRange().getValues() : [];
  
  // --- 1. Processar lista de Equipes ---
  var equipes = [];
  if (equipesRaw.length > 1) {
    for (var e = 1; e < equipesRaw.length; e++) {
      if (equipesRaw[e][0]) equipes.push(equipesRaw[e][0].toString().trim());
    }
  }

  // --- 2. Processar Colaboradores e criar mapa nome → equipe ---
  var colaboradores = [];
  var colabEquipeMap = {};
  if (colaboradoresRaw.length > 1) {
    for (var i = 1; i < colaboradoresRaw.length; i++) {
      if (colaboradoresRaw[i][0]) {
        var nomeC = colaboradoresRaw[i][0].toString().trim();
        var equipeC = colaboradoresRaw[i][2] ? colaboradoresRaw[i][2].toString().trim() : "";
        colaboradores.push({
          nome: nomeC,
          cargo: colaboradoresRaw[i][1] ? colaboradoresRaw[i][1].toString().trim() : "",
          equipe: equipeC
        });
        colabEquipeMap[nomeC] = equipeC;
      }
    }
  }
  
  // --- 3. Processar Atendimentos para exibição na tabela (últimos 50 registros) ---
  var atendimentos = [];
  if (atendimentosRaw.length > 1) {
    var limite = Math.max(1, atendimentosRaw.length - 50);
    for (var j = atendimentosRaw.length - 1; j >= limite; j--) {
      if (atendimentosRaw[j][0]) {
        var rawDate = atendimentosRaw[j][3];
        var fDate = rawDate;
        if (rawDate instanceof Date) {
          fDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }
        
        // Adicionado .trim() para garantir compatibilidade exata com o colabEquipeMap
        var nomeAtend = atendimentosRaw[j][0] != null ? atendimentosRaw[j][0].toString().trim() : "";
        
        atendimentos.push({
          linha: j + 1,
          nome: nomeAtend,
          equipe: colabEquipeMap[nomeAtend] || "Sem Equipe",
          cargo: atendimentosRaw[j][1] != null ? atendimentosRaw[j][1].toString() : "",
          cpf: atendimentosRaw[j][2] != null ? atendimentosRaw[j][2].toString() : "",
          dataContato: fDate,
          apolice: atendimentosRaw[j][4] != null && atendimentosRaw[j][4] !== "" ? atendimentosRaw[j][4].toString() : "-",
          premio: atendimentosRaw[j][5] != null && atendimentosRaw[j][5] !== "" ? atendimentosRaw[j][5].toString() : "0,00",
          tipoContato: atendimentosRaw[j][6] != null && atendimentosRaw[j][6] !== "" ? atendimentosRaw[j][6].toString() : "-",
          retido: atendimentosRaw[j][7] != null ? atendimentosRaw[j][7].toString() : "",
          analista: atendimentosRaw[j][8] != null ? atendimentosRaw[j][8].toString() : "",
          transferencia: atendimentosRaw[j][9] != null ? atendimentosRaw[j][9].toString() : "",
          pontos: atendimentosRaw[j][10] != null && atendimentosRaw[j][10] !== "" ? parseFloat(atendimentosRaw[j][10]) : 0
        });
      }
    }
  }

  // --- 4. Calcular Rankings (unifica pontos de atendimento + lançamentos manuais) ---
  var scoreIndividualMap = {}; // { nomePessoa: { pontos, equipe } }
  var scoreTeamMap       = {}; // { nomeEquipe: totalPontos }

  // Inicializa todos com zero para garantir que apareçam no ranking mesmo sem pontos
  colaboradores.forEach(function(c) {
    scoreIndividualMap[c.nome] = { pontos: 0, equipe: c.equipe };
  });
  equipes.forEach(function(eq) {
    scoreTeamMap[eq] = 0;
  });

  // A) Somar pontos vindos diretamente da aba ATENDIMENTO (coluna 10 = Pontuação)
  if (atendimentosRaw.length > 1) {
    for (var m = 1; m < atendimentosRaw.length; m++) {
      var aNome = atendimentosRaw[m][0] ? atendimentosRaw[m][0].toString().trim() : null;
      // CORRIGIDO: índice [10] = Pontuação Atendimento (anteriormente [8] era o Nome do Analista)
      var aPts = parseFloat(atendimentosRaw[m][10]) || 0;
      if (!aNome) continue;
      
      var aEquipe = colabEquipeMap[aNome] || "";
      
      if (scoreIndividualMap[aNome]) {
        scoreIndividualMap[aNome].pontos += aPts;
      } else {
        scoreIndividualMap[aNome] = { pontos: aPts, equipe: aEquipe };
      }
      
      if (aEquipe && scoreTeamMap[aEquipe] !== undefined) {
        scoreTeamMap[aEquipe] += aPts;
      }
    }
  }

  // B) Somar pontos vindos de lançamentos manuais da aba PONTUACAO
  if (pontuacoesRaw.length > 1) {
    for (var k = 1; k < pontuacoesRaw.length; k++) {
      var pNome = pontuacoesRaw[k][0] ? pontuacoesRaw[k][0].toString().trim() : null;
      var pEquipe = pontuacoesRaw[k][1] ? pontuacoesRaw[k][1].toString().trim() : null;
      var pValor = parseFloat(pontuacoesRaw[k][2]) || 0;
      
      if (!pNome) continue;

      if (scoreIndividualMap[pNome]) {
        scoreIndividualMap[pNome].pontos += pValor;
      } else {
        scoreIndividualMap[pNome] = { pontos: pValor, equipe: pEquipe || colabEquipeMap[pNome] || "" };
      }
      
      var eqFinal = pEquipe || colabEquipeMap[pNome];
      if (eqFinal && scoreTeamMap[eqFinal] !== undefined) {
        scoreTeamMap[eqFinal] += pValor;
      }
    }
  }

  // Remove possíveis cabeçalhos que podem ter sido incluídos acidentalmente
  delete scoreIndividualMap["Nome Completo"];
  delete scoreIndividualMap["Nome Colaborador"];
  delete scoreTeamMap["Nome da Equipe"];
  delete scoreTeamMap["Equipe"];

  var rankingIndividual = Object.keys(scoreIndividualMap).map(function(nome) {
    return { nome: nome, equipe: scoreIndividualMap[nome].equipe, pontos: scoreIndividualMap[nome].pontos };
  }).sort(function(a, b) { return b.pontos - a.pontos; });

  var rankingTimes = Object.keys(scoreTeamMap).map(function(nome) {
    return { nome: nome, pontos: scoreTeamMap[nome] };
  }).sort(function(a, b) { return b.pontos - a.pontos; });

  return {
    usuario: userEmail,
    equipes: equipes,
    colaboradores: colaboradores,
    atendimentos: atendimentos,
    rankingIndividual: rankingIndividual,
    rankingTimes: rankingTimes
  };
}

// =========================================================================
// FUNÇÕES DE ESCRITA (CREATE / UPDATE / DELETE)
// =========================================================================

/**
 * Adiciona uma nova equipe na aba "equipes".
 * @param {string} nome - Nome da equipe.
 * @returns {boolean}
 */
function incluirEquipe(nome) {
  var sheet = obterAbaSegura(obterPlanilha(), "equipes");
  if (sheet) { sheet.appendRow([nome]); return true; }
  return false;
}

/**
 * Adiciona um novo colaborador na aba "colaboradores".
 * @param {string} nome   - Nome completo.
 * @param {string} cargo  - Cargo (ex: "Emissão VI", "Central Vida").
 * @param {string} equipe - Nome da equipe.
 * @returns {boolean}
 */
function incluirColaborador(nome, cargo, equipe) {
  var sheet = obterAbaSegura(obterPlanilha(), "colaboradores");
  if (sheet) { sheet.appendRow([nome, cargo, equipe]); return true; }
  return false;
}

/**
 * Salva ou edita um atendimento na aba "atendimento".
 *
 * Regra de negócio: a Transferência é derivada automaticamente do Tipo de Contato:
 *   - "Potencial"     → Transferência = "Devida"
 *   - "Não potencial" → Transferência = "Indevida"
 * O campo não é aceito do formulário — é sempre recalculado aqui no servidor.
 *
 * @param {Object}      dados     - Dados enviados pelo front-end.
 * @param {number|null} linhaEdit - Linha da planilha para editar; null para inserir.
 * @returns {boolean}
 */
function salvarAtendimento(dados, linhaEdit) {
  try {
    var sheet = obterAbaSegura(obterPlanilha(), "atendimento");
    if (sheet) {

      // Pontuação calculada no servidor — é a fonte da verdade, não aceita do front-end
      var pontosCalculados = calcularPontosAutomatizados(dados.tipoContato, dados.retido, dados.premio);

      // Transferência derivada do tipo de contato (regra de negócio automática)
      var transferenciaCalculada = (dados.tipoContato === "Potencial") ? "Devida" : "Indevida";

      var rowData = [
        dados.nome,
        dados.cargo,
        dados.cpf,
        dados.dataContato,
        dados.apolice,
        dados.premio,
        dados.tipoContato,
        dados.retido,
        dados.analista || "",
        transferenciaCalculada,  // Calculada automaticamente — não vem do formulário
        pontosCalculados,
        new Date()               // Timestamp de registro
      ];
      
      if (linhaEdit) {
        sheet.getRange(linhaEdit, 1, 1, rowData.length).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
      }
      return true;
    }
    return false;
  } catch(e) { 
    Logger.log("Erro ao salvar atendimento: " + e.toString());
    return false; 
  }
}

/**
 * Registra um lançamento manual de pontos na aba "pontuacao".
 * @param {string} nome   - Colaborador que recebe os pontos.
 * @param {string} equipe - Equipe do colaborador.
 * @param {number} pontos - Quantidade (pode ser negativo para penalidade).
 * @param {string} motivo - Justificativa do lançamento.
 * @returns {boolean}
 */
function lancarPontuacao(nome, equipe, pontos, motivo) {
  var sheet = obterAbaSegura(obterPlanilha(), "pontuacao");
  if (sheet) {
    sheet.appendRow([nome, equipe, pontos, motivo, new Date()]);
    return true;
  }
  return false;
}

/**
 * Exclui uma equipe e todos os colaboradores vinculados a ela.
 * A varredura é feita de baixo para cima para não deslocar os índices de linha.
 * @param {string} nomeEquipe - Nome da equipe a excluir.
 * @returns {boolean} true se encontrou e excluiu.
 */
function excluirEquipeSheets(nomeEquipe) {
  var ss = obterPlanilha();
  var sheetEquipes = obterAbaSegura(ss, "equipes");
  var sucesso = false;
  var equipeRef = nomeEquipe.toString().trim().toLowerCase();
  
  if (sheetEquipes) {
    var data = sheetEquipes.getDataRange().getValues();
    // Exclui de baixo para cima para não quebrar os índices
    for (var i = data.length - 1; i >= 0; i--) {
      if (data[i][0] && data[i][0].toString().trim().toLowerCase() === equipeRef) { 
        sheetEquipes.deleteRow(i + 1); 
        sucesso = true;
      }
    }
  }
  
  if (sucesso) {
    var sheetColaboradores = obterAbaSegura(ss, "colaboradores");
    if (sheetColaboradores) {
      var dataColab = sheetColaboradores.getDataRange().getValues();
      for (var j = dataColab.length - 1; j >= 0; j--) {
        if (dataColab[j][2] && dataColab[j][2].toString().trim().toLowerCase() === equipeRef) {
          sheetColaboradores.deleteRow(j + 1);
        }
      }
    }
  }
  
  return sucesso;
}

/**
 * Exclui um colaborador da aba "colaboradores" pelo nome.
 * @param {string} nomeColab - Nome do colaborador.
 * @returns {boolean}
 */
function excluirColaboradorSheets(nomeColab) {
  var sheet = obterAbaSegura(obterPlanilha(), "colaboradores");
  if (!sheet) return false;
  
  var data = sheet.getDataRange().getValues();
  var nomeRef = nomeColab.toString().trim().toLowerCase();
  
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === nomeRef) { 
      sheet.deleteRow(i + 1); 
      return true; 
    }
  }
  return false;
}

/**
 * Exclui uma linha da aba "atendimento" pelo número da linha (1-based).
 * @param {number} linha - Número da linha na planilha.
 * @returns {boolean}
 */
function excluirAtendimentoSheets(linha) {
  try {
    var sheet = obterAbaSegura(obterPlanilha(), "atendimento");
    if (sheet) {
      sheet.deleteRow(linha);
      return true;
    }
    return false;
  } catch(e) { return false; }
}

/**
 * Verifica se todas as abas necessárias existem e as cria se estiver faltando.
 * Também define o cabeçalho formatado (negrito, fundo azul, texto branco).
 * Deve ser chamada no início de buscarDadosIniciais() como verificação de saúde.
 */
function verificarESetupPlanilhas() {
  var ss = obterPlanilha();
  var estruturas = {
    "atendimento": ["Nome Colaborador", "Cargo", "CPF", "Data do Contato", "Nº da Apólice", "Valor do Prêmio", "Tipo de Contato", "Foi Retido?", "Nome do Analista", "Transferência", "Pontuação Atendimento", "Data Registro"],
    "pontuacao": ["Nome Colaborador", "Equipe", "Pontos Atribuídos", "Motivo", "Data do Lançamento"],
    "colaboradores": ["Nome Completo", "Cargo", "Equipe Associada"],
    "equipes": ["Nome da Equipe"]
  };
  
  Object.keys(estruturas).forEach(function(nomeAba) {
    var sheet = obterAbaSegura(ss, nomeAba);
    if (!sheet) {
      sheet = ss.insertSheet(nomeAba);
      sheet.appendRow(estruturas[nomeAba]);
      sheet.getRange(1, 1, 1, estruturas[nomeAba].length).setFontWeight("bold").setBackground("#002776").setFontColor("#ffffff");
    }
  });
}

// =========================================================================
// FUNÇÕES DE ATUALIZAÇÃO (EDIÇÃO DE EQUIPES E COLABORADORES)
// =========================================================================

/**
 * Renomeia uma equipe e atualiza o vínculo de todos os colaboradores associados.
 * @param {string} nomeAntigo - Nome atual da equipe.
 * @param {string} novoNome   - Novo nome desejado.
 * @returns {{ sucesso: boolean, mensagem: string }}
 */
function atualizarEquipeNoSheets(nomeAntigo, novoNome) {
  try {
    var ss = obterPlanilha();
    var nomeAntigoRef = nomeAntigo.toString().trim().toLowerCase();
    
    var sheetEquipes = obterAbaSegura(ss, "equipes"); 
    if (sheetEquipes) {
      var dadosEq = sheetEquipes.getDataRange().getValues();
      for (var i = 0; i < dadosEq.length; i++) {
        if (dadosEq[i][0] && dadosEq[i][0].toString().trim().toLowerCase() === nomeAntigoRef) {
          sheetEquipes.getRange(i + 1, 1).setValue(novoNome.trim());
          break;
        }
      }
    }
    
    var sheetColaboradores = obterAbaSegura(ss, "colaboradores");
    if (sheetColaboradores) {
      var dadosColab = sheetColaboradores.getDataRange().getValues();
      for (var j = 0; j < dadosColab.length; j++) {
        if (dadosColab[j][2] && dadosColab[j][2].toString().trim().toLowerCase() === nomeAntigoRef) {
          sheetColaboradores.getRange(j + 1, 3).setValue(novoNome.trim());
        }
      }
    }
    
    return { sucesso: true, mensagem: "Equipe e vínculos atualizados com sucesso!" };
  } catch (erro) {
    Logger.log("Erro ao atualizar equipe: " + erro.toString());
    throw new Error("Erro no servidor: " + erro.message);
  }
}

/**
 * Atualiza os dados de um colaborador existente (nome, cargo e equipe).
 * @param {string} nomeAntigo - Nome atual para localizar o registro.
 * @param {string} nomeNovo   - Novo nome.
 * @param {string} cargoNovo  - Novo cargo.
 * @param {string} equipeNova - Nova equipe.
 * @returns {{ sucesso: boolean, mensagem: string }}
 */
function atualizarColaboradorNoSheets(nomeAntigo, nomeNovo, cargoNovo, equipeNova) {
  try {
    var ss = obterPlanilha();
    var sheetColaboradores = obterAbaSegura(ss, "colaboradores");
    
    if (!sheetColaboradores) {
      throw new Error("Aba 'Colaboradores' não foi encontrada na planilha.");
    }
    
    var dadosColab = sheetColaboradores.getDataRange().getValues();
    var localizado = false;
    var nomeAntigoRef = nomeAntigo.toString().trim().toLowerCase();
    
    for (var i = 0; i < dadosColab.length; i++) {
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
    
    return { sucesso: true, mensagem: "Colaborador atualizado com sucesso!" };
  } catch (erro) {
    Logger.log("Erro ao atualizar colaborador: " + erro.toString());
    throw new Error("Erro no servidor: " + erro.message);
  }
}

// =========================================================================
// FILTRO DE RANKING POR PERÍODO
// =========================================================================

/**
 * Retorna o Top 3 de colaboradores por pontuação em um período filtrado.
 * Une pontos da aba "atendimento" com lançamentos manuais da aba "pontuacao".
 * Também conta quantos atendimentos foram retidos por colaborador no período.
 *
 * @param {string} dataInicioStr - Data inicial no formato "YYYY-MM-DD".
 * @param {string} dataFimStr    - Data final no formato "YYYY-MM-DD".
 * @returns {{ sucesso: boolean, topAtendentes: Array }|{ erro: string }}
 */
function obterDadosFiltradosAba1(dataInicioStr, dataFimStr) {
  try {
    var ss                = obterPlanilha();
    var sheetAtendimentos = obterAbaSegura(ss, "atendimento");
    var sheetPontuacao    = obterAbaSegura(ss, "pontuacao");
    var sheetColabs       = obterAbaSegura(ss, "colaboradores");

    if (!sheetAtendimentos || !sheetPontuacao) {
      return { erro: "As abas necessárias não foram encontradas." };
    }

    var dadosAtend  = sheetAtendimentos.getDataRange().getValues();
    var dadosPont   = sheetPontuacao.getDataRange().getValues();
    var dataInicio  = new Date(dataInicioStr + "T00:00:00");
    var dataFim     = new Date(dataFimStr    + "T23:59:59");

    var mapaColabEquipeMaster = {};
    if (sheetColabs) {
      var dadosColabs = sheetColabs.getDataRange().getValues();
      for (var c = 1; c < dadosColabs.length; c++) {
        var nomeC = dadosColabs[c][0] ? dadosColabs[c][0].toString().trim() : "";
        var eqC   = dadosColabs[c][2] ? dadosColabs[c][2].toString().trim() : "";
        if (nomeC) mapaColabEquipeMaster[nomeC] = eqC;
      }
    }

    var mapaPontosAtendentes  = {};
    var mapaEquipeDoAtendente = {};
    var mapaRetidosAtendentes = {};

    // Pontos manuais da aba pontuacao
    if (dadosPont.length > 1) {
      for (var p = 1; p < dadosPont.length; p++) {
        var linhaPont = dadosPont[p];
        if (!linhaPont[4]) continue;
        var dataLancamento = new Date(linhaPont[4]);
        if (dataLancamento >= dataInicio && dataLancamento <= dataFim) {
          var nomeAt = linhaPont[0] ? linhaPont[0].toString().trim() : null;
          var nomeEq = linhaPont[1] ? linhaPont[1].toString().trim() : (mapaColabEquipeMaster[nomeAt] || "Sem Equipe");
          var pts    = Number(linhaPont[2]) || 0;
          if (nomeAt) {
            mapaPontosAtendentes[nomeAt]  = (mapaPontosAtendentes[nomeAt]  || 0) + pts;
            mapaEquipeDoAtendente[nomeAt] = nomeEq;
          }
        }
      }
    }

    // Pontos e retenções da aba atendimento
    if (dadosAtend.length > 1) {
      for (var a = 1; a < dadosAtend.length; a++) {
        var linhaA = dadosAtend[a];
        if (!linhaA[3]) continue;
        var dataLinha = new Date(linhaA[3]);
        if (dataLinha >= dataInicio && dataLinha <= dataFim) {
          var colabName = linhaA[0] ? linhaA[0].toString().trim() : null;
          var ptsAtend  = Number(linhaA[10]) || 0;
          var nomeEqA   = mapaColabEquipeMaster[colabName] || "Sem Equipe";
          var isRetido  = linhaA[7] && linhaA[7].toString().trim().toLowerCase() === "sim";
          if (colabName) {
            mapaPontosAtendentes[colabName]  = (mapaPontosAtendentes[colabName]  || 0) + ptsAtend;
            mapaEquipeDoAtendente[colabName] = nomeEqA;
            if (isRetido) {
              mapaRetidosAtendentes[colabName] = (mapaRetidosAtendentes[colabName] || 0) + 1;
            }
          }
        }
      }
    }

    var arrayAtendentes = [];
    for (var at in mapaPontosAtendentes) {
      if (mapaPontosAtendentes[at] !== 0) {
        arrayAtendentes.push({
          nome:    at,
          equipe:  mapaEquipeDoAtendente[at] || "Sem Equipe",
          pontos:  mapaPontosAtendentes[at],
          retidos: mapaRetidosAtendentes[at] || 0
        });
      }
    }
    arrayAtendentes.sort(function(a, b) { return b.pontos - a.pontos; });

    return {
      sucesso:       true,
      topAtendentes: arrayAtendentes.slice(0, 3)
    };
    
  } catch (erro) {
    Logger.log("Erro no processamento do ranking simplificado: " + erro.toString());
    return { erro: erro.toString() };
  }
}

// =========================================================================
// REGRAS DE NEGÓCIO — CÁLCULO DE PONTUAÇÃO
// =========================================================================

/**
 * Calcula a pontuação de um atendimento com base em 3 regras:
 *
 *   Regra 1 — Tipo de Contato:
 *     - "Potencial"     → +10 pts
 *     - "Não potencial" → -5 pts
 *
 *   Regra 2 — Retenção:
 *     - "Sim" (retido)  → +25 pts
 *
 *   Regra 3 — Bônus por prêmio alto (aplica SOMENTE se Potencial + Retido):
 *     - Prêmio > R$ 3.000 → +10 pts adicionais
 *
 * @param {string} tipoContato - "Potencial" ou "Não potencial".
 * @param {string} retido      - "Sim" ou "Não".
 * @param {string} premioStr   - Valor do prêmio como string (ex: "R$ 3.500,00").
 * @returns {number} Total de pontos calculados.
 */
function calcularPontosAutomatizados(tipoContato, retido, premioStr) {
  var totalPontos = 0;

  // Regra 1: Tipo de Contato
  if (tipoContato === "Potencial") {
    totalPontos += 10;
  } else if (tipoContato === "Não potencial") {
    totalPontos -= 5;
  }

  // Regra 2: Retenção
  if (retido === "Sim") {
    totalPontos += 25;
  }

  // Regra 3: Bônus por prêmio alto (somente quando Potencial E Retido)
  if (premioStr) {
    // Remove "R$", pontos de milhar e converte vírgula decimal para ponto
    var limpo = premioStr.toString().replace("R$", "").replace(/\./g, "").replace(",", ".").trim();
    var valorPremio = parseFloat(limpo) || 0;
    if (valorPremio > 3000 && retido === "Sim" && tipoContato === "Potencial") {
      totalPontos += 10;
    }
  }

  return totalPontos;
}