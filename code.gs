/**
 * Sistema de Campanhas Integrado - Backend Google Apps Script
 */

const SPREADSHEET_ID = "1kS-cCzgiUD5XS0WYEkYnpIQzS8qC7pt5hclCvoN7E70";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Sistema de Campanhas - Copa')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obterPlanilha() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function obterAbaSegura(ss, nomeDesejado) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toLowerCase() === nomeDesejado.toLowerCase()) {
      return sheets[i];
    }
  }
  return null;
}

function buscarDadosIniciais() {
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
  
  // 1. Processar Equipes
  var equipes = [];
  if (equipesRaw.length > 0) {
    var startIdx = (equipesRaw[0][0] && equipesRaw[0][0].toString().toLowerCase().includes("nome")) ? 1 : 0;
    for (var e = startIdx; e < equipesRaw.length; e++) {
      if (equipesRaw[e][0]) equipes.push(equipesRaw[e][0].toString().trim());
    }
  }

  // 2. Processar Colaboradores
  var colaboradores = [];
  if (colaboradoresRaw.length > 0) {
    var startIdx = (colaboradoresRaw[0][0] && colaboradoresRaw[0][0].toString().toLowerCase().includes("nome")) ? 1 : 0;
    for (var i = startIdx; i < colaboradoresRaw.length; i++) {
      if (colaboradoresRaw[i][0]) {
        colaboradores.push({
          nome: colaboradoresRaw[i][0].toString().trim(),
          cargo: colaboradoresRaw[i][1] ? colaboradoresRaw[i][1].toString().trim() : "",
          equipe: colaboradoresRaw[i][2] ? colaboradoresRaw[i][2].toString().trim() : ""
        });
      }
    }
  }
  
  // 3. Processar Atendimentos
  var atendimentos = [];
  if (atendimentosRaw.length > 0) {
    var startIdx = (atendimentosRaw[0][0] && atendimentosRaw[0][0].toString().toLowerCase().includes("nome")) ? 1 : 0;
    var limite = Math.max(startIdx, atendimentosRaw.length - 50);
    for (var j = atendimentosRaw.length - 1; j >= limite; j--) {
      if (atendimentosRaw[j][0]) {
        var rawDate = atendimentosRaw[j][3];
        var fDate = rawDate;
        if (rawDate instanceof Date) {
          fDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }
        
        atendimentos.push({
          linha: j + 1,
          nome: atendimentosRaw[j][0] != null ? atendimentosRaw[j][0].toString() : "",
          cargo: atendimentosRaw[j][1] != null ? atendimentosRaw[j][1].toString() : "",
          cpf: atendimentosRaw[j][2] != null ? atendimentosRaw[j][2].toString() : "",
          dataContato: fDate,
          apolice: atendimentosRaw[j][4] != null && atendimentosRaw[j][4] !== "" ? atendimentosRaw[j][4].toString() : "-",
          premio: atendimentosRaw[j][5] != null && atendimentosRaw[j][5] !== "" ? atendimentosRaw[j][5].toString() : "0,00",
          tipoContato: atendimentosRaw[j][6] != null && atendimentosRaw[j][6] !== "" ? atendimentosRaw[j][6].toString() : "-",
          retido: atendimentosRaw[j][7] != null ? atendimentosRaw[j][7].toString() : ""
        });
      }
    }
  }

  // 4. Processar Rankings
  var scoreIndividualMap = {};
  var scoreTeamMap = {};
  
  colaboradores.forEach(function(c) {
    scoreIndividualMap[c.nome] = { pontos: 0, equipe: c.equipe };
  });
  equipes.forEach(function(eq) {
    scoreTeamMap[eq] = 0;
  });

  if (pontuacoesRaw.length > 0) {
    var startIdx = (pontuacoesRaw[0][0] && pontuacoesRaw[0][0].toString().toLowerCase().includes("nome")) ? 1 : 0;
    for (var k = startIdx; k < pontuacoesRaw.length; k++) {
      var pNome = pontuacoesRaw[k][0] ? pontuacoesRaw[k][0].toString().trim() : null;
      var pEquipe = pontuacoesRaw[k][1] ? pontuacoesRaw[k][1].toString().trim() : null;
      var pValor = parseFloat(pontuacoesRaw[k][2]) || 0;
      
      if (!pNome) continue;

      if (scoreIndividualMap[pNome]) {
        scoreIndividualMap[pNome].pontos += pValor;
      } else {
        scoreIndividualMap[pNome] = { pontos: pValor, equipe: pEquipe };
      }
      
      if (pEquipe) {
        if (scoreTeamMap[pEquipe] !== undefined) {
          scoreTeamMap[pEquipe] += pValor;
        } else {
          scoreTeamMap[pEquipe] = pValor;
        }
      }
    }
  }

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

// ========================
// FUNÇÕES DE ESCRITA E DELEÇÃO
// ========================

function incluirEquipe(nome) {
  var sheet = obterAbaSegura(obterPlanilha(), "equipes");
  if (sheet) { sheet.appendRow([nome]); return true; }
  return false;
}

function incluirColaborador(nome, cargo, equipe) {
  var sheet = obterAbaSegura(obterPlanilha(), "colaboradores");
  if (sheet) { sheet.appendRow([nome, cargo, equipe]); return true; }
  return false;
}

function salvarAtendimento(dados, linhaEdit) {
  try {
    var sheet = obterAbaSegura(obterPlanilha(), "atendimento");
    if (sheet) {
      var rowData = [
        dados.nome, dados.cargo, dados.cpf, dados.dataContato,
        dados.apolice, dados.premio, dados.tipoContato, dados.retido, new Date()
      ];
      
      if (linhaEdit) {
        sheet.getRange(linhaEdit, 1, 1, rowData.length).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
      }
      return true;
    }
    return false;
  } catch(e) { return false; }
}

function lancarPontuacao(nome, equipe, pontos, motivo) {
  var sheet = obterAbaSegura(obterPlanilha(), "pontuacao");
  if (sheet) { 
    sheet.appendRow([nome, equipe, pontos, motivo, new Date()]); 
    return true; 
  }
  return false;
}

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

function verificarESetupPlanilhas() {
  var ss = obterPlanilha();
  var estruturas = {
    "atendimento": ["Nome Colaborador", "Cargo", "CPF", "Data do Contato", "Nº da Apólice", "Valor do Prêmio", "Tipo de Contato", "Foi Retido?", "Data Registro"],
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

/**
 * =========================================================================
 * FUNÇÕES DE EDIÇÃO E ATUALIZAÇÃO (CONFIGURAÇÕES)
 * =========================================================================
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