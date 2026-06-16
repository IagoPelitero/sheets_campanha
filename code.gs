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
  if (equipesRaw.length > 1) {
    for (var e = 1; e < equipesRaw.length; e++) {
      if (equipesRaw[e][0]) equipes.push(equipesRaw[e][0].toString().trim());
    }
  }

  // 2. Processar Colaboradores e criar mapa de equipes
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
  
  // 3. Processar Atendimentos para exibição (📋 Atendimentos Auditados)
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
        
        atendimentos.push({
          linha: j + 1,
          nome: atendimentosRaw[j][0] != null ? atendimentosRaw[j][0].toString() : "",
          cargo: atendimentosRaw[j][1] != null ? atendimentosRaw[j][1].toString() : "",
          cpf: atendimentosRaw[j][2] != null ? atendimentosRaw[j][2].toString() : "",
          dataContato: fDate,
          apolice: atendimentosRaw[j][4] != null && atendimentosRaw[j][4] !== "" ? atendimentosRaw[j][4].toString() : "-",
          premio: atendimentosRaw[j][5] != null && atendimentosRaw[j][5] !== "" ? atendimentosRaw[j][5].toString() : "0,00",
          tipoContato: atendimentosRaw[j][6] != null && atendimentosRaw[j][6] !== "" ? atendimentosRaw[j][6].toString() : "-",
          retido: atendimentosRaw[j][7] != null ? atendimentosRaw[j][7].toString() : "",
          pontos: atendimentosRaw[j][8] != null && atendimentosRaw[j][8] !== "" ? parseFloat(atendimentosRaw[j][8]) : 0
        });
      }
    }
  }

  // 4. Processar Rankings (Unificação das duas origens de pontos)
  var scoreIndividualMap = {};
  var scoreTeamMap = {};
  
  colaboradores.forEach(function(c) {
    scoreIndividualMap[c.nome] = { pontos: 0, equipe: c.equipe };
  });
  equipes.forEach(function(eq) {
    scoreTeamMap[eq] = 0;
  });

  // A) Somar pontos vindos diretamente da aba ATENDIMENTO
  if (atendimentosRaw.length > 1) {
    for (var m = 1; m < atendimentosRaw.length; m++) {
      var aNome = atendimentosRaw[m][0] ? atendimentosRaw[m][0].toString().trim() : null;
      var aPts = parseFloat(atendimentosRaw[m][8]) || 0;
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

/**
 * Salva ou edita um atendimento incluindo o novo campo de pontuação direta.
 */
function salvarAtendimento(dados, linhaEdit) {
  try {
    var sheet = obterAbaSegura(obterPlanilha(), "atendimento");
    if (sheet) {
      
      // Força o cálculo automatizado direto no servidor baseado nas regras informadas
      var pontosCalculados = calcularPontosAutomatizados(dados.tipoContato, dados.retido, dados.premio);
      
      var rowData = [
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
      return true;
    }
    return false;
  } catch(e) { 
    Logger.log("Erro ao salvar atendimento: " + e.toString());
    return false; 
  }
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

/**
 * Alinha e garante a estrutura correta de colunas de todas as abas.
 * Nova coluna inserida: "Pontuação Atendimento" na aba de atendimentos.
 */
function verificarESetupPlanilhas() {
  var ss = obterPlanilha();
  var estruturas = {
    "atendimento": ["Nome Colaborador", "Cargo", "CPF", "Data do Contato", "Nº da Apólice", "Valor do Prêmio", "Tipo de Contato", "Foi Retido?", "Pontuação Atendimento", "Data Registro"],
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

/**
 * Realiza o filtro dinâmico por datas integrando a pontuação unificada.
 * Agora incluindo o contador de Atendimentos Retidos.
 */
function obterDadosFiltradosAba1(dataInicioStr, dataFimStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetAtendimentos = ss.getSheetByName("atendimento");
    const sheetPontuacao = ss.getSheetByName("pontuacao");
    const sheetColabs = ss.getSheetByName("colaboradores");
    
    if (!sheetAtendimentos || !sheetPontuacao) {
      return { erro: "As abas necessárias não foram encontradas." };
    }
    
    const dadosAtend = sheetAtendimentos.getDataRange().getValues();
    const dadosPont = sheetPontuacao.getDataRange().getValues();
    
    const dataInicio = new Date(dataInicioStr + "T00:00:00");
    const dataFim = new Date(dataFimStr + "T23:59:59");
    
    let mapaColabEquipeMaster = {};
    if (sheetColabs) {
      const dadosColabs = sheetColabs.getDataRange().getValues();
      for (let c = 1; c < dadosColabs.length; c++) {
        let nomeC = dadosColabs[c][0] ? dadosColabs[c][0].toString().trim() : "";
        let eqC = dadosColabs[c][2] ? dadosColabs[c][2].toString().trim() : "";
        if (nomeC) mapaColabEquipeMaster[nomeC] = eqC;
      }
    }
    
    let mapaPontosAtendentes = {};
    let mapaEquipeDoAtendente = {};
    let mapaRetidosAtendentes = {}; // NOVO: Contador de retenções
    
    // Processa os pontos manuais da aba pontuação
    if (dadosPont.length > 1) {
      for (let p = 1; p < dadosPont.length; p++) {
        let linhaPont = dadosPont[p];
        if (!linhaPont[4]) continue;
        let dataLancamento = new Date(linhaPont[4]);
        
        if (dataLancamento >= dataInicio && dataLancamento <= dataFim) {
          let nomeAt = linhaPont[0] ? linhaPont[0].toString().trim() : null;
          let nomeEq = linhaPont[1] ? linhaPont[1].toString().trim() : mapaColabEquipeMaster[nomeAt] || "Sem Equipe";
          let pts = Number(linhaPont[2]) || 0;
          
          if (nomeAt) {
            mapaPontosAtendentes[nomeAt] = (mapaPontosAtendentes[nomeAt] || 0) + pts;
            mapaEquipeDoAtendente[nomeAt] = nomeEq;
          }
        }
      }
    }
    
    // Processa os pontos e as retenções da aba atendimento
    if (dadosAtend.length > 1) {
      for (let a = 1; a < dadosAtend.length; a++) {
        let linha = dadosAtend[a];
        if (!linha[3]) continue;
        let dataLinha = new Date(linha[3]);
        
        if (dataLinha >= dataInicio && dataLinha <= dataFim) {
          let colabName   = linha[0] ? linha[0].toString().trim() : null;
          let ptsAtend    = Number(linha[8]) || 0;
          let nomeEq      = mapaColabEquipeMaster[colabName] || "Sem Equipe";
          
          // Verifica se a coluna "Foi Retido?" (índice 7) é "Sim"
          let isRetido    = linha[7] && linha[7].toString().trim().toLowerCase() === "sim"; 
          
          if (colabName) {
            mapaPontosAtendentes[colabName] = (mapaPontosAtendentes[colabName] || 0) + ptsAtend;
            mapaEquipeDoAtendente[colabName] = nomeEq;
            
            // Soma +1 no contador se o atendimento foi retido
            if (isRetido) {
                mapaRetidosAtendentes[colabName] = (mapaRetidosAtendentes[colabName] || 0) + 1;
            }
          }
        }
      }
    }
    
    // Prepara o array final para enviar ao front-end
    let arrayAtendentes = [];
    for (let at in mapaPontosAtendentes) {
      if (mapaPontosAtendentes[at] !== 0) {
        arrayAtendentes.push({
          nome: at,
          equipe: mapaEquipeDoAtendente[at] || "Sem Equipe",
          pontos: mapaPontosAtendentes[at],
          retidos: mapaRetidosAtendentes[at] || 0 // Envia os retidos calculados
        });
      }
    }
    arrayAtendentes.sort((a, b) => b.pontos - a.pontos);
    
    return {
      sucesso: true,
      topAtendentes: arrayAtendentes.slice(0, 3)
    };
    
  } catch (erro) {
    Logger.log("Erro no processamento do ranking simplificado: " + erro.toString());
    return { erro: erro.toString() };
  }
}

// LÓGICA DE SOMA AUTOMÁTICA REQUISITADA
function calcularPontosAutomatizados(tipoContato, retido, premioStr) {
  var totalPontos = 0;
  
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
    var limpo = premioStr.toString().replace("R$", "").replace(/\./g, "").replace(",", ".").trim();
    var valorPremio = parseFloat(limpo) || 0;
    if (valorPremio > 3000 && retido === "Sim" && tipoContato === "Potencial") {
      totalPontos += 10;
    }
  }
  
  return totalPontos;
}