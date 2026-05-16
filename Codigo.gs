function doGet() {
  return HtmlService.createHtmlOutputFromFile('Cliente')
      .setTitle('Sistema de Processos')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function obterDadosParaFormulario() {
  const id = "1bx37PatyR6V0BKbYXyI-5rAENmDJGd419n9Ts0AEo6M";
  const ss = SpreadsheetApp.openById(id);
  const clientes = ss.getSheetByName("Clientes").getDataRange().getValues().slice(1).map(r => r[1]);
  const unidades = ss.getSheetByName("Unidades").getDataRange().getValues().slice(1).map(r => r[1]);
  return { clientes: clientes, unidades: unidades };
}


function buscarProcessosComJoin() {
  const id = "1bx37PatyR6V0BKbYXyI-5rAENmDJGd419n9Ts0AEo6M";
  const ss = SpreadsheetApp.openById(id);

  const processos = ss.getSheetByName("Processo").getDataRange().getValues();
  const clientes = ss.getSheetByName("Clientes").getDataRange().getValues();
  const unidades = ss.getSheetByName("Unidades").getDataRange().getValues();

  const mapaClientes = {};
  clientes.slice(1).forEach(linha => mapaClientes[linha[0]] = linha[1]);

  const mapaUnidades = {};
  unidades.slice(1).forEach(linha => mapaUnidades[linha[0]] = linha[1]);

  let resultado = [];
  // Começa no índice 1 para pular o cabeçalho
  for (let i = 1; i < processos.length; i++) {
    let linha = processos[i];
    
    resultado.push({
      linhaPlanilha: i + 1, // Guarda a linha real da planilha para o Update e Delete
      cliente: mapaClientes[linha[0]] || linha[0],
      valor: linha[1],
      garantia: linha[2],
      status: linha[3],
      prioridade: linha[4],
      unidade: mapaUnidades[linha[5]] || linha[5] || "Sem Unidade"
    });
  }
  return resultado;
}


function criarProcesso(nomeCliente, valor, garantia, status, nomeUnidade) {
  const id = "1bx37PatyR6V0BKbYXyI-5rAENmDJGd419n9Ts0AEo6M";
  const ss = SpreadsheetApp.openById(id);
  const sheetSistema = ss.getSheetByName("Processo"); 
  const sheetClientes = ss.getSheetByName("Clientes");
  const sheetUnidades = ss.getSheetByName("Unidades");

  const valorNum = Number(valor);

  // REGRA DE NEGÓCIO
  let prioridade = "Média";
  if (valorNum < 15000) {
    prioridade = "Baixa";
  } else if (garantia !== "Não") {
    prioridade = "Alta"; 
  }

  let idCliente = "";
  const clientes = sheetClientes.getDataRange().getValues();
  let clienteLimpo = nomeCliente.toString().trim().toUpperCase();
  for (let i = 1; i < clientes.length; i++) {
    if (clientes[i][1].toString().trim().toUpperCase() === clienteLimpo) {
      idCliente = clientes[i][0]; break;
    }
  }
  if (!idCliente) {
    let proxId = clientes.length;
    idCliente = "'00" + proxId + "_CLI"; 
    sheetClientes.appendRow([idCliente, clienteLimpo]);
  }

  // BUSCA OU CRIA UNIDADE
  let idUnidade = "";
  const unidades = sheetUnidades.getDataRange().getValues();
  let unidadeLimpa = nomeUnidade.toString().trim().toUpperCase();
  for (let i = 1; i < unidades.length; i++) {
    if (unidades[i][1].toString().trim().toUpperCase() === unidadeLimpa) {
      idUnidade = unidades[i][0]; break;
    }
  }
  if (!idUnidade) {
    let proxUnid = unidades.length;
    idUnidade = "'00" + proxUnid + "_UNI";
    sheetUnidades.appendRow([idUnidade, unidadeLimpa]);
  }


  sheetSistema.appendRow([idCliente, valorNum, garantia, status, prioridade, idUnidade]);
  return true;
}

function atualizarStatus(linha, novoStatus) {
  const ss = SpreadsheetApp.openById("1bx37PatyR6V0BKbYXyI-5rAENmDJGd419n9Ts0AEo6M");
  ss.getSheetByName("Sistema").getRange(linha, 4).setValue(novoStatus);
  return true;
}

function excluirProcesso(linha) {
  const ss = SpreadsheetApp.openById("1bx37PatyR6V0BKbYXyI-5rAENmDJGd419n9Ts0AEo6M");
  ss.getSheetByName("Sistema").deleteRow(linha);
  return true;
}
