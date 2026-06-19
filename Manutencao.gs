/*******************************************************
 * ARQUIVO: Manutencao.gs
 * DESCRIÇÃO: Scripts manuais para correção de permissões e erros
 *******************************************************/

/**
 * Função para aplicar permissão na Coluna A (Lock) 
 * sem resetar as permissões de R-W ou X-Z existentes.
 */
function adicionarPermissaoColunaA() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const EDITORES_COLUNA_A = [
    "atendimentomeuclinicas@hcpa.edu.br", 
    "jhernandez@hcpa.edu.br", 
    "mppadilha@hcpa.edu.br", 
    "rferreira@hcpa.edu.br", 
    "vlacerda@hcpa.edu.br",
    "rvaz@hcpa.edu.br"
  ];

  const ABAS_IGNORAR = new Set([
    "Impressão SAMIS", "Base", "Formulário", "Respostas", 
    "Respostas do formulário 1", "Consolidado 2025", "AUX_DASH"
  ]);

  const sheets = ss.getSheets();
  let log = [];

  const resp = ui.alert(
    "Adicionar Permissão na Coluna A", 
    "Isso vai liberar a Coluna A (A5:A1000) nas abas mensais e dar acesso APENAS ao grupo do SAMIS informado.\n\n" +
    "As outras colunas (R-W, X-Z) NÃO serão alteradas.\n\nDeseja continuar?", 
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    sheets.forEach(sheet => {
      const nome = sheet.getName();
      if (ABAS_IGNORAR.has(nome)) return;

      Logger.log(`Processando: ${nome}...`);
      
      const rangeColA = sheet.getRange("A5:A1000");

      // 1. ACHAR A PROTEÇÃO GERAL DA ABA E LIBERAR A COLUNA A
      const sheetProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      sheetProtections.forEach(prot => {
        let rangesLiberados = prot.getUnprotectedRanges();
        rangesLiberados.push(rangeColA);
        prot.setUnprotectedRanges(rangesLiberados);
      });

      // 2. CRIAR (OU ATUALIZAR) A PROTEÇÃO ESPECÍFICA DA COLUNA A
      const rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      rangeProtections.forEach(p => {
        const r = p.getRange();
        if (r.getColumn() === 1 && r.getLastColumn() === 1 && r.getRow() === 5) {
          p.remove(); 
        }
      });

      const novaProtA = rangeColA.protect().setDescription(`Permissão Lock (SAMIS) - ${nome}`);
      novaProtA.removeEditors(novaProtA.getEditors());
      novaProtA.addEditors(EDITORES_COLUNA_A);

      log.push(`✅ ${nome}: Coluna A liberada e restrita ao grupo SAMIS.`);
    });

    ui.alert("Sucesso!\n\n" + log.join("\n"));

  } catch (e) {
    ui.alert("Erro: " + e.message);
  }
}

/**
 * Função de emergência para destravar uma aba específica
 */
function resgatarAbaTravada() {
  const nomeAbaProblematica = "DEZEMBRO/25"; // Verifique se o nome está EXATO
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(nomeAbaProblematica);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Aba "${nomeAbaProblematica}" não foi encontrada.`);
    return;
  }

  const rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  let countRange = 0;
  rangeProtections.forEach(p => {
    if (p.canEdit()) {
      p.remove();
      countRange++;
    }
  });

  const sheetProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  let countSheet = 0;
  sheetProtections.forEach(p => {
    if (p.canEdit()) {
      p.remove();
      countSheet++;
    }
  });

  SpreadsheetApp.getUi().alert(
    `Resgate concluído na aba "${nomeAbaProblematica}"!\n` +
    `Intervalos removidos: ${countRange} | Proteções de folha removidas: ${countSheet}`
  );
}

/**
 * RESET TOTAL DE PERMISSÕES (USE COM CAUTELA)
 */
function aplicarPermissoesDeSeguranca() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const resposta = ui.alert(
    '⚠️ Atenção: Redefinição de Permissões',
    'Você tem certeza que deseja RESETAR e APLICAR as novas permissões?\n' +
    'Isso removerá editores antigos e aplicará as regras para os grupos de e-mail.',
    ui.ButtonSet.YES_NO
  );

  if (resposta !== ui.Button.YES) {
    ui.alert('Operação cancelada.');
    return;
  }

  const SUPER_ADMINS = ["samisadm@hcpa.edu.br", "lklrocha@hcpa.edu.br"];
  const GRUPO_X_Z = ["gdotto@hcpa.edu.br", "RecepcaoAmbulatorio@hcpa.edu.br"];
  const GRUPO_R_W = [
    "atendimentomeuclinicas@hcpa.edu.br", "jhernandez@hcpa.edu.br", 
    "mppadilha@hcpa.edu.br", "rferreira@hcpa.edu.br", 
    "vlacerda@hcpa.edu.br", "rvaz@hcpa.edu.br"
  ];

  const ABAS_IGNORAR = new Set([
    "Impressão SAMIS", "Base", "Formulário", "Respostas", 
    "Respostas do formulário 1", "Consolidado 2025", "AUX_DASH"
  ]);

  const sheets = ss.getSheets();
  let logRelatorio = []; 
  let abasProcessadas = 0;
  let abasIgnoradas = 0;

  Logger.log(">>> INÍCIO DA AUDITORIA DE PERMISSÕES <<<");

  try {
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();

      if (ABAS_IGNORAR.has(sheetName)) {
        abasIgnoradas++;
        return;
      }

      Logger.log(`[PROCESSANDO] Aba: ${sheetName}...`);

      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protections.forEach(p => p.remove());
      
      const rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      rangeProtections.forEach(p => p.remove());

      const rangeXZ = sheet.getRange("X5:Z1000");
      const rangeRW = sheet.getRange("R5:W1000");

      const protection = sheet.protect().setDescription(`Proteção Geral - ${sheetName}`);
      protection.setUnprotectedRanges([rangeXZ, rangeRW]);
      protection.removeEditors(protection.getEditors());
      protection.addEditors(SUPER_ADMINS);

      const pXZ = rangeXZ.protect().setDescription(`Permissão Grupo X-Z - ${sheetName}`);
      pXZ.removeEditors(pXZ.getEditors());
      pXZ.addEditors(SUPER_ADMINS);
      pXZ.addEditors(GRUPO_X_Z);

      const pRW = rangeRW.protect().setDescription(`Permissão Grupo R-W - ${sheetName}`);
      pRW.removeEditors(pRW.getEditors());
      pRW.addEditors(SUPER_ADMINS);
      pRW.addEditors(GRUPO_R_W);

      abasProcessadas++;
      logRelatorio.push(`✅ ${sheetName}`);
    });

  } catch (e) {
    ui.alert("Erro Crítico durante a execução: " + e.message);
    logRelatorio.push(`❌ ERRO: ${e.message}`);
  }

  const mensagemFinal = `Concluído: ${abasProcessadas} processadas, ${abasIgnoradas} ignoradas.`;
  ui.alert("Relatório de Execução", mensagemFinal, ui.ButtonSet.OK);
}

function TESTE_DE_CONEXAO_URL() {
  // O Link completo da planilha do Robô
  const URL_ALVO = "https://docs.google.com/spreadsheets/d/1xH8OU1QGfgPU1cCvhk8dH50vR9w8yoOP2vDkePVp2XG7rMFSHrHA6B8w/edit";
  
  try {
    // Trocamos openById por openByUrl
    const ss = SpreadsheetApp.openByUrl(URL_ALVO);
    Logger.log("✅ SUCESSO! Conectado à planilha: " + ss.getName());
    
    const sheet = ss.getSheetByName("Fila");
    if (sheet) {
      Logger.log("✅ SUCESSO! Aba 'Fila' encontrada.");
    } else {
      Logger.log("❌ ERRO: A planilha abre, mas não tem a aba 'Fila'.");
    }
    
  } catch (e) {
    Logger.log("❌ FALHA: " + e.message);
  }
}
