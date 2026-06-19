/*******************************************************
 * ARQUIVO: Código.gs
 * DESCRIÇÃO: Core do sistema, comunicação com HTML e Menus
 *******************************************************/


// --- FUNÇÃO PARA INCLUIR ARQUIVOS HTML/CSS EXTERNOS ---
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- FUNÇÃO PARA ABRIR O GUIA DE AJUDA ---
function openHelpGuide() {
  const html = HtmlService.createHtmlOutputFromFile('HelpGuide')
    .setWidth(900)   // Largura aumentada para ficar mais confortável
    .setHeight(750); // Altura aumentada
  
  // O segundo parâmetro é o título da janela. Deixamos um espaço ' ' para ficar "limpo"
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/*******************************************************
 * CONSTANTES GLOBAIS & CONFIGURAÇÕES
 *******************************************************/
const FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSe5bfrtoc7Gohn6NXJfETFBLe4dZEU2Z2GkrleiJG_WXiMLtg/viewform";


// Configurações do Lock (Travamento)
const LOCK_COL = 1; // Coluna A
const LOCK_TTL_MS = 15 * 60 * 1000; // 15 minutos
const LOCK_PREFIX = "LOCK|"; // formato: LOCK|email|isoDate


// Coluna do Nº de folhas (W = 23)
const NUM_FOLHAS_COL = 23; 


// ID da pasta raiz do Google Drive para Anexos
const DRIVE_ROOT_FOLDER_ID = "0ABR5GVMXpSwlUk9PVA"; 


// --- SUPER ADMINISTRADORES (Acesso ao Menu de Configurações) ---
const SUPER_ADMINS = [
  "lklrocha@hcpa.edu.br",
  "samisadm@hcpa.edu.br",
  "ldsmachado@hcpa.edu.br",
  "auryaneborges@hcpa.edu.br"
];


// --- MAPA PADRÃO DE USUÁRIOS (Fallback se não houver configs salvas) ---
const DEFAULT_USER_MAP = {
  "auryaneborges@hcpa.edu.br": "AURYANE",
  "jhernandez@hcpa.edu.br": "JORGIA",
  "ldsmachado@hcpa.edu.br": "LIDIANE",
  "mppadilha@hcpa.edu.br": "MIRIAN",
  "rferreira@hcpa.edu.br": "REGINA",
  "vlacerda@hcpa.edu.br": "VANESSA",
  "lklrocha@hcpa.edu.br": "LEONARDO",
  "samisadm@hcpa.edu.br": "SAMIS ADMIN"
};


// Cabeçalhos esperados para leitura
const ADMIN_HEADERS = [
  "Data e Hora da Solicitação",
  "Nome completo do paciente",
  "E-mail do Paciente",
  "Telefone do Paciente",
  "Número do prontuário:",
  "NOME",
  "CPF",
  "Exames Ambulatoriais",
  "Internação: Sumários de Alta, Prescrição, Descrição Cirúrgica e Obstétrica",
  "Informe o período da solicitação (data final)",
  "Informe o período da solicitação (data inicial)",
  "Declaro que o motivo da solicitação fora do app Meu Clínicas é justificada pelo motivo",
  "Declaro que o motivo da solicitação fora do app Meu Clínicas é justificada pelo motivo ",
  "Forma de Envio do Documento Solicitado:",
  "Outros",
  "Endereço de e-mail colaborador",
  "Pedido realizado por quem?",
  "QUEM ATENDEU (SAMIS)",
  "DATA DO ATENDIMENTO (SAMIS)",
  "LINK",
  "OBSERVAÇÕES SAMIS",
  "TEMPO MÉDIO DISP. SAMIS",
  "DATA DA ENTREGA AO PACIENTE (RECEPÇÃO)",
  "DATA DA ENTREGA AO PACIENTE, (RECEPÇÃO)",
  "QUEM ATENDEU (RECEPÇÃO).",
  "QUEM ATENDEU (RECEPÇÃO)",
  "OBS e Pedidos de Correção" 
];


// Campos permitidos para SALVAR
const SAMIS_EDITABLE_HEADERS = [
  "QUEM ATENDEU (SAMIS)", 
  "DATA DO ATENDIMENTO (SAMIS)",
  "LINK",
  "OBSERVAÇÕES SAMIS",
  "Nº de folhas"
];


// Abas que NÃO são abas mensais administrativas
const EXCLUDE_SHEETS = new Set([
  "Impressão SAMIS",
  "Base",
  "Formulário",
  "Respostas",
  "Respostas do formulário 1",
  "Consolidado 2025",
  "AUX_DASH",
  "Email" // Adicionado para não misturar no dropdown principal
]);


/*******************************************************
 * MENUS E INTERFACE
 *******************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const userEmail = Session.getActiveUser().getEmail();

  ui.createMenu("📋 Formulário")
    .addItem("Abrir Formulário (Google Forms)", "abrirFormulario")
    .addSeparator()
    .addItem("Autorização em Branco", "abrirLinkAutorizacao")
    .addToUi();

  const usuariosBloqueadosMenu = [
    "gdotto@hcpa.edu.br", 
    "RecepcaoAmbulatorio@hcpa.edu.br"
  ];

  if (!usuariosBloqueadosMenu.includes(userEmail)) {
    ui.createMenu("🧩 SAMIS")
      .addItem("Abrir Painel (aba atual)", "adminAbrirPainelAbaAtual")
      .addItem("Editar linha selecionada", "adminEditarLinhaSelecionada")
      .addSeparator()
      .addItem("Ir para próxima pendente", "adminIrParaProximaPendente")
      .addSeparator()
      .addItem("🔄 Verificar Permissões", "verificarPermissoes") 
      .addItem("🎓 Guia de Treinamento", "openHelpGuide") // <--- ADICIONADO AQUI
      .addToUi();
  }
}


function verificarPermissoes() {
  const ui = SpreadsheetApp.getUi();
  // Sem try/catch para forçar o Google a exibir o pop-up de autorização se necessário
  const driveCheck = DriveApp.getRootFolder(); 
  
  ui.alert("✅ Tudo Certo!", 
    "Suas permissões já estão atualizadas e o acesso ao Drive foi confirmado.", 
    ui.ButtonSet.OK);
}


function abrirFormulario() {
  const html = HtmlService.createHtmlOutput(
    `<script>
       window.open("${FORM_URL}", "_blank");
       google.script.host.close();
     </script>`
  )
  .setWidth(250)
  .setHeight(80);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Abrindo Formulário...');
}

function abrirLinkAutorizacao() {
  const url = "https://intranet.hcpa.edu.br/component/jdownloads/?task=download.send&id=502&catid=6&m=0&Itemid=413";
  const html = HtmlService.createHtmlOutput(
    `<script>
       window.open("${url}", "_blank");
       google.script.host.close();
     </script>`
  )
  .setWidth(250)
  .setHeight(80);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Abrindo Documento...');
}

function adminAbrirPainelAbaAtual() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  adminShowPanel_({ sheetName: sheet.getName(), row: null, mode: "open" });
}


function adminEditarLinhaSelecionada() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  const range = sheet.getActiveRange();


  if (!range) {
    SpreadsheetApp.getUi().alert("Selecione uma célula da linha que você quer editar.");
    return;
  }


  const row = range.getRow();
  adminShowPanel_({ sheetName: sheet.getName(), row: row, mode: "edit" });
}


function adminIrParaProximaPendente() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();


  const meta = findHeaderMeta_(sheet);
  if (!meta) {
    SpreadsheetApp.getUi().alert(`A aba "${sheet.getName()}" não parece ser uma aba mensal (cabeçalho não encontrado).`);
    return;
  }


  const pendingRow = findFirstPendingRow_(sheet, meta);
  if (!pendingRow) {
    SpreadsheetApp.getUi().alert(`Nenhuma pendência encontrada na aba "${sheet.getName()}".`);
    return;
  }


  sheet.setActiveSelection(`B${pendingRow}`);
  adminShowPanel_({ sheetName: sheet.getName(), row: pendingRow, mode: "edit" });
}


function adminShowPanel_(opts) {
  const tpl = HtmlService.createTemplateFromFile("AdminForm");
  tpl.bootstrap = JSON.stringify(getAdminBootstrap_(opts));


  const html = tpl.evaluate()
    .setTitle("SAMIS Admin")
    .setWidth(1280)
    .setHeight(820);


  SpreadsheetApp.getUi().showModalDialog(html, "🧩 Gestão de Requerimentos");
}


/*******************************************************
 * BACKEND: COMUNICAÇÃO COM O HTML
 *******************************************************/


function getAdminBootstrap_(opts) {
  const ss = SpreadsheetApp.getActive();
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  
  const adminSheets = getAdminSheets_(ss);


  const defaultSheetName =
    (opts.sheetName && adminSheets.includes(opts.sheetName))
      ? opts.sheetName
      : (adminSheets[0] || ss.getActiveSheet().getName());


  const sheet = ss.getSheetByName(defaultSheetName);
  const meta = sheet ? findHeaderMeta_(sheet) : null;


  const row = (opts.row && Number(opts.row) > 1) ? Number(opts.row) : null;


  const headerMap = meta ? buildHeaderMap_(sheet, meta) : {};
  const values = (meta && row) ? readRow_(sheet, row, meta, headerMap) : {};


  if (sheet && row) {
    values["Nº de folhas"] = String(sheet.getRange(row, NUM_FOLHAS_COL).getDisplayValue() || "").trim();
  }


  const patientNames = (meta && headerMap["Nome completo do paciente"])
    ? getUniqueColumnValues_(sheet, headerMap["Nome completo do paciente"], meta.headerRow + 1)
    : [];


  const lock = (sheet && row) ? acquireSoftLock_(sheet, row) : { ok: true, held: false };


  // Verifica se o usuário é Admin para mostrar o botão de engrenagem
  const isAdmin = SUPER_ADMINS.includes(userEmail);


  return {
    mode: opts.mode || "open",
    sheetName: defaultSheetName,
    row,
    adminSheets,
    meta,
    headerMap,
    values,
    patientNames,
    samisEditableHeaders: SAMIS_EDITABLE_HEADERS,
    lock,
    isAdmin: isAdmin
  };
}


function adminLoad(sheetName, row) {
  return getAdminBootstrap_({ sheetName, row: row || null, mode: row ? "edit" : "open" });
}


function adminFindNextPending(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Aba não encontrada: " + sheetName);


  const meta = findHeaderMeta_(sheet);
  if (!meta) throw new Error("Cabeçalho não encontrado na aba: " + sheetName);


  const r = findFirstPendingRow_(sheet, meta);
  return { row: r || null };
}


function adminListPendentes(sheetName, maxResults) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Aba não encontrada: " + sheetName);


  const meta = findHeaderMeta_(sheet);
  if (!meta) throw new Error("Cabeçalho não encontrado na aba: " + sheetName);


  const startRow = meta.headerRow + 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return { count: 0, items: [] };


  const n = lastRow - startRow + 1;


  const COL_NOME = 3;      // C
  const COL_ATEND = 18;    // R
  const COL_PRONT = 6;     // F


  const nomes = sheet.getRange(startRow, COL_NOME, n, 1).getDisplayValues();
  const atend = sheet.getRange(startRow, COL_ATEND, n, 1).getDisplayValues();
  const pronts = sheet.getRange(startRow, COL_PRONT, n, 1).getDisplayValues();


  const limit = Number(maxResults || 300);


  const items = [];
  let total = 0;


  for (let i = 0; i < n; i++) {
    const nome = String(nomes[i][0] || "").trim();
    const a = String(atend[i][0] || "").trim();
    if (nome && !a) {
      total++;
      if (items.length < limit) {
        items.push({
          row: startRow + i,
          nome,
          prontuario: String(pronts[i][0] || "").trim()
        });
      }
    }
  }


  return { count: total, items };
}

function adminSave(sheetName, row, values) {
  // 1. Verificação de Modo Manutenção
  const props = PropertiesService.getScriptProperties();
  const maintenance = props.getProperty('MAINTENANCE_MODE');
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  
  if (maintenance === 'true' && !SUPER_ADMINS.includes(userEmail)) {
    throw new Error("⚠️ O sistema está em MODO DE MANUTENÇÃO. O salvamento está temporariamente bloqueado.");
  }

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Aba não encontrada: " + sheetName);

  const meta = findHeaderMeta_(sheet);
  if (!meta) throw new Error("Cabeçalho não encontrado.");

  const rowNum = Number(row);
  if (!rowNum || rowNum <= meta.headerRow) throw new Error("Linha inválida.");

  // 2. Lock (Concorrência)
  const lock = acquireSoftLock_(sheet, rowNum);
  if (!lock.ok || lock.heldByOther) throw new Error(lock.message || "Linha em uso por outro usuário.");

  // 3. Leitura dos Dados Atuais da Linha (Para contexto)
  const headerMap = buildHeaderMap_(sheet, meta);
  const currentRowData = readRow_(sheet, rowNum, meta, headerMap); 
  
  // Mescla o que já existe na linha com o que veio do formulário
  const mergedData = { ...currentRowData, ...values }; 
  
  // 4. --- LÓGICA INTELIGENTE DE REGRAS (SAMIS) ---
  const pedidoPor = normalize_(mergedData["Pedido realizado por quem?"]);
  
  // Tenta pegar 'Forma de Envio' de todas as variações possíveis de cabeçalho
  const formaEnvio = normalize_(
      mergedData["Forma de Envio do Documento Solicitado:"] || 
      mergedData["Forma de Envio do Documento Solicitado"] || 
      ""
  );
  
  const isTerceiro = pedidoPor.includes("terceiro");
  const isEmail = formaEnvio.includes("email") || formaEnvio.includes("e-mail");
  
  // Define a observação base: usa a nova se enviada, senão usa a antiga
  // (Nota: values[...] pode ser string vazia se o usuário limpou o campo, então checamos undefined)
  let novaObs = (values["OBSERVAÇÕES SAMIS"] !== undefined) 
      ? values["OBSERVAÇÕES SAMIS"] 
      : currentRowData["OBSERVAÇÕES SAMIS"];
      
  // Garante que seja string para evitar erros
  novaObs = String(novaObs || "").trim();
  
  // --- LÓGICA DE TERCEIRO ---
  if (isTerceiro) {
    try {
      const pront = String(mergedData["Número do prontuário"] || mergedData["Número do prontuário:"] || "").trim().replace(/\D/g, '');
      const nomePac = String(mergedData["Nome completo do paciente"] || mergedData["NOME"] || "").trim().toUpperCase();
      
      if (pront && nomePac) {
        const folderName = `${pront}- ${nomePac}`;
        const rootFolder = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
        const folders = rootFolder.getFoldersByName(folderName);
        
        if (folders.hasNext()) {
          const folder = folders.next();
          const access = folder.getSharingAccess();
          const isLocked = (access === DriveApp.Access.PRIVATE || access === DriveApp.Access.NONE);
          
          if (isLocked) {
            // Se está trancada, força OBS de verificação (segurança)
            novaObs = "Falta Verificação Docs/Contate Paciente";
          } else {
            // Se está liberada (destrancada), aplica regra de envio
            if (isEmail) {
              if (normalize_(novaObs) !== "e-mail enviado") novaObs = "Falta Enviar E-mail";
            } else {
              // CORREÇÃO AQUI TAMBÉM: Só vira "0" se estiver vazio
              if (!novaObs || novaObs === "") {
                 novaObs = "0";
              }
              // Garante que a pasta continue liberada no Drive
              try { folder.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT); } catch(e){}
            }
          }
        } else {
           novaObs = "Falta Verificação Docs/Contate Paciente"; // Pasta nem existe
        }
      }
    } catch (eDrive) {
      console.warn("Erro ao verificar Drive no Save: " + eDrive);
    }
  } 
  // --- LÓGICA DE PRÓPRIO PACIENTE ---
  else {
     if (isEmail) {
        // Se for e-mail, forçamos o lembrete de envio se ainda não foi enviado
        if (normalize_(novaObs) !== "e-mail enviado") novaObs = "Falta Enviar E-mail";
     } else {
        // CORREÇÃO PRINCIPAL:
        // Se NÃO for e-mail (ex: Papel, Presencial), aceita o que o usuário escreveu.
        // Só muda para "0" se o campo estiver vazio.
        if (!novaObs || novaObs === "") {
           novaObs = "0";
        }
     }
  }
  
  // Atualiza o valor final no objeto values para ser gravado
  values["OBSERVAÇÕES SAMIS"] = novaObs;

  // 5. Gravação na Planilha
  const toWrite = {};
  
  // Filtra apenas os cabeçalhos que permitimos editar via SAMIS_EDITABLE_HEADERS
  SAMIS_EDITABLE_HEADERS.forEach((h) => {
    if (h in values) toWrite[h] = values[h];
  });

  Object.keys(toWrite).forEach((header) => {
    // Caso especial: Nº de folhas tem coluna fixa definida na constante
    if (normalize_(header) === normalize_("Nº de folhas")) {
      sheet.getRange(rowNum, NUM_FOLHAS_COL, 1, 1).setValue(toWrite[header]);
      return;
    }
    
    // Outros campos: usa o mapa do cabeçalho
    const col = headerMap[header];
    if (!col) return;
    sheet.getRange(rowNum, col, 1, 1).setValue(toWrite[header]);
  });

  // Renova o Lock para dizer que ainda estamos ativos
  renewSoftLock_(sheet, rowNum);

  return { ok: true, row: rowNum, newObs: novaObs };
}

function adminKeepAlive(sheetName, row) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Aba não encontrada: " + sheetName);
  if (!row) throw new Error("Linha inválida.");


  const res = acquireSoftLock_(sheet, Number(row));
  return res;
}

// --- FUNÇÕES ESPECÍFICAS PARA A ABA EMAIL ---

function adminLoadEmailData() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName("Email");
  if (!sheet) return { items: [] };


  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { items: [] };


  // Lê da Coluna A até M (13 colunas)
  // A=0(dataHora), B=1(nome), C=2(email), D=3(telefone), E=4(prontuario), F=5, G=6, H=7 (solicitações), I=8, J=9(origem), K=10, L=11(protocolo), M=12(statusLabel)
  const dados = sheet.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
  
  const items = dados.map((r, i) => ({
    row: i + 2,
    dataHora: r[0],
    nome: r[1],
    email: r[2],
    telefone: r[3],
    prontuario: r[4],
    // Mescla colunas G(6), H(7), I(8) das solicitações
    solicitacoes: [r[6], r[7], r[8]].filter(x => x && x.trim()).join("\n------------------------------\n"),
    origem: r[9],
    protocolo: r[11], // Coluna L
    statusLabel: r[12] // Coluna M (Labels do Gmail ou Status do Sistema)
  }));


  // FILTRAGEM ROBUSTA
  const pendentes = items.filter(item => {
    const s = (item.statusLabel || "").toUpperCase();
    
    // 1. Se tiver "ATENDIDO" ou "ATENDIDOS" (vindo da Label do Gmail), esconde.
    if (s.includes("ATENDIDO")) return false;


    // 2. Se já foi enviado pelo sistema (status interno), esconde.
    if (s === "ENVIADO" || s === "ENVIAR_AGORA") return false;


    return true; // Se passou pelos filtros, aparece na lista
  });


  return { items: pendentes, total: pendentes.length };
}


// Função para o usuário clicar em "Responder/Atender" no painel novo
function adminAgendarRespostaEmail(payload) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Email");
  const row = Number(payload.row);
  
  // payload: { row, corpoResposta, labelParaAplicar (ex: "ATENDIDOS 01/2026") }
  
  // Grava na planilha para o Robô ler
  // Coluna M (13) = STATUS_ENVIO -> "ENVIAR_AGORA"
  // Coluna N (14) = Corpo da Resposta
  // Coluna O (15) = Label a aplicar (ex: ATENDIDOS 01/2026)
  
  sheet.getRange(row, 13).setValue("ENVIAR_AGORA");
  sheet.getRange(row, 14).setValue(payload.corpoResposta);
  sheet.getRange(row, 15).setValue(payload.labelParaAplicar);
  
  return { success: true };
}

// --- VERSÃO PARA FILA DE ENVIO (USAR ESTA) ---
// --- ATUALIZADO: ATUALIZA STATUS "E-MAIL ENVIADO" ---
const ID_PLANILHA_ROBO = "1LMweHS1jfM4QMCXmwhPmnzaRl2w0Y8mJ5KODNln1iWg"; 

function adminEnviarEmailResposta(destinatario, linkPasta, sheetName, row) {
  try {
    // 1. Gera o Texto Padrão
    const hora = new Date().getHours();
    let saudacao = "Bom dia";
    if (hora >= 12 && hora < 18) saudacao = "Boa tarde";
    if (hora >= 18) saudacao = "Boa noite";

    const assunto = "Resposta a Solicitação";
    const corpo = `${saudacao},\n\n` +
                  `Segue em anexos documentos solicitados.\n\n` +
                  `Atenciosamente`;

    // 2. Conecta na Planilha de Fila
    const ssFila = SpreadsheetApp.openById(ID_PLANILHA_ROBO);
    const sheetFila = ssFila.getSheetByName("Fila");
    
    if (!sheetFila) {
      throw new Error(`A planilha do Robô foi encontrada, mas a aba "Fila" não existe nela.`);
    }
    
    // 3. Adiciona na Fila
    sheetFila.appendRow([
      new Date(),       // A: Data
      destinatario,     // B: Para
      assunto,          // C: Assunto
      corpo,            // D: Mensagem
      linkPasta,        // E: Link Drive
      "PENDENTE"        // F: Status
    ]);

    // 4. ATUALIZA A PLANILHA DE GESTÃO (Status: E-mail Enviado)
    if (sheetName && row) {
      try {
        // Reutiliza a função adminSave para garantir segurança e log
        adminSave(sheetName, row, {
          "OBSERVAÇÕES SAMIS": "E-mail Enviado"
        });
      } catch (eLocal) {
        console.warn("E-mail na fila, mas erro ao atualizar status local: " + eLocal);
        // Não lançamos erro aqui para não assustar o usuário, já que o e-mail foi para a fila com sucesso.
      }
    }
    
    return { success: true };
    
  } catch (e) {
    if (e.toString().includes("openById")) {
      throw new Error("ERRO DE ACESSO: Verifique se você compartilhou a planilha " + ID_PLANILHA_ROBO + " como EDITOR.");
    }
    throw new Error("Falha na Fila: " + e.toString());
  }
}

/*******************************************************
 * HELPERS (LOCK E DADOS)
 *******************************************************/


function acquireSoftLock_(sheet, row) {
  const docLock = LockService.getDocumentLock();
  docLock.waitLock(5000);


  try {
    const now = Date.now();
    const me = getUserId_();


    const cell = sheet.getRange(row, LOCK_COL);
    const raw = String(cell.getDisplayValue() || "").trim();


    const parsed = parseLock_(raw);
    const isExpired = parsed && (now - parsed.tsMs > LOCK_TTL_MS);


    if (!parsed || isExpired) {
      const newVal = formatLock_(me, new Date(now));
      cell.setValue(newVal);
      return {
        ok: true, held: true, mine: true, heldByOther: false,
        owner: me, sinceIso: new Date(now).toISOString(),
        message: "Lock adquirido por você."
      };
    }


    const mine = (normalize_(parsed.owner) === normalize_(me));
    if (mine) {
      const newVal = formatLock_(me, new Date(now));
      cell.setValue(newVal);
      return {
        ok: true, held: true, mine: true, heldByOther: false,
        owner: me, sinceIso: new Date(now).toISOString(),
        message: "Lock renovado por você."
      };
    }


    return {
      ok: false, held: true, mine: false, heldByOther: true,
      owner: parsed.owner || "Outro usuário", sinceIso: new Date(parsed.tsMs).toISOString(),
      message: `Em uso por ${parsed.owner || "outro usuário"} (soft lock).`
    };


  } finally {
    docLock.releaseLock();
  }
}


function renewSoftLock_(sheet, row) {
  const docLock = LockService.getDocumentLock();
  docLock.waitLock(5000);
  try {
    const now = new Date();
    const me = getUserId_();
    sheet.getRange(row, LOCK_COL).setValue(formatLock_(me, now));
  } finally {
    docLock.releaseLock();
  }
}


function parseLock_(s) {
  if (!s) return null;
  const str = String(s);
  if (!str.startsWith(LOCK_PREFIX)) return null;
  const parts = str.split("|");
  if (parts.length < 3) return null;


  const owner = parts[1] || "";
  const iso = parts.slice(2).join("|");
  const ts = Date.parse(iso);
  if (!owner || !isFinite(ts)) return null;


  return { owner, tsMs: ts };
}


function formatLock_(owner, dateObj) {
  return `${LOCK_PREFIX}${owner}|${dateObj.toISOString()}`;
}


function getUserId_() {
  const a = safeEmail_(Session.getActiveUser());
  if (a) return a;
  const e = safeEmail_(Session.getEffectiveUser());
  if (e) return e;
  return "usuario-desconhecido";
}


function safeEmail_(userObj) {
  try {
    const email = userObj && userObj.getEmail ? userObj.getEmail() : "";
    return String(email || "").trim();
  } catch (e) {
    return "";
  }
}


function getAdminSheets_(ss) {
  const out = [];
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if (EXCLUDE_SHEETS.has(name)) return;
    const meta = findHeaderMeta_(sh);
    if (meta) out.push(name);
  });
  return out;
}


function findHeaderMeta_(sheet) {
  const maxR = Math.min(30, sheet.getMaxRows());
  const maxC = Math.min(40, sheet.getMaxColumns());
  const vals = sheet.getRange(1, 1, maxR, maxC).getDisplayValues();
  const target = normalize_("Data e Hora da Solicitação");


  for (let r = 0; r < vals.length; r++) {
    for (let c = 0; c < vals[r].length; c++) {
      if (normalize_(vals[r][c]) === target) {
        return { headerRow: r + 1, headerStartCol: c + 1 };
      }
    }
  }
  return null;
}


function buildHeaderMap_(sheet, meta) {
  const lastCol = sheet.getLastColumn();
  const width = lastCol - meta.headerStartCol + 1;
  if (width <= 0) return {};


  const headerRow = sheet.getRange(meta.headerRow, meta.headerStartCol, 1, width).getDisplayValues()[0];
  const rawMap = {};
  headerRow.forEach((h, i) => {
    const k = normalize_(h);
    if (k) rawMap[k] = meta.headerStartCol + i;
  });


  const expectedMap = {};
  ADMIN_HEADERS.forEach(h => {
    const col = rawMap[normalize_(h)];
    if (col) expectedMap[h] = col;
  });
  return expectedMap;
}


function readRow_(sheet, row, meta, headerMap) {
  const lastCol = sheet.getLastColumn();
  const width = lastCol - meta.headerStartCol + 1;
  const rowVals = sheet.getRange(row, meta.headerStartCol, 1, width).getDisplayValues()[0];
  const out = {};
  Object.keys(headerMap).forEach((header) => {
    const col = headerMap[header];
    const idx = col - meta.headerStartCol;
    out[header] = rowVals[idx] ?? "";
  });
  return out;
}


function getUniqueColumnValues_(sheet, col, startRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];
  const vals = sheet.getRange(startRow, col, lastRow - startRow + 1, 1)
    .getDisplayValues().flat().map(v => String(v || "").trim()).filter(v => v);
  return [...new Set(vals)].sort((a, b) => a.localeCompare(b, "pt-BR"));
}


function findFirstPendingRow_(sheet, meta) {
  const headerMap = buildHeaderMap_(sheet, meta);
  const col = headerMap["QUEM ATENDEU  (SAMIS)"] || headerMap["QUEM ATENDEU (SAMIS)"];
  if (!col) return null;


  const startRow = meta.headerRow + 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return null;


  const vals = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getDisplayValues();
  for (let i = 0; i < vals.length; i++) {
    const v = String(vals[i][0] || "").trim();
    if (!v) return startRow + i;
  }
  return null;
}


function normalize_(s) {
  return String(s || "").trim().replace(/\s+/g, " ").toLowerCase();
}

// --- FUNÇÃO AJUDANTE DE INTELIGÊNCIA DE STATUS ---
function atualizarStatusInteligente_(sheetName, row, folder, isRestricted, formaEnvio) {
  try {
    let novaObs = "";
    const envio = String(formaEnvio || "").toLowerCase();
    const isEmail = envio.includes("email") || envio.includes("e-mail");

    // LÓGICA 1: TERCEIRO AUTORIZADO (Restrito)
    if (isRestricted) {
      // Verifica se a pasta está TRANCADA ou ABERTA
      const access = folder.getSharingAccess();
      const isLocked = (access === DriveApp.Access.PRIVATE || access === DriveApp.Access.NONE);

      if (isLocked) {
        novaObs = "Falta Verificação Docs/Contate Paciente";
      } else {
        // Se está destrancada (foi liberada), define o próximo passo
        novaObs = isEmail ? "Falta Enviar E-mail" : "0";
      }
    } 
    // LÓGICA 2: PRÓPRIO PACIENTE
    else {
      novaObs = isEmail ? "Falta Enviar E-mail" : "0";
    }

    // Grava na Planilha imediatamente
    if (sheetName && row) {
      adminSave(sheetName, row, { "OBSERVAÇÕES SAMIS": novaObs });
    }
    
    return novaObs;
  } catch (e) {
    console.warn("Erro ao atualizar status inteligente: " + e);
    return "Erro Status";
  }
}

/*******************************************************
 * FUNÇÕES DE IDENTIFICAÇÃO E UPLOAD
 *******************************************************/


// 1. Identificar Usuário e Retornar Nome (AGORA LÊ DAS CONFIGURAÇÕES)
function getCurrentUserMap() {
  const email = Session.getActiveUser().getEmail().toLowerCase().trim();
  
  // Tenta ler do banco de propriedades
  let userMap = {};
  try {
    const stored = PropertiesService.getScriptProperties().getProperty('USER_MAP');
    if (stored) {
      userMap = JSON.parse(stored);
    } else {
      userMap = DEFAULT_USER_MAP; // Usa o padrão se nunca foi configurado
    }
  } catch (e) {
    userMap = DEFAULT_USER_MAP;
  }


  let nome = userMap[email];
  if (!nome) {
    nome = email.split("@")[0].toUpperCase();
  }


  return {
    email: email,
    nome: nome
  };
}

// --- ATUALIZAÇÃO DA FUNÇÃO DE UPLOAD PARA ACEITAR RESTRIÇÃO ---
function uploadFileToDrive(data) {
  const props = PropertiesService.getScriptProperties();
  const maintenance = props.getProperty('MAINTENANCE_MODE');
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  
  if (maintenance === 'true' && !SUPER_ADMINS.includes(userEmail)) {
    return { success: false, message: "⚠️ MODO DE MANUTENÇÃO." };
  }

  try {
    const contentType = data.mimeType || "application/octet-stream";
    const folderName = `${data.prontuario}- ${data.nomePaciente.toUpperCase()}`.trim();
    const finalFileName = folderName + ".pdf"; 
    const blob = Utilities.newBlob(Utilities.base64Decode(data.base64), contentType, finalFileName);
    const rootFolder = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
    const folders = rootFolder.getFoldersByName(folderName);
    
    let targetFolder;
    const deveRestringir = data.isRestricted === true; 

    if (folders.hasNext()) {
      targetFolder = folders.next();
      if (deveRestringir) {
         try { targetFolder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch(e){}
      }
    } else {
      targetFolder = rootFolder.createFolder(folderName);
      try {
        if (deveRestringir) {
          targetFolder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
        } else {
          targetFolder.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
        }
      } catch (e) {}
    }
    
    targetFolder.createFile(blob);
    
    // --- ATUALIZAÇÃO AUTOMÁTICA DE STATUS ---
    let obsAtualizada = null;
    // Só roda a lógica se for o último volume E tivermos os dados da linha
    if (data.updateStatus && data.sheetName && data.row) {
       obsAtualizada = atualizarStatusInteligente_(
         data.sheetName, 
         data.row, 
         targetFolder, 
         deveRestringir, 
         data.formaEnvio
       );
    }
    
    return { 
      success: true, 
      folderUrl: targetFolder.getUrl(),
      newObs: obsAtualizada, // Manda a nova OBS pro front-end
      message: deveRestringir ? "Arquivo salvo em pasta RESTRITA." : "Arquivo salvo com sucesso!"
    };
    
  } catch (e) {
    return { success: false, message: "Erro upload: " + e.toString() };
  }
}

// --- NOVAS FUNÇÕES PARA GESTÃO DOCUMENTAL ---

// 1. Salvar documento de verificação na subpasta "Docs" (CORRIGIDO: CRIA PASTA SE NÃO EXISTIR)
function saveVerificationDoc(data) {
  try {
    const rootFolder = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
    const folderName = `${data.prontuario}- ${data.nomePaciente.toUpperCase()}`.trim();
    const folders = rootFolder.getFoldersByName(folderName);
    let patientFolder;
    
    if (folders.hasNext()) {
      patientFolder = folders.next();
    } else {
      patientFolder = rootFolder.createFolder(folderName);
      try { patientFolder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch(e){}
    }
    
    const subFolders = patientFolder.getFoldersByName("Docs");
    let docsFolder;
    if (subFolders.hasNext()) { docsFolder = subFolders.next(); } 
    else { docsFolder = patientFolder.createFolder("Docs"); }
    
    const blob = Utilities.newBlob(Utilities.base64Decode(data.base64), data.mimeType, data.fileName);
    docsFolder.createFile(blob);
    
    // --- FORÇA STATUS "FALTA VERIFICAÇÃO" NA PLANILHA ---
    if (data.sheetName && data.row) {
       adminSave(data.sheetName, data.row, { 
         "OBSERVAÇÕES SAMIS": "Falta Verificação Docs/Contate Paciente" 
       });
    }
    
    return { success: true };
  } catch (e) {
    throw new Error("Falha no servidor: " + e.toString());
  }
}

// 2. Listar documentos na pasta "Docs" (AGORA LÊ O STATUS SALVO)
function listVerificationDocs(prontuario, nomePaciente) {
  try {
    const rootFolder = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
    const folderName = `${prontuario}- ${nomePaciente.toUpperCase()}`.trim();
    const folders = rootFolder.getFoldersByName(folderName);
    
    if (!folders.hasNext()) return [];
    const patientFolder = folders.next();
    
    const subFolders = patientFolder.getFoldersByName("Docs");
    if (!subFolders.hasNext()) return []; 
    
    const docsFolder = subFolders.next();
    const files = docsFolder.getFiles();
    
    let fileList = [];
    while (files.hasNext()) {
      const f = files.next();
      // Lógica de Persistência: Lê a descrição do arquivo no Drive
      const desc = f.getDescription() || "";
      const isVerified = desc.includes("[VERIFICADO]");
      
      fileList.push({ 
        name: f.getName(), 
        url: f.getUrl(),
        id: f.getId(), // Precisamos do ID para marcar depois
        verified: isVerified // Manda true ou false pro front-end
      });
    }
    return fileList;
    
  } catch (e) {
    return [];
  }
}

// 4. (NOVA) Salvar o Check individual do arquivo
function toggleDocStatus(fileId, isChecked) {
  try {
    const file = DriveApp.getFileById(fileId);
    if (isChecked) {
      file.setDescription("[VERIFICADO]"); // Marca como verificado
    } else {
      file.setDescription(""); // Remove a marca
    }
    return { success: true };
  } catch (e) {
    return { success: false };
  }
}

// 3. LIBERAR DOCUMENTAÇÃO (Desbloqueia a pasta e atualiza planilha)
function unlockPatientFolder(prontuario, nomePaciente, sheetName, row, formaEnvio) {
  try {
    const rootFolder = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
    const folderName = `${prontuario}- ${nomePaciente.toUpperCase()}`.trim();
    const folders = rootFolder.getFoldersByName(folderName);
    
    if (!folders.hasNext()) throw new Error("Pasta não encontrada.");
    const targetFolder = folders.next();
    
    // 1. Libera permissão no Drive
    targetFolder.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
    
    // 2. Atualiza Status na Planilha (Inteligente)
    let novaObs = "";
    if (sheetName && row) {
       // Passamos false no 'isRestricted' pois agora a pasta ESTÁ liberada
       novaObs = atualizarStatusInteligente_(sheetName, row, targetFolder, true, formaEnvio);
    }
    
    return { success: true, newObs: novaObs };
  } catch (e) {
    throw new Error("Falha ao liberar: " + e.toString());
  }
}

/*******************************************************
 * FUNÇÕES DE CONFIGURAÇÃO (ADMIN)
 *******************************************************/


function adminGetSettings() {
  const email = Session.getActiveUser().getEmail().toLowerCase();
  if (!SUPER_ADMINS.includes(email)) throw new Error("Acesso negado.");


  const props = PropertiesService.getScriptProperties();
  
  let userMap = props.getProperty('USER_MAP');
  userMap = userMap ? JSON.parse(userMap) : DEFAULT_USER_MAP;


  let maintenance = props.getProperty('MAINTENANCE_MODE');
  maintenance = (maintenance === 'true');


  return { userMap, maintenance };
}


function adminSaveSettings(payload) {
  const email = Session.getActiveUser().getEmail().toLowerCase();
  if (!SUPER_ADMINS.includes(email)) throw new Error("Acesso negado.");


  const props = PropertiesService.getScriptProperties();
  
  if (payload.userMap) {
    props.setProperty('USER_MAP', JSON.stringify(payload.userMap));
  }
  
  props.setProperty('MAINTENANCE_MODE', String(payload.maintenance));


  return { success: true };
}
