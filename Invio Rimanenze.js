// INTRO RIMANENZE
const INTRO_RIMANENZE = 
  "Ciao [COLLABORATORE],<br>" +
  "ti invio l’elenco degli artisti che hanno ancora importi in sospeso relativi al progetto in oggetto.<br>" +
  "Ti chiedo di contattarli per un sollecito.<br>" +
  "Grazie per la collaborazione.";

function inviaEmailRimanenze() {
  if (typeof confermaInvio === "function" && !confermaInvio("Rimanenze")) return;
  inviaEmailFoglioRimanenze(INTRO_RIMANENZE, "Elenco Rimanenze");
}

function inviaEmailFoglioRimanenze(introTesto, oggettoBase) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Rimanenze");
  const shColl = ss.getSheetByName("Collaboratori");

  if (!sh || !shColl) { 
    SpreadsheetApp.getUi().alert("Foglio Rimanenze o Collaboratori non trovato"); 
    return; 
  }

  // --- 1. MAPPATURA COLLABORATORI (Dinamica) ---
  const hColl = shColl.getRange(1, 1, 1, shColl.getLastColumn()).getValues()[0];
  const idxCollNome = hColl.indexOf('ID_COLLABORATORE');
  const idxCollMail = hColl.indexOf('ID_COLLABORATOREMAIL');

  if (idxCollNome === -1 || idxCollMail === -1) {
    throw new Error("ID_COLLABORATORE o ID_COLLABORATOREMAIL non trovati in riga 1 di Collaboratori");
  }

  const collData = shColl.getRange(2, 1, shColl.getLastRow() - 1, shColl.getLastColumn()).getValues();
  const collMap = {};
  collData.forEach(r => {
    const nome = r[idxCollNome];
    const email = r[idxCollMail];
    if (nome && email) collMap[nome.toString().trim().toLowerCase()] = email;
  });

  // --- 2. MAPPATURA E LETTURA RIMANENZE (Dinamica da riga 3) ---
  const hRim = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idxRimProgetto = hRim.indexOf('ID_PROGETTO');
  const idxRimCuratore = hRim.indexOf('ID_CURATORE');
  const idxRimNome = hRim.indexOf('ID_NOME');
  const idxRimRestante = hRim.indexOf('ID_RESTANTE');

  const startRowRim = 3;
  const lastRowRim = sh.getLastRow();
  if (lastRowRim < startRowRim) { 
    SpreadsheetApp.getUi().alert("Nessun dato presente in Rimanenze (riga 3+)"); 
    return; 
  }

  const dataRim = sh.getRange(startRowRim, 1, lastRowRim - startRowRim + 1, sh.getLastColumn()).getValues();

  // --- 3. RAGGRUPPAMENTO PER COLLABORATORE ---
  const gruppi = {};
  dataRim.forEach(r => {
    const collab = r[idxRimCuratore];
    if (!collab) return;
    const collabKey = collab.toString().trim().toLowerCase();

    if (!gruppi[collabKey]) gruppi[collabKey] = { rows: [], displayName: collab };
    gruppi[collabKey].rows.push(r);
  });

  // --- 4. INVIO EMAIL ---
  let emailCount = 0;

  for (const collKey in gruppi) {
    const rows = gruppi[collKey].rows;
    const displayName = gruppi[collKey].displayName;

    // Gestione TEST_MODE definita in Config.gs
    const emailDest = TEST_MODE ? MIO_TEST_EMAIL : collMap[collKey];
    if (!emailDest) {
      Logger.log("Email non trovata per: " + displayName);
      continue;
    }

    const oggetto = (TEST_MODE ? "[TEST] " : "") + oggettoBase + " – " + displayName;
    
    // Passiamo gli indici corretti alla funzione tabella
    const tableHtml = creaTabellaRimanenzeHTML(rows, {
      progetto: idxRimProgetto,
      nome: idxRimNome,
      restante: idxRimRestante
    });

    const body = introTesto.replace("[COLLABORATORE]", displayName) + "<br><br>" + tableHtml +
      '<br><br><small>Messaggio generato automaticamente dal nuovo Gestionale.</small>';

    const ccList = TEST_MODE ? "" : BOSS_EMAIL;

    GmailApp.sendEmail(emailDest, oggetto, "", { htmlBody: body, cc: ccList });
    emailCount++;
    
    // Se siamo in TEST_MODE, evitiamo di inondarti di mail se ci sono 50 collaboratori
    if (TEST_MODE && emailCount >= 1) {
      Logger.log("TEST_MODE attivo: inviata solo la prima mail di prova.");
      break; 
    }
  }

  SpreadsheetApp.getUi().alert("✅ Processo completato.\nEmail inviate: " + emailCount);
}

function creaTabellaRimanenzeHTML(rows, idx) {
  const headers = ["Progetto", "Artista", "Importo Restante"];
  let html = '<table border="1" cellspacing="0" cellpadding="5" style="border-collapse: collapse;">';
  html += '<tr style="background-color: #f2f2f2;">';
  headers.forEach(h => html += "<th>" + h + "</th>");
  html += "</tr>";

  rows.forEach(r => {
    html += "<tr>";
    html += `<td>${r[idx.progetto]}</td>`;
    html += `<td>${r[idx.nome]}</td>`;
    html += `<td style="text-align: right;">${r[idx.restante]}</td>`;
    html += "</tr>";
  });

  html += "</table>";
  return html;
}