/**
 * MENU PERSONALIZZATO UNIFICATO - VERSIONE AGNOSTICA (T1, T2, T3)
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📋 Gestionale")
    .addItem("⬇️ Inserisci nuovo Artista", "mostraForm")
    .addItem("🗑️ Elimina Artista selezionato", "eliminaArtistaSicuro")
    .addSeparator()
    .addItem("📧 1. Elabora e Invia al Reparto", "elaboraEInviaSpedizioni") 
    .addItem("📦 2. Crea Excel per DHL", "creaDistintaDHL_soloValori")
    .addItem('✨ 3. Miracolo DHL', 'ripristinaDHLDaSpediti')
    .addSeparator()
    .addItem("⏳ 4. Estrai Rimanenze", "aggiornaRimanenzeDaDeposito")
    .addItem("🚀 5. Invia Mail Collaboratori", "inviaEmailRimanenze")
    .addToUi();
}

/**
 * FUNZIONE DI UTILITÀ: Reset Orario
 */
function resetToMidnight(date) {
  if (!date || isNaN(date.getTime())) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
}

function mostraForm() {
  const html = HtmlService.createHtmlOutputFromFile("formCliente")
    .setWidth(450)
    .setHeight(680); 
  SpreadsheetApp.getUi().showModalDialog(html, "Inserimento Anagrafica");
}

/**
 * 1. CONTROLLO DUPLICATI LOCALE (Foglio Contabilita)
 */
function verificaDuplicato(nome, cf) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Contabilita");
  if (!sh) return null;

  const dati = sh.getDataRange().getValues();
  const mappa = getMapID(sh);
  
  const colNome = mappa["ID_NOME"] - 1;
  const colCf = mappa["ID_CF"] - 1;
  const colCuratore = mappa["ID_CURATORE"] - 1;

  if (colNome < 0 || colCuratore < 0) return null;

  const nomeRicerca = nome ? nome.toString().trim().toUpperCase() : "";
  const cfRicerca = cf ? cf.toString().trim().toUpperCase() : "";

  for (let i = 6; i < dati.length; i++) {
    let cfEsistente = (dati[i][colCf] || "").toString().trim().toUpperCase();
    let nomeEsistente = (dati[i][colNome] || "").toString().trim().toUpperCase();
    
    if (cfRicerca && cfRicerca === cfEsistente && cfEsistente !== "") {
      return dati[i][colCuratore] || "Nessun collaboratore segnato";
    }
    
    if (nomeRicerca && nomeRicerca === nomeEsistente) {
      return dati[i][colCuratore] || "Nessun collaboratore segnato";
    }
  }
  return null;
}

/**
 * 2. CONTROLLO FILE ESTERNO EU26
 */
function verificaEU26(nome) {
  if (!nome) return null;
  const idEU26 = "1BWz0AYSw04Vaxn76t3axWhIUQJB74GS21QJxibwuWSU";
  const nomeFoglio = "Foglio1";
  try {
    const ssEsterno = SpreadsheetApp.openById(idEU26);
    const shEsterno = ssEsterno.getSheetByName(nomeFoglio);
    const dati = shEsterno.getDataRange().getValues();
    const headers = dati[0].map(h => h.toString().trim().toUpperCase());
    const colNome = headers.indexOf("ID_NOME");
    const colCuratore = headers.indexOf("ID_CURATORE");
    if (colNome === -1 || colCuratore === -1) return null;
    const nomeRicerca = nome.toString().trim().toUpperCase();
    for (let i = 1; i < dati.length; i++) {
      let nomeEsterno = (dati[i][colNome] || "").toString().trim().toUpperCase();
      if (nomeRicerca === nomeEsterno) {
        return dati[i][colCuratore] || "Dato non specificato";
      }
    }
  } catch (e) {
    console.log("Errore accesso EU26: " + e.message);
  }
  return null;
}

/**
 * REGISTRAZIONE INIZIALE: Scrive in Contabilità e Deposito_T1
 */
function inserisciInContabilita(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shCont = ss.getSheetByName("Contabilita");
  const shDep = ss.getSheetByName("Deposito_T1");
  
  const idUnivoco = "ID-" + new Date().getTime();
  d.idRow = idUnivoco;

  const colsDaCopiareCont = ["ID_VERSATO", "ID_RESTANTE", "ID_ULTIMADONAZIONE", "ID_SCADENZA30G", "ID_STATO", "ID_OMAGGIO", "ID_SPESEPAGATE1", "ID_RESTANTE1", "ID_RESTANTE2", "ID_RESTANTE3", "ID_RESTANTE4"];
  const colsDaCopiareDep = ["ID_VERSATO", "ID_RESTANTE", "ID_STATOSPEDIZIONE", "ID_MODPAG", "ID_DOVUTO", "ID_SPEDITO", "ID_MATERIALI", "ID_OMAGGIO"];

  processaFoglioVeloce(shCont, d, colsDaCopiareCont);
  processaFoglioVeloce(shDep, d, colsDaCopiareDep);

  // LOGICA 650€: Se la Prima Donazione è 650, forza il salto in tutti i depositi
  if (parseFloat(d.rata1) === 650) {
    copiaArtistaInTappaSuccessiva(ss, idUnivoco, "Deposito_T2");
    copiaArtistaInTappaSuccessiva(ss, idUnivoco, "Deposito_T3");
    copiaArtistaInTappaSuccessiva(ss, idUnivoco, "Deposito_T4");
  } 
  // AVANZAMENTO STANDARD: Se è stata inserita una Prima Donazione (> 0), copia in T2
  else if (parseFloat(d.rata1) > 0) {
    copiaArtistaInTappaSuccessiva(ss, idUnivoco, "Deposito_T2");
  }

  try {
    aggiungiNuovoCliente(d);
  } catch(e) {
    console.log("Errore DB: " + e.message);
  }
}

function processaFoglioVeloce(sheet, datiInput, colonneDaCopiare) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rigaTarget = getPrimaRigaLibera(sheet);
  const rigaSorgente = sheet.getRange(rigaTarget - 1, 1, 1, lastCol);
  
  const srcFormulas = rigaSorgente.getFormulasR1C1()[0];
  const srcValues = rigaSorgente.getValues()[0];
  const srcFormats = rigaSorgente.getNumberFormats();
  const srcValidations = rigaSorgente.getDataValidations();
  
  let nuovaRigaValues = [];
  let finalFormats = srcFormats;

  for (let i = 0; i < lastCol; i++) {
    let rawHeader = headers[i] ? headers[i].toString() : "";
    let key = rawHeader.trim().toUpperCase().replace(/[\u200B-\u200D\uFEFF]/g, "");
    
    if (key === "ID_CAP") finalFormats[0][i] = "@"; 
    
    let valoreInput = mapInputData(key, datiInput);
    
    if (valoreInput !== null) {
      nuovaRigaValues.push(valoreInput);
    } 
    else if (key === "ID_SPESEPAGATE1" && sheet.getName() === "Contabilita") {
      nuovaRigaValues.push(datiInput.speseGiaPagate ? "SI" : "NO");
    }
    else if (key === "ID_NOTE" && sheet.getName() === "Contabilita" && parseFloat(datiInput.rata1) === 650) {
      nuovaRigaValues.push("Unica Soluzione");
    }
    else if (key === "ID_MATERIALI" && sheet.getName().startsWith("Deposito_T")) {
      nuovaRigaValues.push("Book");
    }
    else if (key === "ID_SPESE" && sheet.getName().startsWith("Deposito_T")) {
      nuovaRigaValues.push(datiInput.speseGiaPagate ? 0 : 25);
    }
    else if (colonneDaCopiare.includes(key)) {
      nuovaRigaValues.push(srcFormulas[i] !== "" ? srcFormulas[i] : srcValues[i]);
    }
    else {
      nuovaRigaValues.push(""); 
    }
  }
  
  const rangeDestinazione = sheet.getRange(rigaTarget, 1, 1, lastCol);
  rangeDestinazione.setValues([nuovaRigaValues]);
  rangeDestinazione.setNumberFormats(finalFormats); 
  rangeDestinazione.setDataValidations(srcValidations);
}

/**
 * MAPPA DATI AGGIORNATA CON LOGICA 650€ SU PRIMA DONAZIONE
 */
function mapInputData(key, d) {
  const is650 = (parseFloat(d.rata1) === 650);

  switch (key) {
    case "ID_ROW": return d.idRow;
    case "ID_NOME": return d.nomeCompleto;
    case "ID_NARTE": return d.nomeArte;
    case "ID_CF": return d.cf;
    case "ID_INDIRIZZO": return d.indirizzo;
    case "ID_CITTA": return d.citta;
    case "ID_PV": return d.provincia;
    case "ID_CELL": return d.telefono;
    case "ID_MAIL": return d.email;
    case "ID_CURATORE": return d.collaboratore;
    case "ID_PROGETTO": return d.progetto;
    case "ID_DATAINSERIMENTO": return resetToMidnight(d.data ? new Date(d.data) : new Date());
    case "ID_IMPORTOPATTUITO": return d.importo;
    
    case "ID_RATA1": return is650 ? 140 : d.rata1;
    case "ID_RATA2": return is650 ? 170 : "";
    case "ID_RATA3": return is650 ? 170 : "";
    case "ID_RATA4": return is650 ? 170 : "";

    case "ID_RATA1DATA": return d.rata1data ? resetToMidnight(new Date(d.rata1data)) : "";
    case "ID_CAP": 
      if (!d.cap) return "";
      let capStr = d.cap.toString().trim();
      return capStr.length < 5 ? capStr.padStart(5, "0") : capStr; 
    default: return null;
  }
}

function getPrimaRigaLibera(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 7) return 7;
  const startCheck = Math.max(7, lastRow - 50);
  const range = sheet.getRange(startCheck, 1, lastRow - startCheck + 5, 3).getValues();
  for (let i = 0; i < range.length; i++) {
    if (!range[i][0] && !range[i][1] && !range[i][2]) return startCheck + i;
  }
  return lastRow + 1;
}

function getMapID(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const riga1 = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let mappa = {};
  riga1.forEach((id, index) => {
    if (id) {
      let chiave = id.toString().trim().toUpperCase().replace(/[\u200B-\u200D\uFEFF]/g, "");
      mappa[chiave] = index + 1;
    }
  });
  return mappa;
}

function aggiungiNuovoCliente(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbNuovo = ss.getSheetByName("Nuovo_Database");
  const dbRicevute = ss.getSheetByName("Database_Ricevute");
  
  if (!d.cf) return;
  const cfCercato = d.cf.toString().trim().toUpperCase();

  // --- LOGICA 1: SCRITTURA SU NUOVO_DATABASE (STRUTTURA STANDARD) ---
  if (dbNuovo) {
    const mapDbNuovo = getMapID(dbNuovo);
    const lastRowNuovo = dbNuovo.getLastRow();
    const colCfNuovo = mapDbNuovo["ID_CF"];
    let rigaTrovataNuovo = -1;

    if (lastRowNuovo >= 2 && colCfNuovo) {
      const cfListNuovo = dbNuovo.getRange(2, colCfNuovo, lastRowNuovo - 1, 1).getValues().flat();
      for (let i = 0; i < cfListNuovo.length; i++) {
        if (cfListNuovo[i].toString().trim().toUpperCase() === cfCercato) {
          rigaTrovataNuovo = i + 2;
          break;
        }
      }
    }

    const lastColNuovo = dbNuovo.getLastColumn();
    if (lastColNuovo > 0) {
      const headersNuovo = dbNuovo.getRange(1, 1, 1, lastColNuovo).getValues()[0];
      let rigaValoriNuovo = [];
      headersNuovo.forEach(h => {
        let key = h.toString().trim().toUpperCase();
        let val = mapInputData(key, d);
        rigaValoriNuovo.push(val !== null ? val : "");
      });

      if (rigaTrovataNuovo !== -1) {
        dbNuovo.getRange(rigaTrovataNuovo, 1, 1, lastColNuovo).setValues([rigaValoriNuovo]);
      } else {
        const rigaLiberaNuovo = getPrimaRigaLibera(dbNuovo);
        dbNuovo.getRange(rigaLiberaNuovo, 1, 1, lastColNuovo).setValues([rigaValoriNuovo]);
      }
    }
  }

  // --- LOGICA 2: SCRITTURA SU DATABASE_RICEVUTE (STRUTTURA COMPATTATA) ---
  if (dbRicevute) {
    const mapDbRicevute = getMapID(dbRicevute);
    const lastRowRicevute = dbRicevute.getLastRow();
    const colCfRicevute = mapDbRicevute["ID_CF"];
    let rigaTrovataRicevute = -1;

    if (lastRowRicevute >= 2 && colCfRicevute) {
      const cfListRicevute = dbRicevute.getRange(2, colCfRicevute, lastRowRicevute - 1, 1).getValues().flat();
      for (let i = 0; i < cfListRicevute.length; i++) {
        if (cfListRicevute[i].toString().trim().toUpperCase() === cfCercato) {
          rigaTrovataRicevute = i + 2;
          break;
        }
      }
    }

    const lastColRicevute = dbRicevute.getLastColumn();
    if (lastColRicevute > 0) {
      const headersRicevute = dbRicevute.getRange(1, 1, 1, lastColRicevute).getValues()[0];
      let rigaValoriRicevute = [];

      // Generazione stringa unificata ID_CCP: Cap Città Provincia con blindatura zeri iniziali
      let strCap = d.cap ? d.cap.toString().trim() : "";
      if (strCap && strCap.length < 5) {
        strCap = strCap.padStart(5, "0");
      }
      const strCitta = d.citta ? d.citta.toString().trim() : "";
      const strPv = d.provincia ? d.provincia.toString().trim() : "";
      const stringaCcp = [strCap, strCitta, strPv].filter(String).join(" ");

      headersRicevute.forEach(h => {
        let key = h.toString().trim().toUpperCase();
        if (key === "ID_NOME") {
          rigaValoriRicevute.push(d.nomeCompleto || "");
        } else if (key === "ID_INDIRIZZO") {
          rigaValoriRicevute.push(d.indirizzo || "");
        } else if (key === "ID_CF") {
          rigaValoriRicevute.push(d.cf || "");
        } else if (key === "ID_CCP") {
          rigaValoriRicevute.push(stringaCcp);
        } else {
          let val = mapInputData(key, d);
          rigaValoriRicevute.push(val !== null ? val : "");
        }
      });

      if (rigaTrovataRicevute !== -1) {
        dbRicevute.getRange(rigaTrovataRicevute, 1, 1, lastColRicevute).setValues([rigaValoriRicevute]);
      } else {
        const rigaLiberaRicevute = getPrimaRigaLibera(dbRicevute);
        dbRicevute.getRange(rigaLiberaRicevute, 1, 1, lastColRicevute).setValues([rigaValoriRicevute]);
      }
    }
  }
}

function inserisciRigaDb(db, d) {
  // Funzione deprecata ma lasciata per consistenza strutturale dello script originale
}

function eliminaArtistaSicuro() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const rigaAttiva = sheet.getActiveCell().getRow();
  const mappa = getMapID(sheet);
  
  if (!mappa["ID_ROW"]) return;
  const idDaCancellare = sheet.getRange(rigaAttiva, mappa["ID_ROW"]).getValue();
  if (!idDaCancellare || rigaAttiva < 7) return;

  const risposta = ui.alert("⚠ CONFERMA", "Eliminare definitivamente l'artista da TUTTI i fogli (Contabilità e Depositi)?", ui.ButtonSet.YES_NO);
  if (risposta == ui.Button.YES) {
    const fogliDaPulire = ["Contabilita", "Deposito_T1", "Deposito_T2", "Deposito_T3", "Deposito_T4"]; 
    fogliDaPulire.forEach(nome => {
      const sh = ss.getSheetByName(nome);
      if (sh) {
        const m = getMapID(sh);
        if (m["ID_ROW"]) {
          const dataIds = sh.getRange(1, m["ID_ROW"], sh.getLastRow(), 1).getValues().flat();
          for (let i = dataIds.length - 1; i >= 0; i--) { 
            if (dataIds[i] === idDaCancellare) sh.deleteRow(i + 1); 
          }
        }
      }
    });
    ss.toast("Artista eliminato ovunque con successo. 🗑️");
  }
}

function getClientiEsistenti() {
  const db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database");
  if (!db) return [];
  const map = getMapID(db);
  return db.getRange(2, map["ID_NOME"] || 1, Math.max(db.getLastRow(), 1), 1).getValues().flat().filter(String);
}

function getCollaboratori() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Collaboratori");
  return (sheet && sheet.getLastRow() > 1) ? sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat().filter(String) : [];
}

function getProgetti() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Progetti");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const dati = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return dati.filter(r => r[0] !== "").map(r => {
    return {
      nome: r[0].toString(),
      prezzo: r[1] ? parseFloat(r[1]) : 0
    };
  });
}

function getDatiCliente(nomeCompleto) {
  const db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database");
  if (!db) return null;
  const data = db.getDataRange().getValues();
  const headers = data[0].map(h=>h.toString().trim().toUpperCase());
  const row = data.find(r => r[headers.indexOf("ID_NOME")] === nomeCompleto);
  if (!row) return null;
  const getVal = (id) => row[headers.indexOf(id)] || "";
  return {
    nomeArte: getVal("ID_NARTE"), cf: getVal("ID_CF"), indirizzo: getVal("ID_INDIRIZZO"),
    citta: getVal("ID_CITTA"), provincia: getVal("ID_PV"), cap: getVal("ID_CAP"),
    telefono: getVal("ID_CELL"), email: getVal("ID_MAIL")
  };
}