/**
 * =================================================================
 * MODULO LOGISTICA UNIFICATO (T1, T2, T3, T4) - VERSIONE INTEGRALE
 * =================================================================
 */

/**
 * FUNZIONE 1: ELABORA, INVIA EMAIL E AGGIORNA STATO
 */
function elaboraEInviaSpedizioni() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shAttivo = ss.getActiveSheet();
  const nomeFoglioAttivo = shAttivo.getName();
  
  if (!nomeFoglioAttivo.startsWith("Deposito_T")) {
    SpreadsheetApp.getUi().alert("Errore: Lancia questa funzione da un foglio Deposito (T1, T2, T3 o T4).");
    return;
  }

  const tappaCorrente = nomeFoglioAttivo.split("_")[1]; 
  const shDhl = ss.getSheetByName("DHL");
  const shCont = ss.getSheetByName("Contabilita"); 
  const ui = SpreadsheetApp.getUi();
  
  if (!shDhl || !shCont) {
    ui.alert("Errore: Assicurati che i fogli 'DHL' e 'Contabilita' esistano.");
    return;
  }

  const conferma = ui.alert("Conferma invio [" + nomeFoglioAttivo + "]", "Procedere con l'estrazione DHL e l'invio delle email?", ui.ButtonSet.YES_NO);
  if (conferma !== ui.Button.YES) return;

  const col = getMapID(shAttivo);
  const colCont = getMapID(shCont);

  const dati = shAttivo.getDataRange().getValues();
  const datiCont = shCont.getDataRange().getValues(); 
  const oggiPuro = new Date().setHours(0,0,0,0);

  const salutoCreativo = getSalutoCasuale();
  let corpoMail = "Buongiorno " + salutoCreativo + ", ecco le spedizioni (" + tappaCorrente + ") da effettuare oggi (" + Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy") + "):<br><br><br>";
  
  let contatore = 0;
  let righeDaProcessare = []; 

  for (let i = 1; i < dati.length; i++) {
    const riga = dati[i];
    const stato = (riga[col["ID_STATOSPEDIZIONE"] - 1] || "").toString().trim();
    const giaSpedito = (riga[col["ID_SPEDITO"] - 1] || "").toString().trim().toUpperCase() === "SI";
    
    let dataVal = riga[col["ID_DATASPEDIZIONE"] - 1];
    let dataCellaPura = (dataVal instanceof Date) ? dataVal.setHours(0,0,0,0) : 0;

    if ((stato === "Da spedire" || stato === "Posticipata") && !giaSpedito && dataCellaPura !== 0 && dataCellaPura <= oggiPuro) {
      
      contatore++;
      const nome = riga[col["ID_NOME"] - 1];
      const presso = (riga[col["ID_PRESSO"] - 1] || "").toString().trim();
      const indirizzo = riga[col["ID_INDIRIZZO"] - 1];
      const cap = (riga[col["ID_CAP"] - 1] || "").toString();
      const citta = riga[col["ID_CITTA"] - 1];
      const prov = riga[col["ID_PV"] - 1];
      const tel = riga[col["ID_CELL"] - 1];
      const email = riga[col["ID_EMAIL"] - 1] || "nomail@gmail.com";
      const materiale = (riga[col["ID_MATERIALI"] - 1] || "").toString().trim();
      const idRowVal = riga[col["ID_ROW"] - 1]; 
      const modPagVal = (riga[col["ID_MODPAG"] - 1] || "").toString().trim(); 
      const noteSpedizione = (riga[col["ID_NOTE_SPEDIZIONE"] - 1] || "").toString().trim();
      const importo = riga[col["ID_DOVUTO"] - 1] || 0;
      const restanteVal = riga[col["ID_RESTANTE"] - 1] || 0; 
      const testoImporto = (importo === 0 || importo === "0") ? "Gratis" : importo;

      let rigaDhl = shDhl.getLastRow() + 1;
      
      shDhl.getRange(rigaDhl, 1).setValue(nome);          
      shDhl.getRange(rigaDhl, 2).setValue(presso !== "" ? "c/o " + presso : ""); 
      shDhl.getRange(rigaDhl, 3).setValue(indirizzo);     
      shDhl.getRange(rigaDhl, 5).setValue(citta);         
      shDhl.getRange(rigaDhl, 7).setNumberFormat("@").setValue(cap); 
      shDhl.getRange(rigaDhl, 6).setValue(prov);          
      shDhl.getRange(rigaDhl, 9).setValue(tel);           
      shDhl.getRange(rigaDhl, 10).setValue(email);        
      shDhl.getRange(rigaDhl, 11).setValue(getNextOrderNumber()); 
      shDhl.getRange(rigaDhl, 13).setValue(importo);      
      shDhl.getRange(rigaDhl, 17).setValue(importo);      
      shDhl.getRange(rigaDhl, 14).setValue(1);            
      shDhl.getRange(rigaDhl, 15).setValue(materiale);    

      shDhl.getRange(2, 18, 1, 13).copyTo(shDhl.getRange(rigaDhl, 18, 1, 13));

      let infoDest = presso !== "" ? "c/o " + presso + "<br>" + indirizzo : indirizzo;
      corpoMail += "<b>" + nome + "</b><br>" + infoDest + "<br>" + citta + " " + cap + " " + prov + "<br>Importo: " + testoImporto + "<br>";
      if (materiale !== "") corpoMail += "Materiale: " + materiale + "<br>";
      if (noteSpedizione !== "") corpoMail += "Note: " + noteSpedizione + "<br>";
      corpoMail += "<br>---<br><br>"; 

      righeDaProcessare.push({ rigaDeposito: i + 1, idRow: idRowVal, modPag: modPagVal, quota: restanteVal });
    }
  }

  if (contatore === 0) {
    ui.alert("Nessuna spedizione pronta per oggi nel foglio " + nomeFoglioAttivo);
    return;
  }

  MailApp.sendEmail({
    to: "effettoartefondazione.amm@gmail.com",
    subject: "📦 Preparazione Spedizioni [" + tappaCorrente + "] - " + Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy"),
    htmlBody: corpoMail
  });

  righeDaProcessare.forEach(item => {
    shAttivo.getRange(item.rigaDeposito, col["ID_SPEDITO"]).setValue("SI");
    if (item.modPag === "Contrassegno") {
      const numTappa = tappaCorrente.replace("T", "");
      const colSpedContrTappa = "ID_SPDCNTR" + numTappa; 
      const colRataCont = "ID_RATA" + numTappa;
      const prossimaTappaNum = parseInt(numTappa) + 1;
      const nomeProssimaTappa = "Deposito_T" + prossimaTappaNum;

      for (let j = 1; j < datiCont.length; j++) {
        if (datiCont[j][colCont["ID_ROW"] - 1] == item.idRow) {
          if (colCont[colRataCont]) {
            shCont.getRange(j + 1, colCont[colRataCont]).setValue(item.quota);
            
            // ATTIVAZIONE CASCATA MANUALE (perché setValue non attiva onEdit)
            if (prossimaTappaNum <= 4) {
              // Richiama la funzione di copia definita nel modulo Automazioni
              if (typeof copiaArtistaInTappaSuccessiva === 'function') {
                copiaArtistaInTappaSuccessiva(ss, item.idRow, nomeProssimaTappa);
              }
            }
          }
          if (colCont[colSpedContrTappa]) {
            shCont.getRange(j + 1, colCont[colSpedContrTappa]).setValue("Spd Cntr");
          }
          break; 
        }
      }
    }
  });

  shAttivo.getRange("D2").setValue("Ultimo invio " + tappaCorrente + ": " + contatore + " spedizioni il " + Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy HH:mm"));
  ui.alert("Operazione completata con successo.");
}

/**
 * FUNZIONE 2: CREA DISTINTA EXCEL
 */
function creaDistintaDHL_soloValori() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('DHL');
  if (!sh) return;
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 3) {
    ui.alert('Non ci sono dati da esportare.');
    return;
  }

  const values = [sh.getRange(1, 1, 1, lastCol).getValues()[0]].concat(sh.getRange(3, 1, lastRow - 2, lastCol).getValues());

  const todayStr = Utilities.formatDate(new Date(), 'GMT+1', 'yyyy-MM-dd');
  const tempSs = SpreadsheetApp.create('Caricamento_DHL_' + todayStr);
  const tempSh = tempSs.getSheets()[0];
  tempSh.getRange(1, 1, values.length, values[0].length).setValues(values);
  tempSh.getDataRange().setNumberFormat("@");

  const url = 'https://docs.google.com/spreadsheets/d/' + tempSs.getId() + '/export?format=xlsx';
  const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob().setName('DHL_' + todayStr + '.xlsx');
  const file = DriveApp.createFile(blob);
  DriveApp.getFileById(tempSs.getId()).setTrashed(true);

  const downloadUrl = file.getUrl().replace(/\/view\?usp=drivesdk/, '/export?format=xlsx');
  const html = HtmlService.createHtmlOutput('<div style="text-align:center;padding:15px;"><a href="'+downloadUrl+'" target="_blank" style="padding:10px;background:#4CAF50;color:white;text-decoration:none;border-radius:5px;">📥 SCARICA FILE EXCEL</a></div>').setWidth(300).setHeight(100);
  ui.showModalDialog(html, 'Download Excel');
  
  if (ui.alert('Export completato', 'Vuoi pulire il foglio DHL?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
    sh.deleteRows(3, sh.getLastRow() - 2);
  }
}

/**
 * FUNZIONE 3: RIPRISTINA DHL DA "SI"
 */
function ripristinaDHLDaSpediti() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shAttivo = ss.getActiveSheet();
  const shDhl = ss.getSheetByName("DHL");
  const ui = SpreadsheetApp.getUi();
  
  if (!shAttivo.getName().startsWith("Deposito_T") || !shDhl) return;

  const col = getMapID(shAttivo);
  const dati = shAttivo.getDataRange().getValues();
  let contatore = 0;

  for (let i = 1; i < dati.length; i++) {
    if ((dati[i][col["ID_SPEDITO"] - 1] || "").toString().toUpperCase() === "SI") {
      contatore++;
      const r = dati[i];
      let rDhl = shDhl.getLastRow() + 1;
      shDhl.getRange(rDhl, 1).setValue(r[col["ID_NOME"]-1]);
      shDhl.getRange(rDhl, 3).setValue(r[col["ID_INDIRIZZO"]-1]);
      shDhl.getRange(rDhl, 5).setValue(r[col["ID_CITTA"]-1]);
      shDhl.getRange(rDhl, 7).setNumberFormat("@").setValue(r[col["ID_CAP"]-1].toString());
      shDhl.getRange(rDhl, 6).setValue(r[col["ID_PV"]-1]);
      shDhl.getRange(rDhl, 9).setValue(r[col["ID_CELL"]-1]);
      shDhl.getRange(rDhl, 13).setValue(r[col["ID_DOVUTO"]-1]);
      shDhl.getRange(rDhl, 17).setValue(r[col["ID_DOVUTO"]-1]);
      shDhl.getRange(2, 18, 1, 13).copyTo(shDhl.getRange(rDhl, 18, 1, 13));
    }
  }
  ui.alert("Recupero terminato. Aggiunti " + contatore + " elementi.");
}

/**
 * UTILITY
 */
function getMapID(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const mappa = {};
  headers.forEach((h, i) => { if (h) mappa[h.toString().trim().toUpperCase()] = i + 1; });
  return mappa;
}

function getSalutoCasuale() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Diario");
  if (!sh) return "ragazzi";
  const saluti = sh.getRange(2, 26, sh.getLastRow(), 1).getValues().flat().filter(String);
  return saluti.length > 0 ? saluti[Math.floor(Math.random() * saluti.length)] : "ragazzi";
}

function getNextOrderNumber() {
  const p = PropertiesService.getScriptProperties();
  let last = p.getProperty('lastOrderNumber') || '20260000';
  let n = parseInt(last) + 1;
  p.setProperty('lastOrderNumber', n.toString());
  return n;
}