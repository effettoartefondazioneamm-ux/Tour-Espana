/**
 * AUTOMAZIONI UNIFICATE - SISTEMA AGNOSTICO A CASCATA
 * Logica: T1 -> T2 -> T3 -> T4 basata su quote e pagamenti totali.
 */

function onEdit(e) {
  const ss = e.source;
  const sheet = ss.getActiveSheet();
  const range = e.range;
  const nomeFoglio = sheet.getName();
  const riga = range.getRow();
  const colModificata = range.getColumn();

  // PROTEZIONE RIGOROSA: Non tocca nulla sopra la riga 7
  if (riga < 7) return; 

  // Configurazione Fogli ammessi
  const fogliDeposito = ["Deposito_T1", "Deposito_T2", "Deposito_T3", "Deposito_T4"];
  if (!fogliDeposito.includes(nomeFoglio) && nomeFoglio !== "Contabilita") return;

  const mappaAttuale = getMapID(sheet);
  const cIDRow = mappaAttuale["ID_ROW"];
  if (!cIDRow) return;

  const idArtista = sheet.getRange(riga, cIDRow).getValue();
  if (!idArtista) return;

  // ===========================================================================
  // === AZIONI DA FOGLIO CONTABILITA ===
  // ===========================================================================
  if (nomeFoglio === "Contabilita") {
    
    // --- GESTIONE MANUALE 650€ ---
    const valoreInserito = e.value ? e.value.toString().trim() : "";
    if (colModificata === mappaAttuale["ID_IMPORTOPATTUITO"] && valoreInserito === "650") {
      
      // Suddivisione quote in Contabilità (140 + 170x3)
      sheet.getRange(riga, mappaAttuale["ID_RATA1"]).setValue(140);
      sheet.getRange(riga, mappaAttuale["ID_RATA2"]).setValue(170);
      sheet.getRange(riga, mappaAttuale["ID_RATA3"]).setValue(170);
      sheet.getRange(riga, mappaAttuale["ID_RATA4"]).setValue(170);
      
      // Attivazione Tour nei Depositi
      copiaArtistaInTappaSuccessiva(ss, idArtista, "Deposito_T2");
      copiaArtistaInTappaSuccessiva(ss, idArtista, "Deposito_T3");
      copiaArtistaInTappaSuccessiva(ss, idArtista, "Deposito_T4");
      
      ss.toast("Ripartizione 650€ e Tour T2-T4 completato.");
      return; 
    }

    // --- A. AZZERAMENTO SPESE (Trigger ID_SPESEPAGATE1, 2, 3, 4) ---
    const triggerSpese = {
      "ID_SPESEPAGATE1": "Deposito_T1",
      "ID_SPESEPAGATE2": "Deposito_T2",
      "ID_SPESEPAGATE3": "Deposito_T3",
      "ID_SPESEPAGATE4": "Deposito_T4"
    };

    for (let trigger in triggerSpese) {
      if (colModificata === mappaAttuale[trigger] && valoreInserito.toUpperCase() === "SI") {
        azzeraSpeseInDeposito(ss, idArtista, triggerSpese[trigger]);
      }
    }

    // --- B. TRIGGER PAGAMENTI ANTICIPATI DA CONTABILITÀ ---
    // CORRETTO: index 0 (Rata1) spinge a T2, index 1 (Rata2) spinge a T3, ecc.
    const rateContabilita = ["ID_RATA1", "ID_RATA2", "ID_RATA3", "ID_RATA4"];
    rateContabilita.forEach((rata, index) => {
      if (colModificata === mappaAttuale[rata]) {
        const cifraRata = Number(valoreInserito) || 0;
        if (cifraRata > 0) {
          const prossimaTappa = index + 2; 
          if (prossimaTappa <= 4) {
            const targetDeposito = "Deposito_T" + prossimaTappa;
            copiaArtistaInTappaSuccessiva(ss, idArtista, targetDeposito);
          }
        }
      }
    });
  }

  // ===========================================================================
  // === AZIONI COMUNI (Sincronizzazione Ritirato) ===
  // ===========================================================================
  const cNote = mappaAttuale["ID_NOTE"];
  const valoreNote = e.value ? e.value.toString().trim().toUpperCase() : "";
  if (colModificata === cNote && valoreNote === "RITIRATO") {
    sincronizzaRitiratoGlobale(ss, idArtista, nomeFoglio);
  }

  // ===========================================================================
  // === AZIONI DA FOGLI DEPOSITO (Automazione Spedizioni) ===
  // ===========================================================================
  if (fogliDeposito.includes(nomeFoglio)) {
    gestisciLogicaSpedizione(ss, sheet, riga, colModificata, mappaAttuale, idArtista, e.value);
  }
}

/**
 * Copia l'artista ereditando i dati aggiornati dal deposito precedente e imposta "Book" di default
 */
function copiaArtistaInTappaSuccessiva(ss, idArtista, foglioDest) {
  const shDest = ss.getSheetByName(foglioDest);
  if (!shDest) return;

  const mDest = getMapID(shDest);
  const lastRowDest = shDest.getLastRow();
  
  // Verifica se l'artista esiste già nel foglio destinazione
  if (lastRowDest >= 7) {
    const dataDest = shDest.getRange(7, mDest["ID_ROW"], lastRowDest - 6, 1).getValues().flat();
    if (dataDest.includes(idArtista)) return;
  }

  // Determina la sorgente dati: Priorità al foglio Deposito corrente, altrimenti Contabilità
  const shAttivo = ss.getActiveSheet();
  const nomeShAttivo = shAttivo.getName();
  let shSorgente, mSorgente, rigaDatiSorgente;

  if (nomeShAttivo.startsWith("Deposito_T")) {
    shSorgente = shAttivo;
    mSorgente = getMapID(shSorgente);
    const datiSorgente = shSorgente.getDataRange().getValues();
    rigaDatiSorgente = datiSorgente.find(r => r[mSorgente["ID_ROW"] - 1] == idArtista);
  }

  // Se non siamo in un deposito o l'artista non è trovato nella sorgente corrente, usiamo la Contabilità
  if (!rigaDatiSorgente) {
    shSorgente = ss.getSheetByName("Contabilita");
    if (!shSorgente) return;
    mSorgente = getMapID(shSorgente);
    const datiSorgente = shSorgente.getDataRange().getValues();
    rigaDatiSorgente = datiSorgente.find(r => r[mSorgente["ID_ROW"] - 1] == idArtista);
  }

  if (rigaDatiSorgente) {
    const lastCol = shDest.getLastColumn();
    const headersDest = shDest.getRange(1, 1, 1, lastCol).getValues()[0];
    const rigaLibera = getPrimaRigaLibera(shDest);
    
    // Prende il modello dalla riga 6 (formule e validazioni)
    const rigaModello = shDest.getRange(6, 1, 1, lastCol);
    const formuleModello = rigaModello.getFormulasR1C1()[0];
    const validazioniModello = rigaModello.getDataValidations()[0];
    const formatiModello = rigaModello.getNumberFormats();

    let nuovaRiga = new Array(lastCol).fill("");

    headersDest.forEach((h, idx) => {
      let key = h.toString().trim().toUpperCase();
      
      // BLINDATURA AGGIORNATA: Stato spedizione, logistica, DOVUTO, SPEDITOCUM e NOTE nascono puliti (pari al modello della riga 6)
      if (key === "ID_STATOSPEDIZIONE" || key === "ID_MODPAG" || key === "ID_SPEDITO" || key === "ID_DATASPEDIZIONE" || key === "ID_DOVUTO" || key === "ID_SPEDITOCUM" || key === "ID_NOTE") {
        nuovaRiga[idx] = formuleModello[idx] !== "" ? formuleModello[idx] : "";
      }
      // PRIORITÀ 1: Formule obbligatorie per logica contabile deposito
      else if (key === "ID_VERSATO" || key === "ID_RESTANTE") {
        nuovaRiga[idx] = formuleModello[idx];
      }
      // PRIORITÀ 2: Spese predefinite
      else if (key === "ID_SPESE") {
        nuovaRiga[idx] = 25; 
      } 
      // PRIORITÀ 3: Materiali predefiniti (Nuovo default: Book)
      else if (key === "ID_MATERIALI") {
        nuovaRiga[idx] = "Book";
      }
      // PRIORITÀ 4: Eredità dati aggiornati dalla sorgente (Deposito precedente o Contabilità)
      else if (mSorgente[key]) {
        nuovaRiga[idx] = rigaDatiSorgente[mSorgente[key] - 1]; 
      } 
      // PRIORITÀ 5: Altre formule presenti nel modello riga 6
      else if (formuleModello[idx] !== "") {
        nuovaRiga[idx] = formuleModello[idx]; 
      }
    });

    const rangeDest = shDest.getRange(rigaLibera, 1, 1, lastCol);
    rangeDest.setValues([nuovaRiga]);
    rangeDest.setNumberFormats(formatiModello);
    
    for (let v = 0; v < validazioniModello.length; v++) {
      if (validazioniModello[v]) rangeDest.getCell(1, v + 1).setDataValidation(validazioniModello[v]);
    }

    ss.toast("Artista aggiunto a " + foglioDest + " ✅");
  }
}

/**
 * Gestione Logica Spedizione, Ritorno in Contabilità e Spedizioni Cumulative
 */
function gestisciLogicaSpedizione(ss, sheet, riga, colModificata, mappa, idArtista, valore) {
  const nomeFoglio = sheet.getName();
  const cStato = mappa["ID_STATOSPEDIZIONE"];
  const cModPag = mappa["ID_MODPAG"];
  const cRestanteDep = mappa["ID_RESTANTE"];
  const cSpese = mappa["ID_SPESE"];
  const cDataSped = mappa["ID_DATASPEDIZIONE"];
  const cSpedito = mappa["ID_SPEDITO"];
  const cSpeditoCum = mappa["ID_SPEDITOCUM"];
  const cDovuto = mappa["ID_DOVUTO"];
  
  const oggi = new Date();
  oggi.setHours(0,0,0,0);

  const valTesto = valore ? valore.toString().trim() : "";

  if (colModificata === cStato) {
    if (valTesto === "Da spedire" || valTesto === "Posticipata") {
      sheet.getRange(riga, cModPag).setValue("Contrassegno");
      if (valTesto === "Da spedire") sheet.getRange(riga, cDataSped).setValue(oggi);
    }
  }

  // --- TRIGGER: GESTIONE SPEDIZIONE CUMULATIVA MANUALE ---
  if (colModificata === cSpeditoCum && valTesto.toUpperCase() === "SI") {
    if (cSpedito) sheet.getRange(riga, cSpedito).setValue("SI"); 
    
    const shCont = ss.getSheetByName("Contabilita");
    const mCont = getMapID(shCont);
    const importoRestante = Number(sheet.getRange(riga, cRestanteDep).getValue()) || 0;
    const importoDovuto = sheet.getRange(riga, cDovuto).getValue() || 0;
    const testoImporto = (importoDovuto === 0 || importoDovuto === "0") ? "Gratis" : importoDovuto;
    
    const datiCont = shCont.getDataRange().getValues();
    let rigaIdx = datiCont.findIndex(r => r[mCont["ID_ROW"] - 1] == idArtista);

    if (rigaIdx !== -1) {
      let rigaCont = rigaIdx + 1;
      const numTappa = nomeFoglio.replace("Deposito_T", "");
      const colRataTarget = mCont["ID_RATA" + numTappa];
      const colSpdCntrTarget = mCont["ID_SPDCNTR" + numTappa];
      const prossimaTappaNum = parseInt(numTappa) + 1;
      const nomeProssimaTappa = "Deposito_T" + prossimaTappaNum;

      // Scrittura Rata in Contabilità
      if (colRataTarget) {
        shCont.getRange(rigaCont, colRataTarget).setValue(importoRestante);
        const headerRata = shCont.getRange(1, colRataTarget).getValue();
        const colData = mCont[headerRata + "DATA"];
        if (colData) shCont.getRange(rigaCont, colData).setValue(oggi);
      }

      // Scrittura Stato Spedizione in Contabilità con Importo
      if (colSpdCntrTarget) {
        shCont.getRange(rigaCont, colSpdCntrTarget).setValue("Spd Cntr " + testoImporto);
      }

      // Attivazione Cascata Avanzamento Tappa
      if (prossimaTappaNum <= 4) {
        copiaArtistaInTappaSuccessiva(ss, idArtista, nomeProssimaTappa);
      }
    }
    return; 
  }

  // --- LOGICA DI RITORNO IN CONTABILITÀ AL PAGAMENTO ---
  if (colModificata === cModPag && valTesto === "Pagato") {
    const shCont = ss.getSheetByName("Contabilita");
    const mCont = getMapID(shCont);
    const importo = Number(sheet.getRange(riga, cRestanteDep).getValue()) || 0;
    const datiCont = shCont.getDataRange().getValues();
    let rigaIdx = datiCont.findIndex(r => r[mCont["ID_ROW"] - 1] == idArtista);

    if (rigaIdx !== -1) {
      let rigaCont = rigaIdx + 1;
      let colRataTarget, colSpesePagateTarget, colSpdCntrTarget;

      // Identificazione colonne in base al foglio Deposito
      if (nomeFoglio === "Deposito_T1") {
        colRataTarget = mCont["ID_RATA1"];
        colSpesePagateTarget = mCont["ID_SPESEPAGATE1"];
        colSpdCntrTarget = mCont["ID_SPDCNTR1"];
      }
      else if (nomeFoglio === "Deposito_T2") {
        colRataTarget = mCont["ID_RATA2"];
        colSpesePagateTarget = mCont["ID_SPESEPAGATE2"];
        colSpdCntrTarget = mCont["ID_SPDCNTR2"];
      }
      else if (nomeFoglio === "Deposito_T3") {
        colRataTarget = mCont["ID_RATA3"];
        colSpesePagateTarget = mCont["ID_SPESEPAGATE3"];
        colSpdCntrTarget = mCont["ID_SPDCNTR3"];
      }
      else if (nomeFoglio === "Deposito_T4") {
        colRataTarget = mCont["ID_RATA4"];
        colSpesePagateTarget = mCont["ID_SPESEPAGATE4"];
        colSpdCntrTarget = mCont["ID_SPDCNTR4"];
      }
      
      // Scrittura dati in Contabilità
      if (colRataTarget) {
        shCont.getRange(rigaCont, colRataTarget).setValue(importo);
        const headerRata = shCont.getRange(1, colRataTarget).getValue();
        const colData = mCont[headerRata + "DATA"];
        if (colData) shCont.getRange(rigaCont, colData).setValue(oggi);
      }
      
      if (colSpesePagateTarget) shCont.getRange(rigaCont, colSpesePagateTarget).setValue("SI");
      if (colSpdCntrTarget) shCont.getRange(rigaCont, colSpdCntrTarget).setValue("Spd Cntr");

      // Azioni sul foglio Deposito attuale
      sheet.getRange(riga, cSpese).setValue(0); 
      sheet.getRange(riga, cStato).setValue("Da spedire");
      sheet.getRange(riga, cDataSped).setValue(oggi);
    }
  }
}

function azzeraSpeseInDeposito(ss, idArtista, nomeFoglioDep) {
  const shDep = ss.getSheetByName(nomeFoglioDep);
  if (!shDep) return;
  const m = getMapID(shDep);
  const dati = shDep.getDataRange().getValues();
  for (let i = 1; i < dati.length; i++) {
    if (dati[i][m["ID_ROW"] - 1] == idArtista) {
      shDep.getRange(i + 1, m["ID_SPESE"]).setValue(0);
      break;
    }
  }
}

/**
 * Sincronizzazione Storica del Ritirato
 */
function sincronizzaRitiratoGlobale(ss, idArtista, foglioOrigine) {
  const fogliT = ["Deposito_T1", "Deposito_T2", "Deposito_T3", "Deposito_T4"];
  const siglaOrigine = foglioOrigine.startsWith("Deposito_") ? foglioOrigine.split("_")[1] : "";
  const notaDaScrivere = siglaOrigine !== "" ? "Ritirato " + siglaOrigine : "Ritirato";

  // 1. Aggiornamento Fogli Deposito
  fogliT.forEach(f => {
    if (f === foglioOrigine) return;
    let sh = ss.getSheetByName(f);
    if (!sh) return;
    let m = getMapID(sh);
    let dati = sh.getDataRange().getValues();
    
    for (let i = dati.length - 1; i >= 1; i--) {
      if (dati[i][m["ID_ROW"] - 1] == idArtista) {
        if (fogliT.indexOf(f) > fogliT.indexOf(foglioOrigine) && foglioOrigine !== "Contabilita") {
          sh.deleteRow(i + 1);
        } else {
          sh.getRange(i + 1, m["ID_NOTE"]).setValue(notaDaScrivere);
        }
      }
    }
  });

  // 2. Aggiornamento Foglio Contabilità
  if (foglioOrigine !== "Contabilita") {
    let shCont = ss.getSheetByName("Contabilita");
    let mCont = getMapID(shCont);
    let datiCont = shCont.getDataRange().getValues();
    for (let i = 1; i < datiCont.length; i++) {
      if (datiCont[i][mCont["ID_ROW"] - 1] == idArtista) {
        shCont.getRange(i + 1, mCont["ID_NOTE"]).setValue(notaDaScrivere);
        break;
      }
    }
  }
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