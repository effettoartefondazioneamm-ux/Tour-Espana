function aggiornaRimanenzeDaDeposito() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSrc = ss.getSheetByName('Deposito');
  const sheetDest = ss.getSheetByName('Rimanenze');

  if (!sheetSrc || !sheetDest) {
    throw new Error('Assicurati che i fogli "Deposito" e "Rimanenze" esistano.');
  }

  // --- 1. MAPPATURA DINAMICA SORGENTE (Deposito) ---
  const headersSrc = sheetSrc.getRange(1, 1, 1, sheetSrc.getLastColumn()).getValues()[0];
  const getColIdxSrc = (id) => {
    const idx = headersSrc.indexOf(id);
    if (idx === -1) throw new Error('ID non trovato in Deposito riga 1: ' + id);
    return idx;
  };

  const colIdxSrc = {
    progetto: getColIdxSrc('ID_PROGETTO'),
    curatore: getColIdxSrc('ID_CURATORE'),
    row: getColIdxSrc('ID_ROW'),
    nome: getColIdxSrc('ID_NOME'),
    restante: getColIdxSrc('ID_RESTANTE'),
    spedito: getColIdxSrc('ID_SPEDITO'),
    note: getColIdxSrc('ID_NOTE')
  };

  // --- 2. MAPPATURA DINAMICA DESTINAZIONE (Rimanenze) ---
  const headersDest = sheetDest.getRange(1, 1, 1, sheetDest.getLastColumn()).getValues()[0];
  
  // Funzione per trovare la posizione della colonna nel foglio Rimanenze (1-based per il comando .sort)
  const getColPosDest = (id) => {
    const idx = headersDest.indexOf(id);
    return idx !== -1 ? idx + 1 : null;
  };

  // Identifichiamo dove devono andare i dati nel foglio Rimanenze
  const mappingDest = [
    { id: 'ID_PROGETTO', srcIdx: colIdxSrc.progetto },
    { id: 'ID_CURATORE', srcIdx: colIdxSrc.curatore },
    { id: 'ID_ROW', srcIdx: colIdxSrc.row },
    { id: 'ID_NOME', srcIdx: colIdxSrc.nome },
    { id: 'ID_RESTANTE', srcIdx: colIdxSrc.restante }
  ];

  // --- 3. ACQUISIZIONE E FILTRAGGIO DATI ---
  const startRowSrc = 7;
  const lastRowSrc = sheetSrc.getLastRow();
  
  // Pulizia preventiva del foglio Rimanenze (A3:E)
  sheetDest.getRange('A3:E').clearContent();

  if (lastRowSrc < startRowSrc) {
    Logger.log('Nessun dato presente nel Deposito.');
    return;
  }

  const dataSrc = sheetSrc.getRange(startRowSrc, 1, lastRowSrc - startRowSrc + 1, headersSrc.length).getValues();
  const out = [];

  for (let i = 0; i < dataSrc.length; i++) {
    const row = dataSrc[i];
    const valSpedito = row[colIdxSrc.spedito];
    const valNote = String(row[colIdxSrc.note] || "").toLowerCase();
    
    if (valSpedito === "" && !valNote.includes("ritirato")) {
      // Creiamo la riga di output seguendo l'ordine richiesto: PROGETTO, CURATORE, ROW, NOME, RESTANTE
      out.push([
        row[colIdxSrc.progetto],
        row[colIdxSrc.curatore],
        row[colIdxSrc.row],
        row[colIdxSrc.nome],
        row[colIdxSrc.restante]
      ]);
    }
  }

  // --- 4. SCRITTURA E ORDINAMENTO ---
  if (out.length > 0) {
    const destStartRow = 3;
    // Incolla i dati nel range A3:E
    const destRange = sheetDest.getRange(destStartRow, 1, out.length, 5);
    destRange.setValues(out);
    
    // Trova dinamicamente la colonna ID_CURATORE nel foglio Rimanenze per l'ordinamento
    const sortColPos = getColPosDest('ID_CURATORE');
    
    if (sortColPos) {
      destRange.sort({column: sortColPos, ascending: true});
      Logger.log('Ordinamento eseguito per ID_CURATORE (Colonna ' + sortColPos + ')');
    } else {
      // Fallback: se non trova l'ID_CURATORE in riga 1, ordina per la seconda colonna (B)
      destRange.sort({column: 2, ascending: true});
    }
    
    Logger.log('Aggiornamento completato: ' + out.length + ' righe inserite.');
  } else {
    Logger.log('Nessuna riga corrisponde ai criteri di filtro.');
  }
}