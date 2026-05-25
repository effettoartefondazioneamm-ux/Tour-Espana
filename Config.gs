// --- CONFIGURAZIONI GLOBALI ---
const TEST_MODE = true; // Se true, invia tutto a MIO_TEST_EMAIL
const MIO_TEST_EMAIL = "effettoartefondazione.amm@gmail.com"; 
const BOSS_EMAIL = "sandroserradifalco@gmail.com"; 

// Funzione di utilità per la conferma (se non l'hai già nel progetto)
function confermaInvio(tipo) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Conferma Invio', 'Vuoi procedere con l\'invio delle email per ' + tipo + '?', ui.ButtonSet.YES_NO);
  return response == ui.Button.YES;
}

// --- EASTER EGG GIORNALIERO ---
function aggiornaFraseRandomDepositi() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Foglio sorgente frasi
  const foglioDiario = ss.getSheetByName("Diario");

  // Legge intestazioni ID dalla riga 1
  const headers = foglioDiario.getRange(1, 1, 1, foglioDiario.getLastColumn()).getValues()[0];

  // Cerca colonna con ID_FRASI
  const colonnaFrasi = headers.indexOf("ID_FRASI") + 1;

  if (colonnaFrasi === 0) {
    throw new Error("ID_FRASI non trovato nella riga 1 del foglio Diario");
  }

  // Recupera tutte le frasi sotto l'intestazione
  const frasi = foglioDiario
    .getRange(2, colonnaFrasi, foglioDiario.getLastRow() - 1, 1)
    .getValues()
    .flat()
    .filter(String);

  if (frasi.length === 0) {
    throw new Error("Nessuna frase trovata nella colonna ID_FRASI");
  }

  // Evita doppioni consecutivi
  const props = PropertiesService.getScriptProperties();
  const ultimaFrase = props.getProperty("ULTIMA_FRASE");

  let fraseRandom;

  do {
    fraseRandom = frasi[Math.floor(Math.random() * frasi.length)];
  } while (frasi.length > 1 && fraseRandom === ultimaFrase);

  props.setProperty("ULTIMA_FRASE", fraseRandom);

  // Elenco fogli deposito
  const fogliDeposito = [
    "Deposito_T1",
    "Deposito_T2",
    "Deposito_T3",
    "Deposito_T4"
  ];

  // Scrive la frase in G2 mantenendo la formattazione
  fogliDeposito.forEach(nomeFoglio => {

    const foglio = ss.getSheetByName(nomeFoglio);

    if (foglio) {
      foglio.getRange("G2").setValue(fraseRandom);
    }

  });

}