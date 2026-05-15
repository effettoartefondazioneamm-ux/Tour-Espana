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