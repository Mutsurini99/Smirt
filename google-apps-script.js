// GOOGLE APPS SCRIPT - VERSIONE JSONP CON FOGLI SEPARATI MUD/ORDINARI
// Questa versione gestisce due fogli separati per MUD e Ordinari
// VERSIONE AGGIORNATA: Sistema a due fasi + gestione firma cliente separata
//
// 🔄 MODIFICHE PRINCIPALI:
// ✅ Due fogli separati: "MUD" e "Ordinari"
// ✅ Colonne diverse per MUD (senza riferimento) e Ordinari (solo colonne base)
// ✅ Invio dati prima, firma cliente dopo
// ✅ Nessuna gestione firma tecnico (processo separato)
// ✅ Rilevamento automatico tipo documento dall'app
//
// ⚠️ SETUP RICHIESTO:
// 1. Crea un Google Sheet chiamato "Registro interventi"
// 2. Crea due fogli: "MUD" e "Ordinari"
// 3. Modifica CONFIG.SHEET_ID con il tuo Google Sheet ID

// 🔧 CONFIGURAZIONE PRINCIPALE - MODIFICA QUESTI VALORI
const CONFIG = {
  SHEET_ID: '1VuWDP84EM0N9lMripB2Ym_rsXGSVKvIEHDAFasDyhsk', // ✅ NUOVO FOGLIO CONFIGURATO
  
  // Nomi dei fogli
  SHEET_NAMES: {
    MUD: 'MUD',
    ORDINARI: 'Ordinari'
  },
  
  // Mappatura colonne per MUD (con MUD e Riferimento - Riferimento non compilato automaticamente)
  COLUMNS_MUD: {
    TIMESTAMP: 1,        // A - Timestamp  
    UTENTE: 2,          // B - Utente
    MUD: 3,             // C - MUD
    RIFERIMENTO: 4,     // D - Riferimento (solo per operatore, non compilato)
    LUOGO: 5,           // E - Luogo
    DATA_INIZIO: 6,     // F - Data inizio
    DATA_FINE: 7,       // G - Data fine
    DESCRIZIONE: 8,     // H - Descrizione
    MATERIALI: 9,       // I - Materiali
    FIRMA_COMMITTENTE: 10, // J - Firma Committente
    BUONO_LAVORO: 11    // K - Buono di lavoro
  },
  
  // Mappatura colonne per Ordinari (solo colonne base come da screenshot)
  COLUMNS_ORDINARI: {
    TIMESTAMP: 1,        // A - Timestamp  
    UTENTE: 2,          // B - Utente
    LUOGO: 3,           // C - Luogo
    DATA_INIZIO: 4,     // D - Data inizio
    DATA_FINE: 5,       // E - Data fine
    DESCRIZIONE: 6,     // F - Descrizione
    MATERIALI: 7,       // G - Materiali
    FIRMA_COMMITTENTE: 8, // H - Firma Committente
    BUONO_LAVORO: 9     // I - Buono di lavoro
  },
  
  // 🎯 SISTEMA BUONI LAVORO: Mappatura utenti -> codice lettera
  USER_CODE_MAPPING: {
    'admin': 'A',
    'tecnico1': 'T', 
    'tecnico2': 'U',
    'valentino': 'V',
    'marco': 'M',
    'giuseppe': 'G',
    'francesco': 'F',
    'antonio': 'N'
    // Aggiungi altri utenti secondo necessità
  }
};

// 🛠️ FUNZIONE DI CONFIGURAZIONE: Imposta il nuovo Google Sheet con due fogli
function configuraGoogleSheetDueFogli(nuovoSheetId) {
  console.log('=== 🛠️ CONFIGURAZIONE GOOGLE SHEET DUE FOGLI ===');
  console.log('📊 Nuovo Sheet ID:', nuovoSheetId);
  
  try {
    // Test di accesso al sheet
    const ss = SpreadsheetApp.openById(nuovoSheetId);
    console.log('✅ Sheet accessibile:', ss.getName());
    
    // Verifica/Crea foglio MUD
    let mudSheet;
    try {
      mudSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MUD);
      if (!mudSheet) {
        mudSheet = ss.insertSheet(CONFIG.SHEET_NAMES.MUD);
        console.log('📋 Foglio MUD creato');
      }
    } catch (error) {
      mudSheet = ss.insertSheet(CONFIG.SHEET_NAMES.MUD);
      console.log('📋 Foglio MUD creato (fallback)');
    }
    
    // Verifica/Crea foglio Ordinari
    let ordinariSheet;
    try {
      ordinariSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ORDINARI);
      if (!ordinariSheet) {
        ordinariSheet = ss.insertSheet(CONFIG.SHEET_NAMES.ORDINARI);
        console.log('📋 Foglio Ordinari creato');
      }
    } catch (error) {
      ordinariSheet = ss.insertSheet(CONFIG.SHEET_NAMES.ORDINARI);
      console.log('📋 Foglio Ordinari creato (fallback)');
    }
    
    // Intestazioni per foglio MUD (con MUD e Riferimento - 11 colonne)
    const headerRowMUD = [
      'Timestamp',
      'Utente',
      'MUD',
      'Riferimento',
      'Luogo',
      'Data inizio',
      'Data fine',
      'Descrizione',
      'Materiali',
      'Firma Committente',
      'Buono di lavoro'
    ];
    
    // Intestazioni per foglio Ordinari (colonne base)
    const headerRowOrdinari = [
      'Timestamp',
      'Utente', 
      'Luogo',
      'Data inizio',
      'Data fine',
      'Descrizione',
      'Materiali',
      'Firma Committente',
      'Buono di lavoro'
    ];
    
    // Imposta intestazioni MUD se il foglio è vuoto
    if (mudSheet.getDataRange().getNumRows() === 0) {
      mudSheet.getRange(1, 1, 1, headerRowMUD.length).setValues([headerRowMUD]);
      console.log('✅ Intestazioni MUD aggiunte');
    }
    
    // Imposta intestazioni Ordinari se il foglio è vuoto
    if (ordinariSheet.getDataRange().getNumRows() === 0) {
      ordinariSheet.getRange(1, 1, 1, headerRowOrdinari.length).setValues([headerRowOrdinari]);
      console.log('✅ Intestazioni Ordinari aggiunte');
    }
    
    return {
      success: true,
      sheetId: nuovoSheetId,
      spreadsheetName: ss.getName(),
      mudSheetName: mudSheet.getName(),
      ordinariSheetName: ordinariSheet.getName(),
      message: 'Configurazione due fogli completata con successo'
    };
    
  } catch (error) {
    console.error('❌ Errore configurazione sheet:', error);
    return {
      success: false,
      error: error.toString(),
      message: 'Impossibile accedere al Google Sheet. Verifica ID e permessi.'
    };
  }
}

// 🎯 FUNZIONE SISTEMA BUONO LAVORO: Genera codice automatico
function generaBuonoLavoro(username, sheet) {
  try {
    console.log('🎫 Generazione Buono Lavoro per utente:', username);
    
    // Ottieni la lettera associata all'utente
    let userLetter = CONFIG.USER_CODE_MAPPING[username.toLowerCase()];
    if (!userLetter) {
      console.warn('⚠️ Utente non trovato nel mapping, uso "X" di default:', username);
      userLetter = 'X'; // Fallback
    }
    
    console.log('🔤 Lettera assegnata:', userLetter);
    
    // Cerca l'ultimo numero utilizzato per questo utente in QUESTO foglio
    const existingData = sheet.getDataRange().getValues();
    let maxNumber = 0;
    
    for (let i = 1; i < existingData.length; i++) { // Skip header
      const buonoLavoro = existingData[i][existingData[0].length - 1]; // Ultima colonna (Buono Lavoro)
      
      if (buonoLavoro && typeof buonoLavoro === 'string' && buonoLavoro.startsWith(userLetter)) {
        // Estrai il numero dal codice (es. "V0005" -> 5)
        const numberPart = buonoLavoro.substring(1);
        const number = parseInt(numberPart, 10);
        
        if (!isNaN(number) && number > maxNumber) {
          maxNumber = number;
        }
      }
    }
    
    // Genera il prossimo numero (incrementa di 1)
    const nextNumber = maxNumber + 1;
    
    // Formatta con 4 cifre (padding con zeri)
    const formattedNumber = nextNumber.toString().padStart(4, '0');
    
    // Crea il codice finale
    const buonoLavoro = userLetter + formattedNumber;
    
    console.log('✅ Buono Lavoro generato:', buonoLavoro);
    console.log('📊 Dettagli: Ultimo numero era', maxNumber, ', nuovo numero:', nextNumber);
    
    return buonoLavoro;
    
  } catch (error) {
    console.error('❌ Errore generazione buono lavoro:', error);
    // Fallback: genera codice casuale
    const fallbackCode = 'X' + Math.floor(Math.random() * 9999).toString().padStart(4, '0');
    console.log('🔄 Uso codice fallback:', fallbackCode);
    return fallbackCode;
  }
}

function doGet(e) {
  try {
    console.log('Richiesta ricevuta:', e.parameter);
    console.log('Headers disponibili:', JSON.stringify(e.parameter));
    console.log('Timestamp richiesta:', new Date().toISOString());
    
    // Se è una richiesta JSONP (con callback)
    if (e.parameter.callback) {
      return handleJsonpRequest(e);
    }
    
    // Richiesta GET normale
    const response = {
      status: 'ok',
      message: 'Script JSONP funzionante - Due fogli MUD/Ordinari',
      timestamp: new Date().toISOString(),
      version: 'JSONP-V4-DUE-FOGLI-CORS-FIX',
      supportedMethods: ['GET-JSONP', 'POST-via-GET'],
      sheets: ['MUD', 'Ordinari'],
      debug: {
        requestTime: new Date().toISOString(),
        parameters: Object.keys(e.parameter || {})
      }
    };
    
    // Risposta con headers CORS espliciti
    const output = ContentService
      .createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
    
    // Aggiungi headers per CORS (anche se non dovrebbe servire per JSONP)
    try {
      output.setHeaders({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Cache-Control': 'no-cache'
      });
    } catch (headerError) {
      console.log('Headers CORS non supportati, ignoro:', headerError);
    }
    
    return output;
      
  } catch (error) {
    console.error('Errore:', error);
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleJsonpRequest(e) {
  try {
    const callback = e.parameter.callback;
    const action = e.parameter.action;
    const simple = e.parameter.simple;
    
    console.log('=== JSONP REQUEST START ===');
    console.log('JSONP Request - Callback:', callback);
    console.log('JSONP Request - Action:', action);
    console.log('JSONP Request - Simple:', simple);
    console.log('JSONP Request - Timestamp:', new Date().toISOString());
    console.log('JSONP Request - All params:', JSON.stringify(e.parameter, null, 2));
    
    let response;
    
    // Test semplice per diagnostica
    if (simple) {
      response = {
        status: 'success',
        message: 'Test semplice JSONP riuscito - CORS FIX',
        timestamp: new Date().toISOString(),
        method: 'JSONP',
        sheets: CONFIG.SHEET_NAMES,
        config: {
          sheetId: CONFIG.SHEET_ID ? 'CONFIGURATO' : 'MANCANTE',
          folderAccess: 'VERIFICARE_MANUALMENTE'
        },
        debug: {
          callback: callback,
          serverTime: new Date().toLocaleString('it-IT'),
          version: 'JSONP-V4-CORS-DEBUG'
        }
      };
    } else if (action === 'test') {
      // Test di connessione
      response = {
        status: 'success',
        message: 'JSONP test successful - Due fogli pronti',
        timestamp: new Date().toISOString(),
        method: 'JSONP',
        sheets: CONFIG.SHEET_NAMES
      };
    } else if (action === 'save') {
      // FASE 1: Salva dati principali SENZA firme
      response = saveDataToCorrectSheet(e.parameter);
    } else if (action === 'upload-client-signature') {
      // FASE 2: Upload solo firma cliente
      response = uploadClientSignature(e.parameter);
    } else if (action === 'ping') {
      // Test di connettività semplice
      response = {
        status: 'pong',
        timestamp: new Date().toISOString(),
        message: 'Server raggiungibile - Due fogli attivi'
      };
    } else {
      response = {
        status: 'error',
        message: 'Azione non riconosciuta: ' + action
      };
    }
    
    // Crea risposta JSONP
    const jsonpResponse = callback + '(' + JSON.stringify(response) + ');';
    
    console.log('JSONP Response length:', jsonpResponse.length);
    console.log('JSONP Response preview:', jsonpResponse.substring(0, 200) + '...');
    console.log('=== JSONP REQUEST END ===');
    
    const output = ContentService
      .createTextOutput(jsonpResponse)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
    
    // Aggiungi headers CORS anche per JSONP
    try {
      output.setHeaders({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Cache-Control': 'no-cache, no-store, must-revalidate',
        'Pragma': 'no-cache',
        'Expires': '0'
      });
    } catch (headerError) {
      console.log('Headers CORS non applicabili per JSONP:', headerError);
    }
    
    return output;
      
  } catch (error) {
    console.error('Errore JSONP:', error);
    const errorResponse = callback + '(' + JSON.stringify({
      status: 'error',
      message: error.toString()
    }) + ');';
    
    return ContentService
      .createTextOutput(errorResponse)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}

// 📊 FUNZIONE PRINCIPALE: Salva dati nel foglio corretto (MUD o Ordinari)
function saveDataToCorrectSheet(params) {
  try {
    console.log('💾 FASE 1: Salvataggio dati nel foglio corretto...');
    
    // ⚠️ Verifica che l'ID del sheet sia configurato
    if (!CONFIG.SHEET_ID || CONFIG.SHEET_ID === 'INSERISCI_QUI_IL_TUO_GOOGLE_SHEET_ID') {
      throw new Error('❌ ERRORE: Devi configurare CONFIG.SHEET_ID nel codice!');
    }
    
    // Decodifica i dati
    const data = JSON.parse(decodeURIComponent(params.data || '{}'));
    
    console.log('📊 Dati ricevuti:', data);
    console.log('🔍 Tipo intervento ricevuto:', data.tipoIntervento);
    console.log('🔍 Campo MUD ricevuto:', data.mud);
    
    // Determina il foglio di destinazione in base al tipo
    const isMUD = data.tipoIntervento === 'mud';
    const sheetName = isMUD ? CONFIG.SHEET_NAMES.MUD : CONFIG.SHEET_NAMES.ORDINARI;
    const columns = isMUD ? CONFIG.COLUMNS_MUD : CONFIG.COLUMNS_ORDINARI;
    
    console.log('🎯 isMUD calcolato:', isMUD);
    console.log('📋 Foglio destinazione determinato:', sheetName);
    console.log('📐 Struttura colonne:', isMUD ? 'MUD (senza riferimento)' : 'Ordinari (base)');
    
    // Apri il foglio
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet;
    
    try {
      sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        throw new Error('Foglio ' + sheetName + ' non trovato');
      }
    } catch (sheetError) {
      console.error('❌ Foglio non trovato:', sheetName);
      throw new Error('Foglio ' + sheetName + ' non esiste. Crealo prima di continuare.');
    }
    
    console.log('✅ Foglio aperto:', sheet.getName());
    
    // CONTROLLO DUPLICATI: Verifica se esiste già un record con lo stesso identificativo
    const identificativo = data.mud || data.riferimento || (data.luogo + '_' + data.dataInizio);
    const existingData = sheet.getDataRange().getValues();
    let existingRowIndex = -1;
    
    for (let i = 1; i < existingData.length; i++) { // Skip header
      // Per MUD controlla campo MUD, per Ordinari controlla luogo+data
      let recordId;
      if (isMUD) {
        recordId = data.mud;
        const existingId = existingData[i][0]; // Assumendo che il MUD sia salvato da qualche parte nei metadati
        // Per ora controllo su luogo+data per semplicità
        const existingLuogo = existingData[i][columns.LUOGO - 1];
        const existingData1 = existingData[i][columns.DATA_INIZIO - 1];
        if (existingLuogo === data.luogo && existingData1 === data.dataInizio) {
          existingRowIndex = i;
          break;
        }
      } else {
        const existingLuogo = existingData[i][columns.LUOGO - 1];
        const existingData1 = existingData[i][columns.DATA_INIZIO - 1];
        if (existingLuogo === data.luogo && existingData1 === data.dataInizio) {
          existingRowIndex = i;
          break;
        }
      }
    }
    
    // Prepara i dati SENZA la firma (verrà aggiunta dopo)
    const timestamp = new Date();
    
    // 🎫 GENERA BUONO LAVORO AUTOMATICO
    let buonoLavoro;
    if (existingRowIndex !== -1) {
      // Record esistente - mantieni buono lavoro se presente
      buonoLavoro = existingData[existingRowIndex][columns.BUONO_LAVORO - 1];
      if (!buonoLavoro || buonoLavoro === 'N/A') {
        buonoLavoro = generaBuonoLavoro(data.user, sheet);
      }
      console.log('🔄 Aggiornamento record esistente, buono lavoro:', buonoLavoro);
    } else {
      // Nuovo record
      buonoLavoro = generaBuonoLavoro(data.user, sheet);
      console.log('🎫 Nuovo buono lavoro generato:', buonoLavoro);
    }
    
    // Crea array con tutti i valori per la riga
    const totalColumns = Object.keys(columns).length;
    const rowData = new Array(totalColumns).fill('N/A');
    
    // Popola i dati usando la mappatura colonne
    rowData[columns.TIMESTAMP - 1] = timestamp.toLocaleString('it-IT');
    rowData[columns.UTENTE - 1] = data.user || 'N/A';
    
    // Solo per MUD: popola anche MUD, lascia Riferimento vuoto
    if (isMUD) {
      rowData[columns.MUD - 1] = data.mud || 'N/A';
      rowData[columns.RIFERIMENTO - 1] = ''; // Vuoto per l'operatore
    }
    
    rowData[columns.LUOGO - 1] = data.luogo || 'N/A';
    rowData[columns.DATA_INIZIO - 1] = data.dataInizio || 'N/A';
    rowData[columns.DATA_FINE - 1] = data.dataFine || 'N/A';
    rowData[columns.DESCRIZIONE - 1] = data.descrizione || 'N/A';
    rowData[columns.MATERIALI - 1] = data.materiali || 'N/A';
    rowData[columns.FIRMA_COMMITTENTE - 1] = 'FIRMA_IN_ATTESA';
    rowData[columns.BUONO_LAVORO - 1] = buonoLavoro;
    
    // Inserisci o aggiorna i dati
    if (existingRowIndex !== -1) {
      // Aggiorna record esistente
      const existingRow = existingRowIndex + 1;
      for (let col = 1; col <= totalColumns; col++) {
        if (col !== columns.FIRMA_COMMITTENTE) { // Non sovrascrivere firma se già presente
          sheet.getRange(existingRow, col).setValue(rowData[col - 1]);
        }
      }
      console.log('🔄 Record aggiornato alla riga:', existingRow);
    } else {
      // Inserisci nuovo record
      sheet.appendRow(rowData);
      console.log('✅ Nuovo record inserito');
    }
    
    console.log('✅ FASE 1 completata - Dati salvati nel foglio:', sheetName);
    
    return {
      status: 'success',
      message: 'Dati salvati con successo nel foglio ' + sheetName,
      timestamp: timestamp.toISOString(),
      phase: 'DATA_SAVED',
      sheetType: isMUD ? 'MUD' : 'Ordinari',
      sheetName: sheetName,
      buonoLavoro: buonoLavoro,
      identificativo: identificativo,
      row: existingRowIndex !== -1 ? existingRowIndex + 1 : 'nuovo'
    };
    
  } catch (error) {
    console.error('❌ Errore FASE 1:', error);
    return {
      status: 'error', 
      message: 'Errore nel salvataggio: ' + error.toString()
    };
  }
}

// 🖊️ FUNZIONE UPLOAD FIRMA CLIENTE: Carica solo firma cliente
function uploadClientSignature(params) {
  try {
    console.log('🖊️ FASE 2: Upload firma cliente...');
    
    // Debug iniziale dettagliato
    console.log('📥 Parametri raw ricevuti:', JSON.stringify(params, null, 2));
    
    let data;
    try {
      data = JSON.parse(decodeURIComponent(params.data || '{}'));
      console.log('✅ Parsing JSON riuscito');
    } catch (parseError) {
      console.error('❌ Errore parsing JSON:', parseError);
      return {
        status: 'error',
        message: 'Errore parsing dati JSON: ' + parseError.toString(),
        debug: { rawData: params.data }
      };
    }
    
    const { identificativo, tipoIntervento, mudValue, luogo, dataInizio, utente, signatureBase64 } = data;
    
    console.log('📥 Parametri estratti:', {
      identificativo,
      tipoIntervento,
      mudValue,
      luogo,
      dataInizio,
      utente,
      signatureLength: signatureBase64 ? signatureBase64.length : 0,
      signaturePrefix: signatureBase64 ? signatureBase64.substring(0, 30) + '...' : 'NONE'
    });
    
    // Validazione parametri più dettagliata
    const missingParams = [];
    if (!identificativo) missingParams.push('identificativo');
    if (!signatureBase64) missingParams.push('signatureBase64');
    if (!luogo) missingParams.push('luogo');
    if (!dataInizio) missingParams.push('dataInizio');
    if (!utente) missingParams.push('utente');
    
    if (missingParams.length > 0) {
      const errorMsg = `Parametri mancanti: ${missingParams.join(', ')}. Ricevuti: ${Object.keys(data).join(', ')}`;
      console.error('❌ ' + errorMsg);
      return {
        status: 'error',
        message: errorMsg,
        debug: { missingParams, receivedData: data }
      };
    }
    
    console.log('✅ Validazione parametri superata');
    console.log('📤 Upload firma cliente per identificativo:', identificativo);
    console.log('📊 Tipo intervento:', tipoIntervento);
    console.log('🏷️ Valore MUD ricevuto:', mudValue);
    console.log('📐 Dimensione firma Base64:', signatureBase64.length, 'caratteri');
    
    // Determina il foglio corretto
    const isMUD = tipoIntervento === 'mud';
    const sheetName = isMUD ? CONFIG.SHEET_NAMES.MUD : CONFIG.SHEET_NAMES.ORDINARI;
    const columns = isMUD ? CONFIG.COLUMNS_MUD : CONFIG.COLUMNS_ORDINARI;
    
    console.log('📋 Foglio target:', sheetName);
    console.log('📊 Configurazione colonne:', columns);
    
    // Test accesso Google Sheet
    console.log('📊 Tentativo apertura Google Sheet:', CONFIG.SHEET_ID);
    let ss, sheet;
    try {
      ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      console.log('✅ Spreadsheet aperto:', ss.getName());
      
      sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        throw new Error('Foglio ' + sheetName + ' non trovato nella lista: ' + ss.getSheets().map(s => s.getName()).join(', '));
      }
      console.log('✅ Foglio aperto:', sheet.getName());
    } catch (sheetError) {
      console.error('❌ Errore accesso sheet:', sheetError);
      return {
        status: 'error',
        message: 'Errore accesso Google Sheet: ' + sheetError.toString(),
        debug: { sheetId: CONFIG.SHEET_ID, targetSheet: sheetName }
      };
    }
    
    // Trova la riga da aggiornare
    console.log('🔍 Inizio ricerca riga nel foglio...');
    const data_range = sheet.getDataRange();
    const values = data_range.getValues();
    let buonoLavoro = null;
    let localMudValue = null;
    let targetRow = -1;
    
    console.log(`🔍 Foglio contiene ${values.length} righe totali (header incluso)`);
    console.log('🔍 Prima riga (header):', values[0]);
    
    // Mostra le prime righe per debug
    if (values.length > 1) {
      console.log('🔍 Alcune righe di esempio:');
      for (let i = 1; i < Math.min(values.length, 4); i++) {
        console.log(`  Riga ${i + 1}:`, values[i]);
      }
    }
    
    // Strategia di ricerca con log dettagliato
    console.log(`🔍 Cercando: utente="${utente}", luogo="${luogo}", dataInizio="${dataInizio}"`);
    
    for (let i = 1; i < values.length; i++) { // Skip header
      const rowUtente = values[i][columns.UTENTE - 1];
      const rowLuogo = values[i][columns.LUOGO - 1];
      const rowDataInizio = values[i][columns.DATA_INIZIO - 1];
      
      console.log(`🔍 Riga ${i + 1}: utente="${rowUtente}", luogo="${rowLuogo}", dataInizio="${rowDataInizio}"`);
      
      // Confronto utente
      const utenteMatch = rowUtente === utente;
      // Confronto luogo
      const luogoMatch = rowLuogo === luogo;
      // Confronto data
      const rowDataStr = typeof rowDataInizio === 'string' ? rowDataInizio : rowDataInizio.toISOString().split('T')[0];
      const targetDataStr = dataInizio;
      const dataMatch = rowDataStr === targetDataStr || rowDataInizio === dataInizio;
      
      console.log(`    Matches: utente=${utenteMatch}, luogo=${luogoMatch}, data=${dataMatch}`);
      
      if (utenteMatch && luogoMatch && dataMatch) {
        console.log('✅ Record trovato alla riga', i + 1);
        targetRow = i + 1;
        
        // Ottieni il buono lavoro dalla riga
        buonoLavoro = values[i][columns.BUONO_LAVORO - 1];
        
        // Per MUD, usa il valore ricevuto dall'app
        if (isMUD) {
          localMudValue = mudValue || identificativo;
        }
        
        console.log('🎫 Buono lavoro trovato:', buonoLavoro);
        console.log('📋 Valore MUD (se applicabile):', localMudValue);
        break;
      }
    }
    
    if (targetRow === -1) {
      const errorMsg = `❌ Record NON TROVATO per utente="${utente}", luogo="${luogo}", dataInizio="${dataInizio}" nel foglio ${sheetName}`;
      console.error(errorMsg);
      console.log('🔍 Righe disponibili nel foglio:');
      for (let i = 1; i < values.length; i++) {
        console.log(`  Riga ${i + 1}: utente="${values[i][columns.UTENTE - 1]}", luogo="${values[i][columns.LUOGO - 1]}", data="${values[i][columns.DATA_INIZIO - 1]}"`);
      }
      return {
        status: 'error',
        message: errorMsg,
        debug: {
          searchCriteria: { utente, luogo, dataInizio },
          sheetName: sheetName,
          totalRows: values.length - 1,
          availableRows: values.slice(1).map((row, idx) => ({
            row: idx + 2,
            utente: row[columns.UTENTE - 1],
            luogo: row[columns.LUOGO - 1],
            dataInizio: row[columns.DATA_INIZIO - 1]
          }))
        }
      };
    }
    
    // Test upload su Google Drive
    console.log('📤 Tentativo caricamento firma su Google Drive...');
    let driveUrl;
    try {
      driveUrl = uploadImageToDrive(
        signatureBase64, 
        'firma_cliente', 
        identificativo,
        tipoIntervento,
        buonoLavoro
      );
      console.log('✅ Upload Drive completato:', driveUrl);
    } catch (driveError) {
      console.error('❌ Errore upload Drive:', driveError);
      return {
        status: 'error',
        message: 'Errore upload su Google Drive: ' + driveError.toString(),
        debug: {
          driveError: driveError.toString(),
          signatureLength: signatureBase64.length,
          buonoLavoro: buonoLavoro
        }
      };
    }
    
    // Aggiorna la cella nel foglio
    console.log('📝 Aggiornamento cella firma nel foglio...');
    try {
      sheet.getRange(targetRow, columns.FIRMA_COMMITTENTE).setValue(driveUrl);
      console.log(`✅ Cella aggiornata: riga ${targetRow}, colonna ${columns.FIRMA_COMMITTENTE}`);
    } catch (cellError) {
      console.error('❌ Errore aggiornamento cella:', cellError);
      return {
        status: 'error',
        message: 'Errore aggiornamento cella: ' + cellError.toString(),
        debug: { targetRow, column: columns.FIRMA_COMMITTENTE, driveUrl }
      };
    }
    
    const successResult = {
      status: 'success',
      message: 'Firma cliente caricata con successo nel foglio ' + sheetName + ' e Drive',
      driveUrl: driveUrl,
      sheetType: isMUD ? 'MUD' : 'Ordinari',
      sheetName: sheetName,
      identificativo: identificativo,
      buonoLavoro: buonoLavoro,
      driveFolder: isMUD ? (localMudValue || buonoLavoro) : buonoLavoro,
      rowUpdated: targetRow,
      searchCriteria: { utente, luogo, dataInizio },
      debug: {
        totalRowsScanned: values.length - 1,
        foundAtRow: targetRow,
        finalDriveUrl: driveUrl
      }
    };
    
    console.log('✅ Upload firma cliente completato con successo:', JSON.stringify(successResult, null, 2));
    return successResult;
    
  } catch (error) {
    console.error('❌ ERRORE GENERALE upload firma cliente:', error);
    console.error('❌ Stack trace:', error.stack);
    return {
      status: 'error',
      message: 'Errore generale upload firma cliente: ' + error.toString(),
      debug: {
        errorDetails: error.toString(),
        stack: error.stack,
        errorName: error.name
      }
    };
  }
}

// 📤 FUNZIONE UPLOAD GOOGLE DRIVE: Carica immagine nella cartella specifica
function uploadImageToDrive(base64Data, fileName, identificativo, tipoIntervento, buonoLavoro) {
  try {
    console.log('📤 Upload su Google Drive nella cartella specifica');
    console.log('📄 File Name:', fileName);
    console.log('🔍 Identificativo:', identificativo);
    console.log('📋 Tipo intervento:', tipoIntervento);
    console.log('🎫 Buono lavoro:', buonoLavoro);
    
    // Rimuovi il prefisso data:image/...;base64, se presente
    const base64 = base64Data.includes(',') ? base64Data.split(',')[1] : base64Data;
    
    // Determina il tipo di immagine dal prefisso
    let mimeType = 'image/png'; // Default
    let extension = '.png';
    if (base64Data.includes('data:image/jpeg')) {
      mimeType = 'image/jpeg';
      extension = '.jpg';
    } else if (base64Data.includes('data:image/jpg')) {
      mimeType = 'image/jpeg'; 
      extension = '.jpg';
    }
    
    // Crea nome file con timestamp per evitare conflitti
    const timestamp = new Date().getTime();
    const fullFileName = fileName + '_' + timestamp + extension;
    
    console.log('📝 Nome file finale:', fullFileName);
    
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64), 
      mimeType, 
      fullFileName
    );
    
    // STEP 1: Trova la cartella specifica usando l'ID fornito
    let mainFolder;
    try {
      // ID della cartella: 13vI7aODCn6CT4soLk73Ki5FxRptttS74
      mainFolder = DriveApp.getFolderById('13vI7aODCn6CT4soLk73Ki5FxRptttS74');
      console.log('📁 Cartella "Firme rapporti" trovata tramite ID:', mainFolder.getName());
    } catch (folderError) {
      console.warn('⚠️ Errore accesso cartella specifica, uso fallback:', folderError);
      // Fallback: cerca per nome
      try {
        const mainFolders = DriveApp.getFoldersByName('Firme rapporti');
        if (mainFolders.hasNext()) {
          mainFolder = mainFolders.next();
          console.log('📁 Cartella "Firme rapporti" trovata per nome');
        } else {
          mainFolder = DriveApp.createFolder('Firme rapporti');
          console.log('📁 Cartella "Firme rapporti" creata');
        }
      } catch (fallbackError) {
        console.error('❌ Errore anche nel fallback, uso root:', fallbackError);
        mainFolder = DriveApp.getRootFolder();
      }
    }
    
    // STEP 2: Determina il nome della sottocartella
    // Per MUD: usa il valore MUD (se disponibile nel parametro identificativo)
    // Per Ordinari: usa il Buono di Lavoro
    let subFolderName;
    if (tipoIntervento === 'mud') {
      // Prova a estrarre il MUD dall'identificativo o usa il buono lavoro
      if (identificativo && identificativo.includes('MUD-')) {
        // Se l'identificativo contiene un pattern MUD, estrailo
        const mudMatch = identificativo.match(/MUD-[A-Za-z0-9-]+/);
        subFolderName = mudMatch ? mudMatch[0] : buonoLavoro;
      } else {
        // Fallback al buono lavoro per MUD
        subFolderName = buonoLavoro;
      }
      console.log('🏷️ MUD - Nome cartella determinato:', subFolderName);
    } else {
      // Per Ordinari usa sempre il Buono di Lavoro
      subFolderName = buonoLavoro;
      console.log('📋 Ordinario - Nome cartella (Buono Lavoro):', subFolderName);
    }
    
    // Pulisci il nome della cartella da caratteri non validi
    const CARTELLA_NOME = subFolderName.replace(/[^a-zA-Z0-9-_]/g, '_');
    console.log('📂 Nome sottocartella determinato:', CARTELLA_NOME);
    console.log('🎯 Logica usata:', tipoIntervento === 'mud' ? 'MUD → usa valore MUD' : 'Ordinario → usa Buono Lavoro');
    
    let subFolder;
    try {
      const existingFolders = mainFolder.getFoldersByName(CARTELLA_NOME);
      if (existingFolders.hasNext()) {
        subFolder = existingFolders.next();
        console.log('📂 ✅ Sottocartella esistente riutilizzata:', CARTELLA_NOME);
      } else {
        subFolder = mainFolder.createFolder(CARTELLA_NOME);
        console.log('📂 🆕 Nuova sottocartella creata:', CARTELLA_NOME);
      }
    } catch (subFolderError) {
      console.error('❌ Errore sottocartella, uso cartella principale:', subFolderError);
      subFolder = mainFolder;
    }
    
    // STEP 3: Carica il file nella sottocartella
    const file = subFolder.createFile(blob);
    console.log('✅ File caricato nella sottocartella:', file.getName());
    console.log('📁 Percorso completo: Firme rapporti/' + CARTELLA_NOME + '/' + file.getName());
    
    // STEP 4: Imposta permessi di visualizzazione pubblica
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      console.log('🔓 Permessi di visualizzazione impostati');
    } catch (permissionError) {
      console.warn('⚠️ Impossibile impostare permessi pubblici:', permissionError);
    }
    
    // STEP 5: Genera URL per visualizzazione diretta
    const viewUrl = 'https://drive.google.com/uc?export=view&id=' + file.getId();
    console.log('🔗 URL generato per visualizzazione:', viewUrl);
    
    return viewUrl;
    
  } catch (error) {
    console.error('❌ Errore upload Google Drive:', error);
    throw new Error('Upload fallito: ' + error.toString());
  }
}

// Funzione doPost - manteniamo per compatibilità
function doPost(e) {
  const response = {
    status: 'info',
    message: 'Usa JSONP invece di POST per evitare CORS',
    suggestion: 'Aggiungi ?callback=yourCallback&action=save&data=encodedJSON'
  };
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// 🔧 FUNZIONE CORREZIONE HEADER: Corregge gli header dei fogli
function correggiHeaderFogli() {
  console.log('=== 🔧 CORREZIONE HEADER FOGLI ===');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    console.log('✅ Sheet aperto:', ss.getName());
    
    // Intestazioni corrette per foglio MUD (11 colonne con MUD e Riferimento)
    const headerRowMUD = [
      'Timestamp',
      'Utente',
      'MUD', 
      'Riferimento',
      'Luogo',
      'Data inizio',
      'Data fine',
      'Descrizione',
      'Materiali',
      'Firma Committente',
      'Buono di lavoro'
    ];
    
    // Intestazioni corrette per foglio Ordinari
    const headerRowOrdinari = [
      'Timestamp',
      'Utente', 
      'Luogo',
      'Data inizio',
      'Data fine',
      'Descrizione',
      'Materiali',
      'Firma Committente',
      'Buono di lavoro'
    ];
    
    // Correggi foglio MUD
    try {
      const mudSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MUD);
      mudSheet.clear(); // Cancella tutto
      mudSheet.getRange(1, 1, 1, headerRowMUD.length).setValues([headerRowMUD]);
      console.log('✅ Header MUD corretto');
    } catch (mudError) {
      console.error('❌ Errore correzione MUD:', mudError);
    }
    
    // Correggi foglio Ordinari
    try {
      const ordinariSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ORDINARI);
      ordinariSheet.clear(); // Cancella tutto
      ordinariSheet.getRange(1, 1, 1, headerRowOrdinari.length).setValues([headerRowOrdinari]);
      console.log('✅ Header Ordinari corretto');
    } catch (ordinariError) {
      console.error('❌ Errore correzione Ordinari:', ordinariError);
    }
    
    return {
      success: true,
      message: 'Header corretti con successo'
    };
    
  } catch (error) {
    console.error('❌ Errore correzione header:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// 🔍 FUNZIONE DEBUG: Verifica contenuto fogli
function verificaContenutoFogli() {
  console.log('=== 🔍 VERIFICA CONTENUTO FOGLI ===');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    console.log('✅ Sheet aperto:', ss.getName());
    
    // Verifica foglio MUD
    try {
      const mudSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MUD);
      const mudData = mudSheet.getDataRange().getValues();
      console.log('📋 Foglio MUD:');
      console.log('  Righe totali:', mudData.length);
      if (mudData.length > 0) {
        console.log('  Header:', mudData[0]);
        if (mudData.length > 1) {
          console.log('  Ultima riga dati:', mudData[mudData.length - 1]);
        }
      }
    } catch (mudError) {
      console.log('❌ Errore accesso foglio MUD:', mudError.toString());
    }
    
    // Verifica foglio Ordinari
    try {
      const ordinariSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ORDINARI);
      const ordinariData = ordinariSheet.getDataRange().getValues();
      console.log('📋 Foglio Ordinari:');
      console.log('  Righe totali:', ordinariData.length);
      if (ordinariData.length > 0) {
        console.log('  Header:', ordinariData[0]);
        if (ordinariData.length > 1) {
          console.log('  Ultima riga dati:', ordinariData[ordinariData.length - 1]);
        }
      }
    } catch (ordinariError) {
      console.log('❌ Errore accesso foglio Ordinari:', ordinariError.toString());
    }
    
    return {
      success: true,
      message: 'Verifica completata'
    };
    
  } catch (error) {
    console.error('❌ Errore verifica:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// 🔍 TEST ACCESSO GOOGLE SHEETS
function testAccessoGoogleSheets() {
  console.log('=== 🔍 TEST ACCESSO GOOGLE SHEETS ===');
  
  try {
    console.log('📊 Tentativo apertura sheet:', CONFIG.SHEET_ID);
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    console.log('✅ Sheet aperto:', ss.getName());
    
    // Lista tutti i fogli esistenti
    const allSheets = ss.getSheets();
    console.log('📋 Fogli esistenti:');
    allSheets.forEach((sheet, index) => {
      console.log(`  ${index + 1}. ${sheet.getName()} (${sheet.getDataRange().getNumRows()} righe)`);
    });
    
    // Test accesso fogli MUD e Ordinari
    let mudSheet = null;
    let ordinariSheet = null;
    
    try {
      mudSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MUD);
      console.log('✅ Foglio MUD trovato:', mudSheet ? mudSheet.getName() : 'NON TROVATO');
    } catch (mudError) {
      console.log('❌ Foglio MUD non trovato:', mudError.toString());
    }
    
    try {
      ordinariSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ORDINARI);
      console.log('✅ Foglio Ordinari trovato:', ordinariSheet ? ordinariSheet.getName() : 'NON TROVATO');
    } catch (ordinariError) {
      console.log('❌ Foglio Ordinari non trovato:', ordinariError.toString());
    }
    
    return {
      success: true,
      sheetName: ss.getName(),
      totalSheets: allSheets.length,
      sheetNames: allSheets.map(s => s.getName()),
      mudExists: mudSheet !== null,
      ordinariExists: ordinariSheet !== null
    };
    
  } catch (error) {
    console.error('❌ Errore accesso sheet:', error);
    return {
      success: false,
      error: error.toString(),
      sheetId: CONFIG.SHEET_ID
    };
  }
}

// 🧪 TEST CONFIGURAZIONE MANUALE
function testConfigurazioneManuale() {
  console.log('=== 🧪 TEST CONFIGURAZIONE MANUALE ===');
  
  try {
    const result = configuraGoogleSheetDueFogli(CONFIG.SHEET_ID);
    console.log('✅ Risultato configurazione:', JSON.stringify(result, null, 2));
    return result;
  } catch (error) {
    console.error('❌ Errore test configurazione:', error);
    return { success: false, error: error.toString() };
  }
}

// 🧪 TEST SCRIPT CON DUE FOGLI
function testScriptDueFogli() {
  console.log('=== 🧪 TEST SCRIPT DUE FOGLI ===');
  
  // Test configurazione
  console.log('📊 Configurazione attuale:');
  console.log('  SHEET_ID:', CONFIG.SHEET_ID);
  console.log('  Fogli:', CONFIG.SHEET_NAMES);
  console.log('  Colonne MUD:', Object.keys(CONFIG.COLUMNS_MUD).length);
  console.log('  Colonne Ordinari:', Object.keys(CONFIG.COLUMNS_ORDINARI).length);
  
  // Test dati MUD
  const testDataMUD = {
    user: 'admin',
    tipoIntervento: 'mud',
    mud: 'TEST-MUD-' + new Date().getTime(),
    luogo: 'Milano - Test MUD',
    dataInizio: '2025-11-15',
    dataFine: '2025-11-15',
    descrizione: 'Test intervento MUD',
    materiali: 'Materiali test MUD'
  };
  
  // Test dati Ordinari
  const testDataOrdinari = {
    user: 'admin',
    tipoIntervento: 'ordinario',
    luogo: 'Roma - Test Ordinario',
    dataInizio: '2025-11-15',
    dataFine: '2025-11-15',
    descrizione: 'Test intervento Ordinario',
    materiali: 'Materiali test Ordinario'
  };
  
  console.log('📤 Test invio MUD:', testDataMUD);
  console.log('📤 Test invio Ordinario:', testDataOrdinari);
  
  // Simula richieste JSONP
  try {
    // Test MUD
    const paramsMUD = {
      parameter: {
        callback: 'testCallback',
        action: 'save',
        data: encodeURIComponent(JSON.stringify(testDataMUD))
      }
    };
    
    const resultMUD = handleJsonpRequest(paramsMUD);
    console.log('✅ Test MUD risultato:', resultMUD.getContent());
    
    // Test Ordinari
    const paramsOrdinari = {
      parameter: {
        callback: 'testCallback',
        action: 'save',
        data: encodeURIComponent(JSON.stringify(testDataOrdinari))
      }
    };
    
    const resultOrdinari = handleJsonpRequest(paramsOrdinari);
    console.log('✅ Test Ordinari risultato:', resultOrdinari.getContent());
    
    return {
      success: true,
      message: 'Test completato con successo',
      mudTest: 'OK',
      ordinariTest: 'OK'
    };
    
  } catch (error) {
    console.error('❌ Errore test:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}