# SETUP E CONFIGURAZIONE SMIRT

## 📋 Checklist Setup Completo

### 1. Google Apps Script
- [ ] Vai su [script.google.com](https://script.google.com)
- [ ] Crea nuovo progetto "SMIRT-Backend"
- [ ] Copia il contenuto di `google-apps-script.js`
- [ ] **CONFIGURA**:
  - `CONFIG.SHEET_ID`: ID del tuo Google Sheet
  - `CONFIG.DRIVE_FOLDER_ID`: ID della cartella Drive "Firme rapporti"
- [ ] Salva e pubblica come Web App
- [ ] Copia URL del Web App generato

### 2. Google Sheets
- [ ] Crea nuovo spreadsheet "Registro interventi"
- [ ] Rinomina primo foglio in "MUD"
- [ ] Crea secondo foglio "Ordinari"
- [ ] **Intestazioni per entrambi** (9 colonne):
  ```
  A: Timestamp
  B: Utente  
  C: Luogo
  D: Data inizio
  E: Data fine
  F: Descrizione
  G: Materiali
  H: Firma Committente
  I: Buono di lavoro
  ```
- [ ] Copia ID del foglio dalla URL

### 3. Google Drive
- [ ] Crea cartella "Firme rapporti"
- [ ] Imposta permessi di modifica per Google Apps Script
- [ ] Copia ID cartella dalla URL
- [ ] Le sottocartelle (MUD/Buono Lavoro) saranno create automaticamente

### 4. Applicazione Web
- [ ] Modifica `index.html`:
  - Aggiorna `SCRIPT_URL` con URL del Google Apps Script
- [ ] Personalizza `users.json` con i tuoi utenti
- [ ] Configura `CONFIG.USER_CODE_MAPPING` nel Google Apps Script

### 5. Deploy
#### GitHub Pages:
- [ ] Carica tutti i file in repository GitHub
- [ ] Abilita GitHub Pages nelle impostazioni
- [ ] Accedi a `https://username.github.io/repository-name`

#### Server locale:
```bash
python -m http.server 8080
```

## 🔍 ID Necessari

### Come trovare Google Sheet ID:
URL: `https://docs.google.com/spreadsheets/d/ID_QUI/edit`
L'ID è la parte tra `/d/` e `/edit`

### Come trovare Google Drive Folder ID:
URL: `https://drive.google.com/drive/folders/ID_QUI`
L'ID è dopo `/folders/`

### Come trovare Google Apps Script URL:
1. Nel progetto Apps Script
2. Deploy → Nuove deployment
3. Tipo: Web app
4. Esegui come: Me
5. Accesso: Chiunque
6. Copia URL generato

## ⚙️ Configurazioni Personalizzate

### Aggiungi nuovi utenti:
1. In `users.json`: aggiungi credenziali
2. Nel Google Apps Script `CONFIG.USER_CODE_MAPPING`: aggiungi mapping utente→lettera

### Modifica colonne:
Aggiorna `CONFIG.COLUMNS_MUD` e `CONFIG.COLUMNS_ORDINARI` nel Google Apps Script

## 🧪 Test

### Test Google Apps Script:
1. Apri Google Apps Script
2. Esegui funzione `testScriptDueFogli()`
3. Controlla log per errori

### Test Applicazione:
1. Login con utente di test
2. Compila rapporto MUD
3. Compila rapporto Ordinario  
4. Verifica che vadano nei fogli corretti
5. Testa firma cliente

## 🔧 Troubleshooting

### Errori comuni:
- **CORS**: Usa JSONP, non fetch/XMLHttpRequest
- **Permessi**: Google Apps Script deve avere accesso a Drive e Sheets
- **ID sbagliati**: Controlla che gli ID siano corretti
- **Fogli mancanti**: Crea manualmente "MUD" e "Ordinari"

### Debug:
- Console browser per errori frontend
- Log Google Apps Script per errori backend
- Network tab per vedere richieste JSONP