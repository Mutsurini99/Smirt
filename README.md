# SMIRT - Sistema Gestione Rapporti di Intervento

Sistema web per la gestione e l'invio di rapporti di intervento tecnico con firma digitale e integrazione Google Sheets.

## 🚀 Funzionalità

- **Gestione utenti**: Sistema di autenticazione con sessioni 24h
- **Due tipologie di intervento**: MUD e Ordinari con fogli separati
- **Firma digitale**: Canvas per firma cliente (touch e mouse)
- **PWA**: Installabile su mobile e desktop
- **Google Sheets**: Invio automatico dati con buoni lavoro
- **Google Drive**: Archiviazione automatica firme cliente

## 📊 Architettura

### Frontend
- **HTML5 + CSS3**: Interface responsive con Tailwind CSS
- **JavaScript Vanilla**: Gestione form, firme, LocalStorage
- **PWA**: Service Worker per funzionalità offline

### Backend
- **Google Apps Script**: Elaborazione dati e JSONP API
- **Google Sheets**: Database con fogli MUD/Ordinari separati
- **Google Drive**: Archivio firme organizzato per MUD/Buono Lavoro

### Flusso dati
1. **FASE 1**: Invio dati principali al foglio corretto (MUD/Ordinari)
2. **FASE 2**: Upload firma cliente nella cartella Drive specifica

## 🔧 Configurazione

### Google Apps Script
1. Crea un nuovo progetto Google Apps Script
2. Copia il contenuto di `google-apps-script.js`
3. Configura `CONFIG.SHEET_ID` con l'ID del tuo Google Sheet
4. Pubblica come Web App con accesso pubblico

### Google Sheets
1. Crea spreadsheet "Registro interventi"
2. Crea foglio "MUD" con 9 colonne:
   - Timestamp, Utente, Luogo, Data inizio, Data fine, Descrizione, Materiali, Firma Committente, Buono di lavoro
3. Crea foglio "Ordinari" con stesse colonne

### Google Drive
1. Crea cartella "Firme rapporti"
2. Imposta permessi di scrittura per Google Apps Script
3. Le sottocartelle saranno create automaticamente (MUD/Buono Lavoro)

## 🚀 Deploy

### GitHub Pages
1. Fork questo repository
2. Abilita GitHub Pages nelle impostazioni
3. L'app sarà disponibile su `https://username.github.io/repository-name`

### Server locale
```bash
python -m http.server 8080
```

## 📱 Utilizzo

1. **Login**: Usa credenziali da `users.json`
2. **Nuovo intervento**: Seleziona MUD o Ordinario
3. **Compila form**: Tutti i campi sono obbligatori
4. **Firma cliente**: Disegna nell'apposita area
5. **Invia**: Dati vanno nel foglio corretto, firma su Drive

## 🔐 Sicurezza

- Autenticazione basata su sessioni LocalStorage (24h)
- Firme archiviate con timestamp unico
- Validazione lato client e server
- HTTPS obbligatorio per PWA

## 🛠️ Sviluppo

### Struttura file
```
/
├── index.html              # App principale (rapporti_intervento.V3.html)
├── manifest.json           # PWA manifest
├── sw.js                   # Service Worker
├── users.json              # Database utenti
├── js/
│   └── user-manager.js     # Gestione autenticazione
└── google-apps-script.js   # Backend Google Apps Script
```

### Personalizzazione
- **Utenti**: Modifica `users.json` e `CONFIG.USER_CODE_MAPPING`
- **Colonne**: Aggiorna `CONFIG.COLUMNS_MUD` e `CONFIG.COLUMNS_ORDINARI`
- **Stili**: CSS inline modificabile in `index.html`

## 📈 Versioning

- **v3.0**: Fogli separati MUD/Ordinari
- **v2.0**: Sistema firme + Google Drive
- **v1.0**: Base con Google Sheets

## 🆘 Supporto

Per problemi o domande, controlla:
1. Console browser per errori JavaScript
2. Log Google Apps Script per errori server
3. Permessi Google Drive e Sheets

---

**Sviluppato con ❤️ per SMIRT**