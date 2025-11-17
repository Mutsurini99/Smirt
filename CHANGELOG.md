# Changelog SMIRT

## [3.0.1] - 2025-11-15

### 🔧 Correzioni GitHub Pages
- **CORREZIONE CRITICA**: Risolto loop infinito di ricaricamento pagina
- **AUTENTICAZIONE**: Separati login (`login.html`) e app principale (`index.html`)
- **REDIRECT**: Corretti tutti i redirect per funzionamento su GitHub Pages
- **PWA**: Aggiornato manifest per avvio da pagina di login
- **ICONE**: Aggiunte nuove icone SVG con "I" e "Smirt S.r.l."

## [3.0.0] - 2025-11-15

### Added
- 🆕 **Fogli separati MUD/Ordinari**: Invio automatico nel foglio corretto basato sul tipo di intervento
- 📁 **Organizzazione Drive**: Firme cliente salvate in sottocartelle (MUD → Buono Lavoro, Ordinari → Buono Lavoro)
- 🔄 **Sistema a due fasi**: Prima invio dati, poi firma cliente separatamente
- 🎫 **Buoni lavoro automatici**: Generazione sequenziale per utente (A0001, V0002, etc.)
- 🏠 **Pulsante Home**: Navigazione migliorata nel menu

### Changed
- ⚡ **Performance**: Invio ottimizzato con JSONP per evitare CORS
- 📱 **UI/UX**: Menu dropdown funzionanti, navigazione fluida
- 🔧 **Configurazione**: Setup centralizzato in CONFIG object
- 📊 **Struttura dati**: 9 colonne ottimizzate (rimossa colonna Riferimento per MUD)

### Fixed
- 🐛 **Menu dropdown**: Risolti problemi di inizializzazione e event listeners
- 🔤 **Apostrofi**: Corretti errori di sintassi JavaScript
- 📐 **Canvas firme**: Supporto touch e mouse migliorato
- 🔗 **Validazione**: Controlli completi prima dell'invio

### Technical
- **Frontend**: HTML5 + Tailwind CSS + Vanilla JavaScript
- **Backend**: Google Apps Script con JSONP API
- **Storage**: Google Sheets (fogli separati) + Google Drive (firme)
- **PWA**: Service Worker per funzionalità offline

## [2.0.0] - 2025-11-14

### Added
- 🖊️ **Firma digitale**: Canvas per firma cliente e tecnico
- 📤 **Google Drive**: Upload automatico firme
- 🔄 **Sync bidirezionale**: Sheets ↔ Drive

### Changed
- 📱 **PWA**: Manifest aggiornato per installazione mobile
- 🎨 **Dark mode**: Tema scuro come default

## [1.0.0] - 2025-11-13

### Added
- ✅ **Sistema base**: Form rapporti con Google Sheets
- 👥 **Autenticazione**: Login con users.json
- 💾 **Bozze**: Salvataggio LocalStorage
- 📊 **Validazione**: Controlli campi obbligatori

---

**Legenda:**
- 🆕 Nuove funzionalità
- ⚡ Miglioramenti performance  
- 🐛 Fix bug
- 📱 Mobile/PWA
- 🔧 Configurazione
- 📊 Dati/Backend