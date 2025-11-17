/**
 * SMIRT User Manager - Sistema di Gestione Utenti
 * Gestisce autenticazione, sessioni e ruoli utente
 */

class UserManager {
    constructor() {
        this.users = new Map();
        this.currentUser = null;
        this.sessionKey = 'smirt_session';
        this.usersKey = 'smirt_users';
        this.sessionDuration = 24 * 60 * 60 * 1000; // 24 ore
    }

    /**
     * Carica gli utenti dal file users.json o dal localStorage
     */
    async loadUsers() {
        try {
            // Prova a caricare dal file users.json
            const response = await fetch('./users.json');
            if (response.ok) {
                const data = await response.json();
                this.users.clear();
                Object.entries(data.users).forEach(([username, userData]) => {
                    this.users.set(username, {
                        username: username,
                        password: userData.password,
                        displayName: userData.displayName,
                        role: userData.role,
                        permissions: userData.permissions || [],
                        active: true,
                        createdAt: new Date().toISOString()
                    });
                });
                console.log('👥 Utenti caricati dal file users.json:', this.users.size);
                this.saveUsers();
                return;
            }
        } catch (error) {
            console.warn('⚠️ File users.json non trovato, uso localStorage:', error.message);
        }

        try {
            const storedUsers = localStorage.getItem(this.usersKey);
            if (storedUsers) {
                const usersArray = JSON.parse(storedUsers);
                this.users.clear();
                usersArray.forEach(user => {
                    this.users.set(user.username, user);
                });
                console.log('👥 Utenti caricati dal localStorage:', this.users.size);
            } else {
                // Inizializza utenti predefiniti
                this.initializeDefaultUsers();
                console.log('👥 Utenti predefiniti inizializzati');
            }
            
            // Verifica che gli utenti siano stati caricati
            if (this.users.size === 0) {
                console.warn('⚠️ Nessun utente trovato, reinizializzo gli utenti predefiniti');
                this.initializeDefaultUsers();
            }
        } catch (error) {
            console.error('❌ Errore caricamento utenti:', error);
            this.initializeDefaultUsers();
        }
    }

    /**
     * Inizializza gli utenti predefiniti
     */
    initializeDefaultUsers() {
        const defaultUsers = [
            {
                username: 'admin',
                password: '2977',
                displayName: 'Amministratore',
                role: 'admin',
                active: true,
                createdAt: new Date().toISOString()
            }
        ];

        this.users.clear();
        defaultUsers.forEach(user => {
            this.users.set(user.username, user);
        });

        this.saveUsers();
    }

    /**
     * Salva gli utenti nel localStorage
     */
    saveUsers() {
        try {
            const usersArray = Array.from(this.users.values());
            localStorage.setItem(this.usersKey, JSON.stringify(usersArray));
            console.log('💾 Utenti salvati nel localStorage');
        } catch (error) {
            console.error('❌ Errore salvataggio utenti:', error);
        }
    }

    /**
     * Autentica un utente
     */
    authenticate(username, password) {
        console.log(`🔍 Tentativo login per: ${username}`);
        console.log(`📊 Utenti disponibili: ${this.users.size}`);
        console.log(`📋 Utenti: ${Array.from(this.users.keys()).join(', ')}`);
        
        const user = this.users.get(username);
        
        if (!user) {
            console.log(`❌ Utente non trovato: ${username}`);
            return { success: false, message: 'Utente non trovato' };
        }

        if (!user.active) {
            console.log(`❌ Utente disattivato: ${username}`);
            return { success: false, message: 'Utente disattivato' };
        }

        if (user.password !== password) {
            console.log(`❌ Password errata per: ${username}`);
            return { success: false, message: 'Password errata' };
        }

        // Crea sessione
        const sessionData = {
            username: user.username,
            displayName: user.displayName,
            role: user.role,
            loginTime: Date.now(),
            expiresAt: Date.now() + this.sessionDuration
        };

        this.currentUser = sessionData;
        localStorage.setItem(this.sessionKey, JSON.stringify(sessionData));
        
        console.log(`✅ Utente autenticato: ${username} (${user.role})`);
        return { success: true, user: sessionData };
    }

    /**
     * Controlla se la sessione corrente è valida
     */
    isSessionValid() {
        try {
            const sessionData = localStorage.getItem(this.sessionKey);
            if (!sessionData) {
                return false;
            }

            const session = JSON.parse(sessionData);
            const now = Date.now();

            if (now > session.expiresAt) {
                console.log('⏰ Sessione scaduta');
                this.logout();
                return false;
            }

            // Estendi la sessione se valida
            session.expiresAt = now + this.sessionDuration;
            localStorage.setItem(this.sessionKey, JSON.stringify(session));
            this.currentUser = session;

            return true;
        } catch (error) {
            console.error('❌ Errore validazione sessione:', error);
            return false;
        }
    }

    /**
     * Ottiene l'utente corrente
     */
    getCurrentUser() {
        if (this.isSessionValid()) {
            return this.currentUser;
        }
        return null;
    }

    /**
     * Effettua il logout
     */
    logout() {
        this.currentUser = null;
        localStorage.removeItem(this.sessionKey);
        console.log('👋 Logout effettuato');
    }

    /**
     * Aggiunge un nuovo utente
     */
    addUser(userData) {
        if (this.users.has(userData.username)) {
            return { success: false, message: 'Utente già esistente' };
        }

        const newUser = {
            ...userData,
            active: true,
            createdAt: new Date().toISOString()
        };

        this.users.set(userData.username, newUser);
        this.saveUsers();

        console.log(`➕ Nuovo utente aggiunto: ${userData.username}`);
        return { success: true, user: newUser };
    }

    /**
     * Aggiorna un utente esistente
     */
    updateUser(username, updates) {
        const user = this.users.get(username);
        if (!user) {
            return { success: false, message: 'Utente non trovato' };
        }

        const updatedUser = {
            ...user,
            ...updates,
            updatedAt: new Date().toISOString()
        };

        this.users.set(username, updatedUser);
        this.saveUsers();

        console.log(`📝 Utente aggiornato: ${username}`);
        return { success: true, user: updatedUser };
    }

    /**
     * Disattiva un utente
     */
    deactivateUser(username) {
        return this.updateUser(username, { active: false });
    }

    /**
     * Attiva un utente
     */
    activateUser(username) {
        return this.updateUser(username, { active: true });
    }

    /**
     * Ottiene tutti gli utenti (solo per admin)
     */
    getAllUsers() {
        const currentUser = this.getCurrentUser();
        if (!currentUser || currentUser.role !== 'admin') {
            return { success: false, message: 'Accesso negato' };
        }

        return {
            success: true,
            users: Array.from(this.users.values()).map(user => ({
                username: user.username,
                displayName: user.displayName,
                role: user.role,
                active: user.active,
                createdAt: user.createdAt,
                updatedAt: user.updatedAt
            }))
        };
    }

    /**
     * Cambia password utente
     */
    changePassword(username, oldPassword, newPassword) {
        const user = this.users.get(username);
        if (!user) {
            return { success: false, message: 'Utente non trovato' };
        }

        if (user.password !== oldPassword) {
            return { success: false, message: 'Password attuale errata' };
        }

        return this.updateUser(username, { password: newPassword });
    }

    /**
     * Verifica se l'utente corrente ha il ruolo specificato
     */
    hasRole(role) {
        const currentUser = this.getCurrentUser();
        return currentUser && currentUser.role === role;
    }

    /**
     * Verifica se l'utente corrente è admin
     */
    isAdmin() {
        return this.hasRole('admin');
    }

    /**
     * Pulisce sessioni scadute (maintenance)
     */
    cleanupSessions() {
        // In questo caso semplice, gestiamo solo una sessione
        if (!this.isSessionValid()) {
            this.logout();
        }
    }
}

// Instanza globale del gestore utenti
const userManager = new UserManager();

// Esporta per uso globale
if (typeof window !== 'undefined') {
    window.userManager = userManager;
}

// Esporta per moduli (se necessario)
if (typeof module !== 'undefined' && module.exports) {
    module.exports = UserManager;
}

console.log('👥 UserManager inizializzato');
