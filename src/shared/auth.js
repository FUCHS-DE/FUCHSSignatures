/* global Office */
const { PublicClientNext } = require('@azure/msal-browser');
const CONFIG = require('./config');

/**
 * Graph-Scopes, für die wir Tokens anfordern.
 * Müssen mit den Azure AD App-Berechtigungen übereinstimmen.
 */
const GRAPH_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.Read.All'
];

// Singleton-Initialisierung: MSAL nur einmal aufbauen
let _msalPromise = null;

function getMsalInstance() {
    if (!_msalPromise) {
        _msalPromise = PublicClientNext.createPublicClientApplication({
            auth: {
                clientId: CONFIG.clientId,
                // 'organizations' = nur Work/School-Accounts (kein persönliches Microsoft-Konto)
                authority: 'https://login.microsoftonline.com/organizations',
                // Nested App Auth: MSAL nutzt den Office-Kontext für SSO,
                // kein Popup, kein Redirect nötig
                supportsNestedAppAuth: true
            }
        });
    }
    return _msalPromise;
}

/**
 * Graph-Zugriffstoken holen.
 * Nutzt NAA (Nested App Auth) – läuft silent, da der User bereits
 * in Office/OWA eingeloggt ist.
 */
async function getAccessToken() {
    const msal = await getMsalInstance();

    const accounts = msal.getAllAccounts();
    const tokenRequest = {
        scopes: GRAPH_SCOPES,
        account: accounts.length > 0 ? accounts[0] : undefined
    };

    try {
        const result = await msal.acquireTokenSilent(tokenRequest);
        return result.accessToken;
    } catch (err) {
        console.error('MSAL Token-Fehler (' + (err.errorCode || err.name) + '):', err.message);
        throw err;
    }
}

module.exports = { getAccessToken };
