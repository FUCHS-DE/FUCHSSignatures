/**
 * SSO-Token über Office.auth holen.
 * Funktioniert auf Desktop, OWA und Mac ohne Login-Prompt,
 * da der bereits eingeloggte M365-User genutzt wird.
 */
async function getAccessToken() {
    try {
        const token = await Office.auth.getAccessToken({
            allowSignInPrompt: false,
            allowConsentPrompt: false,
            forMSGraphAccess: true
        });
        return token;
    } catch (err) {
        // Fehlercode 13003: Nutzer nicht eingeloggt / SSO nicht möglich
        // Fehlercode 13005: App nicht in Azure AD registriert
        console.error('SSO-Fehler (Code ' + err.code + '):', err.message);
        throw err;
    }
}

module.exports = { getAccessToken };
