/**
 * Ersetzt Platzhalter im HTML-Template durch echte Nutzerdaten.
 *
 * Verfügbare Platzhalter im Template:
 *   {{displayName}}   → Voller Name (z.B. "Max Mustermann")
 *   {{jobTitle}}      → Berufsbezeichnung (z.B. "Servicetechniker")
 *   {{department}}    → Abteilung
 *   {{phone}}         → Erste Bürotelefonnummer
 *   {{mobilePhone}}   → Mobilnummer
 *   {{email}}         → Aktuelle Absenderadresse (From-Adresse)
 *   {{domain}}        → Domain des Absenders (z.B. "fuchs-ag.de")
 *
 * @param {string} template      HTML-Template mit Platzhaltern
 * @param {object} user          Objekt aus Graph API (fetchUserProfile)
 * @param {string} fromEmail     Aktuelle Absenderadresse
 * @returns {string}             Fertiges HTML
 */
function fillPlaceholders(template, user, fromEmail) {
    if (!template) return null;

    const phone      = (user.businessPhones && user.businessPhones[0]) || '';
    const mobile     = user.mobilePhone || '';
    const domain     = fromEmail.split('@')[1] || '';

    return template
        .replace(/\{\{displayName\}\}/g,  user.displayName  || '')
        .replace(/\{\{jobTitle\}\}/g,     user.jobTitle     || '')
        .replace(/\{\{department\}\}/g,   user.department   || '')
        .replace(/\{\{phone\}\}/g,        phone)
        .replace(/\{\{mobilePhone\}\}/g,  mobile)
        .replace(/\{\{email\}\}/g,        fromEmail)
        .replace(/\{\{domain\}\}/g,       domain);
}

module.exports = { fillPlaceholders };
