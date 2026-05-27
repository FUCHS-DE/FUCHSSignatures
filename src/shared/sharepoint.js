const { getAccessToken } = require('./auth');
const CONFIG = require('./config');

/**
 * Lädt eine Signatur-HTML-Datei aus der SharePoint-Bibliothek.
 *
 * @param {string} relativePath  z.B. 'personal/fuchs-ag.de.html'
 *                               oder 'shared/service@fuchs-cranes.de.html'
 * @returns {string|null}        HTML-Inhalt oder null wenn nicht gefunden
 */
async function fetchSignatureTemplate(relativePath) {
    const token = await getAccessToken();

    // SharePoint REST API: Dateiinhalt direkt abrufen
    // Beispiel: /FUCHSSignaturen/personal/fuchs-ag.de.html
    const serverRelativeUrl =
        CONFIG.signaturesLibraryPath + '/' + relativePath;

    // encodeURIComponent kodiert den Slash – stattdessen nur Sonderzeichen kodieren
    const encodedPath = serverRelativeUrl.replace(/'/g, "''");

    const apiUrl =
        CONFIG.sharepointSiteUrl +
        "/_api/web/GetFileByServerRelativeUrl('" +
        encodedPath +
        "')/$value";

    const response = await fetch(apiUrl, {
        headers: {
            Authorization: 'Bearer ' + token,
            Accept: 'text/html, */*'
        }
    });

    if (response.status === 404) {
        console.warn('Signatur-Template nicht gefunden: ' + relativePath);
        return null;
    }

    if (!response.ok) {
        throw new Error('SharePoint Fehler: ' + response.status + ' für ' + relativePath);
    }

    return response.text();
}

module.exports = { fetchSignatureTemplate };
