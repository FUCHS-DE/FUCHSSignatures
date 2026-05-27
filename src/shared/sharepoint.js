const { getAccessToken } = require('./auth');
const CONFIG = require('./config');

// Aus config.js ableiten:
// sharepointSiteUrl: 'https://fuchsag.sharepoint.com'  → Hostname: 'fuchsag.sharepoint.com'
// signaturesLibraryPath: '/FUCHSSignaturen'             → Bibliotheksname: 'FUCHSSignaturen'
const SITE_HOSTNAME  = new URL(CONFIG.sharepointSiteUrl).hostname;
const LIBRARY_NAME   = CONFIG.signaturesLibraryPath.replace(/^\//, '');

/**
 * Lädt eine Signatur-HTML-Datei aus der SharePoint-Bibliothek via Microsoft Graph.
 *
 * Graph-Endpoint:
 *   GET /sites/{hostname}/lists/{library}/drive/root:/{path}:/content
 *
 * @param {string} relativePath  z.B. 'personal/fuchs-cranes.de.html'
 * @returns {string|null}        HTML-Inhalt oder null wenn nicht gefunden
 */
async function fetchSignatureTemplate(relativePath) {
    const token = await getAccessToken();

    const apiUrl =
        'https://graph.microsoft.com/v1.0/sites/' + SITE_HOSTNAME +
        '/lists/' + encodeURIComponent(LIBRARY_NAME) +
        '/drive/root:/' + relativePath + ':/content';

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
        throw new Error('Graph/SharePoint Fehler: ' + response.status + ' für ' + relativePath);
    }

    return response.text();
}

module.exports = { fetchSignatureTemplate };
