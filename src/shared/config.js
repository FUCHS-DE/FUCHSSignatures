/**
 * Zentrale Konfiguration – hier alle Mandanten-spezifischen Werte anpassen.
 */
const CONFIG = {

    // SharePoint-Site, auf der die Signatur-Bibliothek liegt
    sharepointSiteUrl: 'https://fuchsag.sharepoint.com',

    // Server-relativer Pfad zur Dokumentenbibliothek
    // Struktur:
    //   /FUCHSSignaturen/personal/<domain>.html     → z.B. fuchs-ag.de.html
    //   /FUCHSSignaturen/shared/<email>.html        → z.B. service@fuchs-cranes.de.html
    //   /FUCHSSignaturen/assets/<logo>.png
    signaturesLibraryPath: '/FUCHSSignaturen',

    // Azure AD App-ID für SSO (muss in Azure AD registriert sein)
    // Berechtigungen: User.Read, Sites.Read.All
    clientId: '3411d032-43f5-4c9e-a2db-c2e5ee57d732'
};

module.exports = CONFIG;
