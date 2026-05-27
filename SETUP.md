# FUCHS.Signatur – Setup-Anleitung

## Übersicht

Dieses Add-in setzt automatisch die richtige Signatur, abhängig von der Absenderadresse:
- **Persönliches Postfach** → Template mit Nutzerdaten aus Azure AD
- **Geteiltes Postfach** → Festes Template ohne persönliche Daten

Wechselt der Nutzer die Absenderadresse im Composer, tauscht das Add-in die Signatur sofort aus.

---

## Schritt 1: Azure AD App registrieren

1. Azure Portal → Azure Active Directory → App-Registrierungen → Neu
2. Name: `FUCHS Signatur Add-in`
3. Unterstützte Kontotypen: `Nur diese Organisation`
4. Redirect URI: `https://YOUR_HOST/commands.html`

**API-Berechtigungen hinzufügen (delegiert):**
- `User.Read` – Nutzerprofil aus Graph API lesen
- `Sites.Read.All` – SharePoint-Templates lesen

**App-ID** in `src/shared/config.js` eintragen → `clientId`.

---

## Schritt 2: SharePoint-Bibliothek aufbauen

Dokumentenbibliothek `Signaturen` auf der SharePoint-Site anlegen.

```
Signaturen/
├── personal/
│   ├── fuchs-ag.de.html          ← Template mit {{Platzhaltern}}
│   ├── fuchs-cranes.de.html
│   └── weitere-domain.de.html
├── shared/
│   ├── service@fuchs-cranes.de.html   ← Festes Template
│   └── info@fuchs-ag.de.html
└── assets/
    ├── logo-fuchs-ag.png
    └── logo-fuchs-cranes.png
```

Beispiel-Templates aus `sharepoint-templates/` in die Bibliothek hochladen.

---

## Schritt 3: Konfiguration anpassen

`src/shared/config.js` öffnen und anpassen:

```js
sharepointSiteUrl:      'https://EUER_TENANT.sharepoint.com/sites/Signaturen'
signaturesLibraryPath:  '/sites/Signaturen/Signaturen'
clientId:               'AZURE_AD_APP_ID'
```

`manifest.xml` – alle `YOUR_HOST` ersetzen durch die Hosting-URL.

---

## Schritt 4: Build & Hosting

```bash
npm install
npm run build
# → /dist enthält alle Dateien
```

`/dist` auf einen HTTPS-Server deployen (z.B. Azure Static Web Apps – kostenlos).

---

## Schritt 5: Add-in deployen

**Microsoft 365 Admin Center** → Einstellungen → Integrierte Apps → App hinzufügen → `manifest.xml` hochladen

→ Add-in steht automatisch für alle zugewiesenen Nutzer bereit:
- Outlook Desktop (Windows & Mac) ✅
- OWA ✅
- Outlook Mobile (manuell, kein Auto-Trigger) ⚠️

---

## Platzhalter in Templates

| Platzhalter       | Quelle          | Beispiel                  |
|-------------------|-----------------|---------------------------|
| `{{displayName}}` | Azure AD / Graph | Max Mustermann            |
| `{{jobTitle}}`    | Azure AD / Graph | Servicetechniker          |
| `{{department}}`  | Azure AD / Graph | Technik                   |
| `{{phone}}`       | Azure AD / Graph | +49 123 456789            |
| `{{mobilePhone}}` | Azure AD / Graph | +49 170 1234567           |
| `{{email}}`       | Absenderadresse  | max@fuchs-ag.de           |
| `{{domain}}`      | Absenderadresse  | fuchs-ag.de               |

---

## Neuen Mandanten hinzufügen

1. HTML-Template erstellen (nach Vorlage von `personal/fuchs-ag.de.html`)
2. Als `personal/NEUE-DOMAIN.de.html` in SharePoint hochladen
3. Logo in `assets/` ablegen
4. Fertig – kein Code-Deployment nötig
