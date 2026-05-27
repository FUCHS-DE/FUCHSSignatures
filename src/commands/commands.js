/* global Office */

const { fetchUserProfile }      = require('../shared/graph');
const { fetchSignatureTemplate } = require('../shared/sharepoint');
const { fillPlaceholders }       = require('../shared/templateFiller');

// ─────────────────────────────────────────────────────────────────────────────
// Haupt-Einstiegspunkt
// ─────────────────────────────────────────────────────────────────────────────
Office.onReady(() => {
    // Pflicht: Event-Handler global registrieren,
    // damit der event-based Runtime ihn findet
    Office.actions.associate('handleSignature', handleSignature);
});


// ─────────────────────────────────────────────────────────────────────────────
// handleSignature
// Wird aufgerufen bei: OnNewMessageCompose, OnMessageFromChanged,
//                      OnNewMessageReply, OnMessageReplyAll, OnMessageForward
// ─────────────────────────────────────────────────────────────────────────────
async function handleSignature(event) {
    try {
        // 1. Aktuelle Absenderadresse ermitteln
        const fromEmail = await getFromEmail();
        const userEmail = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
        const domain    = fromEmail.split('@')[1];

        // 2. Ist es ein geteiltes Postfach?
        //    → From-Adresse != eingeloggter User
        const isSharedMailbox = (fromEmail !== userEmail);

        let signatureHtml = null;

        if (isSharedMailbox) {
            // ── Geteiltes Postfach: festes Template, keine Nutzerdaten ────────
            // Datei in SharePoint: shared/service@fuchs-cranes.de.html
            signatureHtml = await fetchSignatureTemplate('shared/' + fromEmail + '.html');

        } else {
            // ── Persönliches Postfach: Template + Nutzerdaten ─────────────────
            // Beide Requests parallel für bessere Performance
            const [template, userProfile] = await Promise.all([
                fetchSignatureTemplate('personal/' + domain + '.html'),
                fetchUserProfile()
            ]);

            if (template) {
                signatureHtml = fillPlaceholders(template, userProfile, fromEmail);
            }
        }

        // 3. Signatur setzen (falls ein Template gefunden wurde)
        if (signatureHtml) {
            await setSignature(signatureHtml);
        } else {
            console.warn('Kein Template gefunden für: ' + fromEmail);
        }

    } catch (err) {
        // Fehler still loggen – der Nutzer soll nicht gestört werden.
        // Im Fehlerfall: keine Signatur, kein Absturz.
        console.error('handleSignature Fehler:', err);
    } finally {
        // WICHTIG: event.completed() muss immer aufgerufen werden,
        // sonst gilt das Event als fehlgeschlagen
        event.completed();
    }
}


// ─────────────────────────────────────────────────────────────────────────────
// Hilfsfunktionen
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Aktuelle From-Adresse aus dem Compose-Fenster lesen.
 * Fallback: eingeloggter User (falls From nicht ermittelt werden kann)
 */
function getFromEmail() {
    return new Promise((resolve) => {
        Office.context.mailbox.item.from.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
                resolve(result.value.emailAddress.toLowerCase());
            } else {
                // Fallback
                resolve(Office.context.mailbox.userProfile.emailAddress.toLowerCase());
            }
        });
    });
}

/**
 * HTML-Signatur in den Mail-Body schreiben.
 * setSignatureAsync ersetzt eine vorhandene Signatur (kein Duplikat).
 */
function setSignature(html) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.setSignatureAsync(
            html,
            { coercionType: Office.CoercionType.Html },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve();
                } else {
                    reject(new Error('setSignatureAsync: ' + result.error.message));
                }
            }
        );
    });
}
