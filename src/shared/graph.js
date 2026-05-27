const { getAccessToken } = require('./auth');

/**
 * Nutzerprofil aus Microsoft Graph API laden.
 * Gibt zurück: displayName, jobTitle, businessPhones, mobilePhone, department, mail
 */
async function fetchUserProfile() {
    const token = await getAccessToken();

    const response = await fetch(
        'https://graph.microsoft.com/v1.0/me' +
        '?$select=displayName,jobTitle,businessPhones,mobilePhone,department,mail',
        {
            headers: { Authorization: 'Bearer ' + token }
        }
    );

    if (!response.ok) {
        throw new Error('Graph API Fehler: ' + response.status + ' ' + response.statusText);
    }

    return response.json();
}

module.exports = { fetchUserProfile };
