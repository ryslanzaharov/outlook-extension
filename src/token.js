
async function generatePKCE() {
    const encoder = new TextEncoder();
    const randomValues = new Uint8Array(32);
    crypto.getRandomValues(randomValues);
    const codeVerifier = btoa(String.fromCharCode(...randomValues))
        .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

    const data = encoder.encode(codeVerifier);
    const digest = await crypto.subtle.digest('SHA-256', data);
    const codeChallenge = btoa(String.fromCharCode(...new Uint8Array(digest)))
        .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

    return { codeVerifier, codeChallenge };
}


export async function getAccessToken() {
    const authCode = await getAuthCode();
    return await exchangeCodeForTokens(authCode);
}

export async function getStorageAccessToken() {
    const storage = await new Promise(resolve => chrome.storage.local.get(null, resolve));

    if (storage.access_token && Date.now() < storage.expires_at) {
        return storage.access_token;
    }

    if (storage.refresh_token) {
        try {
            return await refreshAccessToken();
        } catch (error) {
            if (error.message.includes("AADSTS70000")) {
                // Refresh token истёк, требуется повторная аутентификация
                return await getAccessToken();
            }
            throw error;
        }
    }
    chrome.runtime.sendMessage({ action: 'closeWindow' });
}


async function getAuthCode() {
    const { codeVerifier, codeChallenge } = await generatePKCE();

    return new Promise((resolve, reject) => {
        chrome.identity.launchWebAuthFlow(
            {
                url: `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=20536967-8923-4d15-8b76-de1a794f46ce&response_type=code&redirect_uri=https://${chrome.runtime.id}.chromiumapp.org/&scope=https://graph.microsoft.com/Calendars.ReadWrite offline_access&code_challenge=${codeChallenge}&code_challenge_method=S256`,
                interactive: true
            },
            async (redirectUrl) => { // Сделаем callback асинхронным
                if (chrome.runtime.lastError || !redirectUrl) {
                    reject(new Error("Authorization failed"));
                    return;
                }

                const url = new URL(redirectUrl);
                const authCode = url.searchParams.get("code");

                if (!authCode) {
                    reject(new Error("Authorization code not found"));
                    return;
                }

                // Дожидаемся сохранения code_verifier
                await new Promise((resolve) =>
                    chrome.storage.local.set({ code_verifier: codeVerifier }, resolve)
                );

                resolve(authCode);
            }
        );
    });
}

async function exchangeCodeForTokens(authCode) {
    const storage = await new Promise(resolve => chrome.storage.local.get(null, resolve));
    const codeVerifier = storage.code_verifier;

    if (!codeVerifier) {
        throw new Error("Code verifier not found");
    }

    const tokenUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

    const params = new URLSearchParams();
    params.append("client_id", "20536967-8923-4d15-8b76-de1a794f46ce");
    params.append("grant_type", "authorization_code");
    params.append("code", authCode);
    params.append("redirect_uri", `https://${chrome.runtime.id}.chromiumapp.org/`);
    params.append("scope", "https://graph.microsoft.com/Calendars.ReadWrite offline_access");
    params.append("code_verifier", codeVerifier);

    const response = await fetch(tokenUrl, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: params
    });

    const data = await response.json();

    if (data.error) {
        throw new Error(`Token exchange failed: ${data.error_description}`);
    }

    // Дожидаемся сохранения токенов
    await new Promise(resolve =>
        chrome.storage.local.set({
            access_token: data.access_token,
            refresh_token: data.refresh_token,
            expires_at: Date.now() + data.expires_in * 1000
        }, resolve)
    );

    return data.access_token;
}

async function refreshAccessToken() {
    const storage = await new Promise(resolve => chrome.storage.local.get(null, resolve));

    if (!storage.refresh_token) {
        throw new Error("No refresh token available");
    }

    const tokenUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

    const params = new URLSearchParams();
    params.append("client_id", "20536967-8923-4d15-8b76-de1a794f46ce");
    params.append("grant_type", "refresh_token");
    params.append("refresh_token", storage.refresh_token);
    params.append("redirect_uri", `https://${chrome.runtime.id}.chromiumapp.org/`);
    params.append("scope", "https://graph.microsoft.com/Calendars.ReadWrite");

    const response = await fetch(tokenUrl, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: params
    });

    const data = await response.json();

    if (data.error) {
        throw new Error(`Refresh token failed: ${data.error_description}`);
    }

    // Ждем сохранения нового токена
    await new Promise(resolve =>
        chrome.storage.local.set({
            access_token: data.access_token,
            refresh_token: data.refresh_token || storage.refresh_token,
            expires_at: Date.now() + data.expires_in * 1000
        }, resolve)
    );

    return data.access_token;
}
