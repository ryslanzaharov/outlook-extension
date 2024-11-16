chrome.runtime.onInstalled.addListener(() => {
    console.log('Расширение установлено');
});

// Функция для аутентификации пользователя через Azure
async function login() {
    const clientId = '20536967-8923-4d15-8b76-de1a794f46ce'; // Замените на ваш client ID
    // const redirectUri = 'https://eodppinalifeiinkhekbfkngcmohnpol.chromiumapp.org/';
    const redirectUri = 'http://localhost:3000/';
    /*
url: `https://login.microsoftonline.com/<TENANT_ID>/oauth2/v2.0/authorize?client_id=<CLIENT_ID>&response_type=token&redirect_uri=https://${chrome.runtime.id}.chromiumapp.org/&scope=openid profile email`,
     */
    // const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}` +
    //     `&response_type=token&redirect_uri=${encodeURIComponent(redirectUri)}` +
    //     `&scope=User.Read%20Calendars.Read%20offline_access` +
    //     `&response_mode=fragment`;

    const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?` +
        `response_type=${encodeURIComponent('token')}` +
        `&response_mode=${encodeURIComponent('fragment')}` +
        `&client_id=${encodeURIComponent(clientId)}` +
        `&redirect_uri=${encodeURIComponent(redirectUri)}` +
        `&scope=${encodeURIComponent('https://graph.microsoft.com/.default')}`;


            chrome.identity.launchWebAuthFlow(
        {
            url: authUrl,
            interactive: true
        },
        (redirectedUrl) => {

            console.info("обрабатываем redirectedUrl: " + redirectedUrl);
            if (chrome.runtime.lastError) {
                console.error("Ошибка входа:", chrome.runtime.lastError);
                return;
            }

            if (!redirectedUrl) {
                console.error("Не удалось получить redirectedUrl.");
                return;
            }

            const url = new URL(redirectedUrl);
            const tokenMatch = url.hash.match(/access_token=([^&]+)/);

            console.info("redirectedUrl: " + redirectedUrl);
            console.info("tokenMatch: " + tokenMatch);

            if (tokenMatch && tokenMatch[1]) {
                const token = tokenMatch[1];
                chrome.storage.local.set({ accessToken: token });
                console.log("Токен успешно получен:", token);
            } else {
                console.error("Не удалось извлечь токен из URL");
            }
        }
    );

}

// Обработка сообщений из popup.js
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.action === 'login') {
        login();
    }
});
