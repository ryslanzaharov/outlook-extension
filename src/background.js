chrome.runtime.onInstalled.addListener(() => {
    console.log("Outlook Calendar Viewer installed.");
});

async function getAccessToken() {
    return new Promise((resolve, reject) => {
        chrome.identity.launchWebAuthFlow(
            {
                url: `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=20536967-8923-4d15-8b76-de1a794f46ce&response_type=token&redirect_uri=https://${chrome.runtime.id}.chromiumapp.org/&scope=https://graph.microsoft.com/Calendars.Read`,
                interactive: true
            },
            redirectUrl => {
                if (chrome.runtime.lastError || !redirectUrl) {
                    reject(new Error("Authorization failed"));
                    return;
                }

                const url = new URL(redirectUrl);
                const accessToken = url.hash.match(/access_token=([^&]*)/)[1];
                chrome.storage.local.set({ token: accessToken }, () => {
                    resolve(accessToken);
                });
            }
        );
    });
}


chrome.action.onClicked.addListener(async () => {
    try {
        const token = await getAccessToken();
        console.log("click and get token" + token);
        chrome.storage.local.set({ token });
    } catch (error) {
        console.error("Error fetching access token:", error);
    }
});



// Функция для удаления всех cookies Microsoft
async function clearMicrosoftCookies() {
    const microsoftDomains = ["login.microsoftonline.com", "microsoft.com"];
    for (const domain of microsoftDomains) {
        const cookies = await chrome.cookies.getAll({ domain });
        for (const cookie of cookies) {
            chrome.cookies.remove({
                url: `https://${domain}${cookie.path}`,
                name: cookie.name
            });
        }
    }
}

async function clearStorage() {
    chrome.storage.local.clear(() => {
        if (chrome.runtime.lastError) {
            console.error("Ошибка при очистке хранилища:", chrome.runtime.lastError);
        } else {
            console.log("Хранилище успешно очищено.");
        }
    });

}

// Функция для выхода из аккаунта Microsoft
async function logoutMicrosoft() {
    await clearMicrosoftCookies();
    await clearStorage();

    // URL для выхода из Microsoft учетной записи
    const logoutUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/logout";

    // Открываем новую вкладку с URL логаута
    chrome.tabs.create({ url: logoutUrl });
}

async function setToken() {
    try {
        const token = await getAccessToken();
        console.log("click and get token" + token);
        chrome.storage.local.set({ token });
    } catch (error) {
        console.error("Error fetching access token:", error);
    }
}

// обработка уведомлений
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.action === "createNotificationWindow") {
        const event = message.event;
        const now = new Date();
        const eventTime = new Date(event.time); // Преобразуем строку ISO обратно в объект Date
        console.log("now event.time description: ", event.description);

        if (isNaN(eventTime.getTime())) {
            console.error("Invalid event time:", event.time);
            sendResponse({ status: "Invalid event time" });
            return;
        }

        // Уведомление за 15 минут
        const timeToNotify15 = new Date(eventTime.getTime() - 15 * 60 * 1000); // За 15 минут до начала
        const delay15 = timeToNotify15.getTime() - now.getTime();

        const timeString = eventTime.toLocaleTimeString(navigator.language, {
            hour: '2-digit',
            minute: '2-digit',
            hour12: false, // 24-часовой формат
        });

        if (delay15 > 0) {

            setTimeout(() => {
                chrome.windows.create({
                    url: `notification.html?title=${encodeURIComponent(event.title)}&time=${encodeURIComponent(timeString)}&location=${encodeURIComponent(event.location)}&description=${encodeURIComponent(event.description)}`,
                    type: "popup",
                    width: 400,
                    height: 300,
                    focused: true,
                    alwaysOnTop: true
                }, () => {
                    console.log("Notification window created for 15 minutes before event.");
                });
            }, delay15);
        }

        // Уведомление за 1 минуту
        const timeToNotify1 = new Date(eventTime.getTime() - 1 * 60 * 1000); // За 1 минуту до начала
        const delay1 = timeToNotify1.getTime() - now.getTime();

        if (delay1 > 0) {
            setTimeout(() => {
                chrome.windows.create({
                    url: `notification.html?title=${encodeURIComponent(event.title)}&time=${encodeURIComponent(timeString)}&location=${encodeURIComponent(event.location)}&description=${encodeURIComponent(event.description)}`,
                    type: "popup",
                    width: 400,
                    height: 300,
                    focused: true,
                    alwaysOnTop: true
                }, () => {
                    console.log("Notification window created for 1 minute before event.");
                });
            }, delay1);
        }

        sendResponse({ status: "Notification windows scheduled" });
    }
});



// Обработка сообщений от popup.js
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "logout") {
        logoutMicrosoft();
        sendResponse({ status: "success" });
    } else if (request.action === "login") {
        setToken();
    } else  if (request.action === 'closeWindow' && sender.tab) {
        chrome.windows.remove(sender.tab.windowId);
    }
});

