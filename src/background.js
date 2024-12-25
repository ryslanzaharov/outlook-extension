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
                chrome.storage.local.set({token: accessToken}, () => {
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
        chrome.storage.local.set({token});
    } catch (error) {
        console.error("Error fetching access token:", error);
    }
});


// Функция для удаления всех cookies Microsoft
async function clearMicrosoftCookies() {
    const microsoftDomains = ["login.microsoftonline.com", "microsoft.com"];
    for (const domain of microsoftDomains) {
        const cookies = await chrome.cookies.getAll({domain});
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
    chrome.tabs.create({url: logoutUrl});
}

async function setToken() {
    try {
        const token = await getAccessToken();
        console.log("click and get token" + token);
        chrome.storage.local.set({token});
    } catch (error) {
        console.error("Error fetching access token:", error);
    }
}


let events = []; // Хранилище для событий

// Устанавливаем события и планируем уведомления
function setEvents(newEvents) {
    const now = new Date();
    // console.log("newEvents " + JSON.stringify(newEvents, null, 2));

    // Фильтруем события за текущий день
    events = newEvents
        .filter(event => isToday(new Date(event.start)))
        .filter(event => new Date(event.start) > now) // Исключаем прошедшие события
        .sort((a, b) => new Date(a.start) - new Date(b.start)); // Сортируем по времени начала

    if (events.length === 0) {
        console.log("No upcoming events for today.");
        chrome.action.setBadgeText({text: ""}); // Очищаем значок
        return;
    }

    console.log("Планируем уведомления " + events);
    // Планируем уведомления
    scheduleNextEvent();
}

// Проверяем, является ли событие сегодняшним
function isToday(date) {
    const today = new Date();
    return (
        date.getFullYear() === today.getFullYear() &&
        date.getMonth() === today.getMonth() &&
        date.getDate() === today.getDate()
    );
}

// Планируем уведомление для следующего события
function scheduleNextEvent() {
    const now = new Date();
    const nextEvent = events.find(event => new Date(event.start) > now);

    if (nextEvent) {
        console.log("Следующее событие:", nextEvent);
        const eventTime = new Date(nextEvent.start);
        const timeUntilEvent = eventTime - now;

        if (timeUntilEvent <= 30 * 60 * 1000) {
            updateBadgeTime(nextEvent); // Обновляем значок за 30 мин
        }
        // Уведомление за 15 минут до события
        if (timeUntilEvent <= 15 * 60 * 1000) {
            showNotification(nextEvent, "15 минут");
        }

        // Уведомление за 1 минуту до события
        if (timeUntilEvent <= 1 * 60 * 1000 && timeUntilEvent > 0) {
            showNotification(nextEvent, "1 минута");
        }
    } else {
        chrome.action.setBadgeText({text: ""}); // Очищаем значок, если событий больше нет
    }
}


// Обновляем значок с оставшимся временем
function updateBadgeTime(event) {
    const now = new Date();
    const eventTime = new Date(event.start);
    const timeLeft = Math.floor((eventTime - now) / 60000); // Остаток времени в минутах

    if (timeLeft > 0) {
        chrome.action.setBadgeText({text: `${timeLeft}m`}); // Устанавливаем текст на значке
        chrome.action.setBadgeBackgroundColor({color: "#FF5733"}); // Цвет значка
    } else {
        chrome.action.setBadgeText({text: ""}); // Очищаем значок
    }
}

// Показываем уведомление
function showNotification(event) {
    console.log("Установка таймаута:", event);
    const eventTime = new Date(event.start); // Преобразуем строку ISO обратно в объект Date
    if (isNaN(eventTime.getTime())) {
        console.error("Invalid event time:", event.start);
        return;
    }

    // Уведомление за 15 минут
    const now = new Date();
    const timeToNotify15 = new Date(eventTime.getTime() - 15 * 60 * 1000); // За 15 минут до начала
    const delay15 = timeToNotify15.getTime() - now.getTime();

    const timeString = eventTime.toLocaleTimeString(navigator.language, {
        hour: '2-digit',
        minute: '2-digit',
        hour12: false, // 24-часовой формат
    });

    if (delay15 > 0) {
        console.log("delay15 ", delay15);
        setTimeout(() => {
            chrome.windows.create({
                url: `notification.html?title=${encodeURIComponent(event.title)}&time=${encodeURIComponent(timeString)}&location=${encodeURIComponent(event.location)}&description=${encodeURIComponent(event.description)}`,
                type: "popup",
                width: 400,
                height: 300,
                focused: true,
                // alwaysOnTop: true
            }, () => {
                console.log("Notification window created for 15 minutes before event.");
            });
        }, delay15);
    }

    // Уведомление за 1 минуту
    const timeToNotify1 = new Date(eventTime.getTime() - 1 * 60 * 1000); // За 1 минуту до начала
    const delay1 = timeToNotify1.getTime() - now.getTime();

    if (delay1 > 0) {
        console.log("delay1 ", delay1);
        setTimeout(() => {
            chrome.windows.create({
                url: `notification.html?title=${encodeURIComponent(event.title)}&time=${encodeURIComponent(timeString)}&location=${encodeURIComponent(event.location)}&description=${encodeURIComponent(event.description)}`,
                type: "popup",
                width: 400,
                height: 300,
                focused: true,
                // alwaysOnTop: true
            }, () => {
                console.log("Notification window created for 1 minute before event.");
            });
        }, delay1);
    }
}

// Устанавливаем будильник на каждые 30 секунд
chrome.alarms.create("checkEvents", { periodInMinutes: 1 });

chrome.alarms.onAlarm.addListener((alarm) => {
    if (alarm.name === "checkEvents") {
        console.log("Alarm triggered: checking events...");
        // Проверяем и обновляем данные
        const now = new Date();
        events = events.filter(event => new Date(event.start) > now); // Удаляем прошедшие события
        scheduleNextEvent();
    }
});

// Получаем события через сообщение
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.action === "setEvents") {
        setEvents(message.notifyEvents);
        sendResponse({status: "Events set and monitoring started"});
    }
});

// Обработка сообщений от popup.js
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "logout") {
        logoutMicrosoft();
        sendResponse({status: "success"});
    } else if (request.action === "login") {
        setToken();
    } else if (request.action === 'closeWindow' && sender.tab) {
        chrome.windows.remove(sender.tab.windowId);
    }
});

