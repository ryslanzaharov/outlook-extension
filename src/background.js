chrome.runtime.onInstalled.addListener(() => {
    console.log("Outlook Calendar Viewer installed.");
});

async function getAccessToken() {
    return new Promise((resolve, reject) => {
        chrome.identity.launchWebAuthFlow(
            {
                url: `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=20536967-8923-4d15-8b76-de1a794f46ce&response_type=token&redirect_uri=https://${chrome.runtime.id}.chromiumapp.org/&scope=https://graph.microsoft.com/Calendars.ReadWrite`,
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

function formatTo12Hour(date) {
    const options = {
        hour: 'numeric',
        hour12: true
    };
    return date.toLocaleString('en-US', options);
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
const timeouts = new Map();

// Устанавливаем события и планируем уведомления
function setEvents(newEvents) {
    const now = new Date();

    // Очистка существующих таймаутов
    clearAllTimeouts();

    // Фильтруем события за текущий день
    events = newEvents
        .filter(event => isToday(new Date(event.start)))
        .filter(event => new Date(event.start) > now) // Исключаем прошедшие события
        .sort((a, b) => new Date(a.start) - new Date(b.start)); // Сортируем по времени начала

    if (events.length === 0) {
        console.log("No upcoming events for today.");
        chrome.action.setBadgeText({ text: "" }); // Очищаем значок
        return;
    }

    // console.log("Планируем уведомления", events);
    scheduleNextBadgeTime();
    // Планируем уведомления
    events.forEach(scheduleEventNotifications);
}

// Планируем уведомления для события
function scheduleEventNotifications(event) {
    const now = new Date();
    const eventTime = new Date(event.start);

    if (eventTime <= now) return; // Пропускаем прошедшие события

    // Уведомление за 15 минут
    const notify15Time = eventTime.getTime() - 15 * 60 * 1000;
    if (notify15Time > now.getTime()) {
        const timeout15 = setTimeout(() => {
            showNotification(event, "15 минут");
        }, notify15Time - now.getTime());
        timeouts.set(`${event.title}_15`, timeout15);
    }

    // Уведомление за 1 минуту
    const notify1Time = eventTime.getTime() - 1 * 60 * 1000;
    if (notify1Time > now.getTime()) {
        const timeout1 = setTimeout(() => {
            showNotification(event, "1 минута");
        }, notify1Time - now.getTime());
        timeouts.set(`${event.title}_1`, timeout1);
    }
}

// Очистка всех таймаутов
function clearAllTimeouts() {
    timeouts.forEach(timeout => clearTimeout(timeout));
    timeouts.clear();
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

// Обновляем значок за 30 мин
function scheduleNextBadgeTime() {
    const now = new Date();
    const nextEvent = events.find(event => new Date(event.start) > now);

    if (nextEvent) {
        // console.log("Следующее событие:", nextEvent);
        const eventTime = new Date(nextEvent.start);
        const timeUntilEvent = eventTime - now;

        if (timeUntilEvent <= 30 * 60 * 1000) {
            updateBadgeTime(nextEvent); // Обновляем значок за 30 мин
        } else {
            chrome.action.setBadgeText({text: formatTo12Hour(eventTime)});
            chrome.action.setBadgeBackgroundColor({ color: "#98908e" }); // Ярко-оранжевый цвет
            chrome.action.setBadgeBackgroundColor({color: "#336dff"}); // Цвет значка
        }

    } else {
        chrome.action.setBadgeText({text: ""}); // Очищаем значок, если событий больше нет
    }
}


// Обновляем значок с оставшимся временем
function updateBadgeTime(event) {
    const now = new Date();
    const eventTime = new Date(event.start);
    const timeLeft = Math.floor((eventTime - now) / 60000) + 1; // Остаток времени в минутах

    if (timeLeft > 0) {
        chrome.action.setBadgeText({text: `${timeLeft}m`}); // Устанавливаем текст на значке
        chrome.action.setBadgeBackgroundColor({ color: "#98908e" }); // Ярко-оранжевый цвет
        chrome.action.setBadgeBackgroundColor({color: "#336dff"}); // Цвет значка
    } else {
        chrome.action.setBadgeText({text: ""}); // Очищаем значок
    }
}

// Показываем уведомление
function showNotification(event, timeLabel) {
    // console.log(`Уведомление: ${timeLabel} до события "${event.title}"`);
    const eventTime = new Date(event.start).toLocaleTimeString(navigator.language, {
        hour: 'numeric',
        minute: 'numeric',
        hour12: true
    });

    chrome.windows.create({
        url: `notification.html?title=${encodeURIComponent(event.title)}&time=${encodeURIComponent(eventTime)}&location=${encodeURIComponent(event.location)}`,
        type: "popup",
        width: 400,
        height: 300,
        focused: true,
    }, () => {
        console.log(`Notification window created: ${timeLabel} до события.`);
    });

}

// Устанавливаем будильник на каждые 30 секунд
chrome.alarms.create("checkEvents", { periodInMinutes: 0.5 });

chrome.alarms.onAlarm.addListener((alarm) => {
    if (alarm.name === "checkEvents") {
        console.log("Alarm triggered: checking events...");
        // Проверяем и обновляем данные
        const now = new Date();
        events = events.filter(event => new Date(event.start) > now); // Удаляем прошедшие события
        scheduleNextBadgeTime();
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

