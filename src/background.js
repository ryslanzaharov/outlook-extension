import { getAccessToken } from './token.js';

chrome.runtime.onInstalled.addListener(() => {
    console.log("Outlook Calendar Viewer installed.");
});

function formatTo12Hour(date) {
    const options = {
        hour: 'numeric',
        hour12: true
    };
    return date.toLocaleString('en-US', options);
}

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

// Устанавливаем события и планируем уведомления
async function setNewEvents(newEvents) {
    chrome.storage.local.remove("todayEvents", () => {
        console.log("todayEvents удалён из chrome.storage.local");
    });
    const now = new Date();

    // Фильтруем события за текущий день
    let events = newEvents
        .filter(event => isToday(new Date(event.start)))
        .filter(event => new Date(event.start) > now) // Исключаем прошедшие события
        .sort((a, b) => new Date(a.start) - new Date(b.start)); // Сортируем по времени начала

    if (events.length === 0) {
        console.log("No upcoming events for today.");
        chrome.action.setBadgeText({ text: "" }); // Очищаем значок
        return;
    }
    chrome.storage.local.set({ todayEvents: events });

    await scheduleNextBadgeTime();
    // Планируем уведомления
    await scheduleEventNotifications();
}

async function getTodayEvents() {
    return new Promise((resolve) => {
        chrome.storage.local.get("todayEvents", (data) => {
            resolve(data.todayEvents || []); // Исправлено: data.todayEvents
        });
    });
}

function clearOldAlarms(callback) {
    chrome.alarms.getAll((alarms) => {
        let count = alarms.length;
        if (count === 0) {
            if (callback) callback(); // Если будильников нет, вызываем колбэк
            return;
        }

        alarms.forEach((alarm) => {
            if (alarm.name.startsWith("event_")) {
                chrome.alarms.clear(alarm.name, () => {
                    console.log(`Будильник удалён: ${alarm.name}`);
                    count--;
                    if (count === 0 && callback) callback(); // Когда все удалены, вызываем колбэк
                });
            }
        });
    });
}

// Планируем уведомления для события
async function scheduleEventNotifications() {
    clearOldAlarms();
    chrome.storage.local.remove("scheduledEvents", () => {
        console.log("scheduledEvents удалён из chrome.storage.local");
    });

    let events = await getTodayEvents();
    const now = Date.now();
    let scheduledEvents = {};
    events.forEach(event => {
        const eventTime = new Date(event.start).getTime();

        if (eventTime <= now) return; // Пропускаем прошедшие события

        // Уведомление за 15 минут
        const notify15Time = eventTime - 15 * 60 * 1000;
        if (notify15Time > now) {
            const alarmName15 = `event_${event.title}_15`;
            chrome.alarms.create(alarmName15, { when: notify15Time });
            scheduledEvents[alarmName15] = event; // Добавляем в мапу
        }

        // Уведомление за 0 минут
        if (eventTime > now) {
            const alarmName1 = `event_${event.title}_1`;
            chrome.alarms.create(alarmName1, { when: eventTime });
            scheduledEvents[alarmName1] = event; // Добавляем в мапу
        }
    });
    chrome.storage.local.set({ scheduledEvents }, () => {
        console.log("Все события сохранены в scheduledEvents", scheduledEvents);
    });
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
async function scheduleNextBadgeTime() {
    const now = new Date();
    let events = await getTodayEvents();
    const nextEvent = events.find(event => new Date(event.start) > now);

    if (nextEvent) {
        const eventTime = new Date(nextEvent.start);
        const timeUntilEvent = eventTime - now;

        if (timeUntilEvent <= 30 * 60 * 1000) {
            updateBadgeTime(nextEvent); // Обновляем значок за 30 мин
        } else {
            chrome.action.setBadgeText({text: formatTo12Hour(eventTime)});
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
        chrome.action.setBadgeBackgroundColor({color: "#336dff"}); // Цвет значка
    } else {
        chrome.action.setBadgeText({text: ""}); // Очищаем значок
    }
}

// Показываем уведомление
function showNotification(event, timeLabel) {
    const eventTime = new Date(event.start).toLocaleTimeString(navigator.language, {
        hour: 'numeric',
        minute: 'numeric',
        hour12: true
    });

    // Проверяем, существует ли уже окно уведомления (опционально)
    chrome.windows.getAll({ populate: true }, (windows) => {
        const existingPopup = windows.find(win => win.type === 'popup' && win.tabs[0].url.includes('notification.html'));
        if (existingPopup) {
            chrome.windows.update(existingPopup.id, { focused: true }); // Фокусируем существующее окно
            return;
        }
        console.log("notify event.outlookUrl", event.outlookUrl);
        // Создаем новое окно уведомления
        chrome.windows.create({
            url: `notification.html?title=${encodeURIComponent(event.title)}&time=${encodeURIComponent(eventTime)}&location=${encodeURIComponent(event.location)}&outlookUrl=${encodeURIComponent(event.outlookUrl)}`,
            type: "popup",
            width: 400,
            height: 230,
            focused: true
        }, (window) => {
            if (chrome.runtime.lastError) {
                console.error("Ошибка при создании окна: ", chrome.runtime.lastError);
            } else {
                console.log(`Notification window created: ${timeLabel} до события.`);
            }
        });
    });
}

function createSyncedAlarm() {
    const now = Date.now();
    const nextMinute = Math.ceil(now / 60000) * 60000; // Округляем до следующей минуты

    chrome.alarms.create("checkEvents", { when: nextMinute, periodInMinutes: 1 });
}

chrome.alarms.onAlarm.addListener((alarm) => {
    if (alarm.name === "checkEvents") {
        scheduleNextBadgeTime();
    } else if (alarm.name.startsWith("event_")) {
        chrome.storage.local.get("scheduledEvents", (data) => {
            console.log(`Notification data` + data + " alarm.name " + alarm.name);
            const event = data.scheduledEvents?.[alarm.name];
            if (event) {
                const timeLabel = alarm.name.includes("_15") ? "15 минут" : "1 минута";
                showNotification(event, timeLabel);
            }
        });
    }
});

// Получаем события через сообщение
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.action === "setEvents") {
        setNewEvents(message.notifyEvents);
        sendResponse({status: "Events set and monitoring started"});
    } else if (message.action === "logout") {
        logoutMicrosoft();
        sendResponse({status: "success"});
    } else if (message.action === 'closeWindow' && sender.tab) {
        chrome.windows.remove(sender.tab.windowId);
    } else if (message.action === "authorization") {
        getAccessToken().then(token => {
            sendResponse({ success: true, token });
        }).catch(error => {
            sendResponse({ success: false, error: error.message });
        });
    }  else if (message.action === 'log') {
        console.log("log message", message);
    }
});

createSyncedAlarm();

