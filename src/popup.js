import { Calendar } from '@fullcalendar/core';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import interactionPlugin from '@fullcalendar/interaction';

 import './styles.css'; // Импорт стилей


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
                chrome.storage.local.set({ token: accessToken }, () => {
                    resolve(accessToken);
                });
            }
        );
    });
}

function showEventDetails(event) {
    // Элементы модального окна
    const modal = document.getElementById('modal');
    const modalBody = document.getElementById('modalBody');

    // Проверяем, если location является URL, либо создаем ссылку для Google Maps
    let locationLink = '';
    let description = event.extendedProps.description || '';

    if (event.extendedProps.location) {
        const isUrl = event.extendedProps.location.startsWith('http://') || event.extendedProps.location.startsWith('https://');
        locationLink = isUrl
            ? `<a href="${event.extendedProps.location}" target="_blank" rel="noopener noreferrer">${event.extendedProps.location}</a>`
            : `<a href="https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(event.extendedProps.location)}" target="_blank" rel="noopener noreferrer">${event.extendedProps.location}</a>`;
    }
    // Наполнение модального окна данными события
    modalBody.innerHTML = `
        <h3>${event.title}</h3>
        <p><strong>Start:</strong> ${event.start.toLocaleString()}</p>
        ${event.end ? `<p><strong>End:</strong> ${event.end.toLocaleString()}</p>` : ''}
        ${locationLink ? `<p><strong>Location:</strong> ${locationLink}</p>` : ''}
        ${description ? `${description}</p>` : ''}
    `;

    // Отображение модального окна
    modal.classList.remove('hidden');
    modal.style.display = 'block';
}

function extractTextFromHTML(htmlString) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(htmlString, "text/html");
    return doc.body.textContent || ""; // Возвращает текст без HTML-тегов
}

function getMonthDateRange(date) {
    const startDate = new Date(date.getFullYear(), date.getMonth(), 1);
    const endDate = new Date(date.getFullYear(), date.getMonth() + 2, 0, 23, 59, 59); // Конец месяца
    return { startDate, endDate };
}


async function fetchEvents(startDate, endDate) {
    const tokenData = await new Promise(resolve => {
        chrome.storage.local.get("token", resolve);
    });

    let accessToken = tokenData?.token;

    if (!accessToken) {
        try {
            accessToken = await getAccessToken();
        } catch (error) {
            console.error("Authorization failed:", error);
            document.getElementById("events").innerText = "Authorization required.";
            return;
        }
    }

    try {
        const events = []; // Список для всех событий
        let url = `https://graph.microsoft.com/v1.0/me/calendar/events?$filter=start/dateTime ge '${startDate.toISOString()}' and end/dateTime le '${endDate.toISOString()}'`;

        do {
            let response = await fetch(url, {
                headers: {
                    "Authorization": `Bearer ${accessToken}`
                }
            });

            if (response.status === 401) {
                try {
                    accessToken = await getAccessToken();
                    response = await fetch(url, {
                        headers: {
                            "Authorization": `Bearer ${accessToken}`
                        }
                    });
                } catch (error) {
                    console.error("Authorization failed:", error);
                    document.getElementById("events").innerText = "Authorization required.";
                    return;
                }
            }

            if (!response.ok) {
                throw new Error(`Error fetching events: ${response.status}`);
            }

            const data = await response.json();
            if (data.value) {
                events.push(...data.value); // Добавляем события из текущей страницы
            }

            url = data["@odata.nextLink"]; // Устанавливаем URL для следующей страницы

        } while (url); // Продолжаем пока есть @odata.nextLink

        if (events.length > 0) {
            // Преобразуем события в формат FullCalendar
            const fullCalendarData = [];
            for (const event of events) {
                if (event.recurrence) {
                    // Обработка повторяющихся событий
                    const occurrences = expandRecurringEvent(event, startDate, endDate);
                    fullCalendarData.push(...occurrences);
                } else {
                    // Обычные события
                    const start = new Date(event.start.dateTime + 'Z');
                    const end = new Date(event.end.dateTime + 'Z');
                    fullCalendarData.push({
                        title: event.subject,
                        start: start,
                        end: end,
                        location: event.location?.displayName || '',
                        description: event.body?.content || ''
                    });
                }
            }

            renderCalendar(fullCalendarData);
            const notifyEvents = fullCalendarData.map(event => {
                return {
                    title: event.title,
                    start: event.start,
                    end: event.end,
                    location: event.location,
                    description: extractTextFromHTML(event.description)
                };
            });
            chrome.runtime.sendMessage({
                action: "setEvents", notifyEvents
            }, response => {
                console.log(response.status);
            });
        } else {
            const eventsContainer = document.getElementById("calendar");
            eventsContainer.innerText = "No upcoming events.";
        }
    } catch (error) {
        console.error("Error fetching events:", error);
        document.getElementById("events").innerText = "Error fetching events.";
    }
}

function expandRecurringEvent(event, rangeStart, rangeEnd) {
    const occurrences = [];
    const rule = event.recurrence.pattern;
    const range = event.recurrence.range;

    const recurrenceStart = new Date(range.startDate + 'T00:00:00Z');
    const recurrenceEnd = new Date(range.endDate + 'T23:59:59Z');

    if (recurrenceEnd < rangeStart || recurrenceStart > rangeEnd) {
        return occurrences; // Диапазоны не пересекаются
    }

    let currentDate = new Date(recurrenceStart);

    while (currentDate <= recurrenceEnd && currentDate <= rangeEnd) {
        if (currentDate >= rangeStart) {
            const dayOfWeek = currentDate.toLocaleString('en-US', { weekday: 'long' }).toLowerCase();
            if (rule.daysOfWeek.includes(dayOfWeek)) {
                const startRecurrence = new Date(event.start.dateTime + 'Z');
                const endRecurrence = new Date(event.end.dateTime + 'Z');
                const start = new Date(currentDate + 'Z');
                const end = new Date(currentDate + 'Z');
                start.setHours(startRecurrence.getHours());
                end.setHours(endRecurrence.getHours());
                occurrences.push({
                    title: event.subject,
                    start: start,
                    end: end,
                    location: event.location?.displayName || '',
                    description: event.body?.content || ''
                });
            }
        }

        switch (rule.type) {
            case "daily":
                currentDate.setDate(currentDate.getDate() + rule.interval);
                break;
            case "weekly":
                currentDate.setDate(currentDate.getDate() + 1); // Переход на следующий день
                break;
            case "absoluteMonthly":
                currentDate.setMonth(currentDate.getMonth() + rule.interval);
                break;
            case "absoluteYearly":
                currentDate.setFullYear(currentDate.getFullYear() + rule.interval);
                break;
            default:
                console.warn("Unknown recurrence type:", rule.type);
                return occurrences;
        }
    }

    return occurrences;
}


function renderCalendar(events) {
    const calendarEl = document.getElementById('calendar');
    const calendar = new Calendar(calendarEl, {
        plugins: [dayGridPlugin, timeGridPlugin, interactionPlugin],
        initialView: 'timeGridDay',
        selectable: true,            // Включаем возможность выделения
        editable: true,              // Позволяем перемещать и изменять события
        height: 'auto', // Убедитесь, что высота адаптируется
        contentHeight: 'auto',
        headerToolbar: {
            left: 'prev,next',
            center: 'title',
            right: 'timeGridDay,timeGridWeek,dayGridMonth'
        },
        // todo next release
        // select: function(info) {
        //     console.info("Start (raw): ", info.start);
        //     console.info("Start (UTC): ", info.start.toISOString());
        //
        //     // Форматируем время с учетом локальной таймзоны
        //     const eventLocalData = {
        //         start: info.start,
        //         end: info.end
        //     };
        //
        //     openEventModal(eventLocalData); // Открываем модальное окно
        //     calendar.unselect(); // Сбрасываем выделение
        // },
        events: events,
        timeZone: 'local',
        dayMaxEvents: 3,
        dateClick: function(info) {
            calendar.changeView('timeGridDay', info.dateStr);
        },
        datesSet: function () {
            // Вызываем при смене дат
            setTimeout(() => scrollToMiddle(), 0);
        },
        eventClick: function(info) {
            showEventDetails(info.event);
            /*
            const event = info.event;
            const eventData = {
                id: event.id, // Уникальный идентификатор события
                title: event.title,
                start: event.start, // Формат для datetime-local
                end: event.end,
                location: event.extendedProps.location || '',
                description: event.extendedProps.description || ''
            };

            openEventModal(eventData); // Открываем модальное окно для редактирования
             */
        },
        views: {
            dayGridMonth: {
                selectable: false // Отключаем выделение для dayGridMonth
            }
        }
    });

    calendar.render();
}

function formatDateForDatetimeLocal(date) {
    // Преобразуем дату в формат "YYYY-MM-DDTHH:mm" с учетом локальной таймзоны
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Месяц от 0 до 11
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${year}-${month}-${day}T${hours}:${minutes}`;
}


function scrollToMiddle() {
    const scroller = document.querySelector('#calendar-container');
    if (scroller) {
        scroller.scrollTop = (scroller.scrollHeight - scroller.clientHeight) / 1.5;
    }
}

// Добавить обработчик для закрытия модального окна
document.getElementById('closeModal').addEventListener('click', () => {
    const modal = document.getElementById('modal');
    modal.classList.add('hidden');
    modal.style.display = 'none';
});

// Закрытие модального окна при клике вне его содержимого
window.addEventListener('click', (event) => {
    const modal = document.getElementById('modal');
    if (event.target === modal) {
        modal.classList.add('hidden');
        modal.style.display = 'none';
    }
});

document.getElementById("logoutButton").addEventListener("click", () => {
    chrome.runtime.sendMessage({ action: "logout" });
});

// document.getElementById("loginButton").addEventListener("click", () => {
//     chrome.runtime.sendMessage({ action: "login" });
// });

async function createEvent(eventData) {
    const tokenData = await new Promise(resolve => {
        chrome.storage.local.get("token", resolve);
    });

    let accessToken = tokenData?.token;

    if (!accessToken) {
        accessToken = await getAccessToken();
    }

    const event = {
        subject: eventData.title,
        start: {
            dateTime: new Date(eventData.start).toISOString(),
            timeZone: "UTC"
        },
        end: {
            dateTime: new Date(eventData.end).toISOString(),
            timeZone: "UTC"
        },
        location: {
            displayName: eventData.location
        },
        body: {
            contentType: "HTML",
            content: eventData.description
        }
    };

    const response = await fetch("https://graph.microsoft.com/v1.0/me/events", {
        method: "POST",
        headers: {
            "Authorization": `Bearer ${accessToken}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify(event)
    });

    if (!response.ok) {
        const errorResponse = await response.json();
        console.error("Error response:", errorResponse);
        throw new Error(`Error creating event: ${response.statusText}`);
    }
}

/*
async function updateEvent(eventId, eventData) {
    // Реализация обновления события (например, через API)
    console.log('Updating event:', eventId, eventData);
    // Здесь вызовите API или обновите локальные данные
}
 */

function openEventModal(eventLocalData = {}) {
    const modal = document.getElementById('modal');
    const modalBody = document.getElementById('modalBody');
    const eventData = {
        id: eventLocalData.id, // Уникальный идентификатор события
        title: eventLocalData.title,
        start: eventLocalData.start.toISOString().slice(0, 16), // Формат для datetime-local
        end: eventLocalData.end.toISOString().slice(0, 16),
        location: eventLocalData.location || '',
        description: eventLocalData.description || ''
    };
    const localDate = {
        start: formatDateForDatetimeLocal(eventLocalData.start), // Локальное время
        end: formatDateForDatetimeLocal(eventLocalData.end)     // Локальное время
    };
    // Наполнение модального окна формой
    modalBody.innerHTML = `
        <form id="createEventForm">
            <h3>${eventData.id ? 'Edit Event' : 'Create New Event'}</h3>
            <div class="createEventForm-row">
                <label for="eventTitle">Title:</label>
                <input type="text" id="eventTitle" name="eventTitle" value="${eventData.title || ''}" required>
            </div>
            <div class="createEventForm-row">
                <label for="eventStart">Start:</label>
                <input type="datetime-local" id="eventStart" name="eventStart" value="${localDate.start || ''}" required>
            </div>
            <div class="createEventForm-row">
                <label for="eventEnd">End:</label>
                <input type="datetime-local" id="eventEnd" name="eventEnd" value="${localDate.end || ''}" required>
            </div>
            <div class="createEventForm-row">
                <label for="eventLocation">Location:</label>
                <input type="text" id="eventLocation" name="eventLocation" value="${eventData.location || ''}">
            </div>
            <div class="createEventForm-row">
                <label for="eventDescription">Description:</label>
                <textarea id="eventDescription" name="eventDescription">${eventData.description || ''}</textarea>
            </div>
            <button type="submit">${eventData.id ? 'Update Event' : 'Create Event'}</button>
        </form>
    `;

    modal.classList.remove('hidden');
    modal.style.display = 'block';

    document.getElementById('createEventForm').addEventListener('submit', async (e) => {
        e.preventDefault();

        const title = document.getElementById('eventTitle').value;
        const start = document.getElementById('eventStart').value;
        const end = document.getElementById('eventEnd').value;
        const location = document.getElementById('eventLocation').value;
        const description = document.getElementById('eventDescription').value;

        try {
            if (eventData.id) {
                // Логика для обновления события
                await updateEvent(eventData.id, { title, start, end, location, description });
            } else {
                // Логика для создания нового события
                await createEvent({ title, start, end, location, description });
            }

            modal.classList.add('hidden');
            modal.style.display = 'none';
            const { startDate, endDate } = getMonthDateRange(eventLocalData.start);
            fetchEvents(startDate, endDate); // Перезагружаем события в календаре
        } catch (error) {
            console.error('Error processing event:', error);
        }
    });
}

document.getElementById('createEventButton').addEventListener('click', () => {
    openEventModal(); // Открытие пустого модального окна для создания события
});



document.addEventListener("DOMContentLoaded", () => {
    const now = new Date();
    const { startDate, endDate } = getMonthDateRange(now);
    fetchEvents(startDate, endDate);
});

