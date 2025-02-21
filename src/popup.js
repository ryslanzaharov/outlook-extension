import { Calendar } from '@fullcalendar/core';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import interactionPlugin from '@fullcalendar/interaction';
import flatpickr from "flatpickr";
import "flatpickr/dist/flatpickr.min.css";

import './styles.css'; // Импорт стилей
import { getStorageAccessToken } from './token.js';

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
        <br>
        <div class="event-row">
            <img src="./images/time-18.png" alt="Time" title="Time">
            <span>${startFormatTo12Hour(event.start)} - ${endFormatTo12Hour(event.end)}</span>
        </div>
        <div class="event-row">
            <img src="./images/location-18.png" alt="Location" title="Location"><span>${locationLink}</span>
        </div>
        <div class="event-row">
            <img src="./images/text-18.png" alt="Description" title="Description"><span>${description}</span>
        </div>
        `;

    // Отображение модального окна
    modal.classList.remove('hidden');
    modal.style.display = 'block';
}

function startFormatTo12Hour(date) {
    const monthDay = date.toLocaleString('en-US', { month: 'long', day: 'numeric' });
    const time = date.toLocaleString('en-US', { hour: 'numeric', minute: 'numeric', hour12: true });
    return `${monthDay}, ${time}`;
}

function endFormatTo12Hour(date) {
    return date.toLocaleString('en-US', { hour: 'numeric', minute: 'numeric', hour12: true });
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

function subtractMonths(date, months) {
    let result = new Date(date);
    result.setMonth(result.getMonth() - months);
    return result;
}

async function getAllEvents() {
    return new Promise((resolve) => {
        chrome.storage.local.get("events", (data) => {
            resolve(Array.isArray(data.events) ? data.events : []); // Гарантируем, что вернётся массив
        });
    });
}

async function fetchWithAuth(url, accessToken) {
    let response = await fetch(url, {
        headers: { "Authorization": `Bearer ${accessToken}` }
    });
    return response;
}


async function fetchEventsFromGraph(url, accessToken, events) {
    do {
        let response = await fetchWithAuth(url, accessToken);
        if (!response.ok) throw new Error(`Error fetching events: ${response.status}`);

        const data = await response.json();
        if (data.value) {
            events.push(...data.value);
        }

        url = data["@odata.nextLink"]; // Следующая страница данных
    } while (url);
}

async function fetchEvents(isSync, startDate, endDate) {
    try {
        let allEvents = await getAllEvents();
        console.log("allEvents", allEvents);
        if (allEvents.length === 0 || isSync) {
            console.log("get events");
            //todo тут скорее всего в кэш не успевает сохраниться
            let accessToken = await getStorageAccessToken();
            if (!accessToken) {
                chrome.runtime.sendMessage({ action: "authorization" }, (response) => {
                    if (response.success) {
                        console.log("Success");
                    } else {
                        console.error("Fail", response.error);
                    }
                });
                accessToken = await getStorageAccessToken();
            }

            const events = [];

            // 1. Загружаем обычные события за указанный диапазон
            const normalEventsUrl = `https://graph.microsoft.com/v1.0/me/calendar/events?$filter=start/dateTime ge '${startDate.toISOString()}' and end/dateTime le '${endDate.toISOString()}'`;
            await fetchEventsFromGraph(normalEventsUrl, accessToken, events);

            // 2. Загружаем повторяемые события (seriesMaster) за последние 2 года, исключая текущий день
            const recurringUrl = `https://graph.microsoft.com/v1.0/me/calendar/events?$filter=type eq 'seriesMaster' and start/dateTime lt '${startDate.toISOString()}' and start/dateTime ge '${subtractMonths(startDate, 24).toISOString()}'`;
            await fetchEventsFromGraph(recurringUrl, accessToken, events);
            chrome.storage.local.set({ events: events });
            allEvents = events;
        }
        processEvents(allEvents, startDate, endDate);
    } catch (error) {
        console.error("Error fetching events:", error);
        document.getElementById("events").innerText = "Error fetching events.";
    }
}

function processEvents(events, startDate, endDate) {
    if (events.length === 0) {
        document.getElementById("calendar").innerText = "No upcoming events.";
        return;
    }

    const fullCalendarData = events.flatMap(event => {
        if (event.recurrence) {
            return expandRecurringEvent(event, startDate, endDate);
        } else {
            return [{
                title: event.subject,
                start: new Date(event.start.dateTime + 'Z'),
                end: new Date(event.end.dateTime + 'Z'),
                location: event.location?.displayName || '',
                description: event.body?.content || ''
            }];
        }
    });

    renderCalendar(fullCalendarData);
    sendEventNotifications(fullCalendarData);
}

function sendEventNotifications(fullCalendarData) {
    const notifyEvents = fullCalendarData.map(event => ({
        title: event.title,
        start: event.start,
        end: event.end,
        location: event.location,
        description: extractTextFromHTML(event.description)
    }));

    chrome.runtime.sendMessage({ action: "setEvents", notifyEvents }, response => {
        console.log("Notification status:", response.status);
    });
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


function rightButtons() {
    // Добавляем иконку support в кнопку
    const supportButton = document.querySelector('.fc-icons-button');
    if (supportButton) {
        supportButton.innerHTML =
            '   <div class="icons">\n' +
            '    <button id="updateButton" class="update-button">\n' +
            '        <img src="./images/updating.png" alt="Update" title="Update" class="icon">\n' +
            '    </button>\n' +
            '        <a id="support" class="link" href="https://mail.google.com/mail/u/0/?view=cm&fs=1&to=ruslan.ext.dev@gmail.com&su=Outlook%20Calendar%20Checker&body=Hello,%20I%20would%20like%20to%20suggest%20you%20to%20do" target="_blank">\n' +
            '            <img src="./images/support.png" alt="Support" title="Support" class="icon">\n' +
            '        </a>\n' +
            '        <a id="owaCalendar" class="link" href="https://outlook.live.com/calendar/0/view/day" target="_blank">\n' +
            '        <img src="./images/external-link.png" alt="Calendar" title="Open OWA" class="icon">\n' +
            '    </a>\n' +
            '    <button id="logoutButton" class="logout-button">\n' +
            '        <img src="./images/logout.png" alt="Logout" title="Logout" class="icon">\n' +
            '    </button>\n' +
            '    </div>';
    }
    document.getElementById("logoutButton").addEventListener("click", () => {
        chrome.runtime.sendMessage({ action: "logout" });
    });
    document.getElementById("updateButton").addEventListener("click", () => {
        chrome.storage.local.remove("token", () => {});
        const now = new Date();
        const { startDate, endDate } = getMonthDateRange(now);
        fetchEvents(true, startDate, endDate);
    });
    loadBar();
}

function viewButtons(toolbar, calendar) {
    // Меняем кнопку на иконку
    const viewButton = toolbar.querySelector('.fc-viewToggleButton-button');
    if (viewButton) {
        viewButton.innerHTML = '<img src="./images/view.png" alt="View" class="view-icon">';
    }
    const nameButton = toolbar.querySelector('.fc-name-button');
    if (nameButton) {
        nameButton.innerHTML = `
        <img src="./images/31-24.png" alt="Calendar" class="calendar-icon">
        <span>Outlook Calendar Checker</span>
    `;
        nameButton.style.display = 'flex';
        nameButton.style.alignItems = 'center';
        nameButton.style.gap = '5px'; // Добавляем небольшой отступ между иконкой и текстом
    }

    const dropdown = document.createElement('div');
    dropdown.id = 'viewDropdown';
    dropdown.classList.add('dropdown-menu');

    // Отображаемые названия и соответствующие представления
    const views = {
        'Day': 'timeGridDay',
        'Week': 'timeGridWeek',
        'Month': 'dayGridMonth'
    };

    Object.entries(views).forEach(([label, view]) => {
        const option = document.createElement('div');
        option.textContent = label;
        option.classList.add('dropdown-item');
        option.onclick = function () {
            calendar.changeView(view);
            dropdown.classList.remove('show'); // Закрываем меню
        };
        dropdown.appendChild(option);
    });

    toolbar.appendChild(dropdown);

    // Закрываем меню при клике вне него
    document.addEventListener('click', function (event) {
        if (!toolbar.contains(event.target) && !event.target.classList.contains('fc-button')) {
            dropdown.classList.remove('show');
        }
    });
}

function renderCalendar(events) {
    const calendarEl = document.getElementById('calendar');
    const calendar = new Calendar(calendarEl, {
        plugins: [dayGridPlugin, timeGridPlugin, interactionPlugin],
        initialView: 'timeGridDay',
        height: 'auto',
        contentHeight: 'auto',
        timeZone: 'local',
        dayMaxEvents: 3,
        customButtons: {
            viewToggleButton: {
                text: '', // Убираем текст
                click: function() {
                    const dropdown = document.getElementById('viewDropdown');
                    dropdown.classList.toggle('show'); // Открываем/закрываем меню
                }
            },
            name: {
                text: ''
            },
            icons: {
                text: ''
            },
        },
        headerToolbar: {
            left: 'viewToggleButton,name, prev,next',
            center: 'title',
            right: 'icons' // Добавляем кнопку support справа
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
        dateClick: function(info) {
            calendar.changeView('timeGridDay', info.dateStr);
        },
        datesSet: function (info) {
            const viewType = info.view.type;
            if (viewType !== 'dayGridMonth') {
                // Вызываем при смене дат
                setTimeout(() => scrollToMiddle(), 0);
            }
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
                selectable: false
            }
        },
        events: events
    });

    calendar.render();

    // Добавляем выпадающее меню
    const toolbar = document.querySelector('.fc-toolbar-chunk:first-child');
    viewButtons(toolbar, calendar);
    rightButtons();
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

// document.getElementById("loginButton").addEventListener("click", () => {
//     chrome.runtime.sendMessage({ action: "login" });
// });

async function createEvent(eventData) {
    let accessToken = await getStorageAccessToken();

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
//todo next release
    // const createdEvent = await response.json(); // Получаем данные созданного события
    //
    // // 1. Загружаем существующие события из памяти
    // const existingEvents = await getAllEvents();
    // const updatedEvents = [...existingEvents, createdEvent];
    //
    // chrome.storage.local.set({ events: updatedEvents }, () => {
    //     console.log("Event added to local storage:", createdEvent);
    // });
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
        id: eventLocalData.id || '',
        title: eventLocalData.title || '',
        start: eventLocalData.start || null,
        end: eventLocalData.end || null,
        location: eventLocalData.location || '',
        description: eventLocalData.description || ''
    };

    // Наполнение модального окна формой
    modalBody.innerHTML = `
        <form id="createEventForm">
            <p></p>
            <div class="createEventForm-row">
                <label for="eventTitle">Title:</label>
                <input type="text" id="eventTitle" name="eventTitle" value="${eventData.title}" required>
            </div>
            <div class="createEventForm-row">
                <label for="eventStart">Start:</label>
                <input type="text" id="eventStart" name="eventStart" placeholder="Select start date and time" required>
            </div>
            <div class="createEventForm-row">
                <label for="eventEnd">End:</label>
                <input type="text" id="eventEnd" name="eventEnd" placeholder="Select end date and time" required>
            </div>
            <div class="createEventForm-row">
                <label for="eventLocation">Location:</label>
                <input type="text" id="eventLocation" name="eventLocation" value="${eventData.location}">
            </div>
            <div class="createEventForm-row">
                <label for="eventDescription">Description:</label>
                <textarea id="eventDescription" name="eventDescription">${eventData.description}</textarea>
            </div>
            <button type="submit">${eventData.id ? 'Update Event' : 'Create Event'}</button>
        </form>
    `;

    modal.classList.remove('hidden');
    modal.style.display = 'block';

    // Инициализация Flatpickr для полей start и end
    flatpickr("#eventStart", {
        enableTime: true,
        dateFormat: "Y-m-d H:i", // ISO формат (чтобы было совместимо с сервером)
        defaultDate: eventData.start || null, // Устанавливаем начальное значение
        time_24hr: false // 12-часовой формат с AM/PM
    });

    flatpickr("#eventEnd", {
        enableTime: true,
        dateFormat: "Y-m-d H:i",
        defaultDate: eventData.end || null,
        time_24hr: false
    });
    // updateAccessTokenByWorker();
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
            const { startDate, endDate } = getMonthDateRange(new Date(start));
            fetchEvents(true, startDate, endDate); // Перезагружаем события в календаре
        } catch (error) {
            console.error('Error processing event:', error);
        }
    });
}

document.getElementById('createEventButton').addEventListener('click', () => {
    openEventModal(); // Открытие пустого модального окна для создания события
});

function loadBar() {
        const updateButton = document.getElementById("updateButton");
        const loadingBar = document.getElementById("loading-bar");

        if (updateButton) {
            updateButton.addEventListener("click", function () {
                showLoadingBar();

                // Здесь можно запустить обновление данных, например, через setTimeout имитируем процесс
                setTimeout(() => {
                    hideLoadingBar();
                }, 3000); // Имитация загрузки (замените на реальный вызов)
            });
        }

        function showLoadingBar() {
            loadingBar.style.width = "100%";
        }

        function hideLoadingBar() {
            setTimeout(() => {
                loadingBar.style.width = "0";
            }, 500);
        }

}


document.addEventListener("DOMContentLoaded", () => {
    const now = new Date();
    const { startDate, endDate } = getMonthDateRange(now);
    fetchEvents(false, startDate, endDate);
});

