import { Calendar } from '@fullcalendar/core';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import interactionPlugin from '@fullcalendar/interaction';
import flatpickr from "flatpickr";
import "flatpickr/dist/flatpickr.min.css";

import './styles.css'; // Импорт стилей
import { getStorageAccessToken } from './token.js';

function showEventDetails(event) {
    console.log("Event ID for Outlook link:", event.id);
    const modal = document.getElementById('modal');
    const modalBody = document.getElementById('modalBody');

    let locationLink = '';
    let description = event.extendedProps.description || '';
    let attendees = event.extendedProps.attendees || [];

    // Проверяем и форматируем location
    if (event.extendedProps.location) {
        const isUrl = event.extendedProps.location.startsWith('http://') || event.extendedProps.location.startsWith('https://');
        locationLink = isUrl
            ? `<a href="${event.extendedProps.location}" target="_blank" rel="noopener noreferrer">${event.extendedProps.location}</a>`
            : `<a href="https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(event.extendedProps.location)}" target="_blank" rel="noopener noreferrer">${event.extendedProps.location}</a>`;
    }

    // Обрабатываем ссылки в description
    if (description) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(description, 'text/html');
        const links = doc.getElementsByTagName('a');
        for (let link of links) {
            link.setAttribute('target', '_blank');
            link.setAttribute('rel', 'noopener noreferrer');
        }
        description = doc.body.innerHTML;
    }

    // Разделяем участников на обязательных и необязательных
    const requiredAttendees = attendees.filter(att => att.type === 'required');
    const optionalAttendees = attendees.filter(att => att.type === 'optional');

    // Форматируем списки участников с ограничением до 2 человек и добавлением скрытых элементов
    const formatAttendees = (attendeesList, type) => {
        if (attendeesList.length === 0) return '';

        const visibleAttendees = attendeesList.slice(0, 2).map(att => att.emailAddress.name || att.emailAddress.address).join(', ');
        const hiddenAttendees = attendeesList.slice(2).map(att => att.emailAddress.name || att.emailAddress.address).join(', ');

        if (hiddenAttendees) {
            return `
                <span class="attendees-list" data-full-list="${visibleAttendees}, ${hiddenAttendees}">
                    ${visibleAttendees}
                    <span class="more-indicator" style="color: #007bff; cursor: pointer;"> (+${attendeesList.length - 2} more)</span>
                    <span class="hidden-attendees" style="display: none;">, ${hiddenAttendees}</span>
                </span>
            `;
        }
        return visibleAttendees;
    };

    const requiredAttendeesList = formatAttendees(requiredAttendees, 'required');
    const optionalAttendeesList = formatAttendees(optionalAttendees, 'optional');

    // Формируем URL для события в Outlook
    const outlookUrl = `https://outlook.live.com/calendar/0/item/${encodeURIComponent(event.id)}`;

    // Наполнение модального окна данными события с добавлением ссылки
    modalBody.innerHTML = `
        <div class="outlook-link">
            <a href="${outlookUrl}" target="_blank" rel="noopener noreferrer">
                <img src="./images/external-link.png" alt="Open in Outlook" title="Open in Outlook">
            </a>
        </div>
        <h3>${event.title}</h3>
        <br>
        <div class="event-row">
            <img src="./images/time-18.png" alt="Time" title="Time">
            <span>${startFormatTo12Hour(event.start)} - ${endFormatTo12Hour(event.end)}</span>
        </div>
        ${locationLink ? `
            <div class="event-row">
                <img src="./images/location-18.png" alt="Location" title="Location">
                <span>${locationLink}</span>
            </div>
        ` : ''}
                ${requiredAttendeesList ? `
        <div class="event-row attendees-row" data-type="required">
            <img src="./images/invite_required-18.png" alt="Required Attendees" title="Required Attendees">
            <span>${requiredAttendeesList}</span>
        </div>
        ` : ''}
        ${optionalAttendeesList ? `
            <div class="event-row attendees-row" data-type="optional">
                <span>${optionalAttendeesList}</span>
            </div>
        ` : ''}
        ${description ? `
            <div class="event-row description-container">
                <img src="./images/text-18.png" alt="Description" title="Description">
                <span>${description}</span>
            </div>
        ` : ''}
    `;

    // Добавляем обработчик клика для показа/скрытия дополнительных участников
    document.querySelectorAll('.attendees-row').forEach(row => {
        row.addEventListener('click', function(e) {
            if (e.target.classList.contains('more-indicator') || e.target.tagName === 'SPAN') {
                const attendeesList = this.querySelector('.attendees-list');
                const hiddenAttendees = attendeesList.querySelector('.hidden-attendees');
                const moreIndicator = attendeesList.querySelector('.more-indicator');

                if (hiddenAttendees.style.display === 'none') {
                    hiddenAttendees.style.display = 'inline';
                    moreIndicator.textContent = ' (hide)';
                } else {
                    hiddenAttendees.style.display = 'none';
                    moreIndicator.textContent = ` (+${attendeesList.dataset.fullList.split(',').length - 2} more)`;
                }
            }
        });
    });

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

// // Перехватываем console.log
// const originalConsoleLog = console.log;
// console.log = function (...args) {
//     const message = args.map(arg => String(arg)).join(' ');
//     const timestamp = new Date().toISOString();
//     const logEntry = `[${timestamp}] INFO: ${message}`;
//     chrome.runtime.sendMessage({ action: 'log', logEntry });
//     originalConsoleLog.apply(console, args); // Сохраняем вывод в консоль popup
// };
//
// // Перехватываем console.error
// const originalConsoleError = console.error;
// console.error = function (...args) {
//     const message = args.map(arg => String(arg)).join(' ');
//     const timestamp = new Date().toISOString();
//     const logEntry = `[${timestamp}] ERROR: ${message}`;
//     chrome.runtime.sendMessage({ action: 'log', logEntry });
//     originalConsoleError.apply(console, args);
// };

async function fetchEvents(isSync, startDate, endDate) {
    try {
        let allEvents = await getAllEvents();
        console.log("allEvents", allEvents);
        if (allEvents.length === 0 || isSync) {
            console.log("get events");
            //todo тут скорее всего в кэш не успевает сохраниться
            let accessToken = await getStorageAccessToken();
            const events = [];

            // 1. Загружаем обычные события
            const normalEventsUrl = `https://graph.microsoft.com/v1.0/me/calendar/events?$filter=start/dateTime ge '${startDate.toISOString()}' and end/dateTime le '${endDate.toISOString()}'`;
            await fetchEventsFromGraph(normalEventsUrl, accessToken, events);

            // 2. Загружаем повторяемые события
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

    const fullCalendarData = events.flatMap(event => {
        if (event.recurrence) {
            return expandRecurringEvent(event, startDate, endDate);
        } else {
            return [{
                id: event.id,
                title: event.subject,
                start: new Date(event.start.dateTime + 'Z'),
                end: new Date(event.end.dateTime + 'Z'),
                location: event.location?.displayName || '',
                description: event.body?.content || '',
                attendees: event.attendees || [] // Добавляем участников
            }];
        }
    });

    renderCalendar(fullCalendarData);
    sendEventNotifications(fullCalendarData);
}

function sendEventNotifications(fullCalendarData) {
    // Формируем URL для события в Outlook
    const notifyEvents = fullCalendarData.map(event => ({
        title: event.title,
        start: event.start,
        end: event.end,
        location: event.location,
        description: extractTextFromHTML(event.description),
        outlookUrl: `https://outlook.live.com/calendar/0/item/${encodeURIComponent(event.id)}`
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
        return occurrences;
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
                    id: event.id,
                    title: event.subject,
                    start: start,
                    end: end,
                    location: event.location?.displayName || '',
                    description: event.body?.content || '',
                    attendees: event.attendees || [] // Добавляем участников
                });
            }
        }

        switch (rule.type) {
            case "daily":
                currentDate.setDate(currentDate.getDate() + rule.interval);
                break;
            case "weekly":
                currentDate.setDate(currentDate.getDate() + 1);
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
        nameButton.style.gap = '5px';
    }
}

function addDatepickerAndViewControls(toolbar, calendar) {
    // Создаём контейнер для datepicker и кнопок
    const controlsContainer = document.createElement('div');
    controlsContainer.id = 'controlsContainer';
    controlsContainer.style.position = 'absolute';
    controlsContainer.style.left = '0';
    controlsContainer.style.top = '40px'; // Под тулбаром
    controlsContainer.style.background = '#fff';
    controlsContainer.style.border = '1px solid #ccc';
    controlsContainer.style.boxShadow = '0 2px 5px rgba(0,0,0,0.2)';
    controlsContainer.style.zIndex = '1000';
    controlsContainer.style.display = 'none'; // Скрыт по умолчанию

    // Datepicker
    const datepickerContainer = document.createElement('div');
    datepickerContainer.id = 'datepickerContainer';
    const datepicker = document.createElement('div');
    datepicker.id = 'datepicker';
    datepickerContainer.appendChild(datepicker);
    controlsContainer.appendChild(datepickerContainer);

    // Инициализация Flatpickr для календаря
    flatpickr(datepicker, {
        inline: true, // Всегда видимый календарь
        onChange: function(selectedDates) {
            const selectedDate = selectedDates[0];
            // Форматируем дату в локальный формат YYYY-MM-DD
            const year = selectedDate.getFullYear();
            const month = String(selectedDate.getMonth() + 1).padStart(2, '0'); // +1, т.к. месяцы с 0
            const day = String(selectedDate.getDate()).padStart(2, '0');
            const dateStr = `${year}-${month}-${day}`;

            console.log("Selected date:", dateStr); // Для отладки
            calendar.changeView('timeGridDay', dateStr);
            controlsContainer.style.display = 'none';
        }
    });

    // Контейнер для кнопок видов
    const viewControls = document.createElement('div');
    viewControls.id = 'viewControls';
    viewControls.style.padding = '10px';
    viewControls.style.borderTop = '1px solid #ccc';

    const views = {
        'Day': 'timeGridDay',
        'Week': 'timeGridWeek',
        'Month': 'dayGridMonth'
    };

    Object.entries(views).forEach(([label, view]) => {
        const button = document.createElement('button');
        button.textContent = label;
        button.classList.add('view-button');
        button.style.marginRight = '5px';
        button.style.padding = '5px 10px';
        button.style.border = '1px solid #0078d4'; // Стиль в духе Outlook
        button.style.background = calendar.view.type === view ? '#0078d4' : '#fff';
        button.style.color = calendar.view.type === view ? '#fff' : '#0078d4';
        button.style.cursor = 'pointer';
        button.onclick = function() {
            calendar.changeView(view);
            updateViewButtons(viewControls, calendar); // Обновляем стили кнопок
            controlsContainer.style.display = 'none'; // Скрываем после выбора
        };
        viewControls.appendChild(button);
    });

    controlsContainer.appendChild(viewControls);
    document.body.appendChild(controlsContainer); // Добавляем в body, чтобы избежать перекрытия toolbar

    // Показ/скрытие при клике на viewToggleButton
    const viewToggleButton = toolbar.querySelector('.fc-viewToggleButton-button');
    viewToggleButton.addEventListener('click', () => {
        controlsContainer.style.display = controlsContainer.style.display === 'none' ? 'block' : 'none';
    });

    // Закрытие при клике вне контейнера
    document.addEventListener('click', function(event) {
        if (!controlsContainer.contains(event.target) && !viewToggleButton.contains(event.target)) {
            controlsContainer.style.display = 'none';
        }
    });
}

// Функция для обновления стилей кнопок видов
function updateViewButtons(viewControls, calendar) {
    const buttons = viewControls.querySelectorAll('.view-button');
    buttons.forEach(button => {
        const view = button.textContent.toLowerCase() === 'day' ? 'timeGridDay' :
            button.textContent.toLowerCase() === 'week' ? 'timeGridWeek' :
                'dayGridMonth';
        button.style.background = calendar.view.type === view ? '#0078d4' : '#fff';
        button.style.color = calendar.view.type === view ? '#fff' : '#0078d4';
    });
}

let calendarInstance;

function renderCalendar(events) {
    const calendarEl = document.getElementById('calendar');
    const savedView = localStorage.getItem('calendarView') || 'timeGridDay';

    const calendar = new Calendar(calendarEl, {
        plugins: [dayGridPlugin, timeGridPlugin, interactionPlugin],
        initialView: savedView,
        height: 'auto',
        contentHeight: 'auto',
        timeZone: 'local',
        dayMaxEvents: 3,
        customButtons: {
            viewToggleButton: {
                text: '',
                click: function() {
                    const datepickerContainer = document.getElementById('datepickerContainer');
                    datepickerContainer.classList.toggle('show'); // Показываем/скрываем календарь
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
            left: 'viewToggleButton,name,prev,next',
            center: 'title',
            right: 'icons'
        },
        dateClick: function(info) {
            console.log('click date', info.dateStr);
            calendar.changeView('timeGridDay', info.dateStr);
        },
        datesSet: function(info) {
            const viewType = info.view.type;
            localStorage.setItem('calendarView', viewType);
            if (viewType !== 'dayGridMonth') {
                setTimeout(() => scrollToMiddle(), 0);
            }
        },
        eventClick: function(info) {
            showEventDetails(info.event);
        },
        views: {
            dayGridMonth: {
                selectable: false
            }
        },
        events: events
    });

    calendar.render();
    calendarInstance = calendar;
    // Добавляем элементы интерфейса
    const toolbar = document.querySelector('.fc-toolbar-chunk:first-child');
    viewButtons(toolbar, calendar);
    rightButtons();

    // Добавляем контейнер для календаря и кнопок видов
    addDatepickerAndViewControls(toolbar, calendar);
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

    const requiredAttendees = eventData.requiredAttendees
        ? eventData.requiredAttendees.split(';').map(email => ({
            emailAddress: { address: email.trim() },
            type: "required"
        }))
        : [];

    const optionalAttendees = eventData.optionalAttendees
        ? eventData.optionalAttendees.split(';').map(email => ({
            emailAddress: { address: email.trim() },
            type: "optional"
        }))
        : [];

    // Базовый объект события
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
        },
        attendees: [...requiredAttendees, ...optionalAttendees]
    };

    // Добавляем Skype-ссылку, если выбрано и есть участники
    if (eventData.addSkype && (requiredAttendees.length > 0 || optionalAttendees.length > 0)) {
        // Генерируем простую гостевую ссылку Skype
        const skypeGuestLink = `https://join.skype.com/invite/${generateRandomId()}`;
        const skypeText = `<br><br><strong>Join Skype Meeting:</strong><br><a href="${skypeGuestLink}" target="_blank">${skypeGuestLink}</a>`;
        event.body.content += skypeText;
    }

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

    const createdEvent = await response.json();
    console.log("Created event:", createdEvent);
    return createdEvent;
}

// Функция для генерации случайного ID (пример)
function generateRandomId() {
    return Math.random().toString(36).substring(2, 10); // Простой случайный ID
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
    const loadingBar = document.getElementById("loading-bar");

    const eventData = {
        id: eventLocalData.id || '',
        title: eventLocalData.title || '',
        start: eventLocalData.start || null,
        end: eventLocalData.end || null,
        location: eventLocalData.location || '',
        description: eventLocalData.description || '',
        requiredAttendees: eventLocalData.requiredAttendees || '',
        optionalAttendees: eventLocalData.optionalAttendees || ''
    };

    const currentDate = calendarInstance ? calendarInstance.getDate() : new Date();
    const startDate = eventData.start ? new Date(eventLocalData.start) : currentDate;
    const endDate = eventData.end ? new Date(eventData.end) : new Date(startDate.getTime() + 30 * 60 * 1000);

    const formatDate = (date) => {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    };
    const formatTime = (date) => `${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
    const formatTime12 = (hour, minute) => {
        const period = hour < 12 ? 'AM' : 'PM';
        const displayHour = hour % 12 || 12;
        return `${displayHour}:${String(minute).padStart(2, '0')} ${period}`;
    };
    const parseTimeInput = (input) => {
        const trimmed = input.trim();
        const timeMatch = trimmed.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
        if (!timeMatch) return null;

        let [_, hours, minutes, period] = timeMatch;
        hours = parseInt(hours);
        minutes = parseInt(minutes);

        if (minutes >= 60) return null;
        if (period) {
            period = period.toUpperCase();
            if (hours > 12) return null;
            if (period === 'PM' && hours < 12) hours += 12;
            if (period === 'AM' && hours === 12) hours = 0;
        } else if (hours > 23) {
            return null;
        }

        return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
    };

    const timeOptions = [];
    for (let hour = 0; hour < 24; hour++) {
        for (let minute = 0; minute < 60; minute += 30) {
            const valueStr = `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
            const displayStr = formatTime12(hour, minute);
            timeOptions.push({ value: valueStr, display: displayStr });
        }
    }

    modalBody.innerHTML = `
        <form id="createEventForm">
            <p id="formError" style="color: red; display: none;"></p>
            <div class="createEventForm-row">
                <label for="eventTitle"></label>
                <input type="text" id="eventTitle" name="eventTitle" value="${eventData.title}" placeholder="Add a title" required>
            </div>
            <div class="createEventForm-row">
                <label for="eventStartDate"><img src="./images/time-18.png" alt="Start date" title="Start date"></label>
                <input type="text" id="eventStartDate" name="eventStartDate" placeholder="Start date" value="${formatDate(startDate)}" required>
                <input type="text" id="eventStartTime" name="eventStartTime" value="${eventData.start ? formatTime12(startDate.getHours(), startDate.getMinutes()) : '9:00 AM'}" required>
                <div id="startTimeDropdown" class="time-dropdown" style="display: none;"></div>
            </div>
            <div class="createEventForm-row">
                <label for="eventEndDate"></label>
                <input type="text" id="eventEndDate" name="eventEndDate" placeholder="End date" value="${formatDate(endDate)}" required>
                <input type="text" id="eventEndTime" name="eventEndTime" value="${eventData.end ? formatTime12(endDate.getHours(), endDate.getMinutes()) : '9:30 AM'}" required>
                <div id="endTimeDropdown" class="time-dropdown" style="display: none;"></div>
            </div>
            <div class="createEventForm-row">
                <label for="eventLocation"><img src="./images/location-18.png" alt="Location" title="Location"></label>
                <input type="text" id="eventLocation" name="eventLocation" value="${eventData.location}" placeholder="Location">
            </div>
            <div class="createEventForm-row">
                <label for="requiredAttendees"><img src="./images/invite_required-18.png" alt="Required attendees" title="Required attendees"></label>
                <input type="text" id="requiredAttendees" name="requiredAttendees" value="${eventData.requiredAttendees}" placeholder="Required attendees (email1;email2)">
            </div>
            <div class="createEventForm-row">
                <label for="optionalAttendees"></label>
                <input type="text" id="optionalAttendees" name="optionalAttendees" value="${eventData.optionalAttendees}" placeholder="Optional attendees (email1;email2)">
            </div>
            <div class="createEventForm-row">
                <label for="eventDescription"><img src="./images/text-18.png" alt="Description" title="Description"></label>
                <textarea id="eventDescription" name="eventDescription" placeholder="Description">${eventData.description}</textarea>
            </div>
            <div class="createEventForm-row checkbox-container">
                <label for="addSkype"><img src="./images/skype-18.png" alt="Add Skype" title="Add Skype meeting"></label>
                <div class="custom-checkbox">
                    <input type="checkbox" id="addSkype" name="addSkype">
                    <label for="addSkype"></label>
                </div>
                <span class="checkbox-text">Add Skype meeting</span>
            </div>
            <button type="submit">${eventData.id ? 'Update Event' : 'Create Event'}</button>
        </form>
    `;

    modal.classList.remove('hidden');
    modal.style.display = 'block';

    flatpickr("#eventStartDate", {
        dateFormat: "Y-m-d",
        defaultDate: formatDate(startDate)
    });
    flatpickr("#eventEndDate", {
        dateFormat: "Y-m-d",
        defaultDate: formatDate(endDate)
    });

    // Логика выпадающего списка
    const startTimeInput = document.getElementById('eventStartTime');
    const endTimeInput = document.getElementById('eventEndTime');
    const startTimeDropdown = document.getElementById('startTimeDropdown');
    const endTimeDropdown = document.getElementById('endTimeDropdown');

    // Заполняем списки в 12-часовом формате
    startTimeDropdown.innerHTML = timeOptions.map(opt =>
        `<div class="time-option" data-value="${opt.value}">${opt.display}</div>`
    ).join('');
    endTimeDropdown.innerHTML = timeOptions.map(opt =>
        `<div class="time-option" data-value="${opt.value}">${opt.display}</div>`
    ).join('');

    // Показ/скрытие списка
    function toggleDropdown(input, dropdown) {
        const isVisible = dropdown.style.display === 'block';
        dropdown.style.display = isVisible ? 'none' : 'block';
        if (!isVisible) {
            const inputRect = input.getBoundingClientRect();
            const rowRect = input.parentElement.getBoundingClientRect();
            dropdown.style.top = `${inputRect.bottom - rowRect.top}px`;
            dropdown.style.left = `${inputRect.left - rowRect.left}px`;
            const parsedTime = parseTimeInput(input.value);
            if (parsedTime) {
                dropdown.scrollTop = Array.from(dropdown.children).findIndex(opt => opt.dataset.value === parsedTime) * 20;
            }
        }
    }

    function handleTimeInput(input, dropdown, otherInput) {
        input.addEventListener('click', (e) => {
            e.preventDefault();
            toggleDropdown(input, dropdown);
        });

        input.addEventListener('input', () => {
            const parsedTime = parseTimeInput(input.value);
            if (parsedTime && input === startTimeInput) {
                const [hour, minute] = parsedTime.split(':').map(Number);
                let endHour = hour;
                let endMinute = minute + 30;
                if (endMinute >= 60) {
                    endHour = (endHour + 1) % 24;
                    endMinute = 0;
                }
                endTimeInput.value = formatTime12(endHour, endMinute);
            }
        });

        input.addEventListener('blur', () => {
            const parsedTime = parseTimeInput(input.value);
            if (!parsedTime) {
                input.value = input === startTimeInput ? '9:00 AM' : '9:30 AM';
            } else {
                const [hour, minute] = parsedTime.split(':').map(Number);
                input.value = formatTime12(hour, minute);
            }
        });
    }

    startTimeDropdown.addEventListener('click', (e) => {
        const time = e.target.dataset.value;
        if (time) {
            startTimeInput.value = timeOptions.find(opt => opt.value === time).display;
            const [startHour, startMinute] = time.split(':').map(Number);
            let endHour = startHour;
            let endMinute = startMinute + 30;
            if (endMinute >= 60) {
                endHour = (endHour + 1) % 24;
                endMinute = 0;
            }
            endTimeInput.value = formatTime12(endHour, endMinute);
            startTimeDropdown.style.display = 'none';
        }
    });

    endTimeDropdown.addEventListener('click', (e) => {
        const time = e.target.dataset.value;
        if (time) {
            endTimeInput.value = timeOptions.find(opt => opt.value === time).display;
            endTimeDropdown.style.display = 'none';
        }
    });

    document.addEventListener('click', (e) => {
        if (!startTimeInput.contains(e.target) && !startTimeDropdown.contains(e.target)) {
            startTimeDropdown.style.display = 'none';
        }
        if (!endTimeInput.contains(e.target) && !endTimeDropdown.contains(e.target)) {
            endTimeDropdown.style.display = 'none';
        }
    });

    handleTimeInput(startTimeInput, startTimeDropdown, endTimeInput);
    handleTimeInput(endTimeInput, endTimeDropdown, startTimeInput);

    document.getElementById('createEventForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        const title = document.getElementById('eventTitle').value.trim();
        const startDate = document.getElementById('eventStartDate').value;
        const startTimeDisplay = document.getElementById('eventStartTime').value;
        const endDate = document.getElementById('eventEndDate').value;
        const endTimeDisplay = document.getElementById('eventEndTime').value;
        const location = document.getElementById('eventLocation').value.trim();
        const requiredAttendees = document.getElementById('requiredAttendees').value.trim();
        const optionalAttendees = document.getElementById('optionalAttendees').value.trim();
        const description = document.getElementById('eventDescription').value.trim();
        const addSkype = document.getElementById('addSkype').checked;
        const errorElement = document.getElementById('formError');

        const startTime = parseTimeInput(startTimeDisplay);
        const endTime = parseTimeInput(endTimeDisplay);

        if (!startTime) {
            showError('Please enter a valid start time (e.g., 9:00 AM or 13:00)');
            return;
        }
        if (!endTime) {
            showError('Please enter a valid end time (e.g., 9:30 AM or 13:30)');
            return;
        }

        const start = `${startDate}T${startTime}:00`;
        const end = `${endDate}T${endTime}:00`;

        if (!title) { showError('Please enter a title'); return; }
        if (!startDate || isNaN(new Date(start).getTime())) {
            showError('Please enter a valid start date and time'); return;
        }
        if (!endDate || isNaN(new Date(end).getTime())) {
            showError('Please enter a valid end date and time'); return;
        }
        const startDateTime = new Date(start);
        const endDateTime = new Date(end);
        if (endDateTime <= startDateTime) {
            showError('End time must be after start time'); return;
        }

        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (requiredAttendees) {
            const requiredEmails = requiredAttendees.split(';');
            for (let email of requiredEmails) {
                if (email.trim() && !emailRegex.test(email.trim())) {
                    showError('Please enter valid email addresses for required attendees');
                    return;
                }
            }
        }
        if (optionalAttendees) {
            const optionalEmails = optionalAttendees.split(';');
            for (let email of optionalEmails) {
                if (email.trim() && !emailRegex.test(email.trim())) {
                    showError('Please enter valid email addresses for optional attendees');
                    return;
                }
            }
        }

        showLoadingBar();
        try {
            if (eventData.id) {
                await updateEvent(eventData.id, {
                    title,
                    start,
                    end,
                    location,
                    requiredAttendees,
                    optionalAttendees,
                    description,
                    addSkype
                });
            } else {
                await createEvent({
                    title,
                    start,
                    end,
                    location,
                    requiredAttendees,
                    optionalAttendees,
                    description,
                    addSkype
                });
            }

            modal.classList.add('hidden');
            modal.style.display = 'none';
            const {startDate: monthStart, endDate: monthEnd} = getMonthDateRange(new Date(start));
            fetchEvents(true, monthStart, monthEnd);
            hideLoadingBar();
        } catch (error) {
            console.error('Error processing event:', error);
            hideLoadingBar();
            showError('An error occurred while processing the event');
        }

        function showError(message) {
            errorElement.textContent = message;
            errorElement.style.display = 'block';
            setTimeout(() => {
                errorElement.style.display = 'none';
            }, 3000);
        }
    });
}

document.getElementById('createEventButton').addEventListener('click', () => {
    openEventModal(); // Открытие пустого модального окна для создания события
});

function loadBar() {
        const updateButton = document.getElementById("updateButton");

        if (updateButton) {
            updateButton.addEventListener("click", function () {
                showLoadingBar();

                // Здесь можно запустить обновление данных, например, через setTimeout имитируем процесс
                setTimeout(() => {
                    hideLoadingBar();
                }, 3000); // Имитация загрузки (замените на реальный вызов)
            });
        }

}

function showLoadingBar() {
    const loadingBar = document.getElementById("loading-bar");
    loadingBar.style.width = "100%";
}

function hideLoadingBar() {
    const loadingBar = document.getElementById("loading-bar");
    setTimeout(() => {
        loadingBar.style.width = "0";
    }, 500);
}


document.addEventListener("DOMContentLoaded", () => {
    const now = new Date();
    const { startDate, endDate } = getMonthDateRange(now);
    fetchEvents(false, startDate, endDate);
});

