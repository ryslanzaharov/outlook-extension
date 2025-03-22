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
            : ``;
    }

    // Обрабатываем description
    if (description) {
        console.log("Original description:", description);

        // Парсим HTML
        const parser = new DOMParser();
        const doc = parser.parseFromString(description, 'text/html');

        // Извлекаем содержимое body
        let bodyContent = doc.body.innerHTML;

        // Проверяем, является ли это Skype Meeting и убираем два последовательных <br>
        if (bodyContent.includes('Join Skype Meeting')) {
            // Удаляем два последовательных <br> в начале
            bodyContent = bodyContent.replace(/^(\s*<br\s*\/?>\s*){2}/i, '');
            console.log("Cleaned body content:", bodyContent);
        }

        // Добавляем атрибуты к ссылкам
        const tempDoc = parser.parseFromString(bodyContent, 'text/html');
        const links = tempDoc.getElementsByTagName('a');
        for (let link of links) {
            link.setAttribute('target', '_blank');
            link.setAttribute('rel', 'noopener noreferrer');
        }

        description = tempDoc.body.innerHTML;
        console.log("Final description:", description);
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
                <svg class="theme-icon" width="18" height="18" viewBox="0 0 22 22" fill="currentColor" alt="Open in Outlook" title="Open in Outlook">
                    <path d="M14 3v2h5.59L4 20.59 5.41 22 20 7.41V13h2V3z"></path>
                </svg>
            </a>
        </div>
        <h3>${event.title}</h3>
        <br>
        <div class="event-row">
            <svg class="theme-icon" width="22" height="22" viewBox="0 0 22 22" alt="Time" title="Time">
                <circle cx="11" cy="11" r="10" fill="none" stroke="black"/>
                <path d="M11 11V6M11 11h5" fill="none" stroke="black"/>
            </svg>
            <span>${startFormatTo12Hour(event.start)} - ${endFormatTo12Hour(event.end)}</span>
        </div>
        ${locationLink ? `
            <div class="event-row">
                <svg class="theme-icon" width="22" height="22" viewBox="0 0 22 22" alt="Location" title="Location">
                    <path d="M11 2a8 8 0 0 1 8 8c0 5-8 10-8 10s-8-5-8-10a8 8 0 0 1 8-8z" fill="none" stroke="black"/>
                    <circle cx="11" cy="10" r="3" fill="none" stroke="black"/>
                </svg>
                <span>${locationLink}</span>
            </div>
        ` : ''}
        ${requiredAttendeesList ? `
        <div class="event-row attendees-row" data-type="required">
            <svg class="theme-icon" width="22" height="22" viewBox="0 0 22 22" alt="Required attendees" title="Required attendees">
                <circle cx="9" cy="5" r="3" fill="none" stroke="black"/>
                <path d="M4 12c0-2 2-4 5-4s5 2 5 4" fill="none" stroke="black"/>
            </svg>
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
                <svg class="theme-icon" width="22" height="22" viewBox="0 0 22 22" alt="Description" title="Description">
                    <path d="M4 4h14v14H4z" fill="none" stroke="black"/>
                    <path d="M7 7h8M7 10h8M7 13h4" fill="none" stroke="black"/>
                </svg>
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
    const supportButton = document.querySelector('.fc-icons-button');
    if (supportButton) {
        supportButton.innerHTML = `
<div class="icons">
    <button id="updateButton" class="icon-button" title="Update">
        <svg class="icon">
            <path d="M12 6V2L8 6l4 4V6c3.31 0 6 2.69 6 6s-2.69 6-6 6a5.99 5.99 0 0 1-5.3-3H4.26A7.99 7.99 0 0 0 12 20c4.42 0 8-3.58 8-8s-3.58-8-8-8z"></path>
        </svg>
    </button>

    <button id="themeButton" class="icon-button" title="Toggle Day/Night">
        <!-- SVG для солнца (показывается в тёмной теме) -->
        <svg id="sunIcon" class="icon" width="24" height="24" viewBox="0 0 24 24" style="display: none;">
            <circle cx="12" cy="12" r="5" fill="none" stroke="currentColor" stroke-width="2"/>
            <path d="M12 2v4M12 18v4M2 12h4M18 12h4M5.64 5.64l2.83 2.83M15.54 15.54l2.83 2.83M5.64 18.36l2.83-2.83M15.54 8.46l2.83-2.83" fill="none" stroke="currentColor" stroke-width="2"/>
        </svg>
        <!-- SVG для полумесяца (показывается в светлой теме) -->
        <svg id="moonIcon" class="icon" width="24" height="24" viewBox="0 0 24 24" style="display: block;">
            <path d="M9.37 5.51c-.18.64-.27 1.31-.27 1.99 0 4.08 3.32 7.4 7.4 7.4.68 0 1.35-.09 1.99-.27C17.45 17.19 14.93 19 12 19c-3.86 0-7-3.14-7-7 0-2.93 1.81-5.45 4.37-6.49M12 3c-4.97 0-9 4.03-9 9s4.03 9 9 9 9-4.03 9-9c0-.46-.04-.92-.1-1.36-.98 1.37-2.58 2.26-4.4 2.26-2.98 0-5.4-2.42-5.4-5.4 0-1.81.89-3.42 2.26-4.4-.44-.06-.90-.1-1.36-.1"></path>
        </svg>
    </button>

    <a id="support" class="link" href="https://mail.google.com/mail/u/0/?view=cm&fs=1&to=ruslan.ext.dev@gmail.com&su=Outlook%20Calendar%20Checker&body=Hello,%20I%20would%20like%20to%20suggest%20you%20to%20do" target="_blank" title="Support">
        <svg class="icon">
            <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm-1-13h2v6h-2zm1 8c-.83 0-1.5-.67-1.5-1.5S11.17 14 12 14s1.5.67 1.5 1.5S12.83 17 12 17z"></path>
        </svg>
    </a>

    <a id="owaCalendar" class="link" href="https://outlook.live.com/calendar/0/view/day" target="_blank" title="Open OWA">
        <svg class="icon">
            <path d="M14 3v2h5.59L4 20.59 5.41 22 20 7.41V13h2V3z"></path>
        </svg>
    </a>

    <button id="logoutButton" class="icon-button" title="Logout">
        <svg class="icon">
            <path d="M17 7l-1.41 1.41L18.17 11H8v2h10.17l-2.58 2.58L17 17l5-5zM4 5h8V3H4c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h8v-2H4z"></path>
        </svg>
    </button>
</div>
        `;
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

    const themeButton = document.getElementById("themeButton");
    const sunIcon = document.getElementById("sunIcon");
    const moonIcon = document.getElementById("moonIcon");

    themeButton.addEventListener("click", () => {
        const rootElement = document.body;
        const isDark = rootElement.classList.contains("dark-theme");

        if (isDark) {
            // Переключаем на светлую тему (показываем полумесяц)
            rootElement.classList.remove("dark-theme");
            rootElement.classList.add("light-theme");
            chrome.storage.local.set({ theme: "light" });
            sunIcon.style.display = "none"; // Скрываем солнце
            moonIcon.style.display = "block"; // Показываем полумесяц
        } else {
            // Переключаем на тёмную тему (показываем солнце)
            rootElement.classList.remove("light-theme");
            rootElement.classList.add("dark-theme");
            chrome.storage.local.set({ theme: "dark" });
            sunIcon.style.display = "block"; // Показываем солнце
            moonIcon.style.display = "none"; // Скрываем полумесяц
        }
    });

    // Загрузка сохранённой темы
    chrome.storage.local.get(["theme"], (result) => {
        const rootElement = document.body;
        if (result.theme === "dark") {
            rootElement.classList.remove("light-theme");
            rootElement.classList.add("dark-theme");
            sunIcon.style.display = "block"; // Солнце в тёмной теме
            moonIcon.style.display = "none";
        } else {
            rootElement.classList.remove("dark-theme");
            rootElement.classList.add("light-theme");
            sunIcon.style.display = "none"; // Полумесяц в светлой теме
            moonIcon.style.display = "block";
        }
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

    // Datepicker
    const datepickerContainer = document.createElement('div');
    datepickerContainer.id = 'datepickerContainer';
    const datepicker = document.createElement('div');
    datepicker.id = 'datepicker';
    datepickerContainer.appendChild(datepicker);
    controlsContainer.appendChild(datepickerContainer);

    // Инициализация Flatpickr для календаря
    flatpickr(datepicker, {
        inline: true,
        onChange: function(selectedDates) {
            const selectedDate = selectedDates[0];
            const year = selectedDate.getFullYear();
            const month = String(selectedDate.getMonth() + 1).padStart(2, '0');
            const day = String(selectedDate.getDate()).padStart(2, '0');
            const dateStr = `${year}-${month}-${day}`;

            console.log("Selected date:", dateStr);
            calendar.changeView('timeGridDay', dateStr);
            controlsContainer.style.display = 'none';
        }
    });

    // Контейнер для кнопок видов
    const viewControls = document.createElement('div');
    viewControls.id = 'viewControls';

    const views = {
        'Day': 'timeGridDay',
        'Week': 'timeGridWeek',
        'Month': 'dayGridMonth'
    };

    Object.entries(views).forEach(([label, view]) => {
        const button = document.createElement('button');
        button.textContent = label;
        button.classList.add('view-button');
        if (calendar.view.type === view) {
            button.classList.add('active');
        }
        button.onclick = function() {
            calendar.changeView(view);
            updateViewButtons(viewControls, calendar);
            controlsContainer.style.display = 'none';
        };
        viewControls.appendChild(button);
    });

    controlsContainer.appendChild(viewControls);
    document.body.appendChild(controlsContainer);

    // Показ/скрытие при клике на viewToggleButton
    const viewToggleButton = toolbar.querySelector('.fc-viewToggleButton-button');
    if (viewToggleButton) {
        viewToggleButton.addEventListener('click', (e) => {
            e.stopPropagation(); // Предотвращаем всплытие события
            const isHidden = controlsContainer.style.display === 'none' || controlsContainer.style.display === '';
            controlsContainer.style.display = isHidden ? 'block' : 'none';
        });
    }

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
        button.classList.toggle('active', calendar.view.type === view);
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
            displayName: eventData.location || "Online Meeting"
        },
        body: {
            contentType: "HTML",
            content: eventData.description || ""
        },
        attendees: [...requiredAttendees, ...optionalAttendees]
    };

    if (eventData.addSkype && (requiredAttendees.length > 0 || optionalAttendees.length > 0)) {
        const skypeGuestLink = `https://join.skype.com/invite/${generateRandomId()}`;
        const skypeText = `<br><br><strong>Join Skype Meeting:</strong><br><a href="${skypeGuestLink}" target="_blank">${skypeGuestLink}</a>`;
        event.body.content += skypeText;
        event.location.displayName = "Skype Meeting";
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

function generateRandomId() {
    return Math.random().toString(36).substring(2, 10);
}

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
        optionalAttendees: eventLocalData.optionalAttendees || '',
        addSkype: eventLocalData.addSkype || false
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
            <label for="eventStartDate">
                <svg class="theme-icon" width="22" height="22" viewBox="0 0 22 22" alt="Start date" title="Start date">
                    <rect x="2" y="4" width="14" height="12" stroke="black" fill="none"/>
                    <line x1="2" y1="8" x2="16" y2="8" stroke="black"/>
                </svg>
            </label>
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
        <div class="createEventForm-row location-row">
            <label for="eventLocation">
                <svg class="theme-icon" width="22" height="22" viewBox="0 0 22 22" alt="Location" title="Location">
                    <path d="M9 2C6.24 2 4 4.24 4 7c0 4 5 9 5 9s5-5 5-9c0-2.76-2.24-5-5-5z" fill="none" stroke="black"/>
                    <circle cx="9" cy="7" r="2" fill="none" stroke="black"/>
                </svg>
            </label>
            <input type="text" id="eventLocation" name="eventLocation" value="${eventData.location}" placeholder="Location">
            <div class="custom-checkbox skype-checkbox">
                <input type="checkbox" id="addSkype" name="addSkype" ${eventData.addSkype ? 'checked' : ''}>
                <label for="addSkype"></label>
                <span class="checkbox-text">Skype meeting</span>
            </div>
        </div>
        <div class="createEventForm-row">
            <label for="requiredAttendees">
                <svg class="theme-icon" width="22" height="22" viewBox="0 0 22 22" alt="Required attendees" title="Required attendees">
                    <circle cx="9" cy="5" r="3" fill="none" stroke="black"/>
                    <path d="M4 12c0-2 2-4 5-4s5 2 5 4" fill="none" stroke="black"/>
                </svg>
            </label>
            <input type="text" id="requiredAttendees" name="requiredAttendees" value="${eventData.requiredAttendees}" placeholder="Required attendees (email1;email2)">
        </div>
        <div class="createEventForm-row">
            <label for="optionalAttendees"></label>
            <input type="text" id="optionalAttendees" name="optionalAttendees" value="${eventData.optionalAttendees}" placeholder="Optional attendees (email1;email2)">
        </div>
        <div class="createEventForm-row">
            <label for="eventDescription">
                <svg class="theme-icon" width="22" height="22" viewBox="0 0 22 22" alt="Description" title="Description">
                    <rect x="2" y="2" width="14" height="14" fill="none" stroke="black"/>
                    <line x1="5" y1="6" x2="13" y2="6" stroke="black"/>
                    <line x1="5" y1="9" x2="13" y2="9" stroke="black"/>
                    <line x1="5" y1="12" x2="9" y2="12" stroke="black"/>
                </svg>
            </label>
            <textarea id="eventDescription" name="eventDescription" placeholder="Description">${eventData.description}</textarea>
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

    const startTimeInput = document.getElementById('eventStartTime');
    const endTimeInput = document.getElementById('eventEndTime');
    const startTimeDropdown = document.getElementById('startTimeDropdown');
    const endTimeDropdown = document.getElementById('endTimeDropdown');
    const locationInput = document.getElementById('eventLocation');
    const addSkypeCheckbox = document.getElementById('addSkype');

    startTimeDropdown.innerHTML = timeOptions.map(opt =>
        `<div class="time-option" data-value="${opt.value}">${opt.display}</div>`
    ).join('');
    endTimeDropdown.innerHTML = timeOptions.map(opt =>
        `<div class="time-option" data-value="${opt.value}">${opt.display}</div>`
    ).join('');

    // Логика для чекбокса Skype meeting
    addSkypeCheckbox.addEventListener('change', () => {
        if (addSkypeCheckbox.checked) {
            locationInput.value = "Link will be generated";
            locationInput.readOnly = true;
        } else {
            locationInput.value = eventData.location || "";
            locationInput.readOnly = false;
            locationInput.placeholder = "Location";
        }
    });

    // Установка начального состояния
    if (addSkypeCheckbox.checked) {
        locationInput.value = "Link will be generated";
        locationInput.readOnly = true;
    }

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

        if (!title) {
            showError('Please enter a title');
            return;
        }
        if (!startDate || isNaN(new Date(start).getTime())) {
            showError('Please enter a valid start date and time');
            return;
        }
        if (!endDate || isNaN(new Date(end).getTime())) {
            showError('Please enter a valid end date and time');
            return;
        }
        const startDateTime = new Date(start);
        const endDateTime = new Date(end);
        if (endDateTime <= startDateTime) {
            showError('End time must be after start time');
            return;
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
            const { startDate: monthStart, endDate: monthEnd } = getMonthDateRange(new Date(start));
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


async function updateEvent(eventId, eventData) {
    // Реализация обновления события
    console.log('Updating event:', eventId, eventData);
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

