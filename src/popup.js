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
        let response = await fetch(`https://graph.microsoft.com/v1.0/me/calendar/events?$filter=start/dateTime ge '${startDate.toISOString()}' and end/dateTime le '${endDate.toISOString()}'`, {
            headers: {
                "Authorization": `Bearer ${accessToken}`
            }
        });
        if (response.status === 401) {
            try {
                accessToken = await getAccessToken();
                response = await fetch(`https://graph.microsoft.com/v1.0/me/calendar/events?$filter=start/dateTime ge '${startDate.toISOString()}' and end/dateTime le '${endDate.toISOString()}'`, {
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

        const data = await response.json();
        if (data.value && data.value.length > 0) {
            // Преобразуем события в формат FullCalendar
            const fullCalendarData = [];
            for (const event of data.value) {
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
        height: 'auto', // Убедитесь, что высота адаптируется
        contentHeight: 'auto',
        headerToolbar: {
            left: 'prev,next',
            center: 'title',
            right: 'timeGridDay,timeGridWeek,dayGridMonth'
        },
        events: events,
        timeZone: 'local',
        dayMaxEvents: 3,
        dateClick: function(info) {
            calendar.changeView('timeGridDay', info.dateStr);
        },
        // dateDidMount: function () {
        //     // Убедитесь, что прокрутка происходит после рендера
        //     setTimeout(() => scrollToMiddle(), 0);
        // },
        datesSet: function () {
            // Вызываем при смене дат
            setTimeout(() => scrollToMiddle(), 0);
        },
        eventClick: function(info) {
            showEventDetails(info.event);
        },
    });

    calendar.render();
}

function scrollToMiddle() {
    const scroller = document.querySelector('.fc-timegrid-body');
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

document.getElementById('createEventButton').addEventListener('click', () => {
    const modal = document.getElementById('modal');
    const modalBody = document.getElementById('modalBody');

    // Наполнение модального окна формой
    modalBody.innerHTML = `
        
        <form id="createEventForm">
        <h3>Create New Event</h3>
    <div class="createEventForm-row">
        <label for="eventTitle">Title:</label>
        <input type="text" id="eventTitle" name="eventTitle" required>
    </div>

    <div class="createEventForm-row">
        <label for="eventStart">Start:</label>
        <input type="datetime-local" id="eventStart" name="eventStart" required>
    </div>

    <div class="createEventForm-row">
        <label for="eventEnd">End:</label>
        <input type="datetime-local" id="eventEnd" name="eventEnd" required>
    </div>

    <div class="createEventForm-row">
        <label for="eventLocation">Location:</label>
        <input type="text" id="eventLocation" name="eventLocation">
    </div>

    <div class="createEventForm-row">
        <label for="eventDescription">Description:</label>
        <textarea id="eventDescription" name="eventDescription"></textarea>
    </div>

    <button type="submit">Create Event</button>
</form>
    <p style="color: red; font-weight: bold;">
        Note: This feature will become a paid service starting from 01.06.2025.
    </p>
    `;

    modal.classList.remove('hidden');
    modal.style.display = 'block';

    // Добавляем обработчик для отправки формы
    document.getElementById('createEventForm').addEventListener('submit', async (e) => {
        e.preventDefault();

        const title = document.getElementById('eventTitle').value;
        const start = document.getElementById('eventStart').value;
        const end = document.getElementById('eventEnd').value;
        const location = document.getElementById('eventLocation').value;
        const description = document.getElementById('eventDescription').value;

        try {
            await createEvent({ title, start, end, location, description });
            alert('Event created successfully!');
            modal.classList.add('hidden');
            modal.style.display = 'none';
            const now = new Date();
            const { startDate, endDate } = getMonthDateRange(new Date(start));
            fetchEvents(startDate, endDate); // Перезагружаем события в календаре
        } catch (error) {
            console.error('Error creating event:', error);
            alert('Failed to create event.');
        }
    });
});


document.addEventListener("DOMContentLoaded", () => {
    const now = new Date();
    const { startDate, endDate } = getMonthDateRange(now);
    fetchEvents(startDate, endDate);
});

