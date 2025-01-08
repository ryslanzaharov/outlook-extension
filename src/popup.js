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

async function fetchEvents() {
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
        const response = await fetch("https://graph.microsoft.com/v1.0/me/calendar/events", {
            headers: {
                "Authorization": `Bearer ${accessToken}`
            }
        });

        if (response.status === 401) {
            try {
                accessToken = await getAccessToken();
            } catch (error) {
                console.error("Authorization failed:", error);
                document.getElementById("events").innerText = "Authorization required.";
                return;
            }
        }

        const data = await response.json()
        if (data.value && data.value.length > 0) {
            // Преобразуем события в формат FullCalendar
            const fullCalendarData = data.value.map(event => {
                const start = new Date(event.start.dateTime + 'Z');
                const end = new Date(event.end.dateTime + 'Z');

                return {
                    title: event.subject,
                    start: start,
                    end: end,
                    location: event.location?.displayName || '',
                    description: event.body?.content || ''
                };
            });

            const calendarEl = document.getElementById('calendar');

            // Создание календаря
            const calendar = new Calendar(calendarEl, {
                plugins: [dayGridPlugin, timeGridPlugin, interactionPlugin], // Подключение необходимых плагинов
                initialView: 'timeGridDay', // Вид календаря (месяц, неделя и т.д.)
                headerToolbar: {
                    left: 'prev,next today',
                    center: 'title',
                    right: 'timeGridDay,timeGridWeek,dayGridMonth'
                },
                events: fullCalendarData, // Передаем события в FullCalendar
                timeZone: 'local',
                height: 'auto', // Или фиксированная высота, например '600px'
                dayMaxEvents: 3, // Ограничение на количество событий в день
                eventDidMount: function(info) {
                    // console.log(`Event title: ${info.event.title}`);
                },
                moreLinkClick: function(info) {
                    // Обработка клика на "more"
                    info.view.calendar.changeView('timeGridDay', info.date);
                    return false; // Возвращаем false, если не хотим выполнять стандартное поведение
                },
                dateClick: function (info) {
                    // Устанавливаем дату для timeGridDay и переключаемся
                    calendar.changeView('timeGridDay', info.dateStr);
                },
                eventClick: function (info) {
                    // Вызов модального окна для отображения данных события
                    showEventDetails(info.event);
                    // notifyEvent(info.event);
                }
            });

            calendar.render();
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
            fetchEvents(); // Перезагружаем события в календаре
        } catch (error) {
            console.error('Error creating event:', error);
            alert('Failed to create event.');
        }
    });
});



document.addEventListener("DOMContentLoaded", fetchEvents);
