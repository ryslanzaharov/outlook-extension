import { Calendar } from '@fullcalendar/core';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import interactionPlugin from '@fullcalendar/interaction';


 import './styles.css'; // Импорт стилей


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

function showEventDetails(event) {
    // Элементы модального окна
    const modal = document.getElementById('modal');
    const modalBody = document.getElementById('modalBody');

    // Проверяем, если location является URL, либо создаем ссылку для Google Maps
    let locationLink = '';
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
    `;

    // Отображение модального окна
    modal.classList.remove('hidden');
    modal.style.display = 'block';
}


async function hiddenElements() {
    // const eventsContainer = document.getElementById("loginButton");
    // eventsContainer.innerHTML = "";
    // eventsContainer.classList.add("hidden");
}

async function fetchEvents() {
    const tokenData = await new Promise(resolve => {
        chrome.storage.local.get("token", resolve);
    });

    let accessToken = tokenData?.token;
    console.log("accessToken1: " + accessToken);

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

        await hiddenElements();

        const data = await response.json()
        if (data.value && data.value.length > 0) {
            // Преобразуем события в формат FullCalendar
            const fullCalendarData = data.value.map(event => ({
                title: event.subject,
                start: event.start.dateTime,
                end: event.end.dateTime,
                location: event.location?.displayName || ''
            }));
            const calendarEl = document.getElementById('calendar');

            // Создание календаря
            const calendar = new Calendar(calendarEl, {
                plugins: [dayGridPlugin, timeGridPlugin, interactionPlugin], // Подключение необходимых плагинов
                initialView: 'dayGridMonth', // Вид календаря (месяц, неделя и т.д.)
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
                    console.log(`Event title: ${info.event.title}`);
                },
                moreLinkClick: function(info) {
                    // Обработка клика на "more"
                    console.log("More" + info);
                    info.view.calendar.changeView('timeGridDay', info.date);
                    return false; // Возвращаем false, если не хотим выполнять стандартное поведение
                },
                dateClick: function (info) {
                    // Устанавливаем дату для timeGridDay и переключаемся
                    calendar.changeView('timeGridDay', info.dateStr);
                },
                eventClick: function (info) {
                    // Вызов функции для получения созвонов
                    // showCallsForDate(info.dateStr, calendar);
                    // Вызов модального окна для отображения данных события
                    showEventDetails(info.event);
                }
            });

            calendar.render();
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

// document.getElementById("logoutButton").addEventListener("click", () => {
//     chrome.runtime.sendMessage({ action: "logout" });
// });

// document.getElementById("loginButton").addEventListener("click", () => {
//     chrome.runtime.sendMessage({ action: "login" });
// });

document.addEventListener("DOMContentLoaded", fetchEvents);
