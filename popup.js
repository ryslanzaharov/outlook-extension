document.addEventListener('DOMContentLoaded', () => {
    const loginButton = document.getElementById('loginButton');
    const calendarElement = document.getElementById('calendar');

    // Обработчик нажатия на кнопку авторизации
    loginButton.addEventListener('click', () => {
        chrome.runtime.sendMessage({ action: 'login' });
    });

    // Проверка наличия токена и загрузка событий календаря
    chrome.storage.local.get(['accessToken'], async (result) => {
        const token = result.accessToken;
        if (!token) {
            calendarElement.textContent = 'Пожалуйста, авторизуйтесь, чтобы увидеть события календаря.';
            return;
        }

        // Запрос событий календаря
        try {
            const response = await fetch('https://graph.microsoft.com/v1.0/me/events', {
                headers: {
                    'Authorization': `Bearer ${token}`
                }
            });

            if (response.ok) {
                const data = await response.json();
                displayEvents(data.value);
            } else {
                calendarElement.textContent = 'Не удалось получить события.';
            }
        } catch (error) {
            calendarElement.textContent = 'Ошибка при получении данных календаря.';
            console.error(error);
        }
    });

    // Функция для отображения событий календаря
    function displayEvents(events) {
        calendarElement.innerHTML = '';
        events.forEach(event => {
            const eventElement = document.createElement('div');
            eventElement.textContent = `${event.subject} в ${new Date(event.start.dateTime).toLocaleTimeString()}`;
            calendarElement.appendChild(eventElement);
        });
    }
});
