document.addEventListener("DOMContentLoaded", () => {
    const params = new URLSearchParams(window.location.search);

    const title = params.get("title") || "Без названия";
    const time = params.get("time") || "Не указано";
    const location = params.get("location") || "Не указано";
    const description = params.get("description") || "Не указано";

    document.getElementById("eventTitle").innerText = title;
    document.getElementById("eventTime").innerText = `Время: ${time}`;
    document.getElementById("eventLocation").innerText = `Место: ${location}`;
    document.getElementById("eventDescription").innerText = `Описание: ${description}`;
});
