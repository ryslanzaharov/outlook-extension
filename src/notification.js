document.addEventListener("DOMContentLoaded", () => {
    const params = new URLSearchParams(window.location.search);
    const title = decodeURIComponent(params.get("title") || '');
    const time = decodeURIComponent(params.get("time") || '');
    const location = decodeURIComponent(params.get("location") || '');
    // const outlookUrl = decodeURIComponent(params.get("outlookUrl") || ''); // Декодируем URL
    const outlookUrl = params.get("outlookUrl") || ''; // Декодируем URL
    console.log("params.get('outlookUrl'):", outlookUrl);

    document.getElementById("eventTitle").textContent = title || "No Title";
    document.getElementById("eventTime").textContent = "Time: " + (time || "No Time");

    const locationElement = document.getElementById("eventLocation");
    if (location && (location.startsWith("http://") || location.startsWith("https://"))) {
        locationElement.innerHTML = `<a href="${location}" target="_blank" rel="noopener noreferrer">${location}</a>`;
    } else {
        locationElement.textContent = location || "No Location";
    }

    const outlookLink = document.getElementById("outlookLink");
    if (outlookUrl) {
        outlookLink.href = outlookUrl; // Устанавливаем декодированный URL
        outlookLink.target = "_blank"; // Открытие в новой вкладке
        outlookLink.rel = "noopener noreferrer"; // Безопасность
    } else {
        outlookLink.style.display = "none"; // Скрываем, если ссылки нет
    }
    console.log("outlookLink:", outlookLink);
    document.getElementById("closeButton").addEventListener("click", () => {
        chrome.runtime.sendMessage({ action: "closeWindow" });
    });
});