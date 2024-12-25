document.addEventListener("DOMContentLoaded", () => {
    const params = new URLSearchParams(window.location.search);
    const title = params.get("title");
    const time = params.get("time");
    const location = params.get("location");
    // const description = params.get("description");

    document.getElementById("eventTitle").textContent = title || "No Title";
    document.getElementById("eventTime").textContent = "Time:" + time || "No Time";

    const locationElement = document.getElementById("eventLocation");
    if (location && (location.startsWith("http://") || location.startsWith("https://"))) {
        locationElement.innerHTML = `<a href="${location}" target="_blank" rel="noopener noreferrer">${location}</a>`;
    } else {
        locationElement.textContent = location || "No Location";
    }

    // document.getElementById("eventDescription").textContent = "Description: " + description || "No Description";
});

document.getElementById('closeButton').addEventListener('click', () => {
    chrome.runtime.sendMessage({ action: 'closeWindow' });
});

