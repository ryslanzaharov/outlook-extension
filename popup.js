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

async function fetchEvents() {
    const tokenData = await new Promise(resolve => {
        chrome.storage.local.get("token", resolve);
    });

    let accessToken = tokenData?.token;
    console.log("accessToken1: " + accessToken);

    if (!accessToken) {
        try {
            accessToken = await getAccessToken();
            console.log("accessToken2: " + accessToken);
        } catch (error) {
            console.error("Authorization failed:", error);
            document.getElementById("events").innerText = "Authorization required.";
            return;
        }
    }

    try {
        console.log("accessToken3: " + accessToken);
        const response = await fetch("https://graph.microsoft.com/v1.0/me/calendar/events", {
            headers: {
                "Authorization": `Bearer ${accessToken}`
            }
        });

        const data = await response.json();
        const eventsContainer = document.getElementById("events");
        eventsContainer.innerHTML = "";

        if (data.value && data.value.length > 0) {
            data.value.forEach(event => {
                const eventElement = document.createElement("div");
                eventElement.className = "event";
                eventElement.innerHTML = `<strong>${event.subject}</strong><br>${new Date(event.start.dateTime).toLocaleString()} - ${new Date(event.end.dateTime).toLocaleString()}`;
                eventsContainer.appendChild(eventElement);
            });
        } else {
            eventsContainer.innerText = "No upcoming events.";
        }
    } catch (error) {
        console.error("Error fetching events:", error);
        document.getElementById("events").innerText = "Error fetching events.";
    }
}

document.getElementById("logoutButton").addEventListener("click", () => {
    chrome.runtime.sendMessage({ action: "logout" });
});

document.getElementById("loginButton").addEventListener("click", () => {
    chrome.runtime.sendMessage({ action: "login" });
});

document.addEventListener("DOMContentLoaded", fetchEvents);
