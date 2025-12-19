// Status Check on Load
document.addEventListener('DOMContentLoaded', async () => {
    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    if (tabs[0]) {
        chrome.tabs.sendMessage(tabs[0].id, { action: "check_status" }, (response) => {
            if (chrome.runtime.lastError) {
                // Content script not ready or not injected
                updateStatus("Waiting for page... (Try Refreshing Page)");
            } else if (response && response.message) {
                updateStatus(response.message);
                if (response.status === 'ready') {
                    // Maybe enable button or change color?
                }
            }
        });
    }
});

// Download PDF Handler
document.getElementById('dlPdfBtn').addEventListener('click', async () => {
    // Logic similar to startBtn, find tab and send message
    let targetTab = null;
    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    const activeTab = tabs[0];

    if (activeTab && activeTab.url && activeTab.url.includes("chrome-extension://")) {
        try {
            const matchedTabs = await chrome.tabs.query({ url: "*://coretaxdjp.pajak.go.id/*" });
            if (matchedTabs.length > 0) {
                targetTab = matchedTabs[0];
            } else {
                updateStatus("Error: Tab Pajak tidak ditemukan.");
                return;
            }
        } catch (err) {
            updateStatus('Error finding tab.');
            return;
        }
    } else {
        targetTab = activeTab;
    }

    if (!targetTab) {
        updateStatus('Error: Target tab not found.');
        return;
    }

    // Send PDF Download Signal
    chrome.tabs.sendMessage(targetTab.id, {
        action: "start_download_pdf"
    }, (response) => {
        if (chrome.runtime.lastError) {
            updateStatus('Error: ' + chrome.runtime.lastError.message);
        } else {
            updateStatus('Mulai download PDF...');
        }
    });
});


// Popup Logic
// document.getElementById('openTabBtn').addEventListener('click', () => {
//     chrome.tabs.create({ url: 'popup.html' });
// });

document.getElementById('startBtn').addEventListener('click', async () => {
    // Find the correct tab
    let targetTab = null;
    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    const activeTab = tabs[0];

    if (activeTab && activeTab.url && activeTab.url.includes("chrome-extension://")) {
        // Pop-out mode: search for tax page
        try {
            // Find tabs with coretaxdjp
            const matchedTabs = await chrome.tabs.query({ url: "*://coretaxdjp.pajak.go.id/*" });
            if (matchedTabs.length > 0) {
                targetTab = matchedTabs[0];
                updateStatus(`Targeting tab: ${targetTab.title}`);
            } else {
                updateStatus(`Error: Tax tab not found. Please open coretaxdjp.`);
                return;
            }
        } catch (err) {
            updateStatus('Error finding tab.');
            return;
        }
    } else {
        // Normal mode
        targetTab = activeTab;
    }

    if (!targetTab) {
        updateStatus('Error: No target tab found.');
        return;
    }

    // Send "AUTO" to content script
    chrome.tabs.sendMessage(targetTab.id, {
        action: "start_fetch",
        apiUrl: "AUTO",
        authToken: "AUTO",
        payload: null
    }, (response) => {
        if (chrome.runtime.lastError) {
            updateStatus('Error: ' + chrome.runtime.lastError.message + '. Try refreshing the Tax page.');
        } else {
            updateStatus('Request sent. Waiting for capture...');
        }
    });
});

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "update_status") {
        updateStatus(request.message);
    }
});

function updateStatus(msg) {
    const el = document.getElementById('status');
    if (el) el.innerText = msg;
}
