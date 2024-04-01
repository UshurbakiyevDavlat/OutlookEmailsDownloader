document.getElementById('downloadButton').addEventListener('click', () => {
    chrome.tabs.query({active: true, currentWindow: true}, (tabs) => {
        chrome.tabs.sendMessage(tabs[0].id, {action: "downloadEmail"}, (response) => {
            if (response.content) {
                document.getElementById('status').textContent = 'Download initiated!';
            } else {
                document.getElementById('status').textContent = 'Failed to download.';
            }
        });
    });
});
