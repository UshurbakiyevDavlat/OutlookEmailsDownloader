document.getElementById('downloadButton').addEventListener('click', () => {
    const amountOfEmails = Number(document.getElementById('amountOfEmails').value);
    if (amountOfEmails) {
        chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
            const messageData = {
                action: "downloadEmail",
                amountOfEmails: amountOfEmails,
            };
            document.getElementById('status').textContent = 'Working ...';

            chrome.tabs.sendMessage(tabs[0].id, messageData, (response) => {
                console.log(response)
                if (response.content) {
                    document.getElementById('status').textContent = 'Download initiated!';
                } else {
                    document.getElementById('status').textContent = 'Failed to download.';
                }
            });
        });
    } else {
        document.getElementById('status').textContent = 'Enter amount of emails';
    }
});
