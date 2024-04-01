chrome.runtime.onMessage.addListener(
    async (request, sender, sendResponse) => {
        if (request.action === "downloadEmail") {
            for (let i = 0; i < 3; i++) {
                console.log('Step: ' + i)
                await downloadEmail(request, sender, sendResponse, 2000 * i)
            }
        }
    },
);

async function downloadEmail(request, sender, sendResponse, scrollDown) {
    const listOfEmailItems = document.querySelectorAll('#MailList .customScrollBar > div > div > div > div');

    const emailsContent = [];
    for (let email of listOfEmailItems) {
        if (
            email.role !== 'heading'
            && email.clientHeight > 0
        ) {
            email.click();
            await delay(1000);
            console.log('downloadEmail accept');
            const emailContent = await getEmailContent();
            console.log(emailContent);
            emailsContent.push(emailContent);
            await downloadEML(emailContent)
        }
    }
    console.log(emailsContent);
    sendResponse({ content: emailsContent });
    const positionY = 1000 + scrollDown;
    document.querySelector('#MailList .customScrollBar').scroll(0, positionY);
}

async function getEmailContent() {
    let emailContent = {
        from: "",
        subject: "",
        body: "",
        attachments: [],
    };

    const headerElement = document.querySelector("#ConversationReadingPaneContainer [title]").getAttribute('title');
    if (headerElement) {
        emailContent.subject = headerElement;
    }

    const bodyElement = document.querySelector("#UniqueMessageBody");
    if (bodyElement) {
        emailContent.body = bodyElement.innerHTML;
    }

    let attachmentsFiles = document.querySelectorAll('.wide-content-host [role="listbox"] > div > div');

    if (attachmentsFiles.length > 0) {
        for (let i = 0; i < attachmentsFiles.length; i++) {
            document.querySelectorAll('.wide-content-host [role="listbox"] > div > div')[i].click()
            console.log('Open File Preview')
            await delay(1000)

            const attachmentHash = await getAttachment()
            emailContent.attachments.push(attachmentHash);

            // Close preview
            const mainBlock = document.querySelectorAll('.ms-Modal-scrollableContent > div > div')[0]
            let closeBlock = mainBlock.querySelector("button[aria-label='Закрыть']") ?? mainBlock.querySelector("button[aria-label='Close']")
            closeBlock = mainBlock.querySelector("button[aria-label='Жабу']") ?? closeBlock

            closeBlock.click()
            console.log('Close File Preview')
            await delay(1000)
        }
    }

    return emailContent;
}

async function getAttachment() {
    let previewWindow = document.querySelectorAll('.ms-Modal-scrollableContent > div > div');
    let fileName = previewWindow[0].querySelector('div').title;
    let filePath = '';

    let filePathElement = previewWindow[1].querySelector('object');
    let filePathImageElement = previewWindow[1].querySelector('img');
    if (filePathElement) {
        filePath = filePathElement.data;
    } else if (filePathImageElement) {
        filePath = filePathImageElement.src;
    }

    return {
        name: fileName,
        url: filePath,
    };
}

function delay(time) {
    return new Promise(resolve => setTimeout(resolve, time));
}

async function downloadEML(content) {
    let emlContent = `
From: ${content.from}
Subject: ${content.subject}
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary="boundary123"

--boundary123
Content-Type: text/html; charset="UTF-8"
Content-Transfer-Encoding: 7bit

!<DOCTYPE html>
<html>
<head>
    <style>
        body {
            color: #000000 !important;
            background-color: #FFFFFF !important;
        }
</style>
</head>
<body>
${content.body}
</body>
</html>

`;

    for (let attachment of content.attachments) {
        let mimeType;

        const extension = attachment.name.split('.').pop().toLowerCase();
        switch (extension) {
            case 'pdf':
                mimeType = 'application/pdf';
                break;
            case 'doc':
            case 'docx':
                mimeType = 'application/vnd.ms-word';
                break;
            case 'jpg':
            case 'jpeg':
                mimeType = 'image/jpeg';
                break;
            case 'xls':
            case 'xlsx':
                mimeType = 'application/vnd.ms-excel';
                break;
            case 'ppt':
            case 'pptx':
                mimeType = 'application/vnd.ms-powerpoint';
                break;
            default:
                mimeType = 'application/octet-stream';
        }

        try {
            const fileBase64 = await fetchFileAsBase64(attachment.url);
            emlContent += `--boundary123
Content-Type: ${mimeType}; name="${attachment.name}"
Content-Disposition: attachment; filename="${attachment.name}"
Content-Transfer-Encoding: base64

${fileBase64}

`;
        } catch (error) {
            console.error(`Failed to process ${attachment.name}:`, error);
        }
    }

    emlContent += `--boundary123--`;

    const blob = new Blob([emlContent], { type: 'message/rfc822' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'mail';
    document.body.appendChild(a);
    a.click();

    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}


function fetchFileAsBase64(url) {
    return fetch(url)
        .then(response => {
            if (!response.ok) {
                throw new Error(`Failed to fetch ${url}: ${response.statusText}`);
            }
            return response.arrayBuffer();
        })
        .then(buffer => {
            return arrayBufferToBase64(buffer);
        })
        .catch(error => {
            console.error('Error fetching file:', error);
            throw error;
        });
}

function arrayBufferToBase64(buffer) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
}

