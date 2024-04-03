const DELAY_MS = 1000;
let directoryHandle;

chrome.runtime.onMessage.addListener(async (request, sender, sendResponse) => {
    if (request.action === "downloadEmail") {
        const emailsContent = await downloadEmails(request, sender, sendResponse, request.amountOfEmails)
        sendResponse({content: emailsContent});
    }
});

async function downloadEmails(request, sender, sendResponse, amountOfEmails) {
    const emailsContent = [];

    for (let i = 0; i < amountOfEmails; i++) {
        document.querySelector("#MailList .customScrollBar div[aria-selected='true']");
        await delay(DELAY_MS);
        console.log('downloadEmail accept');

        const emailContent = await getEmailContent();
        console.log(emailContent);

        emailsContent.push(emailContent);
        await downloadEML(emailContent, i)
        document.querySelector("#MailList .customScrollBar div[aria-selected='true']").parentElement.parentElement.nextElementSibling.querySelector("div[aria-selected='false']").click()
    }


    return emailsContent;
}

async function getEmailContent() {
    const emailContent = {
        from: "",
        subject: "",
        body: "",
        attachments: []
    };

    const headerElement = document.querySelector("#ConversationReadingPaneContainer [title]");
    if (headerElement) {
        emailContent.subject = headerElement.getAttribute('title');
    }

    const bodyElement = document.querySelector("#UniqueMessageBody");
    if (bodyElement) {
        emailContent.body = bodyElement.innerHTML;
        const embeddedImages = bodyElement.querySelectorAll('img[data-imagetype="AttachmentByCid"]');
        let counter = 0;

        embeddedImages.forEach(image => {
            const attachmentInfo = {
                name: "Embedded Image" + counter++,
                url: image.src,
                isImage: true
            };
            emailContent.attachments.push(attachmentInfo);
        });
    }

    const attachments = document.querySelectorAll('.wide-content-host [role="listbox"] > div > div');
    for (let attachment of attachments) {
        attachment.click();
        await delay(DELAY_MS);
        const attachmentInfo = await getAttachment();
        emailContent.attachments.push(attachmentInfo);
    }

    return emailContent;
}

async function getAttachment() {
    const previewWindow = document.querySelectorAll('.ms-Modal-scrollableContent > div > div');
    const fileName = previewWindow[0].querySelector('div').title;

    let filePath = '';
    const filePathElement = previewWindow[1].querySelector('object');
    const filePathImageElement = previewWindow[1].querySelector('img');

    if (filePathElement) {
        filePath = filePathElement.data;
    } else if (filePathImageElement) {
        filePath = filePathImageElement.src;
    }

    return {
        name: fileName,
        url: filePath
    };
}

async function delay(time) {
    return new Promise(resolve => setTimeout(resolve, time));
}

async function downloadEML(content, counter) {
    if (!directoryHandle) {
        // If directoryHandle is not yet set, prompt user to choose a directory
        try {
            directoryHandle = await window.showDirectoryPicker();
        } catch (err) {
            console.error('Error when selecting directory:', err);
            return;
        }
    }

    let emlContent = `
From: ${content.from}
Subject: ${content.subject}
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary="boundary123"

--boundary123
Content-Type: text/html; charset="UTF-8"
Content-Transfer-Encoding: 7bit

<!DOCTYPE html>
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
        if (attachment.url) {
            try {
                let attachmentContent = '';
                if (attachment.isImage) {
                    attachmentContent = await fetchImageAsBase64(attachment.url);
                    emlContent += `--boundary123
Content-Type: image/jpeg; name="${attachment.name}"
Content-Disposition: inline; filename="${attachment.name}"
Content-Transfer-Encoding: base64

${attachmentContent}

`;
                } else {
                    attachmentContent = await fetchFileAsBase64(attachment.url);
                    emlContent += `--boundary123
Content-Type: application/octet-stream; name="${attachment.name}"
Content-Disposition: attachment; filename="${attachment.name}"
Content-Transfer-Encoding: base64

${attachmentContent}

`;
                }
            } catch (error) {
                console.error(`Failed to process ${attachment.name}:`, error);
            }
        }
    }

    emlContent += `--boundary123--`;

    const blob = new Blob([emlContent], { type: 'message/rfc822' });

    try {
        const fileHandle = await directoryHandle.getFileHandle(`mail-${counter}.eml`, { create: true });
        const writableStream = await fileHandle.createWritable();
        await writableStream.write(blob);
        await writableStream.close();
    } catch (err) {
        console.error('Error when saving file:', err);
    }
}

async function fetchImageAsBase64(url) {
    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(`Failed to fetch ${url}: ${response.statusText}`);
    }
    const blob = await response.blob();
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result.split(',')[1]);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

async function fetchFileAsBase64(url) {
    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(`Failed to fetch ${url}: ${response.statusText}`);
    }
    const buffer = await response.arrayBuffer();
    return arrayBufferToBase64(buffer);
}

function arrayBufferToBase64(buffer) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
}

