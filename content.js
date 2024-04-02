const DELAY_MS = 1000;

chrome.runtime.onMessage.addListener(async (request, sender, sendResponse) => {
    if (request.action === "downloadEmail") {
        const emailsContent = await downloadEmails();
        sendResponse({ content: emailsContent });
    }
});

async function downloadEmails() {
    const emailsContent = [];

    await scrollToBottom('#MailList .customScrollBar > div > div > div > div');

    const listOfEmailItems = document.querySelectorAll('#MailList .customScrollBar > div > div > div > div');
    for (let email of listOfEmailItems) {
        if (email.role !== 'heading' && email.clientHeight > 0) {
            await ensureElementVisible(email);

            email.click();
            await delay(DELAY_MS);
            const emailContent = await getEmailContent();
            emailsContent.push(emailContent);
            await downloadEML(emailContent);
        }
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

    const attachmentInfo = {
        name: fileName,
        url: filePath
    };

    if (isImageAttachment(fileName)) {
        attachmentInfo.isImage = true;
    }

    return attachmentInfo;
}

function isImageAttachment(fileName) {
    const imageExtensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp'];
    const extension = fileName.split('.').pop().toLowerCase();
    return imageExtensions.includes(`.${extension}`);
}

async function scrollToBottom(selector) {
    return new Promise((resolve) => {
        const element = document.querySelector(selector);
        if (!element) {
            resolve();
            return;
        }
        const scrollHeight = element.scrollHeight;
        element.scrollTo({ top: scrollHeight, behavior: 'smooth' });
        setTimeout(() => {
            resolve();
        }, DELAY_MS);
    });
}

async function delay(time) {
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
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'mail.eml';
    document.body.appendChild(a);
    a.click();

    document.body.removeChild(a);
    URL.revokeObjectURL(url);
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

async function ensureElementVisible(element) {
    return new Promise((resolve) => {
        if (element.getBoundingClientRect().top >= 0 && element.getBoundingClientRect().bottom <= window.innerHeight) {
            resolve();
        } else {
            element.scrollIntoView({ behavior: "smooth" });
            setTimeout(() => {
                resolve();
            }, DELAY_MS);
        }
    });
}
