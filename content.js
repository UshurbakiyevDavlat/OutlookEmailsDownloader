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

            emailContent.attachments.push(await getAttachment());

            // Close preview
            document.querySelectorAll('.ms-Modal-scrollableContent > div > div')[0].querySelector(
                'button[title="Close"]').click()
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
    // Replace image URLs with base64-encoded image data
    const contentWithEmbeddedImages = await replaceImageURLsWithBase64(content);

    // Generate EML content with embedded images
    const emlContent = generateEMLContent(contentWithEmbeddedImages);

    // Download the EML file
    downloadFile(emlContent, 'mail.eml');
}

async function replaceImageURLsWithBase64(content) {
    // Copy the content object to avoid modifying the original
    const modifiedContent = { ...content };

    // Replace image URLs with base64-encoded image data
    for (let attachment of modifiedContent.attachments) {
        if (attachment.url.startsWith('cid:')) {
            const base64ImageData = await fetchImageAsBase64(attachment.url);
            attachment.url = `data:image/jpeg;base64,${base64ImageData}`;
        }
    }

    return modifiedContent;
}

function generateEMLContent(content) {
    // Generate EML content with embedded images
    let emlContent = `
From: ${content.from}
Subject: ${content.subject}
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary="boundary123"

--boundary123
Content-Type: text/html; charset="UTF-8"
Content-Transfer-Encoding: 7bit

<DOCTYPE html>
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
        emlContent += `--boundary123
Content-Type: ${attachment.type}; name="${attachment.name}"
Content-Disposition: attachment; filename="${attachment.name}"
Content-Transfer-Encoding: base64

${attachment.data}

`;
    }

    emlContent += `--boundary123--`;

    return emlContent;
}

function downloadFile(content, filename) {
    // Create a Blob object from the content
    const blob = new Blob([content], { type: 'message/rfc822' });

    // Create a temporary URL for the Blob
    const url = URL.createObjectURL(blob);

    // Create a link element to trigger the download
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;

    // Trigger the download
    document.body.appendChild(a);
    a.click();

    // Cleanup
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

async function fetchImageAsBase64(url) {
    // Fetch the image and convert it to base64
    const response = await fetch(url);
    const blob = await response.blob();
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result.split(',')[1]);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
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

