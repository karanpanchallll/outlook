Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Display current email subject
        displayEmailInfo();
    }
});

function displayEmailInfo() {
    Office.context.mailbox.item.subject.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            document.getElementById('emailSubject').textContent = result.value;
            document.getElementById('emailInfo').style.display = 'block';
        }
    });
}

async function triggerFlow() {
    const button = document.getElementById('triggerButton');
    const statusDiv = document.getElementById('statusMessage');
    
    // Disable button during request
    button.disabled = true;
    button.textContent = 'Triggering...';
    
    // Show info message
    showStatus('Sending request to Power Automate...', 'info');
    
    try {
        // Get email details to send to the flow
        const emailData = await getEmailData();
        
        // Replace with your Power Automate HTTP POST URL
        const flowUrl = 'https://default74afe875305e4ab4ba4ac1359a7629.ae.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/98694a6fe5ce4b1d8389f23d378bd9e0/triggers/manual/paths/invoke?api-version=1';
        
        // Make the request to Power Automate
        const response = await fetch(flowUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(emailData)
        });
        
        if (response.ok) {
            showStatus('Flow triggered successfully!', 'success');
        } else {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
    } catch (error) {
        console.error('Error triggering flow:', error);
        showStatus('Failed to trigger flow: ' + error.message, 'error');
    } finally {
        // Re-enable button
        button.disabled = false;
        button.textContent = 'Trigger Flow';
    }
}

async function getEmailData() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;
        
        // Get email subject
        item.subject.getAsync((subjectResult) => {
            if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error('Failed to get email subject'));
                return;
            }
            
            // Get sender email
            const from = item.from ? item.from.emailAddress : 'Unknown';
            
            // Get email ID
            const itemId = item.itemId;
            
            // Prepare data to send to Power Automate
            const emailData = {
                subject: subjectResult.value,
                from: from,
                itemId: itemId,
                triggeredAt: new Date().toISOString(),
                userEmail: Office.context.mailbox.userProfile.emailAddress
            };
            
            resolve(emailData);
        });
    });
}

function showStatus(message, type) {
    const statusDiv = document.getElementById('statusMessage');
    statusDiv.textContent = message;
    statusDiv.className = 'status ' + type;
    statusDiv.style.display = 'block';
    
    // Auto-hide after 5 seconds for success messages
    if (type === 'success') {
        setTimeout(() => {
            statusDiv.style.display = 'none';
        }, 5000);
    }
}