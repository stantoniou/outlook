/* global Office */

let contactsData = [];
let allEmailsData = [];

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("scanButton").onclick = scanEmails;
        document.getElementById("clearButton").onclick = clearResults;
        document.getElementById("searchInput").addEventListener("input", filterContacts);
        
        showStatus("Ready to scan emails", "info");
    }
});

async function scanEmails() {
    const scanButton = document.getElementById("scanButton");
    const contactsList = document.getElementById("contactsList");
    const statusEl = document.getElementById("status");
    
    scanButton.disabled = true;
    contactsData = [];
    allEmailsData = [];
    
    contactsList.innerHTML = `
        <div class="loading">
            <div class="spinner"></div>
            <p>Scanning your emails...</p>
        </div>
    `;
    
    showStatus("Scanning emails...", "info");
    
    try {
        // Get the mailbox
        const mailbox = Office.context.mailbox;
        
        // For Outlook Web, we need to use REST API
        // Get access token
        mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
            if (result.status === "succeeded") {
                const accessToken = result.value;
                
                try {
                    await fetchEmailsFromREST(accessToken);
                    processContacts();
                    displayContacts();
                    showStatus(`Successfully scanned ${allEmailsData.length} emails`, "success");
                } catch (error) {
                    showStatus(`Error: ${error.message}`, "error");
                    contactsList.innerHTML = `
                        <div class="empty-state">
                            <div class="empty-state-icon">⚠️</div>
                            <p>Error scanning emails. Please try again.</p>
                        </div>
                    `;
                }
            } else {
                showStatus("Failed to get access token", "error");
                contactsList.innerHTML = `
                    <div class="empty-state">
                        <div class="empty-state-icon">⚠️</div>
                        <p>Authentication failed. Please try again.</p>
                    </div>
                `;
            }
            
            scanButton.disabled = false;
        });
        
    } catch (error) {
        showStatus(`Error: ${error.message}`, "error");
        scanButton.disabled = false;
        contactsList.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">⚠️</div>
                <p>Error scanning emails. Please try again.</p>
            </div>
        `;
    }
}

async function fetchEmailsFromREST(accessToken) {
    const restUrl = Office.context.mailbox.restUrl;
    
    // Fetch messages from Inbox (limit to 500 for performance)
    const getMessagesUrl = `${restUrl}/v2.0/me/messages?$top=500&$select=toRecipients,ccRecipients,bccRecipients,subject,receivedDateTime`;
    
    const response = await fetch(getMessagesUrl, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        }
    });
    
    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const data = await response.json();
    allEmailsData = data.value || [];
}

function processContacts() {
    const contactMap = new Map();
    
    allEmailsData.forEach(email => {
        // Process TO recipients
        if (email.toRecipients) {
            email.toRecipients.forEach(recipient => {
                addContactToMap(contactMap, recipient.emailAddress, 'to', email);
            });
        }
        
        // Process CC recipients
        if (email.ccRecipients) {
            email.ccRecipients.forEach(recipient => {
                addContactToMap(contactMap, recipient.emailAddress, 'cc', email);
            });
        }
        
        // Process BCC recipients
        if (email.bccRecipients) {
            email.bccRecipients.forEach(recipient => {
                addContactToMap(contactMap, recipient.emailAddress, 'bcc', email);
            });
        }
    });
    
    // Convert map to array and sort by email count
    contactsData = Array.from(contactMap.values()).sort((a, b) => b.totalCount - a.totalCount);
}

function addContactToMap(contactMap, emailAddress, recipientType, email) {
    if (!emailAddress || !emailAddress.address) return;
    
    const address = emailAddress.address.toLowerCase();
    
    if (!contactMap.has(address)) {
        contactMap.set(address, {
            email: emailAddress.address,
            name: emailAddress.name || emailAddress.address,
            toCount: 0,
            ccCount: 0,
            bccCount: 0,
            totalCount: 0,
            emails: []
        });
    }
    
    const contact = contactMap.get(address);
    
    if (recipientType === 'to') contact.toCount++;
    else if (recipientType === 'cc') contact.ccCount++;
    else if (recipientType === 'bcc') contact.bccCount++;
    
    contact.totalCount++;
    
    // Add email reference
    contact.emails.push({
        subject: email.subject || '(No subject)',
        receivedDateTime: email.receivedDateTime,
        recipientType: recipientType
    });
}

function displayContacts() {
    const contactsList = document.getElementById("contactsList");
    const statsContainer = document.getElementById("statsContainer");
    
    if (contactsData.length === 0) {
        contactsList.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">📭</div>
                <p>No contacts found in your emails</p>
            </div>
        `;
        statsContainer.style.display = 'none';
        return;
    }
    
    // Update statistics
    const totalEmails = allEmailsData.length;
    const totalContacts = contactsData.length;
    const avgPerContact = (totalEmails / totalContacts).toFixed(1);
    
    document.getElementById("totalContacts").textContent = totalContacts;
    document.getElementById("totalEmails").textContent = totalEmails;
    document.getElementById("avgPerContact").textContent = avgPerContact;
    statsContainer.style.display = 'grid';
    
    // Display contacts
    let html = '';
    
    contactsData.forEach(contact => {
        html += createContactCard(contact);
    });
    
    contactsList.innerHTML = html;
}

function createContactCard(contact) {
    const displayName = contact.name !== contact.email ? contact.name : '';
    
    return `
        <div class="contact-card" onclick="showContactDetails('${escapeHtml(contact.email)}')">
            <div class="contact-header">
                <div>
                    ${displayName ? `<div style="font-size: 13px; color: #6b7280; margin-bottom: 2px;">${escapeHtml(displayName)}</div>` : ''}
                    <div class="contact-email">${escapeHtml(contact.email)}</div>
                </div>
                <div class="contact-count">${contact.totalCount}</div>
            </div>
            <div class="contact-details">
                ${contact.toCount > 0 ? `<span class="detail-badge to">To: ${contact.toCount}</span>` : ''}
                ${contact.ccCount > 0 ? `<span class="detail-badge cc">CC: ${contact.ccCount}</span>` : ''}
                ${contact.bccCount > 0 ? `<span class="detail-badge bcc">BCC: ${contact.bccCount}</span>` : ''}
            </div>
        </div>
    `;
}

function filterContacts() {
    const searchTerm = document.getElementById("searchInput").value.toLowerCase();
    
    if (!searchTerm) {
        displayContacts();
        return;
    }
    
    const filteredContacts = contactsData.filter(contact => 
        contact.email.toLowerCase().includes(searchTerm) || 
        contact.name.toLowerCase().includes(searchTerm)
    );
    
    const contactsList = document.getElementById("contactsList");
    
    if (filteredContacts.length === 0) {
        contactsList.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">🔍</div>
                <p>No contacts match your search</p>
            </div>
        `;
        return;
    }
    
    let html = '';
    filteredContacts.forEach(contact => {
        html += createContactCard(contact);
    });
    
    contactsList.innerHTML = html;
}

function showContactDetails(email) {
    const contact = contactsData.find(c => c.email === email);
    
    if (!contact) return;
    
    // Sort emails by date (most recent first)
    const sortedEmails = contact.emails.sort((a, b) => 
        new Date(b.receivedDateTime) - new Date(a.receivedDateTime)
    );
    
    let emailsHtml = sortedEmails.slice(0, 10).map(email => {
        const date = new Date(email.receivedDateTime).toLocaleDateString();
        const typeColor = email.recipientType === 'to' ? '#1e40af' : 
                         email.recipientType === 'cc' ? '#92400e' : '#6b21a8';
        
        return `
            <div style="padding: 8px; border-bottom: 1px solid #f3f4f6; font-size: 12px;">
                <div style="color: ${typeColor}; font-weight: 600; text-transform: uppercase; font-size: 10px; margin-bottom: 4px;">
                    ${email.recipientType}
                </div>
                <div style="color: #1a1a1a; margin-bottom: 2px;">${escapeHtml(email.subject)}</div>
                <div style="color: #6b7280;">${date}</div>
            </div>
        `;
    }).join('');
    
    if (sortedEmails.length > 10) {
        emailsHtml += `<div style="padding: 12px; text-align: center; color: #6b7280; font-size: 12px;">And ${sortedEmails.length - 10} more emails...</div>`;
    }
    
    const contactsList = document.getElementById("contactsList");
    contactsList.innerHTML = `
        <div style="padding: 20px;">
            <button onclick="displayContacts()" style="margin-bottom: 16px; padding: 8px 16px; background: #f3f4f6; border: none; border-radius: 6px; cursor: pointer; font-size: 13px;">
                ← Back to all contacts
            </button>
            
            <div style="margin-bottom: 20px;">
                <div style="font-size: 18px; font-weight: 600; color: #1a1a1a; margin-bottom: 4px;">
                    ${escapeHtml(contact.name)}
                </div>
                <div style="color: #6b7280; margin-bottom: 12px;">
                    ${escapeHtml(contact.email)}
                </div>
                <div style="display: flex; gap: 12px;">
                    ${contact.toCount > 0 ? `<span class="detail-badge to">To: ${contact.toCount}</span>` : ''}
                    ${contact.ccCount > 0 ? `<span class="detail-badge cc">CC: ${contact.ccCount}</span>` : ''}
                    ${contact.bccCount > 0 ? `<span class="detail-badge bcc">BCC: ${contact.bccCount}</span>` : ''}
                </div>
            </div>
            
            <div style="background: white; border-radius: 8px; overflow: hidden; border: 1px solid #f3f4f6;">
                <div style="padding: 12px; background: #f9fafb; font-weight: 600; font-size: 13px;">
                    Recent Emails
                </div>
                ${emailsHtml}
            </div>
        </div>
    `;
}

function clearResults() {
    contactsData = [];
    allEmailsData = [];
    
    document.getElementById("contactsList").innerHTML = `
        <div class="empty-state">
            <div class="empty-state-icon">📬</div>
            <p>Click "Scan Emails" to analyze your inbox</p>
        </div>
    `;
    
    document.getElementById("statsContainer").style.display = 'none';
    document.getElementById("searchInput").value = '';
    showStatus("Results cleared", "info");
}

function showStatus(message, type) {
    const statusEl = document.getElementById("status");
    statusEl.textContent = message;
    statusEl.className = type;
    
    if (type === "success" || type === "info") {
        setTimeout(() => {
            statusEl.className = '';
            statusEl.style.display = 'none';
        }, 5000);
    }
}

function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
}
