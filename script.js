// Register Service Worker for PWA
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('./sw.js').catch(err => console.error(err));
    });
}


// =========================================================
// 2. CORE UI ELEMENTS
// =========================================================
const sizeSelect = document.getElementById('envelope-size');
const customSizeInputs = document.getElementById('custom-size-inputs');
const customWidth = document.getElementById('custom-width');
const customHeight = document.getElementById('custom-height');
const logoUpload = document.getElementById('logo-upload');
const returnInput = document.getElementById('return-address');
const saveSenderBtn = document.getElementById('save-sender-btn'); 
const recipientInput = document.getElementById('recipient-address');
const docRefInput = document.getElementById('doc-ref-input'); 
const docStampSelect = document.getElementById('doc-stamp-select'); 
const fontSizeSelect = document.getElementById('font-size-select');
const clearBtn = document.getElementById('clear-btn');
const smartClearCb = document.getElementById('smart-clear-cb');
const printBtn = document.getElementById('print-btn');

// D-PAD State
let currentPosX = 40;
let currentPosY = 35;
const nudgeDisplay = document.getElementById('nudge-display');

// PREVIEW ELEMENTS
const previewBox = document.getElementById('preview-box');
const previewRecipient = document.getElementById('preview-recipient');
const previewReturn = document.getElementById('preview-return');
const previewLogo = document.getElementById('preview-logo');
const previewDocRef = document.getElementById('preview-doc-ref'); 
const previewReturnCb = document.getElementById('preview-return-cb'); 
const previewGrid = document.getElementById('preview-grid'); 
const previewDocStamp = document.getElementById('preview-doc-stamp'); 

// Inject Safe Margin Div into Preview
if (!document.getElementById('preview-safe-margin')) {
    const smBox = document.createElement('div');
    smBox.id = 'preview-safe-margin';
    smBox.style.position = 'absolute';
    smBox.style.top = '4%'; smBox.style.bottom = '4%';
    smBox.style.left = '2%'; smBox.style.right = '2%';
    smBox.style.border = '2px dashed rgba(255, 0, 0, 0.4)';
    smBox.style.pointerEvents = 'none';
    smBox.style.zIndex = '10';
    previewBox.appendChild(smBox);
}
const safeMarginBox = document.getElementById('preview-safe-margin');

// TOGGLES
const enableSafeMargin = document.getElementById('enable-safe-margin');
const enableReturnCb = document.getElementById('enable-return-cb'); 
const enableGrid = document.getElementById('enable-grid'); 

// MODALS, ADDRESS BOOK & SEARCH
const clientSearch = document.getElementById('client-search');
const searchResults = document.getElementById('search-results');
const inlineAutocomplete = document.getElementById('inline-autocomplete');
const manageClientsBtn = document.getElementById('manage-clients-btn');
const clientModal = document.getElementById('client-modal');
const closeModalBtn = document.getElementById('close-modal-btn');
const saveClientBtn = document.getElementById('save-client-btn');
const newClientName = document.getElementById('new-client-name');
const newClientPhone = document.getElementById('new-client-phone'); 
const newClientAddress = document.getElementById('new-client-address');
const editClientId = document.getElementById('edit-client-id');
const clientListDiv = document.getElementById('client-list');

// BACKUP & RESTORE
const exportBackupBtn = document.getElementById('export-backup-btn');
const importBackupFile = document.getElementById('import-backup-file');

// LEDGER
const viewLedgerBtn = document.getElementById('view-ledger-btn');
const ledgerModal = document.getElementById('ledger-modal');
const closeLedgerBtn = document.getElementById('close-ledger-btn');
const ledgerBody = document.getElementById('ledger-body');
const clearLedgerBtn = document.getElementById('clear-ledger-btn');
const exportLedgerBtn = document.getElementById('export-ledger-btn');

let logoBase64 = null;


// =========================================================
// 3. UTILITIES & ROUTING
// =========================================================
function showToast(message) {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = 'toast'; toast.textContent = message;
    container.appendChild(toast);
    setTimeout(() => toast.classList.add('show'), 10);
    setTimeout(() => { toast.classList.remove('show'); setTimeout(() => toast.remove(), 300); }, 3000);
}

window.toggleCard = function(contentId) {
    const content = document.getElementById(contentId);
    const icon = document.getElementById(contentId + '-icon');
    if (content.classList.contains('hidden')) {
        content.classList.remove('hidden'); icon.classList.add('open');
    } else {
        content.classList.add('hidden'); icon.classList.remove('open');
    }
};

// PURE DOM ISOLATION
function openAddressBook() { clientModal.classList.add('active'); }
function closeAddressBook() { clientModal.classList.remove('active'); }
function openLedger() { renderLedger(); ledgerModal.classList.add('active'); }
function closeLedger() { ledgerModal.classList.remove('active'); }

manageClientsBtn.addEventListener('click', openAddressBook);
closeModalBtn.addEventListener('click', closeAddressBook);
viewLedgerBtn.addEventListener('click', openLedger);
closeLedgerBtn.addEventListener('click', closeLedger);

// =========================================================
// 4. LEDGER REUSE LOGIC
// =========================================================
let ledger = [];
try { ledger = JSON.parse(localStorage.getItem('courier_ledger')) || []; } catch(e) { ledger = []; }

function saveToLedger(recipientName, rawAddress, reference, stampVal) {
    ledger.unshift({ 
        id: Date.now(),
        date: new Date().toLocaleString(),
        name: (recipientName || "Manual Entry").split('\n')[0].substring(0, 40), 
        address: rawAddress || "",
        ref: reference || "None",
        stamp: stampVal || ""
    });
    localStorage.setItem('courier_ledger', JSON.stringify(ledger));
    renderLedger();
}

function renderLedger() {
    ledgerBody.innerHTML = '';
    if (ledger.length === 0) {
        ledgerBody.innerHTML = '<tr><td colspan="3" style="text-align:center; color:var(--text-hint); padding: 30px 10px;">No print history yet.</td></tr>';
        return;
    }
    ledger.forEach(entry => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="font-size:13px; color:var(--text-hint);">${entry.date.split(',')[0]}</td>
            <td style="color:var(--on-surface); font-weight:500;">
                ${entry.name}
                ${entry.ref !== "None" ? `<br><span style="font-size:12px; color:var(--primary); font-weight:bold;">Ref: ${entry.ref}</span>` : ''}
            </td>
            <td style="text-align:right;">
                <button class="text-btn" style="color:var(--primary); padding: 6px 12px; margin-left:auto; background: var(--primary-container);" onclick="reuseLedgerEntry(${entry.id})">
                    <span class="material-symbols-rounded" style="font-size: 16px;">refresh</span> Reuse
                </button>
            </td>
        `;
        ledgerBody.appendChild(tr);
    });
}

window.reuseLedgerEntry = function(id) {
    const entry = ledger.find(e => e.id === id);
    if (entry && entry.address) {
        recipientInput.value = entry.address;
        docRefInput.value = entry.ref !== "None" ? entry.ref : "";
        if (entry.stamp) docStampSelect.value = entry.stamp;
        closeLedger();
        updatePreview();
        showToast("♻️ Exact envelope loaded");
    } else {
        showToast("⚠️ This old entry cannot be reused.");
    }
};

clearLedgerBtn.addEventListener('click', () => {
    if(confirm("Clear the entire dispatch history?")) {
        ledger = []; localStorage.setItem('courier_ledger', JSON.stringify(ledger));
        renderLedger(); showToast("🗑️ History cleared");
    }
});

exportLedgerBtn.addEventListener('click', () => {
    if (ledger.length === 0) { showToast("⚠️ Nothing to export"); return; }
    let csvContent = "Date,Recipient,Reference\n"; // Removed data URI prefix
    ledger.forEach(e => { 
        const safeName = (e.name || "").replace(/"/g, '""');
        const safeRef = (e.ref || "").replace(/"/g, '""');
        csvContent += `"${e.date}","${safeName}","${safeRef}"\n`; 
    });
    
    // Use Blob to safely handle # and & symbols in addresses
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", "dispatch_ledger_" + new Date().toISOString().split('T')[0] + ".csv");
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
});

// =========================================================
// 5. ADDRESS BOOK & SEARCH
// =========================================================
let clients = [];
try { clients = JSON.parse(localStorage.getItem('courier_clients')) || []; } catch(e) { clients = []; }

function saveClientsToStorage() { localStorage.setItem('courier_clients', JSON.stringify(clients)); renderClientModalList(); }

// Main Screen Search Bar Logic
function renderSearchResults(query = "") {
    searchResults.innerHTML = '';
    const filtered = clients.filter(c => c.name.toLowerCase().includes(query.toLowerCase()));
    if (filtered.length === 0) {
        searchResults.innerHTML = '<div style="padding: 16px; color: var(--on-surface-variant); font-size: 14px;">No clients found.</div>';
    } else {
        filtered.forEach(client => {
            const div = document.createElement('div');
            div.className = 'search-result-item';
            div.textContent = client.name;
            div.addEventListener('click', () => {
                recipientInput.value = client.address;
                clientSearch.value = ''; 
                searchResults.classList.add('hidden');
                updatePreview();
                showToast("👤 Client selected");
            });
            searchResults.appendChild(div);
        });
    }
}
clientSearch.addEventListener('input', (e) => { searchResults.classList.remove('hidden'); renderSearchResults(e.target.value); });
clientSearch.addEventListener('focus', () => { searchResults.classList.remove('hidden'); renderSearchResults(clientSearch.value); });

// Inline Autocomplete Logic
recipientInput.addEventListener('input', (e) => {
    const val = e.target.value.trim();
    if (val.length < 2 || val.includes('\n')) { inlineAutocomplete.classList.add('hidden'); updatePreview(); return; }
    const filtered = clients.filter(c => c.name.toLowerCase().includes(val.toLowerCase()));
    
    if (filtered.length > 0) {
        inlineAutocomplete.innerHTML = '';
        filtered.forEach(client => {
            const div = document.createElement('div');
            div.className = 'autocomplete-item'; div.textContent = client.name;
            div.addEventListener('click', () => {
                recipientInput.value = client.address;
                inlineAutocomplete.classList.add('hidden'); 
                updatePreview(); 
                showToast("👤 Auto-filled address");
            });
            inlineAutocomplete.appendChild(div);
        });
        inlineAutocomplete.classList.remove('hidden');
    } else { inlineAutocomplete.classList.add('hidden'); }
    updatePreview();
});

// Close dropdowns when clicking outside
document.addEventListener('click', (e) => { 
    if (!e.target.closest('.search-dropdown-container')) { searchResults.classList.add('hidden'); }
    if (!e.target.closest('#recipient-address') && !e.target.closest('#inline-autocomplete')) { inlineAutocomplete.classList.add('hidden'); } 
});

// Modal Address Book Add/Edit
saveClientBtn.addEventListener('click', () => {
    const name = newClientName.value.trim(); const phone = newClientPhone.value.trim(); const address = newClientAddress.value.trim();
    if (!name || !address) { showToast("⚠️ Enter both Name and Address"); return; }
    const idToEdit = editClientId.value;
    if (idToEdit) {
        const index = clients.findIndex(c => c.id == idToEdit);
        if (index > -1) { clients[index].name = name; clients[index].phone = phone; clients[index].address = address; }
        showToast("✅ Client updated");
    } else {
        clients.push({ id: Date.now(), name: name, phone: phone, address: address }); showToast("✅ Client saved");
    }
    newClientName.value = ''; newClientPhone.value = ''; newClientAddress.value = ''; editClientId.value = '';
    saveClientBtn.textContent = 'Save Client'; saveClientsToStorage();
});

window.editClient = function(id) {
    const client = clients.find(c => c.id == id);
    if (client) {
        newClientName.value = client.name; newClientPhone.value = client.phone || ""; newClientAddress.value = client.address;
        editClientId.value = client.id; saveClientBtn.textContent = 'Update Client';
        document.querySelector('#client-modal .modal-content').scrollTo({ top: 0, behavior: 'smooth' });
    }
};

window.deleteClient = function(id) {
    if(confirm("Delete this client?")) { clients = clients.filter(c => c.id != id); saveClientsToStorage(); showToast("🗑️ Client deleted"); }
};

function renderClientModalList() {
    clientListDiv.innerHTML = '';
    if(clients.length === 0) { clientListDiv.innerHTML = '<p style="color:var(--text-hint); text-align:center;">No clients saved.</p>'; return; }
    [...clients].sort((a, b) => a.name.localeCompare(b.name)).forEach(client => {
        const item = document.createElement('div'); item.className = 'client-item';
        item.innerHTML = `
            <div class="client-name-disp">${client.name}</div>
            <div class="action-wrap">
                <button class="action-btn edit" onclick="editClient(${client.id})">Edit</button>
                <button class="action-btn del" onclick="deleteClient(${client.id})">Del</button>
            </div>
        `;
        clientListDiv.appendChild(item);
    });
}
renderClientModalList();

// --- BULK DOCX & TXT UPLOAD LOGIC ---
const bulkDocUpload = document.getElementById('bulk-doc-upload');

function processExtractedText(rawText) {
    // Clean up empty lines and extra spaces
    let lines = rawText.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    
    // Remove "TO" if it is the very first line
    if (lines.length > 0 && lines[0].toUpperCase() === 'TO') {
        lines.shift();
    }
    
    if (lines.length < 2) return false; // Needs at least a name and address
    
    // The first line is the name (strip trailing commas if present)
    let rawName = lines[0];
    let cleanName = rawName.endsWith(',') ? rawName.slice(0, -1) : rawName;
    
    // Check if client already exists to prevent duplicates
    if (!clients.some(c => c.name.toLowerCase() === cleanName.toLowerCase())) {
        clients.push({
            id: Date.now() + Math.floor(Math.random() * 1000), // Randomize ID for fast loops
            name: cleanName,
            phone: "",
            address: lines.join('\n') // Keep the exact address structure
        });
        return true;
    }
    return false;
}

if (bulkDocUpload) {
    bulkDocUpload.addEventListener('change', async (e) => {
        const files = e.target.files;
        if (!files || files.length === 0) return;
        
        let addedCount = 0;
        showToast(`⏳ Processing ${files.length} file(s)...`);

        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            
            try {
                if (file.name.endsWith('.docx')) {
                    const arrayBuffer = await file.arrayBuffer();
                    const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
                    if (processExtractedText(result.value)) addedCount++;
                } else if (file.name.endsWith('.txt')) {
                    const text = await file.text();
                    if (processExtractedText(text)) addedCount++;
                }
            } catch (err) {
                console.error(`Failed to process ${file.name}:`, err);
            }
        }
        
        saveClientsToStorage();
        showToast(`✅ Added ${addedCount} new addresses!`);
        bulkDocUpload.value = ''; // Reset input
    });
}


// --- JSON BACKUP & RESTORE LOGIC ---
exportBackupBtn.addEventListener('click', () => {
    const backupData = { clients: clients, ledger: ledger, senderText: returnInput.value, senderLogo: logoBase64 || "" };
    const blob = new Blob([JSON.stringify(backupData)], {type: 'application/json'});
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a'); link.href = url;
    link.download = `cover_printer_backup_${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
    showToast("💾 Backup downloaded safely!");
});

importBackupFile.addEventListener('change', (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
        try {
            const data = JSON.parse(event.target.result);
            if (!data || typeof data !== 'object') throw new Error("Invalid format");
            if (Array.isArray(data.clients)) { clients = data.clients; localStorage.setItem('courier_clients', JSON.stringify(clients)); }
            if (Array.isArray(data.ledger)) { ledger = data.ledger; localStorage.setItem('courier_ledger', JSON.stringify(ledger)); }
            if (data.senderText !== undefined) { returnInput.value = data.senderText; localStorage.setItem('courier_sender_text', data.senderText); }
            if (data.senderLogo) { logoBase64 = data.senderLogo; localStorage.setItem('courier_sender_logo', data.senderLogo); previewLogo.src = logoBase64; previewLogo.style.display = 'block'; } 
            else { logoBase64 = null; localStorage.removeItem('courier_sender_logo'); previewLogo.style.display = 'none'; }
            renderClientModalList(); renderLedger(); updatePreview(); showToast("🔄 Data restored successfully!");
        } catch(err) { showToast("❌ Error: Invalid backup file."); console.error(err); }
        importBackupFile.value = '';
    };
    reader.readAsText(file);
});

// =========================================================
// 6. SENDER DEFAULTS LOGIC
// =========================================================
const savedSenderText = localStorage.getItem('courier_sender_text');
const savedSenderLogo = localStorage.getItem('courier_sender_logo');
if (savedSenderText) returnInput.value = savedSenderText;
if (savedSenderLogo) { logoBase64 = savedSenderLogo; previewLogo.src = logoBase64; previewLogo.style.display = 'block'; }

saveSenderBtn.addEventListener('click', () => {
    localStorage.setItem('courier_sender_text', returnInput.value);
    if (logoBase64) localStorage.setItem('courier_sender_logo', logoBase64); else localStorage.removeItem('courier_sender_logo');
    showToast("💾 Defaults saved");
});

logoUpload.addEventListener('change', function(e) {
    if (e.target.files[0]) {
        const reader = new FileReader();
        reader.onload = function(event) { 
            logoBase64 = event.target.result; 
            previewLogo.src = logoBase64; 
            previewLogo.style.display = 'block'; 
            // Wait for the image to decode in the DOM before calculating naturalWidth
            previewLogo.onload = () => updatePreview(); 
        };
        reader.readAsDataURL(e.target.files[0]);
    } else { logoBase64 = null; previewLogo.style.display = 'none'; previewLogo.src = ''; }
});

// =========================================================
// 7. CORE PRINTING & PREVIEW LOGIC
// =========================================================
const sizes = { dl: { w: 220, h: 110 }, c5: { w: 229, h: 162 }, c4: { w: 324, h: 229 } };
function getCurrentSize() {
    if (sizeSelect.value === 'custom') { 
        const w = parseFloat(customWidth.value); const h = parseFloat(customHeight.value);
        return { width: isNaN(w) || w <= 0 ? 250 : w, height: isNaN(h) || h <= 0 ? 150 : h }; 
    }
    return { width: sizes[sizeSelect.value].w, height: sizes[sizeSelect.value].h };
}

function updatePreview() {
    // Add "To," permanently to the preview formatting
    if (recipientInput.value.trim() !== '') {
        // Auto-capitalize and indent every line of the address
        const uppercaseAddress = recipientInput.value.toUpperCase();
        const indentedAddress = uppercaseAddress.split('\n').map(line => '   ' + line).join('\n');
        previewRecipient.textContent = "To,\n" + indentedAddress;
    } else {
        previewRecipient.textContent = "";
    }

    const fromLines = returnInput.value.split('\n').filter(line => line.trim() !== '');
    if (fromLines.length > 0) {
        // Darkens the company name and lightens the address text for contrast
        previewReturn.innerHTML = `<strong style="color: #000; font-size: 1.1em;">${fromLines[0]}</strong><br><span>${fromLines.slice(1).join(', ')}</span>`; 
    } else {
        previewReturn.innerHTML = '';
    }

    if (sizeSelect.value === 'custom') customSizeInputs.classList.remove('hidden'); else customSizeInputs.classList.add('hidden');

    const selectedSize = getCurrentSize();
    previewBox.style.aspectRatio = `${selectedSize.width} / ${selectedSize.height}`;
    
    // Responsive S/M/L Font Sizing Match
    const currentWidth = previewBox.clientWidth || previewBox.getBoundingClientRect().width || 300;
    const fsMap = { 'S': 0.018, 'M': 0.025, 'L': 0.035 };
    const selectedFs = fontSizeSelect ? fontSizeSelect.value : 'M';
    previewBox.style.fontSize = `${currentWidth * fsMap[selectedFs]}px`;

    // Apply D-Pad Position
    previewRecipient.style.left = `${currentPosX}%`; 
    previewRecipient.style.top = `${currentPosY}%`;
    nudgeDisplay.innerHTML = `X: ${currentPosX}<br>Y: ${currentPosY}`;

    // Hardware Safe Margin Toggle
    if (safeMarginBox && enableSafeMargin) { safeMarginBox.style.display = enableSafeMargin.checked ? 'block' : 'none'; }

    if (docRefInput.value.trim() !== '') { previewDocRef.textContent = "Ref: " + docRefInput.value.trim(); previewDocRef.style.display = 'block'; } else previewDocRef.style.display = 'none';
    if (enableReturnCb.checked) previewReturnCb.style.display = 'block'; else previewReturnCb.style.display = 'none';
    if (enableGrid.checked) previewGrid.style.display = 'block'; else previewGrid.style.display = 'none';
    if (docStampSelect.value !== '') { previewDocStamp.textContent = `[ ${docStampSelect.value} ]`; previewDocStamp.style.display = 'block'; } else previewDocStamp.style.display = 'none';
}

// D-Pad Event Listeners
document.getElementById('nudge-left').addEventListener('click', () => { currentPosX = Math.max(2, currentPosX - 1); updatePreview(); });
document.getElementById('nudge-right').addEventListener('click', () => { currentPosX = Math.min(90, currentPosX + 1); updatePreview(); });
document.getElementById('nudge-up').addEventListener('click', () => { currentPosY = Math.max(5, currentPosY - 1); updatePreview(); });
document.getElementById('nudge-down').addEventListener('click', () => { currentPosY = Math.min(90, currentPosY + 1); updatePreview(); });

// Smart Clear Logic
clearBtn.addEventListener('click', () => {
    recipientInput.value = ''; docRefInput.value = '';
    if (!smartClearCb.checked) {
        currentPosX = 40; currentPosY = 35; sizeSelect.value = 'dl'; docStampSelect.value = 'TAX INVOICE';
        if(fontSizeSelect) fontSizeSelect.value = 'M';
    }
    updatePreview(); recipientInput.focus(); 
});

enableReturnCb.addEventListener('change', updatePreview); 
enableGrid.addEventListener('change', updatePreview); 
docStampSelect.addEventListener('change', updatePreview);
sizeSelect.addEventListener('change', updatePreview); 
customWidth.addEventListener('input', updatePreview); 
customHeight.addEventListener('input', updatePreview); 
returnInput.addEventListener('input', updatePreview);
if (fontSizeSelect) fontSizeSelect.addEventListener('change', updatePreview);
if (enableSafeMargin) enableSafeMargin.addEventListener('change', updatePreview);

window.addEventListener('resize', updatePreview);
updatePreview(); 

// --- PDF ENGINE ---
function drawEnvelopeLayout(doc, selectedSize, toAddressText, docRef, docStampVal) {
    if (enableGrid.checked) {
        doc.setDrawColor(200, 200, 200); doc.setLineWidth(0.1);
        for(let i = 0; i <= selectedSize.width; i += 10) doc.line(i, 0, i, selectedSize.height);
        for(let i = 0; i <= selectedSize.height; i += 10) doc.line(0, i, selectedSize.width, i);
    }

    const bottomY = selectedSize.height - 15; 
    doc.setDrawColor(150, 150, 150); doc.setLineWidth(0.3);
    doc.line(10, bottomY - 5, selectedSize.width - 10, bottomY - 5);

    let textStartX = 10;
    if (logoBase64) {
        const imageFormat = logoBase64.startsWith('data:image/png') ? 'PNG' : 'JPEG';
        // Smart scaling logic for natural aspect ratio
        const imgW = previewLogo.naturalWidth || 10; 
        const imgH = previewLogo.naturalHeight || 10;
        const aspect_ratio = imgW / imgH;
        
        const targetHeight = 10;
        const targetWidth = targetHeight * aspect_ratio; 
        
        doc.addImage(logoBase64, imageFormat, 10, bottomY - 2, targetWidth, targetHeight);
        textStartX = 10 + targetWidth + 4; // Pushes sender text securely to the right
    }

    const fromLines = returnInput.value.split('\n').filter(line => line.trim() !== '');
    if (fromLines.length > 0) {
        doc.setFont("helvetica", "bold"); doc.setFontSize(10); doc.setTextColor(0, 0, 0); 
        doc.text(fromLines[0], textStartX, bottomY);
        
        if (fromLines.length > 1) { 
            // Use normal weight for the address to make the Company Name stand out
            doc.setFont("helvetica", "normal"); doc.setFontSize(8); doc.setTextColor(80, 80, 80); 
            doc.text(fromLines.slice(1).join('  |  '), textStartX, bottomY + 5); 
        }
    }

    // "To," formatting with dynamic indent
    const pdfFsMap = { 'S': 11, 'M': 14, 'L': 18 };
    const selectedFs = fontSizeSelect ? fontSizeSelect.value : 'M';
    
        doc.setFont("helvetica", "bold");
    doc.setFontSize(pdfFsMap[selectedFs]); doc.setTextColor(0, 0, 0);
    
    const startX = (selectedSize.width * currentPosX) / 100;
    const startY = (selectedSize.height * currentPosY) / 100;
    
    if (toAddressText && toAddressText.trim() !== '') {
        doc.text("To,", startX, startY);
        // Push the entire address block 3mm to the right
        doc.text(toAddressText, startX + 3, startY + 6, { 
            maxWidth: selectedSize.width - (startX + 3) - 10, 
            lineHeightFactor: 1.5 
        }); 
    }

    if (docRef !== '') { doc.setFont("helvetica", "bold"); doc.setFontSize(10); doc.setTextColor(0, 0, 0); doc.text("Ref: " + docRef, selectedSize.width - 10, bottomY - 8, { align: 'right' }); }

    if (enableReturnCb.checked) {
        doc.setFont("helvetica", "normal"); doc.setFontSize(8); doc.setTextColor(100, 100, 100); const cbY = bottomY - 11;
        doc.text("If undelivered, tick reason:", 10, cbY); doc.setDrawColor(100, 100, 100); doc.setLineWidth(0.2);
        doc.rect(10, cbY + 2, 3, 3); doc.text("Shifted", 14, cbY + 4.5); doc.rect(26, cbY + 2, 3, 3); doc.text("Locked", 30, cbY + 4.5);
        doc.rect(42, cbY + 2, 3, 3); doc.text("Wrong Addr", 46, cbY + 4.5); doc.rect(64, cbY + 2, 3, 3); doc.text("Refused", 68, cbY + 4.5);
    }

    if (docStampVal && docStampVal !== '') {
        doc.setFont("helvetica", "bold"); doc.setFontSize(14); doc.setTextColor(100, 100, 100);
        doc.text(`[ ${docStampVal} ]`, selectedSize.width / 2, 15, { align: 'center' });
    }
}

// =========================================================
// 8. NATIVE ANDROID SHARE API (SHARING INTEGRATION)
// =========================================================
function triggerNativePrint(doc, filename) {
    const pdfBlob = doc.output('blob');
    const file = new File([pdfBlob], filename, { type: "application/pdf" });

    // Attempt to force the Native Share Menu directly (bypassing strict checks)
    if (navigator.share) {
        navigator.share({
            files: [file],
            title: 'Share Cover',
            text: 'Cover Document'
        }).then(() => {
            showToast("📤 Share menu opened");
        }).catch((err) => {
            console.error("Native share blocked by wrapper:", err);
            forceDownload(pdfBlob, filename); // Fallback if wrapper blocks it
        });
    } else {
        forceDownload(pdfBlob, filename);
    }
}

function forceDownload(pdfBlob, filename) {
    const pdfUrl = URL.createObjectURL(pdfBlob);
    const link = document.createElement("a");
    link.href = pdfUrl;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast("💾 Saved to Downloads folder");
}

printBtn.addEventListener('click', () => {
    showToast("⏳ Generating PDF...");
    try {
        const { jsPDF } = window.jspdf; const selectedSize = getCurrentSize();
        // Drop the explicit orientation string to prevent format array conflicts
        const doc = new jsPDF({ unit: 'mm', format: [selectedSize.width, selectedSize.height] });
        
        const refText = docRefInput.value.trim(); const stampVal = docStampSelect.value;
        drawEnvelopeLayout(doc, selectedSize, recipientInput.value, refText, stampVal);
        
        const filename = `Envelope_${Date.now()}.pdf`;
        triggerNativePrint(doc, filename);
        
        saveToLedger(recipientInput.value, recipientInput.value, refText, stampVal);
    } catch(err) { showToast("❌ Error generating PDF"); console.error(err); }
});
