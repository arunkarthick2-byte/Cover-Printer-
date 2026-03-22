// ==========================================
// SOLLO ERP - MAIN APPLICATION CONTROLLER (v5.2 Enterprise)
// ==========================================

// --- NEW CODE: Import all our modules ---
import { 
    initDB, getAllRecords, getRecordById, saveRecord, deleteRecordById, 
    getAllFirms, saveInvoiceTransaction, getNextDocumentNumber, 
    getKhataStatement, getGlobalTimeline, exportDatabase, importDatabase, generateGSTReport 
} from './db.js?v=3';
import Utils from './utils.js?v=3';
import UI from './ui.js?v=3';
// --- END OF NEW CODE ---

const app = {
    state: { currentEditId: null, currentReceiptId: null, currentDocType: 'invoice', firmId: 'firm1' },

    // ==========================================
    // 1. BOOT SEQUENCE & FIRM MANAGEMENT
    // ==========================================
    init: async () => {
        try {
            // --- NEW CODE: Apply theme immediately on boot ---
            UI.initTheme();
            
            // NEW: Load our dynamic dropdowns
            app.loadDropdowns();
            // --- END OF NEW CODE ---

            // (Service Worker registration moved to index.html for PWABuilder)

            // --- NEW CODE: Request Persistent Storage ---
            if (navigator.storage && navigator.storage.persist) {
                const isPersisted = await navigator.storage.persist();
                console.log(`Persistent storage granted: ${isPersisted}`);
            } else {
                console.warn("Persistent storage API is not supported in this browser.");
            }
            // --- END OF NEW CODE ---

            // 2. Initialize the Database and load data
            await initDB();
            await app.loadFirms();
            
            await app.loadAllData(); 
            app.setupForms();
            
            // 3. Render the UI
            UI.renderDashboard();
            UI.applyFilters('sales');
            UI.applyFilters('purchases');
            UI.applyFilters('masters');
            UI.applyFilters('expenses');
            UI.applyFilters('cashbook');
            UI.applyFilters('timeline');
            
            UI.hideSplash();

            // FIX: Parse PWA Home Screen Shortcuts and route the user!
            const urlParams = new URLSearchParams(window.location.search);
            const action = urlParams.get('action');
            if (action === 'new_sale') {
                setTimeout(() => app.openForm('sales', null, 'invoice'), 300);
            } else if (action === 'cashbook') {
                setTimeout(() => { if(window.UI) window.UI.switchTab('cashbook'); }, 300);
            }

        } catch (e) { 
            console.error(e);
            alert("Critical Error: Offline storage could not be accessed."); 
            UI.hideSplash();
        }
    },

    loadFirms: async () => {
        let firms = typeof getAllFirms === 'function' ? await getAllFirms() : [];
        
        // Auto-initialize a default firm if the database is completely empty
        if (!firms || firms.length === 0) {
            const defaultFirm = {
                // --- NEW CODE: Dynamically generate a unique ID to prevent cloud sync collisions ---
                id: typeof Utils !== 'undefined' ? Utils.generateId() : 'firm-' + Date.now(),
                // --- END OF NEW CODE ---
                name: 'Main Company',
                phone: '',
                email: '',
                gst: '',
                address: '',
                state: ''
            };
            await saveRecord('firms', defaultFirm);
            firms = [defaultFirm];
        }
        
        app.state.firmId = firms[0].id;
        
        const firmData = await getRecordById('businessProfile', app.state.firmId);
        if (firmData) {
            document.getElementById('profile-name').value = firmData.name || '';
            document.getElementById('profile-phone').value = firmData.phone || '';
            document.getElementById('profile-email').value = firmData.email || '';
            document.getElementById('profile-gst').value = firmData.gst || '';
            document.getElementById('profile-address').value = firmData.address || '';
            document.getElementById('profile-state').value = firmData.state || '';
            document.getElementById('profile-bank').value = firmData.bankDetails || '';
            document.getElementById('profile-terms').value = firmData.terms || '';
            
            if (firmData.logo) {
                const img = document.getElementById('profile-logo-preview');
                img.src = firmData.logo;
                img.classList.remove('hidden');
            }
            if (firmData.signature) {
                const img = document.getElementById('profile-signature-preview');
                img.src = firmData.signature;
                img.classList.remove('hidden');
            }
        }
    },

    // ==========================================
    // 2. DATA HYDRATION & STATE MANAGEMENT
    // ==========================================
    loadAllData: async () => {
        // Load core data specifically for the active firml
        const stripBloat = (arr) => arr.map(r => { const c = {...r}; delete c.image; delete c.attachment; return c; });
        UI.state.rawData.sales = (await getAllRecords('sales')).filter(r => r.firmId === app.state.firmId);
        UI.state.rawData.purchases = (await getAllRecords('purchases')).filter(r => r.firmId === app.state.firmId);
        UI.state.rawData.items = stripBloat((await getAllRecords('items')).filter(r => r.firmId === app.state.firmId));
        UI.state.rawData.ledgers = (await getAllRecords('ledgers')).filter(r => r.firmId === app.state.firmId);
        UI.state.rawData.expenses = stripBloat((await getAllRecords('expenses')).filter(r => r.firmId === app.state.firmId));
        UI.state.rawData.cashbook = (await getAllRecords('receipts')).filter(r => r.firmId === app.state.firmId);
        UI.state.rawData.timeline = typeof getGlobalTimeline === 'function' ? await getGlobalTimeline(app.state.firmId) : [];
        
        // Load Bank Accounts
        UI.state.rawData.accounts = (await getAllRecords('accounts')).filter(r => r.firmId === app.state.firmId);

        // UPGRADE 1: Load Recycle Bin from LocalStorage
        const localTrash = JSON.parse(localStorage.getItem('sollo_trash') || '[]');
        UI.state.rawData.trash = localTrash.filter(t => t.firmId === app.state.firmId);

        // Refresh our new dynamic dropdowns
        app.loadDropdowns();

        // Dynamic Account Dropdown Hydration
        const accountDropdowns = ['sales-account-id', 'purchase-account-id', 'expense-account-id', 'pay-in-account', 'pay-out-account'];
        let accountOptions = UI.state.rawData.accounts.map(acc => 
            `<option value="${acc.id}">${acc.name} (${acc.type})</option>`
        ).join('');

        // CRITICAL FIX: Always include the default Cash Drawer if a custom one doesn't exist
        if (!UI.state.rawData.accounts.some(acc => acc.id === 'cash')) {
            accountOptions = `<option value="cash">Default Cash Drawer</option>` + accountOptions;
        }

        accountDropdowns.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.innerHTML = accountOptions;
        });

        // ==========================================
        // TRIGGER SILENT MEMORY OPTIMIZATION
        // ==========================================
        if (window.UI && typeof window.UI.optimizeMemory === 'function') {
            window.UI.optimizeMemory();
        }
    },

    refreshAll: async () => {
        await app.loadAllData();
        UI.applyFilters('sales');
        UI.applyFilters('purchases');
        UI.applyFilters('masters');
        UI.applyFilters('expenses');
        UI.applyFilters('cashbook');
        UI.applyFilters('timeline');
        UI.renderDashboard();
        
        // NEW: Check if we should run a silent background backup!
        if (typeof Cloud !== 'undefined' && Cloud.autoBackup) {
            Cloud.autoBackup();
        }
    },

    // ==========================================
    // NEW: SIMPLE MASTER CRUD ENGINE
    // ==========================================
    loadDropdowns: async () => {
        // 1. Auto-seed and Load Units
        let units = await getAllRecords('units');
        if (units.length === 0) {
            const defaults = ['Pcs', 'Kg', 'Mtr', 'Ltr', 'Box', 'Dozen', 'Tonnes'];
            for (let u of defaults) await saveRecord('units', { id: Utils.generateId(), name: u });
            units = await getAllRecords('units');
        }
        const uomSelect = document.getElementById('product-uom-select');
        if (uomSelect) uomSelect.innerHTML = units.map(u => `<option value="${u.name}">${u.name}</option>`).join('');

        // 2. Auto-seed and Load Categories
        let cats = await getAllRecords('expenseCategories');
        if (cats.length === 0) {
            const defaults = ['Salary', 'Rent', 'Electricity', 'Transport', 'Office Supplies', 'Marketing', 'Maintenance', 'Other'];
            for (let c of defaults) await saveRecord('expenseCategories', { id: Utils.generateId(), name: c });
            cats = await getAllRecords('expenseCategories');
        }
        const catSelect = document.getElementById('expense-category-select');
        if (catSelect) catSelect.innerHTML = '<option value="">Select Category...</option>' + cats.map(c => `<option value="${c.name}">${c.name}</option>`).join('');
    },

    manageSimpleMaster: async (storeName, title) => {
        const records = await getAllRecords(storeName);
        let listText = records.map((r, i) => `${i + 1}. ${r.name}`).join('\n');
        
        let action = prompt(`=== Manage ${title} ===\n\nCurrent List:\n${listText}\n\nType 'ADD' to create a new one, or 'DELETE' to remove one. (Leave blank to cancel)`);
        
        if (!action) return;
        
        if (action.toUpperCase() === 'ADD') {
            let newName = prompt(`Enter new ${title} name:`);
            if (newName && newName.trim()) {
                await saveRecord(storeName, { id: Utils.generateId(), name: newName.trim() });
                Utils.showToast("Added successfully! ✅");
                app.loadDropdowns(); // Refresh the UI instantly
            }
        } else if (action.toUpperCase() === 'DELETE') {
            let delName = prompt(`Enter the exact name to delete:\n\n${listText}`);
            if (!delName) return;
            let match = records.find(r => r.name.toLowerCase() === delName.toLowerCase().trim());
            if (match) {
                await deleteRecordById(storeName, match.id);
                Utils.showToast("Deleted successfully! 🗑️");
                app.loadDropdowns(); // Refresh the UI instantly
            } else {
                alert("Name not found. Please check the spelling.");
            }
        } else {
            alert("Invalid action. Please type ADD or DELETE.");
        }
    },

    // ==========================================
    // UPGRADE: BULK CSV IMPORT ENGINE
    // ==========================================
    downloadCSVTemplate: () => {
        const type = UI.state.currentMasterType;
        let csvContent = "";
        let filename = "";

        if (type === 'products') {
            csvContent = "Name,SellPrice,BuyPrice,OpeningStock,MinStock,UOM,GST_Percent,HSN_Code\n\"Premium Widget\",150.50,100,50,5,Pcs,18,123456\n\"Standard Item\",200,150,20,2,Kg,5,654321";
            filename = "SOLLO_Products_Template.csv";
        } else if (type === 'customers' || type === 'suppliers') {
            const partyType = type === 'customers' ? 'Customer' : 'Supplier';
            csvContent = `Name,Phone,GSTIN,City,State,Address,OpeningBalance,BalanceType\n"Acme Corp",9876543210,29ABCDE1234F1Z5,Bangalore,Karnataka,"123 Main St, Tech Park",1500,${partyType === 'Customer' ? 'To Receive' : 'To Pay'}\n"Global Traders",1234567890,,,Tamil Nadu,,0,Advance`;
            filename = `SOLLO_${partyType}s_Template.csv`;
        }

        if (csvContent && window.Utils) {
            window.Utils.downloadFile(csvContent, filename, "text/csv");
            window.Utils.showToast("Template Downloaded!");
        }
    },

    handleCSVImport: async (event) => {
        const file = event.target.files[0];
        if (!file) return;
        const type = UI.state.currentMasterType;
        
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const text = e.target.result;
                
                // FIX: Bulletproof CSV Parser that gracefully handles line-breaks INSIDE quotes
                const parseCSV = (str) => {
                    let p = '', row = [''], ret = [row], i = 0, r = 0, s = !0, l;
                    for (l of str) {
                        if ('"' === l) { if (s && l === p) row[i] += l; s = !s; }
                        else if (',' === l && s) l = row[++i] = '';
                        else if ('\n' === l && s) {
                            if ('\r' === p) row[i] = row[i].slice(0, -1);
                            row = ret[++r] = [l = '']; i = 0;
                        } else row[i] += l;
                        p = l;
                    }
                    return ret.map(row => row.map(col => col.trim().replace(/^"|"$/g, ''))).filter(row => row.length > 1 || row[0] !== '');
                };

                const parsedData = parseCSV(text);
                if (parsedData.length < 2) throw new Error("File is empty or invalid format.");

                const headers = parsedData[0].map(h => h.toLowerCase().replace(/[^a-z0-9]/g, ''));
                const rows = parsedData; // Map the new parser to the existing logic loop
                
                const parseCSVRow = (row) => row; // Bypass the old regex parser since we already parsed the whole file
                
                let successCount = 0;

                if (window.Utils) window.Utils.showToast("Importing records, please wait...");
                
                // FIX: Fetch existing database records ONCE outside the loop to prevent massive freezing!
                const existingItems = type === 'products' ? await getAllRecords('items') : [];
                
                for (let i = 1; i < rows.length; i++) {
                    const cols = parseCSVRow(rows[i]);
                    if (cols.length < 2 || !cols[0]) continue; // Skip bad rows

                    if (type === 'products') {
                        const nameIdx = headers.indexOf('name');
                        const sellIdx = headers.indexOf('sellprice');
                        const buyIdx = headers.indexOf('buyprice');
                        const stockIdx = headers.indexOf('openingstock');
                        const minStockIdx = headers.indexOf('minstock');
                        const uomIdx = headers.indexOf('uom');
                        const gstIdx = headers.indexOf('gstpercent');
                        const hsnIdx = headers.indexOf('hsncode');

                        if (nameIdx === -1 || !cols[nameIdx]) continue;

                        // NEW: Prevent duplicates by finding existing items by name
                        const match = existingItems.find(i => i.name.toLowerCase() === cols[nameIdx].toLowerCase() && i.firmId === app.state.firmId);

                        const data = {
                            id: match ? match.id : Utils.generateId(),
                            firmId: app.state.firmId,
                            name: cols[nameIdx],
                            sellPrice: parseFloat(cols[sellIdx]) || 0,
                            buyPrice: parseFloat(cols[buyIdx]) || 0,
                            // FIX: Preserve live stock if the item exists, otherwise use CSV opening stock
                            stock: match ? match.stock : (parseFloat(cols[stockIdx]) || 0),
                            minStock: parseFloat(cols[minStockIdx]) || 0,
                            uom: cols[uomIdx] || 'Pcs',
                            gst: parseFloat(cols[gstIdx]) || 0,
                            hsn: cols[hsnIdx] || ''
                        };
                        await saveRecord('items', data);
                        
                        // FIX: Update the memory array so if the CSV has the same item twice, it doesn't create a duplicate!
                        if (!match) existingItems.push(data);
                        else Object.assign(match, data);
                        
                        successCount++;
                    } 
                    else if (type === 'customers' || type === 'suppliers') {
                        const nameIdx = headers.indexOf('name');
                        const phoneIdx = headers.indexOf('phone');
                        const gstIdx = headers.indexOf('gstin');
                        const cityIdx = headers.indexOf('city');
                        const stateIdx = headers.indexOf('state');
                        const addrIdx = headers.indexOf('address');
                        const obIdx = headers.indexOf('openingbalance');
                        const typeIdx = headers.indexOf('balancetype');

                        if (nameIdx === -1 || !cols[nameIdx]) continue;

                        const pType = type === 'customers' ? 'Customer' : 'Supplier';
                        const balTypeRaw = cols[typeIdx] ? cols[typeIdx].toLowerCase() : '';
                        let balType = pType === 'Customer' ? 'To Receive / Debit' : 'To Pay / Credit';
                        
                        if (balTypeRaw.includes('pay') || balTypeRaw.includes('credit')) balType = 'To Pay / Credit';
                        else if (balTypeRaw.includes('receive') || balTypeRaw.includes('debit')) balType = 'To Receive / Debit';

                        const data = {
                            id: Utils.generateId(),
                            firmId: app.state.firmId,
                            name: cols[nameIdx],
                            type: pType,
                            phone: cols[phoneIdx] || '',
                            gst: cols[gstIdx] ? cols[gstIdx].toUpperCase() : '',
                            city: cols[cityIdx] || '',
                            state: cols[stateIdx] || '',
                            address: cols[addrIdx] || '',
                            openingBalance: parseFloat(cols[obIdx]) || 0,
                            balanceType: balType
                        };
                        await saveRecord('ledgers', data);
                        successCount++;
                    }
                }
                
                alert(`✅ Successfully imported ${successCount} records!`);
                event.target.value = ''; // Reset file input
                app.refreshAll();
            } catch (err) {
                console.error(err);
                alert("Error importing CSV. Please ensure you are using the correct template.");
                event.target.value = '';
            }
        };
        reader.readAsText(file);
    },

    // ==========================================
    // 3. LEDGER & RETURNS ENGINE
    // ==========================================
    openPartyLedger: async (partyId, partyType, partyName) => {
        document.getElementById('report-party-name').innerText = partyName;
        const timelineContainer = document.getElementById('timeline-list');
        
        // UPGRADE 3: Inject Skeleton Loaders instead of text
        timelineContainer.innerHTML = '<div class="skeleton-card"></div><div class="skeleton-card" style="animation-delay: 0.1s"></div><div class="skeleton-card" style="animation-delay: 0.2s"></div>';
        
        // NEW: Bind the edit button to the currently open party
        const editBtn = document.getElementById('btn-edit-ledger');
        if(editBtn) editBtn.onclick = () => { 
            UI.closeActivity('activity-report-viewer'); // FIX: Slide down the ledger statement first
            setTimeout(() => app.openForm('ledger', partyId), 150); // Open the form smoothly
        };
        
        UI.openActivity('activity-report-viewer');

        try {
            const statement = await getKhataStatement(partyId, partyType);
            
            // BULLETPROOF MATH: Filter out Draft Invoices to perfectly match the Address Book!
            const validTimeline = [];
            let runningBal = 0;
            
            statement.timeline.forEach(t => {
                if (t.isInvoice && t.id !== 'open-bal') {
                    const doc = partyType === 'Customer' ? 
                        UI.state.rawData.sales.find(s => s.id === t.id) : 
                        UI.state.rawData.purchases.find(p => p.id === t.id);
                    if (doc && doc.status === 'Open') return; // Skip Drafts completely!
                }
                runningBal += (t.impact || 0);
                t.runningBalance = runningBal;
                validTimeline.push(t);
            });
            
            statement.timeline = validTimeline;
            const bal = runningBal; // Use the true calculated balance

            const balEl = document.getElementById('report-party-balance');
            let balText = '';
            let balColor = 'var(--md-text-muted)';

            if (partyType === 'Customer') {
                if (bal > 0.01) { balText = `Closing Balance: \u20B9${bal.toFixed(2)} (To Receive)`; balColor = 'var(--md-error)'; }
                else if (bal < -0.01) { balText = `Closing Balance: \u20B9${Math.abs(bal).toFixed(2)} (Advance)`; balColor = 'var(--md-success)'; }
                else { balText = `Closing Balance: \u20B90.00`; balColor = 'var(--md-on-surface)'; }
            } else {
                if (bal > 0.01) { balText = `Closing Balance: \u20B9${bal.toFixed(2)} (To Pay)`; balColor = 'var(--md-error)'; }
                else if (bal < -0.01) { balText = `Closing Balance: \u20B9${Math.abs(bal).toFixed(2)} (Advance)`; balColor = 'var(--md-success)'; }
                else { balText = `Closing Balance: \u20B90.00`; balColor = 'var(--md-on-surface)'; }
            }
            
            balEl.innerText = balText;
            balEl.style.color = balColor;

            // Move state assignment ABOVE the early return to clear previous memory!
            UI.state.rawData.timeline = statement.timeline;

            if (statement.timeline.length === 0) {
                timelineContainer.innerHTML = '<p class="empty-state">No transactions found for this party.</p>';
                return;
            }
            
            timelineContainer.innerHTML = statement.timeline.map(t => {
                const isPayment = !t.isInvoice && t.id !== 'open-bal';
                const icon = t.id === 'open-bal' ? 'account_balance' : (isPayment ? 'payments' : 'receipt_long');
                const iconBg = t.id === 'open-bal' ? '#e3f2fd' : (isPayment ? '#e8f5e9' : '#fff0f2');
                const iconColor = t.id === 'open-bal' ? '#0061a4' : (isPayment ? '#2e7d32' : '#ba1a1a');
                const amtColor = t.isInvoice ? 'var(--md-error)' : 'var(--md-success)';

                return `
                <div class="m3-card" style="display:flex; align-items:center; gap: 12px; padding: 12px; margin-bottom: 8px;">
                    <div class="icon-circle" style="background: ${iconBg}; color: ${iconColor}; width: 40px; height: 40px; border-radius: 50%; display: flex; justify-content: center; align-items: center; flex-shrink: 0;">
                        <span class="material-symbols-outlined" style="font-size: 20px;">${icon}</span>
                    </div>
                    <div style="flex: 1;">
                        <strong class="large-text">${t.desc}</strong><br>
                        <small style="color: var(--md-text-muted);">${t.date}</small>
                    </div>
                    <div style="text-align:right;">
                        <strong style="font-size: 14px; color: ${amtColor};">${t.isInvoice ? '+' : '-'}\u20B9${(t.amount || 0).toFixed(2)}</strong><br>
                        <small style="color: var(--md-text-muted);">Bal: \u20B9${(t.runningBalance || 0).toFixed(2)}</small>
                    </div>
                </div>`;
            }).join('');

            const chips = document.querySelectorAll('#activity-report-viewer .filter-chips .chip');
            if(chips.length > 0) {
                chips.forEach(c => c.classList.remove('active'));
                chips[0].classList.add('active'); 
            }
            UI.state.activeFilters['timeline'] = 'All';

        } catch (error) {
            console.error("Error loading ledger:", error);
            timelineContainer.innerHTML = '<p class="empty-state" style="color:var(--md-error);">Failed to load statement.</p>';
        }
    },
        openAccountLedger: async (accountId) => {
        // Fetch Account Details
        let account = { name: 'Cash Drawer', openingBalance: 0 };
        if (accountId !== 'cash') {
            account = await getRecordById('accounts', accountId) || account;
        }

        document.getElementById('report-party-name').innerText = account.name;
        const timelineContainer = document.getElementById('timeline-list');
        
        // UPGRADE 3: Inject Skeleton Loaders instead of text
        timelineContainer.innerHTML = '<div class="skeleton-card"></div><div class="skeleton-card" style="animation-delay: 0.1s"></div><div class="skeleton-card" style="animation-delay: 0.2s"></div>';
        
        // Bind the Edit Button in the top right corner
        const editBtn = document.getElementById('btn-edit-ledger');
        if(editBtn) {
            if(accountId === 'cash') {
                editBtn.onclick = () => alert('The Default Cash Drawer is a system account and cannot be modified directly.');
            } else {
                editBtn.onclick = () => { 
                    UI.closeActivity('activity-report-viewer'); // FIX: Slide down the ledger statement first
                    setTimeout(() => app.openForm('account', accountId), 150); // Open the form smoothly
                };
            }
        }
        
        UI.openActivity('activity-report-viewer');

        try {
            let openingBalance = parseFloat(account.openingBalance) || 0;
            let timeline = [];

            if (openingBalance !== 0) {
                timeline.push({
                    id: 'open-bal',
                    date: 'Opening',
                    desc: 'Opening Balance',
                    amount: Math.abs(openingBalance),
                    impact: openingBalance,
                    isInvoice: false
                });
            }

            const receipts = await getAllRecords('receipts');
            const firmId = app.state.firmId;
            
            // Filter out only the transactions for this specific Bank/Cash account
            const accountReceipts = receipts.filter(r => {
                if (r.firmId !== firmId) return false;
                if (accountId === 'cash') return r.accountId === 'cash' || !r.accountId;
                return r.accountId === accountId;
            });

            accountReceipts.forEach(r => {
                const isMoneyIn = r.type === 'in';
                const impact = isMoneyIn ? parseFloat(r.amount) : -parseFloat(r.amount);
                
                // FIX: Translate cross-linked Invoice/PO/Expense IDs into readable numbers for Bank Verification
                let displayRefs = '';
                if (r.invoiceRef) {
                    const refs = String(r.invoiceRef).split(',').map(x => x.trim());
                    const names = refs.map(ref => {
                        const s = UI.state.rawData.sales.find(doc => doc.id === ref || doc.invoiceNo === ref || doc.orderNo === ref);
                        if (s) return s.orderNo || s.invoiceNo || String(s.id).slice(-4).toUpperCase();
                        
                        const p = UI.state.rawData.purchases.find(doc => doc.id === ref || doc.invoiceNo === ref || doc.poNo === ref || doc.orderNo === ref);
                        if (p) return p.orderNo || p.poNo || p.invoiceNo || String(p.id).slice(-4).toUpperCase();
                        
                        const e = UI.state.rawData.expenses.find(doc => doc.id === ref || doc.expenseNo === ref);
                        if (e) return e.expenseNo || String(e.id).slice(-4).toUpperCase();
                        
                        return ref.startsWith('sollo-') ? ref.slice(-4).toUpperCase() : ref;
                    });
                    displayRefs = [...new Set(names)].join(', ');
                }
                
                let finalRef = r.ref || '';
                if (displayRefs) {
                    finalRef = finalRef ? `${finalRef} | Docs: ${displayRefs}` : `Docs: ${displayRefs}`;
                }

                timeline.push({
                    id: r.id,
                    date: r.date,
                    // UPGRADE: Prepend the custom receipt/voucher number for even better tracking
                    desc: r.receiptNo ? `${r.receiptNo} - ${r.desc || (isMoneyIn ? 'Money In' : 'Money Out')}` : (r.desc || (isMoneyIn ? 'Money In' : 'Money Out')),
                    partyName: r.ledgerName || '',
                    amount: parseFloat(r.amount),
                    impact: impact,
                    isInvoice: false,
                    ref: finalRef // This automatically passes into the PDF engine!
                });
            });

            // Sort chronologically (BULLETPROOF)
            timeline.sort((a, b) => {
                if (a.id === 'open-bal') return -1;
                if (b.id === 'open-bal') return 1;
                return String(a.date || '').localeCompare(String(b.date || ''));
            });

            // Calculate Running Balance
            let runningBalance = 0;
            timeline.forEach(t => {
                runningBalance += t.impact;
                t.runningBalance = runningBalance;
            });

            const balEl = document.getElementById('report-party-balance');
            balEl.innerText = `Available Balance: \u20B9${runningBalance.toFixed(2)}`;
            balEl.style.color = runningBalance >= 0 ? 'var(--md-success)' : 'var(--md-error)';

            // Move state assignment ABOVE the early return!
            UI.state.rawData.timeline = timeline; // Hooks into your existing PDF generator!

            if (timeline.length === 0) {
                timelineContainer.innerHTML = '<p class="empty-state">No transactions found for this account.</p>';
                return;
            }
            
            // Render the timeline
            timelineContainer.innerHTML = timeline.map(t => {
                const isPaymentOut = t.impact < 0;
                const icon = t.id === 'open-bal' ? 'account_balance' : (isPaymentOut ? 'arrow_upward' : 'arrow_downward');
                const iconBg = t.id === 'open-bal' ? '#e3f2fd' : (isPaymentOut ? '#fff0f2' : '#e8f5e9');
                const iconColor = t.id === 'open-bal' ? '#0061a4' : (isPaymentOut ? '#ba1a1a' : '#2e7d32');
                const amtColor = isPaymentOut ? 'var(--md-error)' : 'var(--md-success)';
                const sign = isPaymentOut ? '-' : '+';

                return `
                <div class="m3-card" style="display:flex; align-items:center; gap: 12px; padding: 12px; margin-bottom: 8px;">
                    <div class="icon-circle" style="background: ${iconBg}; color: ${iconColor}; width: 40px; height: 40px; border-radius: 50%; display: flex; justify-content: center; align-items: center; flex-shrink: 0;">
                        <span class="material-symbols-outlined" style="font-size: 20px;">${icon}</span>
                    </div>
                    <div style="flex: 1;">
                        <strong class="large-text">${t.desc}</strong><br>
                        <small style="color: var(--md-text-muted);">${t.date} ${t.partyName ? '| ' + t.partyName : ''} ${t.ref ? '<br><span style="color:var(--md-primary); font-size:10px; font-weight:bold;">Ref: ' + t.ref + '</span>' : ''}</small>
                    </div>
                    <div style="text-align:right;">
                        <strong style="font-size: 14px; color: ${amtColor};">${sign}\u20B9${Math.abs(t.amount || 0).toFixed(2)}</strong><br>
                        <small style="color: var(--md-text-muted);">Bal: \u20B9${(t.runningBalance || 0).toFixed(2)}</small>
                    </div>
                </div>`;
            }).join('');

            // Reset filters
            const chips = document.querySelectorAll('#activity-report-viewer .filter-chips .chip');
            if(chips.length > 0) {
                chips.forEach(c => c.classList.remove('active'));
                chips[0].classList.add('active'); 
            }
            UI.state.activeFilters['timeline'] = 'All';

        } catch (error) {
            console.error("Error loading account ledger:", error);
            timelineContainer.innerHTML = '<p class="empty-state" style="color:var(--md-error);">Failed to load statement.</p>';
        }
    },


    loadOriginalDocuments: (partyId, type) => {
        if (app.state.currentDocType !== 'return') return;

        const storeName = type === 'sales' ? 'sales' : 'purchases';
        const rawData = storeName === 'sales' ? UI.state.rawData.sales : UI.state.rawData.purchases;
        const partyKey = type === 'sales' ? 'customerId' : 'supplierId';

        const validInvoices = rawData.filter(doc => 
            doc[partyKey] === partyId && 
            doc.documentType !== 'return' && 
            doc.status !== 'Open'
        );

        const selectEl = document.getElementById(`${type}-original-ref`);
        if (selectEl) {
            const options = validInvoices.map(doc => 
                `<option value="${doc.id}">${doc.invoiceNo || doc.poNo} (${doc.date}) - \u20B9${doc.grandTotal}</option>`
            ).join('');
            selectEl.innerHTML = options ? `<option value="">Select Original Document...</option>` + options : `<option value="">No valid invoices found.</option>`;
        }
    },

    loadReturnItems: async (originalDocId, type) => {
        if (!originalDocId) return;
        
        const storeName = type === 'sales' ? 'sales' : 'purchases';
        const originalDoc = await getRecordById(storeName, originalDocId);
        if (!originalDoc) return;

        const allDocs = storeName === 'sales' ? UI.state.rawData.sales : UI.state.rawData.purchases;
        const previousReturns = allDocs.filter(d => d.documentType === 'return' && d.orderNo === originalDoc.invoiceNo);
        
        const returnedQtyMap = {};
        previousReturns.forEach(ret => {
            ret.items.forEach(item => {
                returnedQtyMap[item.itemId] = (returnedQtyMap[item.itemId] || 0) + parseFloat(item.qty);
            });
        });

        const tbody = document.getElementById(`${type}-items-body`);
        tbody.innerHTML = ''; 

        originalDoc.items.forEach(item => {
            const previouslyReturned = returnedQtyMap[item.itemId] || 0;
            const maxAllowable = parseFloat(item.qty) - previouslyReturned;

            if (maxAllowable > 0) {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>
                        <div style="font-weight:500;">${item.name}</div>
                        <small style="color:var(--md-text-muted);">Max Return: ${maxAllowable}</small>
                        <input type="hidden" class="row-item-id" value="${item.itemId}">
                        <input type="hidden" class="row-item-name" value="${(item.name || '').replace(/"/g, '&quot;')}">                        <input type="hidden" class="row-item-buyprice" value="${item.buyPrice || 0}">
                    </td>
                    <td><input type="text" class="row-hsn" value="${item.hsn || ''}" readonly style="width:60px; text-align:center; padding:4px;"></td>
                    <td><input type="number" inputmode="decimal" class="row-qty" value="0" min="0" max="${maxAllowable}" step="any" oninput="UI.calc${type.charAt(0).toUpperCase() + type.slice(1)}Totals()" style="width:60px; padding:4px; border-color:var(--md-error);"></td>
                    <td><input type="text" class="row-uom" value="${item.uom || ''}" readonly style="width:50px; padding:4px;"></td>
                    <td><input type="number" inputmode="decimal" inputmode="decimal" class="row-rate" value="${item.rate}" step="any" readonly style="width:80px; padding:4px; background:var(--md-surface-variant);"></td>
                    <td><input type="number" inputmode="decimal" inputmode="decimal" class="row-gst" value="${item.gstPercent || 0}" step="any" oninput="UI.calc${type.charAt(0).toUpperCase() + type.slice(1)}Totals()" style="width:50px; padding:4px;"></td>
                    <td class="row-total" style="font-weight:bold; text-align:right;">0.00</td>
                    <td style="text-align:center;"><span class="material-symbols-outlined tap-target" style="color:var(--md-error);" onclick="this.closest('tr').remove(); UI.calc${type.charAt(0).toUpperCase() + type.slice(1)}Totals()">cancel</span></td>
                `;
                tbody.appendChild(tr);
            }
        });
        
        const orderNoEl = document.getElementById(`${type}-order-no`);
        if (orderNoEl) orderNoEl.value = originalDoc.invoiceNo;
        
        type === 'sales' ? UI.calcSalesTotals() : UI.calcPurchaseTotals();
    },
            loadPendingInvoices: async (partyId, type) => {
        const isMoneyIn = type === 'in';
        const storeName = isMoneyIn ? 'sales' : 'purchases';
        const partyKey = isMoneyIn ? 'customerId' : 'supplierId';
        const selectId = isMoneyIn ? 'pay-in-invoice-ref' : 'pay-out-invoice-ref';
        
        const selectEl = document.getElementById(selectId);
        if (!selectEl) return;
        
        if (!partyId) {
            selectEl.innerHTML = '<option value="">On Account / Advance</option>';
            return;
        }

        const allDocs = await getAllRecords(storeName);
        const activeFirmId = app.state ? app.state.firmId : 'firm1';
        
        // Calculate true balance to only show genuinely unpaid invoices
        const receipts = await getAllRecords('receipts');
        const paymentMap = {};
        receipts.forEach(c => {
            if (c.firmId === activeFirmId && c.invoiceRef && c.ledgerId === partyId) {
                let amt = parseFloat(c.amount) || 0;
                if (storeName === 'sales') amt = c.type === 'in' ? amt : -amt;
                if (storeName === 'purchases') amt = c.type === 'out' ? amt : -amt;
                
                // NEW: Split amount evenly if linked to multiple invoices
                const refs = String(c.invoiceRef).split(',').map(r => r.trim());
                if (refs.length > 0) {
                    let splitAmt = amt / refs.length;
                    refs.forEach(r => {
                        paymentMap[r] = (paymentMap[r] || 0) + splitAmt;
                    });
                }
            }
        });

        const pendingDocs = allDocs.filter(doc => {
            if (doc.firmId !== activeFirmId || doc[partyKey] !== partyId) return false;
            if (doc.status === 'Open' || doc.documentType === 'return') return false;
            
            // BULLETPROOF MATH: Safely catches ghost IDs and clean Order Numbers
            // FIX: Use a Set to prevent double-counting if orderNo and invoiceNo are identical strings
            const uniqueRefs = [...new Set([doc.orderNo, doc.invoiceNo, doc.poNo, doc.id].filter(Boolean))];
            const paid = uniqueRefs.reduce((sum, ref) => sum + (paymentMap[ref] || 0), 0);
            
            const balance = (parseFloat(doc.grandTotal) || 0) - paid;
            return balance > 0.01; 
        });

        if (pendingDocs.length === 0) {
            selectEl.innerHTML = '<option value="">No pending invoices (On Account)</option>';
        } else {
            const options = pendingDocs.map(doc => {
                const uniqueRefs = [...new Set([doc.orderNo, doc.invoiceNo, doc.poNo, doc.id].filter(Boolean))];
                const paid = uniqueRefs.reduce((sum, ref) => sum + (paymentMap[ref] || 0), 0);
                
                const balance = (parseFloat(doc.grandTotal) || 0) - paid;
                
                // BULLETPROOF: Directly save the clean Order/PO number to the database!
                const docNo = (isMoneyIn ? (doc.orderNo || doc.invoiceNo) : (doc.orderNo || doc.poNo || doc.invoiceNo)) || doc.id;
                
                let displayNo = '';
                if (isMoneyIn) {
                    displayNo = doc.orderNo || doc.invoiceNo || String(doc.id).slice(-4).toUpperCase();
                } else {
                    displayNo = doc.orderNo || doc.poNo || doc.invoiceNo || String(doc.id).slice(-4).toUpperCase();
                }
                
                return `<option value="${docNo}" data-bal="${balance.toFixed(2)}">${displayNo} (Due: \u20B9${balance.toFixed(2)})</option>`;
            }).join('');
            selectEl.innerHTML = '<option value="">On Account / Advance</option>' + options;
        }
    },

    // ==========================================
    // 4. FORM INITIALIZATION & ROUTING
    // ==========================================
    openForm: async (type, id = null, docType = 'invoice') => {
        UI.closeAllBottomSheets();
        app.state.currentEditId = id;
        app.state.currentDocType = docType;
        
        const form = document.getElementById(`form-${type}`);
        if(form) {
            form.reset();
            // Wipe ghost images from previous sessions
            form.querySelectorAll('img').forEach(img => {
                img.src = '';
                img.classList.add('hidden');
            });
        }
        
        // Hide delete buttons for new records
        UI.toggleDeleteButton(type, !!id);

        if (type === 'sales' || type === 'purchase') {
            document.getElementById(`${type}-items-body`).innerHTML = '';
            document.getElementById(`${type}-subtotal`).innerText = '\u20B90.00';
            document.getElementById(`${type}-gst-total`).innerText = '\u20B90.00';
            document.getElementById(`${type}-grand-total`).innerText = '\u20B90.00';
            document.getElementById(`${type}-${type === 'sales' ? 'customer' : 'supplier'}-display`).innerText = `Select ${type === 'sales' ? 'Customer' : 'Supplier'}...`;
            document.getElementById(`${type}-${type === 'sales' ? 'customer' : 'supplier'}-display`).style.color = 'var(--md-text-muted)';
            
            // NEW: Hide payment history when creating a brand new record
            const historyCard = document.getElementById(`${type}-payment-history-card`);
            if (historyCard) historyCard.classList.add('hidden');
            const expenseCard = document.getElementById(`${type}-expense-history-card`);
            if (expenseCard) expenseCard.classList.add('hidden');
            const titleEl = document.getElementById(`form-title-${type}`);
            const btnEl = document.getElementById(`btn-save-${type}`);
            const headerEl = document.getElementById(`activity-${type}-form`).querySelector('.activity-header');

            if (docType === 'return') {
                if (titleEl) titleEl.innerText = id ? `Edit ${type === 'sales' ? 'Credit Note' : 'Debit Note'}` : `New ${type === 'sales' ? 'Credit Note' : 'Debit Note'}`;
                if (btnEl) btnEl.innerText = `Save ${type === 'sales' ? 'Credit Note' : 'Debit Note'}`;
                if (headerEl) headerEl.style.backgroundColor = type === 'sales' ? '#fff0f2' : '#e8f5e9';
                
                const refGroup = document.getElementById(`${type}-return-ref-group`);
                if (refGroup) refGroup.classList.remove('hidden');
            } else {
                if (titleEl) titleEl.innerText = id ? `Edit ${type === 'sales' ? 'Sales Invoice' : 'Purchase Bill'}` : `New ${type === 'sales' ? 'Sales Invoice' : 'Purchase Bill'}`;
                if (btnEl) btnEl.innerText = `Save ${type === 'sales' ? 'Invoice' : 'Purchase'}`;
                if (headerEl) headerEl.style.backgroundColor = 'var(--md-surface)';
                
                const refGroup = document.getElementById(`${type}-return-ref-group`);
                if (refGroup) refGroup.classList.add('hidden');
            }
            
            UI.toggleDates(type);
        }

        if (id) {
            await app.populateEditForm(type, id);
        } else {
            const dateInput = document.getElementById(`${type}-date`);
            if(dateInput && typeof Utils !== 'undefined' && Utils.getLocalDate) {
                dateInput.value = Utils.getLocalDate();
                if (dateInput._flatpickr) dateInput._flatpickr.setDate(Utils.getLocalDate()); // FIX: Sync Flatpickr
            }
            
            // Injecting the Editable Auto-Numbering Engine
            if (type === 'sales' && typeof getNextDocumentNumber === 'function') {
                const prefix = docType === 'return' ? 'CN' : 'INV';
                document.getElementById('sales-invoice-no').value = await getNextDocumentNumber('sales', prefix);
                // NEW: Auto increment Order Ref No as well
                document.getElementById('sales-order-no').value = await getNextDocumentNumber('sales', 'ORD', 'orderNo');
            } else if (type === 'purchase' && typeof getNextDocumentNumber === 'function') {
                document.getElementById('purchase-po-no').value = ''; // Keep Supplier Bill blank
                // NEW: Auto-increment Our PO No instead!
                document.getElementById('purchase-order-no').value = await getNextDocumentNumber('purchases', 'PO', 'orderNo');
            }
        }
        
        // NEW: Expense form specific logic
        if (type === 'expense') {
            if (!id && typeof getNextDocumentNumber === 'function') {
                document.getElementById('expense-no').value = await getNextDocumentNumber('expenses', 'EXP', 'expenseNo');
            }
            
            // Reset the display label
            const displayEl = document.getElementById('expense-linked-display');
            if (displayEl) {
                displayEl.innerText = '-- No Link (General Expense) --';
                displayEl.style.color = 'var(--md-text-muted)';
            }
            
            // Build the new search screen
            if (typeof app.loadLinkedDocsList === 'function') app.loadLinkedDocsList();
        }
        
        UI.openActivity(`activity-${type}-form`);
    },

    // ==========================================
    // 5. EDIT MODE HYDRATION
    // ==========================================
    populateEditForm: async (type, id) => {
        // FIXED: Added 'account': 'accounts' to ensure bank records can be retrieved for editing
        const storeMap = { 'product': 'items', 'ledger': 'ledgers', 'expense': 'expenses', 'sales': 'sales', 'purchase': 'purchases', 'account': 'accounts' };
        const record = await getRecordById(storeMap[type], id);
        if (!record) return;

        if (type === 'sales' || type === 'purchase') {
            const setDateSafe = (id, val) => {
                const el = document.getElementById(id);
                if (el) {
                    el.value = val;
                    if (el._flatpickr) el._flatpickr.setDate(val); // FIX: Sync Flatpickr
                }
            };
            setDateSafe(`${type}-date`, record.date || '');
            setDateSafe(`${type}-order-date`, record.orderDate || '');
            setDateSafe(`${type}-shipped-date`, record.shippedDate || '');
            setDateSafe(`${type}-completed-date`, record.completedDate || '');
            document.getElementById(`${type}-order-no`).value = record.orderNo || '';
            document.getElementById(`${type}-order-status`).value = record.status || 'Completed';
            document.getElementById(`${type}-freight`).value = record.freightAmount || 0;
            document.getElementById(`${type}-discount`).value = record.discount || 0;
            
            const discTypeEl = document.getElementById(`${type}-discount-type`);
            if (discTypeEl) discTypeEl.value = record.discountType || '\u20B9';
            
            const invTypeEl = document.getElementById(`${type}-invoice-type`);
            if (invTypeEl) invTypeEl.value = record.invoiceType || 'B2B';

            const notesEl = document.getElementById(`${type}-internal-notes`);
            if (notesEl) notesEl.value = record.internalNotes || '';

            const partyKey = type === 'sales' ? 'customer' : 'supplier';
            const partyIdKey = type === 'sales' ? 'customerId' : 'supplierId';
            const partyNameKey = type === 'sales' ? 'customerName' : 'supplierName';
            
            document.getElementById(`${type}-${partyKey}-id`).value = record[partyIdKey] || '';
            document.getElementById(`${type}-${partyKey}-display`).innerText = record[partyNameKey] || `Select ${type === 'sales' ? 'Customer' : 'Supplier'}...`;
            document.getElementById(`${type}-${partyKey}-display`).style.color = 'var(--md-on-surface)';

            if (type === 'sales') {
                document.getElementById('sales-invoice-no').value = record.invoiceNo || '';
                // Removed the deleted instant payment fields from the hydration engine
            } else {
                document.getElementById('purchase-po-no').value = record.invoiceNo || record.poNo || '';
                // Removed the deleted instant payment fields from the hydration engine
            }

            const tbody = document.getElementById(`${type}-items-body`);
            tbody.innerHTML = '';

            // --- FETCH RETURN MAX ALLOWABLE LOGIC ---
            let returnedQtyMap = {};
            let originalDoc = null;
            if (record.documentType === 'return' && record.orderNo) {
                const allDocs = storeMap[type] === 'sales' ? UI.state.rawData.sales : UI.state.rawData.purchases;
                originalDoc = allDocs.find(d => (d.invoiceNo === record.orderNo || d.poNo === record.orderNo) && d.documentType !== 'return');
                const previousReturns = allDocs.filter(d => d.documentType === 'return' && d.orderNo === record.orderNo && d.id !== record.id);
                previousReturns.forEach(ret => ret.items.forEach(i => returnedQtyMap[i.itemId] = (returnedQtyMap[i.itemId] || 0) + parseFloat(i.qty)));
            }
            (record.items || []).forEach(item => {
                let maxHtml = 'min="0.01" step="any"';
                let maxLabel = '';
                
                if (record.documentType === 'return' && originalDoc) {
                    const origItem = originalDoc.items.find(i => i.itemId === item.itemId);
                    if (origItem) {
                        const maxAllowable = parseFloat(origItem.qty) - (returnedQtyMap[item.itemId] || 0);
                        maxHtml = `min="0" max="${maxAllowable}" step="any"`;
                        maxLabel = `<br><small style="color:var(--md-text-muted);">Max Return: ${maxAllowable}</small>`;
                    }
                }

                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>
                        <div style="font-weight:500;">${item.name}</div>
                        ${maxLabel}
                        ${type === 'sales' && record.documentType !== 'return' ? `
                        <div style="display:flex; align-items:center; gap:4px; margin-top:4px;">
                            <span style="font-size:11px; color:var(--md-text-muted);">Buy: ₹</span>
                            <input type="number" inputmode="decimal" inputmode="decimal" class="row-item-buyprice" value="${item.buyPrice || 0}" step="any" oninput="UI.calcSalesTotals()" style="width:60px; padding:2px 4px; font-size:11px; border:1px solid var(--md-outline-variant); border-radius:4px; background:var(--md-surface);">
                        </div>
                        <small class="live-margin" style="font-size:10px; display:block; margin-top:4px;"></small>
                        ` : `<input type="hidden" class="row-item-buyprice" value="${item.buyPrice || 0}">`}
                        <input type="hidden" class="row-item-id" value="${item.itemId}">
                        <input type="hidden" class="row-item-name" value="${(item.name || '').replace(/"/g, '&quot;')}">
                    </td>
                    <td><input type="text" class="row-hsn" value="${item.hsn || ''}" readonly style="width:60px; text-align:center; padding:4px;"></td>
                    <td><input type="number" inputmode="decimal" class="row-qty" value="${item.qty}" ${maxHtml} oninput="UI.calc${type.charAt(0).toUpperCase() + type.slice(1)}Totals()" style="width:60px; padding:4px; ${record.documentType === 'return' ? 'border-color:var(--md-error);' : ''}"></td>
                    <td><input type="text" class="row-uom" value="${item.uom || ''}" readonly style="width:50px; padding:4px;"></td>
                    <td><input type="number" inputmode="decimal" inputmode="decimal" class="row-rate" value="${item.rate}" step="any" ${record.documentType === 'return' ? 'readonly' : ''} oninput="UI.calc${type.charAt(0).toUpperCase() + type.slice(1)}Totals()" style="width:80px; padding:4px; ${record.documentType === 'return' ? 'background:var(--md-surface-variant);' : ''}"></td>
                    <td><input type="number" inputmode="decimal" inputmode="decimal" class="row-gst" value="${item.gstPercent || 0}" step="any" oninput="UI.calc${type.charAt(0).toUpperCase() + type.slice(1)}Totals()" style="width:50px; padding:4px;"></td>
                    <td class="row-total" style="font-weight:bold; text-align:right;">0.00</td>
                    <td style="text-align:center;"><span class="material-symbols-outlined tap-target" style="color:var(--md-error);" onclick="this.closest('tr').remove(); UI.calc${type.charAt(0).toUpperCase() + type.slice(1)}Totals()">cancel</span></td>
                `;
                tbody.appendChild(tr);
            });
            if (record.documentType === 'return') {
                const refEl = document.getElementById(`${type}-original-ref`);
                if (refEl) {
                    refEl.innerHTML = `<option value="">${record.orderNo || 'Linked Reference'}</option>`;
                    refEl.disabled = true; 
                }
            }

            type === 'sales' ? UI.calcSalesTotals() : UI.calcPurchaseTotals();
            UI.toggleDates(type);

            // NEW: Fetch and Display all Linked Receipts / Vouchers
            const historyCard = document.getElementById(`${type}-payment-history-card`);
            const historyList = document.getElementById(`${type}-payment-history-list`);
            if (historyCard && historyList) {
                // FIX: Check ALL references to catch cross-linked payments in the history view!
                const uniqueRefs = [...new Set([record.orderNo, record.invoiceNo, record.poNo, record.id].filter(Boolean))];
                const partyId = type === 'sales' ? record.customerId : record.supplierId;
                const allReceipts = await getAllRecords('receipts');
                
                const linkedReceipts = allReceipts.filter(r => {
                    if (r.firmId !== app.state.firmId || r.ledgerId !== partyId) return false;
                    const refs = String(r.invoiceRef || '').split(',').map(x => x.trim());
                    return refs.some(ref => uniqueRefs.includes(ref));
                });

                if (linkedReceipts.length > 0) {
                    historyCard.classList.remove('hidden');
                    historyList.innerHTML = linkedReceipts.map(r => {
                        // Split the display amount if it's a multi-invoice payment
                        const refCount = String(r.invoiceRef || '').split(',').length || 1;
                        const splitAmt = (parseFloat(r.amount) || 0) / refCount;
                        return `
                        <div style="display:flex; justify-content:space-between; align-items:center; padding:8px 0; border-bottom:1px dashed var(--md-outline-variant);">
                            <div><div style="font-weight:bold; color:var(--md-primary);">${r.receiptNo || (r.isAutoGenerated ? 'Auto-Generated Receipt' : 'Manual Receipt')}</div><small style="color:var(--md-text-muted);">${r.date} | Mode: ${r.mode}</small></div>
                            <div style="font-weight:bold; font-size:16px; color:${type === 'sales' ? 'var(--md-success)' : 'var(--md-error)'};">${type === 'sales' ? '+' : '-'}&#8377;${splitAmt.toFixed(2)}</div>
                        </div>`;
                    }).join('');
                } else {
                    historyCard.classList.add('hidden');
                    historyList.innerHTML = '';
                }
            }

            // NEW: Fetch and Display all Linked Expenses (Job Costing)
            const expCard = document.getElementById(`${type}-expense-history-card`);
            const expList = document.getElementById(`${type}-expense-history-list`);
            if (expCard && expList) {
                const uniqueRefs = [...new Set([record.orderNo, record.invoiceNo, record.poNo, record.id].filter(Boolean))];
                const allExpenses = await getAllRecords('expenses');
                // FIX: Trim whitespace and check against ALL possible IDs to prevent broken links!
                const linkedExpenses = allExpenses.filter(e => {
                    if (e.firmId !== app.state.firmId || !e.linkedInvoice) return false;
                    const cleanLinks = e.linkedInvoice.split(',').map(link => link.trim());
                    return cleanLinks.some(link => uniqueRefs.includes(link));
                });

                if (linkedExpenses.length > 0) {
                    expCard.classList.remove('hidden');
                    expList.innerHTML = linkedExpenses.map(e => `
                        <div style="display:flex; justify-content:space-between; align-items:center; padding:8px 0; border-bottom:1px dashed var(--md-outline-variant);">
                            <div><div style="font-weight:bold; color:var(--md-error);">${e.expenseNo || 'EXP'} - ${e.category}</div><small style="color:var(--md-text-muted);">${e.date} | ${e.notes || 'No notes'}</small></div>
                            <div style="font-weight:bold; font-size:16px; color:var(--md-error);">\u20B9${parseFloat(e.amount).toFixed(2)}</div>
                        </div>
                    `).join('');
                } else {
                    expCard.classList.add('hidden');
                    expList.innerHTML = '';
                }
            }

        } else {
            const form = document.getElementById(`form-${type}`);
            if(!form) return;
            const elements = form.elements;
            for (let i = 0; i < elements.length; i++) {
                const el = elements[i];
                if (el.name && record[el.name] !== undefined) {
                    if (el.type === 'checkbox') {
                        el.checked = record[el.name];
                    } else {
                        el.value = record[el.name];
                        if (el._flatpickr) el._flatpickr.setDate(record[el.name]); // FIX: Sync Flatpickr dynamically
                    }
                }
            }
            
            // Recover images so they aren't erased on save
            if (type === 'product' && record.image) {
                const img = document.getElementById('product-image-preview');
                if (img) { img.src = record.image; img.classList.remove('hidden'); }
            }
            if (type === 'expense') {
                if (record.attachment) {
                    const img = document.getElementById('expense-attachment-preview');
                    if (img) { img.src = record.attachment; img.classList.remove('hidden'); }
                }
                // NEW: Recover the document names for the UI Display (Multi-Link)
                if (record.linkedInvoice) {
                    const links = record.linkedInvoice.split(',').map(x => x.trim()).filter(x => x);
                    const displayNames = links.map(linkId => {
                        // SELF-HEALING: Catch broken fragments like '8965' or '0778'
                        const sDoc = UI.state.rawData.sales.find(s => s.id === linkId || s.invoiceNo === linkId || s.orderNo === linkId || s.id.endsWith(linkId));
                        const pDoc = UI.state.rawData.purchases.find(p => p.id === linkId || p.poNo === linkId || p.invoiceNo === linkId || p.orderNo === linkId || p.id.endsWith(linkId));
                        
                        if (sDoc) return sDoc.orderNo || sDoc.invoiceNo || sDoc.id.slice(-4).toUpperCase();
                        if (pDoc) return pDoc.orderNo || pDoc.poNo || pDoc.invoiceNo || pDoc.id.slice(-4).toUpperCase();
                        
                        return linkId.startsWith('sollo-') ? linkId.slice(-4).toUpperCase() : linkId;
                    });
                    
                    const displayEl = document.getElementById('expense-linked-display');
                    if(displayEl) {
                        // SELF-HEALING: Remove duplicates
                        displayEl.innerText = [...new Set(displayNames)].join(' | ');
                        displayEl.style.color = 'var(--md-primary)';
                    }
                }
            }
        }
    },

    // ==========================================
    // 6. MASTER FORM SUBMISSIONS & EVENT LISTENERS
    // ==========================================
    setupForms: () => {
        // ------------------ SALES & PURCHASE SUBMISSIONS ------------------
        ['sales', 'purchase'].forEach(type => {
            const form = document.getElementById(`form-${type}`);
            if (form) {
                form.addEventListener('submit', async (e) => {
                    e.preventDefault();
                    
                    // NEW: Lock the submit button to prevent duplicate entries
                    const submitBtn = document.getElementById(`btn-save-${type}`);
                    // FIXED: Capture the dynamic text (e.g., "Save Credit Note") before overriding it
                    const originalText = submitBtn ? submitBtn.innerText : `Save ${type === 'sales' ? 'Invoice' : 'Purchase'}`;
                    
                    if (submitBtn) {
                        submitBtn.disabled = true;
                        submitBtn.innerText = "Saving...";
                        submitBtn.style.opacity = "0.7";
                    }

                    try {
                        const partyKey = type === 'sales' ? 'customer' : 'supplier';
                    const partyId = document.getElementById(`${type}-${partyKey}-id`).value;
                    if (!partyId) return alert(`Please select a ${partyKey}.`);

                    const items = [];
                    const rows = document.querySelectorAll(`#${type}-items-body tr`);
                    if (rows.length === 0) return alert("Please add at least one item.");

                    rows.forEach(tr => {
                        const qty = parseFloat(tr.querySelector('.row-qty').value) || 0;
                        if (qty <= 0) return; // FIX: Prevent saving empty "0 qty" items from Returns or accidental inputs

                        items.push({
                            itemId: tr.querySelector('.row-item-id').value,
                            name: tr.querySelector('.row-item-name').value,
                            hsn: tr.querySelector('.row-hsn').value,
                            qty: qty,
                            uom: tr.querySelector('.row-uom').value,
                            rate: parseFloat(tr.querySelector('.row-rate').value) || 0,
                            gstPercent: parseFloat(tr.querySelector('.row-gst').value) || 0,
                            buyPrice: parseFloat(tr.querySelector('.row-item-buyprice').value) || 0
                        });
                    });
                    
                    if (items.length === 0) {
                        if (submitBtn) { submitBtn.disabled = false; submitBtn.innerText = originalText; submitBtn.style.opacity = "1"; }
                        return alert("Please add at least one item with a quantity greater than 0.");
                    }

                    const isReturn = app.state.currentDocType === 'return';
                    const discTypeEl = document.getElementById(`${type}-discount-type`);
                    const accEl = document.getElementById(`${type}-account-id`);

                    // FIX: Prevent Duplicate Numbers, but safetly IGNORE blank numbers!
                    const proposedDocNo = type === 'sales' ? document.getElementById('sales-invoice-no').value : document.getElementById('purchase-po-no').value;
                    const allExistingDocs = await getAllRecords(type === 'sales' ? 'sales' : 'purchases');
                    
                    if (proposedDocNo && proposedDocNo.trim() !== '') {
                        const isDuplicate = allExistingDocs.some(d => 
                            d.firmId === app.state.firmId && 
                            (type === 'sales' ? true : d.supplierId === partyId) &&
                            (d.invoiceNo === proposedDocNo || d.poNo === proposedDocNo) && 
                            d.id !== app.state.currentEditId
                        );
                        if (isDuplicate) {
                            if (submitBtn) { submitBtn.disabled = false; submitBtn.innerText = originalText; submitBtn.style.opacity = "1"; }
                            return alert(`Error: Document number "${proposedDocNo}" already exists! Please use a unique number.`);
                        }
                    }

                    // Negative Inventory Warning (With Edit Recovery Math & Purchase Returns)
                    const isReducingStock = (type === 'sales' && !isReturn) || (type === 'purchase' && isReturn);
                    if (isReducingStock) {
                        const allItems = await getAllRecords('items');
                        let existingInvoice = null;
                        if (app.state.currentEditId) {
                            existingInvoice = await getRecordById(type === 'sales' ? 'sales' : 'purchases', app.state.currentEditId);
                        }
                        
                        for (const row of items) {
                            const dbItem = allItems.find(i => i.id === row.itemId);
                            if (dbItem) {
                                let effectiveStock = parseFloat(dbItem.stock) || 0;
                                // If editing an already-completed invoice, temporarily add the old qty back to the pool
                                if (existingInvoice && existingInvoice.status !== 'Open') {
                                    const oldItem = existingInvoice.items.find(i => i.itemId === row.itemId);
                                    if (oldItem) effectiveStock += (parseFloat(oldItem.qty) || 0);
                                }
                                
                                if (effectiveStock < parseFloat(row.qty)) {
                                    if (!confirm(`Warning: You are trying to deduct ${row.qty} of "${row.name}", but your effective stock is only ${effectiveStock}. This will cause negative inventory. Continue anyway?`)) {
                                        return; 
                                    }
                                }
                            }
                        }
                    }

                    const data = {
                        id: app.state.currentEditId || Utils.generateId(),
                        firmId: app.state.firmId,
                        documentType: app.state.currentDocType,
                        date: document.getElementById(`${type}-date`).value,
                        orderDate: document.getElementById(`${type}-order-date`).value,
                        shippedDate: document.getElementById(`${type}-shipped-date`).value,
                        completedDate: document.getElementById(`${type}-completed-date`).value,
                        
                        [type === 'sales' ? 'customerId' : 'supplierId']: partyId,
                        [type === 'sales' ? 'customerName' : 'supplierName']: document.getElementById(`${type}-${partyKey}-display`).innerText,
                        
                        invoiceNo: type === 'sales' ? document.getElementById('sales-invoice-no').value : (document.getElementById('purchase-po-no').value), 
                        poNo: type === 'purchase' ? document.getElementById('purchase-po-no').value : '',
                        orderNo: document.getElementById(`${type}-order-no`).value,
                        status: document.getElementById(`${type}-order-status`).value,
                        freightAmount: parseFloat(document.getElementById(`${type}-freight`).value) || 0,
                        invoiceType: document.getElementById(`${type}-invoice-type`) ? document.getElementById(`${type}-invoice-type`).value : 'B2B',
                        
                        items: items,
                        
                        subtotal: parseFloat(document.getElementById(`${type}-subtotal`).innerText.replace(/[^\d.-]/g, '')) || 0,
                        discount: parseFloat(document.getElementById(`${type}-discount`).value) || 0,
                        discountType: discTypeEl ? discTypeEl.value : '\u20B9',
                        totalGst: parseFloat(document.getElementById(`${type}-gst-total`).innerText.replace(/[^\d.-]/g, '')) || 0,
                        grandTotal: parseFloat(document.getElementById(`${type}-grand-total`).innerText.replace(/[^\d.-]/g, '')) || 0,
                        internalNotes: document.getElementById(`${type}-internal-notes`) ? document.getElementById(`${type}-internal-notes`).value : '',
                        
                        // NEW: Instant payments removed. Hardcoded to 0. All payments use manual receipts now.
                        amountPaid: 0,
                        paymentMode: 'Cash',
                        accountId: 'cash'
                    };

                        const storeName = type === 'sales' ? 'sales' : 'purchases';
                        await saveInvoiceTransaction(storeName, data);
                        
                        // NEW: Auto-Complete Advance Payments Engine
                        if (typeof app.autoCompleteInvoices === 'function') {
                            await app.autoCompleteInvoices(partyId, type);
                        }
                        
                        UI.showSuccess(); // UPGRADE: Trigger GPay Animation!
                        UI.closeActivity(`activity-${type}-form`);
                        app.refreshAll();
                    } catch (error) {
                        console.error("Save failed:", error);
                        alert("An error occurred while saving. Please try again.");
                    } finally {
                        // NEW: Always unlock the button, even if the save fails
                        if (submitBtn) {
                            submitBtn.disabled = false;
                            submitBtn.innerText = originalText; // FIX: Restores dynamic text (e.g. "Save Credit Note")
                            submitBtn.style.opacity = "1";
                        }
                    }
                });
            }
        });

        // ------------------ GENERAL CRUD FORMS ------------------
        ['product', 'ledger', 'expense', 'account'].forEach(type => {
            const form = document.getElementById(`form-${type}`);
            if (!form) return;

            form.addEventListener('submit', async (e) => {
                e.preventDefault();
                
                const submitBtn = form.querySelector('button[type="submit"]');
                const originalText = submitBtn ? submitBtn.innerText : 'Save';
                if (submitBtn) {
                    submitBtn.disabled = true;
                    submitBtn.innerText = "Saving...";
                    submitBtn.style.opacity = "0.7";
                }

                try {
                    const formData = new FormData(form);
                    
                    let storeName = `${type}s`;
                    if (type === 'product') storeName = 'items';
                    else if (type === 'account') storeName = 'accounts';

                    // Initialize data and merge with existing database record to prevent wiping hidden fields
                    let data = { id: app.state.currentEditId || Utils.generateId(), firmId: app.state.firmId };
                    
                    if (app.state.currentEditId) {
                        const existingRecord = await getRecordById(storeName, app.state.currentEditId);
                        if (existingRecord) data = { ...existingRecord };
                    }
                    
                    formData.forEach((value, key) => { data[key] = value; });
                
                if (type === 'product') {
                    data.sellPrice = parseFloat(data.sellPrice) || 0;
                    data.buyPrice = parseFloat(data.buyPrice) || 0;
                    data.stock = parseFloat(data.stock) || 0;
                    data.minStock = parseFloat(data.minStock) || 0;
                    data.gst = parseFloat(data.gst) || 0;
                    const img = document.getElementById('product-image-preview');
                    if (img && !img.classList.contains('hidden')) data.image = img.src;
                } 
                else if (type === 'ledger') {
                    data.openingBalance = parseFloat(data.openingBalance) || 0;
                } 
                else if (type === 'expense') {
                    data.amount = parseFloat(data.amount) || 0;
                    const accEl = document.getElementById('expense-account-id');
                    data.accountId = accEl ? accEl.value : 'cash';
                    const img = document.getElementById('expense-attachment-preview');
                    if (img && !img.classList.contains('hidden')) data.attachment = img.src;
                }
                else if (type === 'account') {
                    storeName = 'accounts';
                    data.openingBalance = parseFloat(data.openingBalance) || 0;
                }

                await saveRecord(storeName, data);

                // Automatically log Expenses as "Money Out" in the Cashbook
                if (type === 'expense') {
                    const expenseReceipt = {
                        id: 'exp-rec-' + data.id,
                        firmId: app.state.firmId,
                        date: data.date,
                        ledgerId: 'internal-expense',
                        ledgerName: 'Internal Expense',
                        type: 'out',
                        amount: data.amount,
                        mode: 'Expense',
                        accountId: data.accountId,
                        ref: '',
                        desc: `Expense: ${data.category}`,
                        isAutoGenerated: true,
                        invoiceRef: data.id 
                    };
                    await saveRecord('receipts', expenseReceipt);
                }
                
                UI.closeActivity(`activity-${type}-form`);
                app.refreshAll();
                
                } catch (error) {
                    console.error("Save failed:", error);
                    alert("An error occurred while saving. Please try again.");
                } finally {
                    if (submitBtn) {
                        submitBtn.disabled = false;
                        submitBtn.innerText = originalText;
                        submitBtn.style.opacity = "1";
                    }
                }
            });
        });

        // ------------------ STOCK ADJUSTMENT SUBMISSION ------------------
                const adjForm = document.getElementById('form-stock-adjustment');
        if (adjForm) {
            adjForm.addEventListener('submit', async (e) => {
                e.preventDefault();

                // FIXED: Button locking implemented to prevent duplicate stock adjustments
                const submitBtn = adjForm.querySelector('button[type="submit"]');
                const originalText = submitBtn ? submitBtn.innerText : 'Save Adjustment';
                if (submitBtn) {
                    submitBtn.disabled = true;
                    submitBtn.innerText = "Saving...";
                    submitBtn.style.opacity = "0.7";
                }

                try {
                    const itemId = document.getElementById('adj-product-id').value;
                    if (!itemId) throw new Error("Please select a product.");
                    
                    const type = document.getElementById('adj-type').value;
                    const qty = parseFloat(document.getElementById('adj-qty').value) || 0;
                    
                    const item = await getRecordById('items', itemId);
                    if (!item) throw new Error("Product not found in database.");
                    
                    const currentStock = parseFloat(item.stock) || 0;
                    
                    if (type === 'reduce') {
                        if (currentStock - qty < 0) {
                            if (!confirm(`Warning: This adjustment will drop your stock below zero (Current: ${currentStock}). Continue anyway?`)) {
                                return; 
                            }
                        }
                    }
                    
                    const adjData = {
                        id: Utils.generateId(),
                        firmId: app.state.firmId,
                        itemId: itemId,
                        type: type,
                        qty: qty,
                        date: document.getElementById('adj-date').value,
                        notes: document.getElementById('adj-notes').value
                    };
                    
                    item.stock += type === 'add' ? qty : -qty;
                    
                    await saveRecord('adjustments', adjData);
                    await saveRecord('items', item);
                    
                    alert("Stock adjusted successfully!");
                    UI.closeBottomSheet('sheet-stock-adjustment');
                    app.refreshAll();
                } catch (error) {
                    alert(error.message || "An error occurred while saving.");
                } finally {
                    if (submitBtn) {
                        submitBtn.disabled = false;
                        submitBtn.innerText = originalText;
                        submitBtn.style.opacity = "1";
                    }
                }
            });
        }

        // ------------------ BUSINESS PROFILE SUBMISSION ------------------
        const profileForm = document.getElementById('form-business-profile');
        if (profileForm) {
            profileForm.addEventListener('submit', async (e) => {
                e.preventDefault();
                const data = {
                    firmId: app.state.firmId,
                    name: document.getElementById('profile-name').value,
                    phone: document.getElementById('profile-phone').value,
                    email: document.getElementById('profile-email').value,
                    gst: document.getElementById('profile-gst').value,
                    address: document.getElementById('profile-address').value,
                    state: document.getElementById('profile-state').value,
                    bankDetails: document.getElementById('profile-bank').value,
                    terms: document.getElementById('profile-terms').value
                };

                const logoImg = document.getElementById('profile-logo-preview');
                if (logoImg && !logoImg.classList.contains('hidden')) data.logo = logoImg.src;

                const sigImg = document.getElementById('profile-signature-preview');
                if (sigImg && !sigImg.classList.contains('hidden')) data.signature = sigImg.src;

                await saveRecord('businessProfile', data);
                
                const firmRecord = await getRecordById('firms', app.state.firmId) || { id: app.state.firmId };
                firmRecord.name = data.name;
                firmRecord.phone = data.phone;
                firmRecord.email = data.email;
                firmRecord.gst = data.gst;
                firmRecord.address = data.address;
                firmRecord.state = data.state;
                await saveRecord('firms', firmRecord);

                alert("Business Profile Saved Successfully!");
                UI.closeActivity('activity-business-profile');
            });
        }

        // ------------------ MANUAL RECEIPT SUBMISSIONS ------------------
        ['in', 'out'].forEach(type => {
            const form = document.getElementById(`form-payment-${type}`);
            if (form) {
                form.addEventListener('submit', async (e) => {
                    e.preventDefault();

                    const submitBtn = form.querySelector('button[type="submit"]');
                    const originalText = submitBtn ? submitBtn.innerText : 'Save';
                    if (submitBtn) {
                        submitBtn.disabled = true;
                        submitBtn.innerText = "Processing...";
                        submitBtn.style.opacity = "0.7";
                    }

                    try {
                        const targetPartyId = type === 'in' ? document.getElementById('pay-in-customer').value : document.getElementById('pay-out-supplier').value;
                        if (!targetPartyId) {
                            if (submitBtn) { submitBtn.disabled = false; submitBtn.innerText = originalText; submitBtn.style.opacity = "1"; }
                            return alert("Please select a party.");
                        }

                    const ledger = await getRecordById('ledgers', targetPartyId);
                    const accountId = document.getElementById(`pay-${type}-account`).value;

                    // NEW: Prioritize the dropdown selection over the manual text input
                    const invoiceRefEl = document.getElementById(`pay-${type}-invoice-ref`);
                    // NEW: Capture multiple selected invoices and join them
                    const selectedOptions = invoiceRefEl ? Array.from(invoiceRefEl.selectedOptions).map(opt => opt.value).filter(v => v !== '') : [];
                    const selectedInvoiceRef = selectedOptions.join(', ');
                    const manualRef = document.getElementById(`pay-${type}-ref`).value;
                    const docNoInput = document.getElementById(`pay-${type}-no`).value;

                    // FIX: Duplicate Number Protection for Receipts
                    const allReceipts = await getAllRecords('receipts');
                    const isDuplicate = allReceipts.some(r => 
                        r.firmId === app.state.firmId && 
                        r.receiptNo === docNoInput && 
                        r.id !== app.state.currentReceiptId
                    );
                    if (isDuplicate) {
                        if (submitBtn) { submitBtn.disabled = false; submitBtn.innerText = originalText; submitBtn.style.opacity = "1"; }
                        return alert(`Error: Document number "${docNoInput}" already exists! Please use a unique number.`);
                    }

                    const data = {
                        id: app.state.currentReceiptId || Utils.generateId(),
                        receiptNo: docNoInput, // NEW: Save the custom document number
                        firmId: app.state.firmId,
                        date: document.getElementById(`pay-${type}-date`).value,
                        ledgerId: targetPartyId,
                        ledgerName: ledger ? ledger.name : 'Unknown',
                        type: type,
                        amount: parseFloat(document.getElementById(`pay-${type}-amount`).value) || 0,
                        mode: document.getElementById(`pay-${type}-mode`).value,
                        accountId: accountId,
                        ref: manualRef,
                        invoiceRef: selectedInvoiceRef, 
                        desc: document.getElementById(`pay-${type}-notes`).value || (type === 'in' ? 'Payment Received' : 'Payment Made'),
                        isAutoGenerated: false 
                    };

                        await saveRecord('receipts', data);
                        // NEW: Link manual payment to the saved invoices and mathematically update their statuses
                        if (selectedInvoiceRef) {
                            const storeName = type === 'in' ? 'sales' : 'purchases';
                            const allDocs = await getAllRecords(storeName);
                            const allReceipts = await getAllRecords('receipts');
                            const selectedRefs = selectedInvoiceRef.split(',').map(r => r.trim());

                            for (const ref of selectedRefs) {
                                const linkedInvoice = allDocs.find(doc => {
                                    if (doc.firmId !== app.state.firmId) return false;
                                    if ((type === 'in' ? doc.customerId : doc.supplierId) !== targetPartyId) return false;
                                    const docNo = type === 'in' ? doc.invoiceNo : (doc.poNo || doc.invoiceNo);
                                    const refToMatch = docNo || doc.id;
                                    return refToMatch === ref;
                                });

                                if (linkedInvoice && linkedInvoice.status !== 'Completed') {
                                    let totalPaid = 0;
                                    allReceipts.forEach(r => {
                                        if (r.ledgerId === targetPartyId && r.firmId === app.state.firmId) {
                                            const rRefs = String(r.invoiceRef || '').split(',').map(x => x.trim());
                                            if (rRefs.includes(ref)) {
                                                // Split the receipt amount mathematically
                                                totalPaid += (parseFloat(r.amount) || 0) / (rRefs.length || 1);
                                            }
                                        }
                                    });

                                    // If the split manual payments cover the grand total, mark as Completed
                                    if (totalPaid >= parseFloat(linkedInvoice.grandTotal) - 0.5) { 
                                        linkedInvoice.status = 'Completed';
                                        await saveRecord(storeName, linkedInvoice);
                                    }
                                }
                            }
                        }
                        
                        // NEW: Auto-Complete Advance Payments Engine
                        if (typeof app.autoCompleteInvoices === 'function') {
                            await app.autoCompleteInvoices(targetPartyId, type === 'in' ? 'sales' : 'purchases');
                        }

                        UI.showSuccess(); // UPGRADE: Trigger GPay Animation!
                        UI.closeBottomSheet(`sheet-payment-${type}`);
                        app.refreshAll();
                    } catch (error) {
                        console.error("Payment save failed:", error);
                        alert("An error occurred. Please try again.");
                    } finally {
                        if (submitBtn) {
                            submitBtn.disabled = false;
                            submitBtn.innerText = originalText;
                            submitBtn.style.opacity = "1";
                        }
                    }
                });
            }
        });
    },

      openNewPayment: async (type) => {
        app.state.currentReceiptId = null;
        const form = document.getElementById(`form-payment-${type}`);
        if (form) {
            form.reset();
            const dateInput = document.getElementById(`pay-${type}-date`);
            if (dateInput && typeof Utils !== 'undefined') {
                dateInput.value = Utils.getLocalDate();
                if (dateInput._flatpickr) dateInput._flatpickr.setDate(Utils.getLocalDate());
            }

            // NEW: Auto-Increment Engine specifically for Receipts & Vouchers
            const prefix = type === 'in' ? 'REC' : 'VOU';
            const receipts = await getAllRecords('receipts');
            const firmRecords = receipts.filter(r => r.firmId === app.state.firmId);
            let maxNum = 0;
            firmRecords.forEach(r => {
                const rNo = r.receiptNo || '';
                if (rNo.startsWith(prefix + '-')) {
                    const num = parseInt(rNo.replace(prefix + '-', ''), 10);
                    if (!isNaN(num) && num > maxNum) maxNum = num;
                }
            });
            const nextNo = `${prefix}-${String(maxNum + 1).padStart(4, '0')}`;
            const noInput = document.getElementById(`pay-${type}-no`);
            if (noInput) noInput.value = nextNo;
        }
        UI.toggleDeleteButton(`receipt-${type}`, false);
        UI.openBottomSheet(`sheet-payment-${type}`);
    },

    // ==========================================
    // CASHBOOK EDIT & DELETE ENGINE
    // ==========================================
    openReceipt: async (id, type) => {
        const record = await getRecordById('receipts', id);
        if (!record) return alert("Receipt not found.");

        // SAFEGUARD: Prevent tampering with invoice-linked auto-receipts
        if (record.isAutoGenerated) {
            return alert("This is an auto-generated receipt linked to a document. To modify or delete it, please edit the original invoice or bill.");
        }

        app.state.currentReceiptId = id;
        
        await UI.openBottomSheet(`sheet-payment-${type}`);
        
        setTimeout(async () => {
            // Restore the custom document number (Fallback to ID for older unpatched records)
            const noInput = document.getElementById(`pay-${type}-no`);
            if (noInput) noInput.value = record.receiptNo || String(record.id).substring(0, 12).toUpperCase();

            document.getElementById(`pay-${type}-date`).value = record.date || '';
            
            // UPGRADE: Hydrate the new custom Tap-to-Select UI display
            const partyInput = document.getElementById(type === 'in' ? 'pay-in-customer' : 'pay-out-supplier');
            const partyDisplay = document.getElementById(type === 'in' ? 'pay-in-customer-display' : 'pay-out-supplier-display');
            
            if (partyInput) partyInput.value = record.ledgerId || '';
            if (partyDisplay) {
                partyDisplay.innerText = record.ledgerName || 'Unknown Party';
                partyDisplay.style.color = 'var(--md-on-surface)';
            }
            
            if (partyInput) {
                // NEW: Load the pending invoices into the dropdown for the selected party
                await app.loadPendingInvoices(record.ledgerId, type);
                                const invoiceRefEl = document.getElementById(`pay-${type}-invoice-ref`);
                if (invoiceRefEl && record.invoiceRef) {
                    const savedRefs = String(record.invoiceRef).split(',').map(r => r.trim());
                    const store = type === 'in' ? 'sales' : 'purchases';
                    const allDocs = await getAllRecords(store);

                    savedRefs.forEach(ref => {
                        // Check if this specific ref is already in the dropdown options
                        let optionExists = Array.from(invoiceRefEl.options).some(opt => opt.value === ref);
                        
                        // If not, it means it's fully paid and hidden, so we must inject it back
                        if (!optionExists) {
                            let displayRef = ref;
                            // SELF-HEALING: Map perfectly to Order No and catch ghost IDs!
                            const linkedDoc = allDocs.find(d => d.id === ref || d.invoiceNo === ref || d.poNo === ref || d.orderNo === ref || d.id.endsWith(ref));
                            if (linkedDoc) {
                                if (type === 'in') {
                                    displayRef = linkedDoc.orderNo || linkedDoc.invoiceNo || String(linkedDoc.id).slice(-4).toUpperCase();
                                } else {
                                    displayRef = linkedDoc.orderNo || linkedDoc.poNo || linkedDoc.invoiceNo || String(linkedDoc.id).slice(-4).toUpperCase();
                                }
                            }
                            invoiceRefEl.innerHTML += `<option value="${ref}">${displayRef} (Fully Paid)</option>`;
                        }
                    });

                    // Loop through all options and natively select the ones that match our saved refs
                    Array.from(invoiceRefEl.options).forEach(opt => {
                        if (savedRefs.includes(opt.value)) {
                            opt.selected = true;
                        }
                    });
                }
            }
            
            document.getElementById(`pay-${type}-amount`).value = record.amount || 0;
            document.getElementById(`pay-${type}-mode`).value = record.mode || 'Cash';
            
            const accEl = document.getElementById(`pay-${type}-account`);
            if (accEl) accEl.value = record.accountId || 'cash';
            
            document.getElementById(`pay-${type}-ref`).value = record.ref || '';
            document.getElementById(`pay-${type}-notes`).value = record.desc || '';
            
            UI.toggleDeleteButton(`receipt-${type}`, true);
        }, 50);
    },

    // ==========================================
    // PO TO INVOICE CONVERTER
    // ==========================================
    convertPO: async (id) => {
        if (!confirm("Convert this Draft PO into a Completed Purchase Bill? This will officially add the items to your inventory and update your payable ledger.")) return;
        
        const record = await getRecordById('purchases', id);
        if (!record) return alert("Record not found.");

        record.status = 'Completed';
        const today = typeof Utils !== 'undefined' && Utils.getLocalDate ? Utils.getLocalDate() : new Date().toISOString().split('T')[0];
        record.completedDate = today;

        await saveInvoiceTransaction('purchases', record);
        
        alert("PO successfully converted to Purchase Bill!");
        app.refreshAll();
    },

    // ==========================================
    // AUTO-COMPLETE ADVANCE PAYMENT ENGINE
    // ==========================================
    autoCompleteInvoices: async (partyId, type) => {
        const isSales = (type === 'sales');
        const storeName = isSales ? 'sales' : 'purchases';
        const partyKey = isSales ? 'customerId' : 'supplierId';
        
        const allDocs = await getAllRecords(storeName);
        const allReceipts = await getAllRecords('receipts');
        
        // 1. Calculate total money received/paid for this party (including Advances)
        let totalPaid = 0;
        allReceipts.forEach(r => {
            if (r.firmId === app.state.firmId && r.ledgerId === partyId) {
                const amt = parseFloat(r.amount) || 0;
                if (isSales) totalPaid += (r.type === 'in' ? amt : -amt);
                else totalPaid += (r.type === 'out' ? amt : -amt);
            }
        });
        
        // 2. Sort documents chronologically (oldest first)
        const partyDocs = allDocs.filter(d => d.firmId === app.state.firmId && d[partyKey] === partyId && d.documentType !== 'return');
        partyDocs.sort((a, b) => new Date(a.date) - new Date(b.date));
        
        // 3. Smart Allocation - Prioritize explicit payments first, then use FIFO for advances
        let remainingMoney = totalPaid;
        
        for (const doc of partyDocs) {
            const docTotal = parseFloat(doc.grandTotal) || 0;
            
            // Calculate if THIS specific invoice was explicitly paid via manual receipt
            const uniqueRefs = [...new Set([doc.orderNo, doc.invoiceNo, doc.poNo, doc.id].filter(Boolean))];
            let explicitPaid = 0;
            allReceipts.forEach(r => {
                if (r.firmId === app.state.firmId && r.ledgerId === partyId) {
                    const rRefs = String(r.invoiceRef || '').split(',').map(x => x.trim());
                    if (rRefs.some(ref => uniqueRefs.includes(ref))) {
                        explicitPaid += (parseFloat(r.amount) || 0) / (rRefs.length || 1);
                    }
                }
            });
            
            // Mark completed if explicit payments cover it, OR if leftover FIFO advance covers it
            if (explicitPaid >= docTotal - 0.5 || remainingMoney >= docTotal - 0.5) { 
                remainingMoney -= docTotal; // Consume from global pool
                
                if (doc.status === 'Open') {
                    doc.status = 'Completed';
                    if (typeof Utils !== 'undefined' && Utils.getLocalDate) {
                        doc.completedDate = doc.completedDate || Utils.getLocalDate();
                    }
                    await saveRecord(storeName, doc); // Auto-save!
                }
            } else {
                remainingMoney -= Math.min(remainingMoney, docTotal);
            }
        }
    },

    // ==========================================
    // 7. ORPHAN PROTECTION & DELETION ENGINE
    // ==========================================
    deleteRecord: async (type) => {
        const id = type === 'receipt-in' || type === 'receipt-out' ? app.state.currentReceiptId : app.state.currentEditId;

        let storeMap = {
            'sales': 'sales', 'purchase': 'purchases',
            'product': 'items', 'ledger': 'ledgers',
            'expense': 'expenses', 'account': 'accounts',
            'receipt-in': 'receipts', 'receipt-out': 'receipts'
        };
        const storeName = storeMap[type];
        
        const record = await getRecordById(storeName, id);
        if(!record) return;

        let warningMsg = "Are you sure you want to delete this record?";
        if (type === 'sales' || type === 'purchase') {
            warningMsg = "Deleting this document will return all inventory to its previous state. Continue?";
        }

        if (!confirm(warningMsg)) return;

        // ORPHAN PROTECTION: Check for linked manual receipts safely
        if (type === 'sales' || type === 'purchase') {
            // FIX: Safely capture ALL possible linked IDs to prevent orphaned receipts
            const uniqueRefs = [...new Set([record.orderNo, record.invoiceNo, record.poNo, record.id].filter(Boolean))];
            const partyId = type === 'sales' ? record.customerId : record.supplierId;
            
            if (uniqueRefs.length > 0 && partyId) {
                const receipts = await getAllRecords('receipts');
                // Check if any of the manual receipt refs match ANY of the document's IDs
                const linkedManualReceipts = receipts.filter(r => {
                    const refs = String(r.invoiceRef || '').split(',').map(x => x.trim());
                    return refs.some(ref => uniqueRefs.includes(ref)) && r.ledgerId === partyId && r.isAutoGenerated === false;
                });
                if (linkedManualReceipts.length > 0) {
                    if (!confirm(`Warning: This document has ${linkedManualReceipts.length} manual payment(s) linked to it. Delete them as well to keep the cashbook balanced?`)) {
                        return; // Abort the whole deletion if they cancel here to prevent imbalance
                    }
                    for (const r of linkedManualReceipts) {
                        await deleteRecordById('receipts', r.id);
                    }
                }
            }
        }

        // Clean up Cashbook if an Expense is deleted
        if (type === 'expense') {
            await deleteRecordById('receipts', 'exp-rec-' + id);
        }

        // Prevent Bank Account Deletion if transactions are tied to it
        if (type === 'account') {
            const allReceipts = await getAllRecords('receipts');
            const linkedTransactions = allReceipts.filter(r => r.accountId === id);
            if (linkedTransactions.length > 0) {
                return alert(`Cannot delete this account. It has ${linkedTransactions.length} transaction(s) linked to it in the Cashbook.`);
            }
        }

        // FIX: Re-open invoices if their linked manual payment is deleted!
        if ((type === 'receipt-in' || type === 'receipt-out') && record.invoiceRef && !record.isAutoGenerated) {
            const storeToUpdate = type === 'receipt-in' ? 'sales' : 'purchases';
            const linkedDocs = await getAllRecords(storeToUpdate);
            const allReceipts = await getAllRecords('receipts');
            const refsToCheck = String(record.invoiceRef).split(',').map(r => r.trim());

            for (const ref of refsToCheck) {
                // Find the linked invoice
                const invoice = linkedDocs.find(d => d.id === ref || d.invoiceNo === ref || d.poNo === ref || d.orderNo === ref);
                if (invoice && invoice.status === 'Completed') {
                    // Calculate how much has been paid EXCLUDING the receipt we are about to delete
                    let totalPaidNow = 0;
                    allReceipts.forEach(r => {
                        if (r.id !== record.id && r.ledgerId === record.ledgerId) {
                            const rRefs = String(r.invoiceRef || '').split(',').map(x => x.trim());
                            if (rRefs.includes(ref)) totalPaidNow += (parseFloat(r.amount) || 0) / (rRefs.length || 1);
                        }
                    });
                    
                    // If the remaining payments don't cover the total, revert it to Open!
                    if (totalPaidNow < (parseFloat(invoice.grandTotal) || 0) - 0.5) {
                        invoice.status = 'Open';
                        delete invoice.completedDate;
                        await saveRecord(storeToUpdate, invoice);
                    }
                }
            }
        }
        record.deletedAt = new Date().toLocaleString();
        
        // FIX: Delete heavy base64 images before saving to LocalStorage to prevent 5MB Quota crashes!
        if (record.image) delete record.image;
        if (record.attachment) delete record.attachment;
        
        const trashBin = JSON.parse(localStorage.getItem('sollo_trash') || '[]');
        trashBin.push(record);
        localStorage.setItem('sollo_trash', JSON.stringify(trashBin));

        // Delete from main database
        await deleteRecordById(storeName, id);

        if (type === 'sales' || type === 'purchase') UI.closeActivity(`activity-${type}-form`);
        else if (type === 'receipt-in' || type === 'receipt-out') UI.closeBottomSheet(`sheet-payment-${type.split('-')[1]}`);
        else UI.closeActivity(`activity-${type}-form`);
        
        app.refreshAll();
    },

    // ==========================================
    // UPGRADE: RECYCLE BIN RESTORE ENGINE
    // ==========================================
    restoreRecord: async (id, storeName) => {
        if (!confirm("Are you sure you want to restore this item?")) return;

        let trashBin = JSON.parse(localStorage.getItem('sollo_trash') || '[]');
        const recordIndex = trashBin.findIndex(t => t.id === id);
        
        if (recordIndex > -1) {
            const record = trashBin[recordIndex];
            
            // Remove the trash tags
            delete record._module;
            delete record.deletedAt;
            
            // Save it back to the active IndexedDB
            await saveRecord(storeName, record); 
            
            // FIX: Recreate the Cashbook entry if restoring an Expense
            if (storeName === 'expenses') {
                const expenseReceipt = {
                    id: 'exp-rec-' + record.id,
                    firmId: record.firmId,
                    date: record.date,
                    ledgerId: 'internal-expense',
                    ledgerName: 'Internal Expense',
                    type: 'out',
                    amount: record.amount,
                    mode: 'Expense',
                    accountId: record.accountId || 'cash',
                    ref: '',
                    desc: `Expense: ${record.category}`,
                    isAutoGenerated: true,
                    invoiceRef: record.id 
                };
                await saveRecord('receipts', expenseReceipt);
            }
            
            // Remove it from the Trash Vault
            trashBin.splice(recordIndex, 1);
            localStorage.setItem('sollo_trash', JSON.stringify(trashBin));
            
            if (window.Utils) window.Utils.showToast("Record Restored Successfully! ♻️");
            app.refreshAll();
        } else {
            alert("Error: Could not find the record in the Recycle Bin.");
        }
    },

    // ==========================================
    // 8. PDF EXPORT HOOKS
    // ==========================================
        generatePDF: async (type) => {
        if (!app.state.currentEditId) return alert("Please save the document first.");
        const storeName = type === 'sales' ? 'sales' : 'purchases';
        const record = await getRecordById(storeName, app.state.currentEditId);
        if (!record) return;

        // NEW: Calculate true total paid (Instant + Manual) to link it to the PDF
        const receipts = await getAllRecords('receipts');
        // FIX: Match against ALL references so the printed PDF never misses a cross-linked payment!
        const uniqueRefs = [...new Set([record.orderNo, record.invoiceNo, record.poNo, record.id].filter(Boolean))];
        const partyId = type === 'sales' ? record.customerId : record.supplierId;
        
        let totalPaid = 0;
        let linkedReceipts = []; // NEW: Array to hold the detailed receipts
        
        receipts.forEach(r => {
            if (r.ledgerId === partyId && r.firmId === app.state.firmId) {
                const refs = String(r.invoiceRef || '').split(',').map(x => x.trim());
                if (refs.some(ref => uniqueRefs.includes(ref))) {
                    // Mathematically split the payment value so the PDF balance is perfectly accurate
                    const splitAmt = (parseFloat(r.amount) || 0) / (refs.length || 1);
                    totalPaid += splitAmt;
                    
                    // Clone the receipt so we can safely alter the displayed amount on the PDF
                    const clonedReceipt = { ...r, amount: splitAmt };
                    linkedReceipts.push(clonedReceipt);
                }
            }
        });
        
        // Inject the total paid AND the receipt details into the record object
        record.trueTotalPaid = totalPaid;
        record.linkedReceipts = linkedReceipts;

        const profile = await getRecordById('businessProfile', app.state.firmId);
        const party = await getRecordById('ledgers', type === 'sales' ? record.customerId : record.supplierId);

        Utils.generateInvoicePDF(record, profile || {}, party || {}, type);
    },

    // ==========================================
    // 9. GST DASHBOARD CONTROLLER
    // ==========================================
    openGSTReport: async () => {
        // Default to the current month
        const dateInput = document.getElementById('gst-month-selector');
        if (dateInput && !dateInput.value) {
            const today = new Date();
            const yyyy = today.getFullYear();
            const mm = String(today.getMonth() + 1).padStart(2, '0');
            dateInput.value = `${yyyy}-${mm}`;
        }
        
        await app.refreshGSTReport();
        if (typeof UI !== 'undefined') UI.openActivity('activity-gst-report');
    },

    refreshGSTReport: async () => {
        const monthVal = document.getElementById('gst-month-selector').value;
        if (!monthVal) return;
        
        if (typeof generateGSTReport !== 'function') {
            return alert("GST engine not found in db.js! Please make sure you saved the db.js file updates.");
        }
        
        // 1. Run the heavy math engine we built in db.js
        const report = await generateGSTReport(monthVal, app.state.firmId);
        
        // 2. Update GSTR-3B UI Cards
        document.getElementById('gst-out-tax').innerText = `\u20B9${report.gstr3b.outputTax.toFixed(2)}`;
        document.getElementById('gst-in-tax').innerText = `\u20B9${report.gstr3b.inputTax.toFixed(2)}`;
        document.getElementById('gst-net-payable').innerText = `\u20B9${report.gstr3b.netPayable.toFixed(2)}`;
        
        // 3. Update GSTR-1 UI Cards
        document.getElementById('gst-b2b-taxable').innerText = `\u20B9${report.gstr1.b2bTaxable.toFixed(2)}`;
        document.getElementById('gst-b2b-tax').innerText = `\u20B9${report.gstr1.b2bTax.toFixed(2)}`;
        document.getElementById('gst-b2c-taxable').innerText = `\u20B9${report.gstr1.b2cTaxable.toFixed(2)}`;
        document.getElementById('gst-b2c-tax').innerText = `\u20B9${report.gstr1.b2cTax.toFixed(2)}`;
        
        // 4. Bind the Export CSV Button
        const exportBtn = document.getElementById('btn-export-gst');
        if (exportBtn) {
            exportBtn.onclick = () => { Utils.exportGSTCSV(report); };
        }
    }, // <--- ADDED THE MISSING COMMA HERE!

    // ==========================================
    // NEW: EXPENSE LINK SEARCH ENGINE
    // ==========================================
    loadLinkedDocsList: async () => {
        const listEl = document.getElementById('list-linked-docs');
        if (!listEl) return;
        
        const sales = await getAllRecords('sales');
        const purchases = await getAllRecords('purchases');
        
        const currentSelected = (document.getElementById('expense-linked-invoice').value || '').split(',');
        let html = '';
        
        const recentSales = sales.filter(s => s.firmId === app.state.firmId && s.status !== 'Open').slice(-50).reverse();
        if (recentSales.length > 0) {
            html += `<li style="background: var(--md-surface-variant); font-weight: bold; pointer-events: none; padding: 8px 16px; border-radius: 4px;">Sales Invoices</li>`;
            recentSales.forEach(s => {
                // BULLETPROOF: Save the exact Order/Invoice number directly!
                const docNo = s.orderNo || s.invoiceNo || s.id;
                const displayNo = s.orderNo || s.invoiceNo || s.id.slice(-4).toUpperCase();
                const isSelected = currentSelected.includes(docNo);
                const bg = isSelected ? 'var(--md-surface-variant)' : 'transparent';
                const checked = isSelected ? 'checked' : '';

                html += `<li style="background: ${bg};" onclick="app.selectLinkedDoc('${docNo}', '${displayNo}', this)">
                    <div><strong style="color:var(--md-primary);">${displayNo}</strong><br><small>${s.customerName}</small></div>
                    <input type="checkbox" ${checked} style="width: 20px; height: 20px; pointer-events: none;">
                </li>`;
            });
        }
        
        const recentPurch = purchases.filter(p => p.firmId === app.state.firmId && p.status !== 'Open').slice(-50).reverse();
        if (recentPurch.length > 0) {
            html += `<li style="background: var(--md-surface-variant); font-weight: bold; pointer-events: none; padding: 8px 16px; border-radius: 4px; margin-top: 12px;">Purchase Bills</li>`;
            recentPurch.forEach(p => {
                // BULLETPROOF: Save the exact PO/Order number directly!
                const docNo = p.orderNo || p.poNo || p.invoiceNo || p.id;
                const displayNo = p.orderNo || p.poNo || p.invoiceNo || p.id.slice(-4).toUpperCase();
                
                const isSelected = currentSelected.includes(docNo);
                const bg = isSelected ? 'var(--md-surface-variant)' : 'transparent';
                const checked = isSelected ? 'checked' : '';

                html += `<li style="background: ${bg};" onclick="app.selectLinkedDoc('${docNo}', '${displayNo}', this)">
                    <div><strong style="color:var(--md-error);">${displayNo}</strong><br><small>${p.supplierName}</small></div>
                    <input type="checkbox" ${checked} style="width: 20px; height: 20px; pointer-events: none;">
                </li>`;
            });
        }
        listEl.innerHTML = html;
    },

    selectLinkedDoc: (id, displayName, liElement = null) => {
        const inputEl = document.getElementById('expense-linked-invoice');
        const displayEl = document.getElementById('expense-linked-display');

        if (!id) {
            inputEl.value = '';
            displayEl.innerText = '-- No Link (General Expense) --';
            displayEl.style.color = 'var(--md-text-muted)';
            UI.closeBottomSheet('sheet-linked-docs');
            document.querySelectorAll('#list-linked-docs li input[type="checkbox"]').forEach(cb => cb.checked = false);
            document.querySelectorAll('#list-linked-docs li').forEach(li => {
                if(li.style.pointerEvents !== 'none') li.style.background = 'transparent';
            });
            return;
        }

        let currentIds = inputEl.value ? inputEl.value.split(',') : [];
        let currentNames = displayEl.innerText !== '-- No Link (General Expense) --' ? displayEl.innerText.split(' | ') : [];

        if (currentIds.includes(id)) {
            currentIds = currentIds.filter(i => i !== id);
            currentNames = currentNames.filter(n => n !== displayName);
            if (liElement) {
                liElement.style.background = 'transparent';
                const cb = liElement.querySelector('input');
                if (cb) cb.checked = false;
            }
        } else {
            currentIds.push(id);
            currentNames.push(displayName);
            if (liElement) {
                liElement.style.background = 'var(--md-surface-variant)';
                const cb = liElement.querySelector('input');
                if (cb) cb.checked = true;
            }
        }

        inputEl.value = currentIds.join(',');
        displayEl.innerText = currentNames.length > 0 ? currentNames.join(' | ') : '-- No Link (General Expense) --';
        displayEl.style.color = currentIds.length > 0 ? 'var(--md-primary)' : 'var(--md-text-muted)';
    },

    fetchPincode: async (pincode, prefix) => {
        if (!pincode || pincode.toString().length !== 6) return;
        
        try {
            if (window.Utils) window.Utils.showToast("Fetching Location...");
            
            const response = await fetch(`https://api.postalpincode.in/pincode/${pincode}`);
            const data = await response.json();
            
            if (data && data[0] && data[0].Status === 'Success') {
                const postOffice = data[0].PostOffice[0];
                
                const cityEl = document.getElementById(`${prefix}-city`);
                const stateEl = document.getElementById(`${prefix}-state`);
                
                // Set the values, but they remain editable by the user!
                if (cityEl) cityEl.value = postOffice.District;
                if (stateEl) stateEl.value = postOffice.State;
                
                if (window.Utils) window.Utils.showToast(`📍 Found: ${postOffice.District}, ${postOffice.State}`);
            } else {
                if (window.Utils) window.Utils.showToast("⚠️ Invalid Pincode");
            }
        } catch (error) {
            console.error("Pincode API Error:", error);
            if (window.Utils) window.Utils.showToast("⚠️ Could not fetch location");
        }
    },

    filterLinkedDocs: (term) => {
        const listItems = document.querySelectorAll('#list-linked-docs li');
        const lowerTerm = term.toLowerCase();
        listItems.forEach(li => {
            if(li.style.pointerEvents === 'none') return; // Skip headers
            const text = li.innerText.toLowerCase();
            li.style.display = text.includes(lowerTerm) ? 'flex' : 'none';
        });
    },

}; // <--- THIS IS THE CRITICAL CLOSING BRACKET FOR THE APP OBJECT

// --- NEW CODE: Module Initialization ---
window.onload = app.init;
window.app = app; // Explicitly map to window to protect your HTML buttons

// BRIDGE FOR CLOUD.JS: Expose the database engine so Google Drive can access it
window.exportDatabase = exportDatabase;
window.importDatabase = importDatabase;
// --- END OF NEW CODE ---
