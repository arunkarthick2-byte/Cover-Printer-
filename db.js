// ==========================================
// SOLLO ERP - DATABASE ENGINE (v5.2 Enterprise)
// ==========================================
const DB_NAME = 'SOLLO_ERP_DB';
const DB_VERSION = 9; // Upgraded version for Units & Categories
let db;

const initDB = () => {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(DB_NAME, DB_VERSION);

        request.onupgradeneeded = (event) => {
            db = event.target.result;
            
            // Core Settings Stores
            if (!db.objectStoreNames.contains('firms')) db.createObjectStore('firms', { keyPath: 'id' });
            if (!db.objectStoreNames.contains('businessProfile')) db.createObjectStore('businessProfile', { keyPath: 'firmId' });
            if (!db.objectStoreNames.contains('counters')) db.createObjectStore('counters', { keyPath: 'id' });
            
            // Master Stores
            if (!db.objectStoreNames.contains('items')) {
                let s = db.createObjectStore('items', { keyPath: 'id' });
                s.createIndex('firmId', 'firmId', { unique: false });
            }
            if (!db.objectStoreNames.contains('ledgers')) {
                let s = db.createObjectStore('ledgers', { keyPath: 'id' });
                s.createIndex('firmId', 'firmId', { unique: false });
            }
            if (!db.objectStoreNames.contains('accounts')) {
                let s = db.createObjectStore('accounts', { keyPath: 'id' });
                s.createIndex('firmId', 'firmId', { unique: false });
            }
            
            // Transaction Stores
            if (!db.objectStoreNames.contains('sales')) {
                let s = db.createObjectStore('sales', { keyPath: 'id' });
                s.createIndex('firmId', 'firmId', { unique: false });
            }
            if (!db.objectStoreNames.contains('purchases')) {
                let s = db.createObjectStore('purchases', { keyPath: 'id' });
                s.createIndex('firmId', 'firmId', { unique: false });
            }
            if (!db.objectStoreNames.contains('receipts')) {
                let s = db.createObjectStore('receipts', { keyPath: 'id' });
                s.createIndex('firmId', 'firmId', { unique: false });
            }
            if (!db.objectStoreNames.contains('expenses')) {
                let s = db.createObjectStore('expenses', { keyPath: 'id' });
                s.createIndex('firmId', 'firmId', { unique: false });
            }
            if (!db.objectStoreNames.contains('adjustments')) {
                let s = db.createObjectStore('adjustments', { keyPath: 'id' });
                s.createIndex('firmId', 'firmId', { unique: false });
                s.createIndex('itemId', 'itemId', { unique: false });
            }
            // NEW: Simple Master Stores
            if (!db.objectStoreNames.contains('units')) db.createObjectStore('units', { keyPath: 'id' });
            if (!db.objectStoreNames.contains('expenseCategories')) db.createObjectStore('expenseCategories', { keyPath: 'id' });
        };

        request.onsuccess = (event) => { db = event.target.result; resolve(); };
        request.onerror = (event) => { console.error("IndexedDB error:", event.target.errorCode); reject(event.target.error); };
    });
};

// ==========================================
// STANDARD CRUD OPERATIONS
// ==========================================
const getAllRecords = (storeName) => {
    return new Promise((resolve, reject) => {
        const transaction = db.transaction(storeName, 'readonly');
        const store = transaction.objectStore(storeName);
        const request = store.getAll();
        request.onsuccess = () => resolve(request.result || []);
        request.onerror = () => reject(request.error);
    });
};

const getRecordById = (storeName, id) => {
    return new Promise((resolve, reject) => {
        const transaction = db.transaction(storeName, 'readonly');
        const store = transaction.objectStore(storeName);
        const request = store.get(id);
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
};

const saveRecord = (storeName, data) => {
    return new Promise((resolve, reject) => {
        const transaction = db.transaction(storeName, 'readwrite');
        const store = transaction.objectStore(storeName);
        const request = store.put(data);
        request.onsuccess = () => resolve(data.id || data.firmId);
        request.onerror = () => reject(request.error);
    });
};

const deleteRecordById = (storeName, id) => {
    return new Promise(async (resolve, reject) => {
        try {
            // 1. Fetch the record first to identify linked impacts
            const oldRecord = await getRecordById(storeName, id);
            
            if (oldRecord && (storeName === 'sales' || storeName === 'purchases')) {
                // 2. CRITICAL: Await the stock reversal COMPLETELY before proceeding
                await reverseStockImpact(storeName, oldRecord);
                
                // 3. Batched cleanup of auto-generated receipts (Optimized & Collision Protected)
                // FIX: Gather ALL possible IDs to safely catch and delete orphaned auto-receipts
                const uniqueRefs = [...new Set([oldRecord.orderNo, oldRecord.invoiceNo, oldRecord.poNo, oldRecord.id].filter(Boolean))];
                // Grab the party ID based on whether it's a sale or purchase
                const partyId = storeName === 'sales' ? oldRecord.customerId : oldRecord.supplierId;
                
                if (uniqueRefs.length > 0 && partyId) {
                    const receipts = await getAllRecords('receipts');
                    // FIX: Check if the receipt's invoiceRef matches ANY of the document's unique IDs
                    const receiptsToDelete = receipts.filter(r => uniqueRefs.includes(r.invoiceRef) && r.ledgerId === partyId && r.isAutoGenerated && r.firmId === oldRecord.firmId);

                    if (receiptsToDelete.length > 0) {
                        await new Promise((resolveBatch, rejectBatch) => {
                            const t = db.transaction('receipts', 'readwrite');
                            const store = t.objectStore('receipts');
                            
                            // Delete all matched receipts within one single transaction
                            receiptsToDelete.forEach(r => store.delete(r.id));
                            
                            t.oncomplete = () => resolveBatch();
                            t.onerror = () => rejectBatch(t.error);
                        });
                    }
                }
            }

            // 4. Finally, delete the actual document only after all reversals finish
            const transaction = db.transaction(storeName, 'readwrite');
            const request = transaction.objectStore(storeName).delete(id);
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        } catch (err) {
            console.error("Deletion engine failure:", err);
            reject(err);
        }
    });
};

const getAllFirms = () => getAllRecords('firms');

// ==========================================
// STRICT INVENTORY & INVOICE ENGINE
// ==========================================
const reverseStockImpact = async (storeName, record) => {
    if (record.status === 'Open') return; // Drafts do not reverse Stock
    const isReturn = record.documentType === 'return';
    const items = await getAllRecords('items');
    
    // CRITICAL FIX: Fallback to empty array if record items are missing (corrupted data)
    for (const row of (record.items || [])) {
        const dbItem = items.find(i => i.id === row.itemId);
        if (dbItem) {
            let qty = parseFloat(row.qty) || 0;
            // Reverse the math
            if (storeName === 'sales') {
                dbItem.stock = Math.round((dbItem.stock + (isReturn ? -qty : qty)) * 100) / 100;
            } else if (storeName === 'purchases') {
                dbItem.stock = Math.round((dbItem.stock + (isReturn ? qty : -qty)) * 100) / 100;
            }
            await saveRecord('items', dbItem);
        }
    }
};

const applyStockImpact = async (storeName, record) => {
    if (record.status === 'Open') return; // Drafts do not apply Stock
    const isReturn = record.documentType === 'return';
    const items = await getAllRecords('items');
    
    for (const row of (record.items || [])) {
        const dbItem = items.find(i => i.id === row.itemId);
        if (dbItem) {
            let qty = parseFloat(row.qty) || 0;
            // Apply the math
            if (storeName === 'sales') {
                dbItem.stock = Math.round((dbItem.stock + (isReturn ? qty : -qty)) * 100) / 100; 
            } else if (storeName === 'purchases') {
                dbItem.stock = Math.round((dbItem.stock + (isReturn ? -qty : qty)) * 100) / 100; 
                
                // Automatically update buy price to the latest rate (factoring in discounts!)
                if (!isReturn && row.rate > 0) {
                    let discountRatio = 0;
                    if (record.discount > 0 && record.subtotal > 0) {
                        discountRatio = record.discountType === '%' ? (record.discount / 100) : (record.discount / record.subtotal);
                    }
                    dbItem.buyPrice = row.rate * (1 - discountRatio);
                }
            }
            await saveRecord('items', dbItem);
        }
    }
};

const saveInvoiceTransaction = async (storeName, data) => {
    // 1. If Editing, Reverse previous stock impacts first
    const existingRecord = await getRecordById(storeName, data.id);
    if (existingRecord) {
        await reverseStockImpact(storeName, existingRecord);
        
        // Delete old auto-receipts to recreate them fresh (Collision Protected)
        const docNo = existingRecord.invoiceNo || existingRecord.poNo || existingRecord.id;
        const partyId = storeName === 'sales' ? existingRecord.customerId : existingRecord.supplierId;
        
        const newDocNo = data.invoiceNo || data.poNo || data.id; // The corrected invoice number
        if (docNo && partyId) {
            const receipts = await getAllRecords('receipts');
            for (const r of receipts) {
                // FIX: Convert legacy auto-receipts to manual so they aren't permanently lost!
                if (r.invoiceRef === docNo && r.ledgerId === partyId && r.isAutoGenerated) {
                    r.isAutoGenerated = false;
                    await saveRecord('receipts', r);
                }
                // FIX: Migrate Manual Receipts to the new invoice number if it was changed
                else if (r.ledgerId === partyId && !r.isAutoGenerated && docNo !== newDocNo) {
    const refs = String(r.invoiceRef || '').split(',').map(x => x.trim());
    if (refs.includes(docNo)) {
        const newRefs = refs.map(ref => ref === docNo ? newDocNo : ref);
        r.invoiceRef = newRefs.join(', ');
        if (r.desc && r.desc.includes(docNo)) {
            r.desc = r.desc.replace(docNo, newDocNo); 
        }
        await saveRecord('receipts', r);
    }
}
            }
        }
    }

    // 2. Apply New Stock Impacts
    await applyStockImpact(storeName, data);

    // 3. Save the actual Invoice/PO document
    await saveRecord(storeName, data);
    
    // INSTANT PAYMENT CREATION REMOVED! All payments must now be done via Payment In/Out forms.
};

// ==========================================
// SMART AUTO-NUMBERING ENGINE
// ==========================================
// NEW: Added targetField parameter to allow Order Ref incrementing
const getNextDocumentNumber = async (storeName, prefix, targetField = null) => {
    // NEW: Removed hardcoding so it can accept 'expenses' as a valid database store
    const records = await getAllRecords(storeName);
    const firmId = typeof app !== 'undefined' && app.state ? app.state.firmId : 'firm1';
    const firmRecords = records.filter(r => r.firmId === firmId);
    
    let maxNumber = 0;

    firmRecords.forEach(record => {
        let docNo = '';
        if (targetField) {
            docNo = record[targetField] || '';
        } else {
            // FIX: Changed 'type' to 'storeName' so the engine knows what to search for
            docNo = (storeName === 'sales' ? record.invoiceNo : (record.poNo || record.invoiceNo)) || ''; 
        }
        
        if (docNo.startsWith(prefix + '-')) {
            const numPart = docNo.replace(prefix + '-', '');
            const parsedNum = parseInt(numPart, 10);
            
            if (!isNaN(parsedNum) && parsedNum > maxNumber) {
                maxNumber = parsedNum;
            }
        }
    });

    const nextNumber = maxNumber + 1;
    const paddedNumber = String(nextNumber).padStart(4, '0');
    return `${prefix}-${paddedNumber}`;
};

// ==========================================
// LEDGER / KHATA ENGINE
// ==========================================
const getKhataStatement = async (partyId, partyType) => {
    const ledger = await getRecordById('ledgers', partyId);
    const firmId = typeof app !== 'undefined' && app.state ? app.state.firmId : 'firm1';
    
    let openingBalance = parseFloat(ledger ? ledger.openingBalance : 0) || 0;
    
    // FIX: Bulletproof case-insensitive string parsing
    const isCustomer = String(partyType).toLowerCase() === 'customer';
    const balType = (ledger && ledger.balanceType) ? ledger.balanceType.toLowerCase() : '';
    
    // Mathematical signs for ledger: Positive means "We are owed (Receive)", Negative means "We owe them (Pay)"
    if (isCustomer) {
        openingBalance = (balType.includes('pay') || balType.includes('credit')) ? -openingBalance : openingBalance;
    } else {
        openingBalance = (balType.includes('receive') || balType.includes('debit')) ? -openingBalance : openingBalance;
    }

    let timeline = [];
    
    if (openingBalance !== 0) {
        timeline.push({
            id: 'open-bal',
            date: 'Opening',
            desc: 'Opening Balance',
            amount: Math.abs(openingBalance),
            isInvoice: openingBalance > 0, 
            impact: openingBalance
        });
    }

    const sales = (await getAllRecords('sales')).filter(s => s.firmId === firmId && s.customerId === partyId && s.status !== 'Open');
    const purchases = (await getAllRecords('purchases')).filter(p => p.firmId === firmId && p.supplierId === partyId && p.status !== 'Open');
    const receipts = (await getAllRecords('receipts')).filter(r => r.firmId === firmId && r.ledgerId === partyId);

    // Build Sales timeline
    sales.forEach(s => {
        const isReturn = s.documentType === 'return';
        const isNonGST = s.invoiceType === 'Non-GST';
        // FIX: Dynamically label Bill of Supply vs Tax Invoice
        const docLabel = isReturn ? 'Credit Note' : (isNonGST ? 'Bill of Supply' : 'Sales Invoice');
        // FIX: Bulletproof ID fallback so it never prints ()
        const docNo = s.invoiceNo || s.orderNo || String(s.id).slice(-4).toUpperCase();
        
        timeline.push({
            id: s.id,
            date: s.date,
            desc: `${docLabel} (${docNo})`,
            amount: s.grandTotal,
            isInvoice: !isReturn, 
            impact: isReturn ? -s.grandTotal : s.grandTotal
        });
    });

    // Build Purchase timeline
    purchases.forEach(p => {
        const isReturn = p.documentType === 'return';
        const isNonGST = p.invoiceType === 'Non-GST';
        // FIX: Dynamically label Bill of Supply vs Tax Invoice
        const docLabel = isReturn ? 'Debit Note' : (isNonGST ? 'Bill of Supply' : 'Purchase Bill');
        // FIX: Bulletproof ID fallback
        const docNo = p.poNo || p.invoiceNo || p.orderNo || String(p.id).slice(-4).toUpperCase();
        
        timeline.push({
            id: p.id,
            date: p.date,
            desc: `${docLabel} (${docNo})`,
            amount: p.grandTotal,
            isInvoice: !isReturn,
            impact: isReturn ? -p.grandTotal : p.grandTotal // CRITICAL FIX: Purchase increases debt to supplier
        });
    });

    // Build Receipts timeline
    receipts.forEach(r => {
        const isMoneyIn = r.type === 'in';
        let impact = 0;
        
        // CRITICAL FIX: Math accurately reflects debt depending on party type
        if (isCustomer) {
            impact = isMoneyIn ? -r.amount : r.amount;
        } else {
            impact = isMoneyIn ? r.amount : -r.amount; 
        }
        
        timeline.push({
            id: r.id,
            date: r.date,
            // FIX: Prepend the custom Receipt/Voucher Number to the description
            desc: (r.receiptNo ? r.receiptNo + ' - ' : '') + (r.desc || (isMoneyIn ? 'Payment Received' : 'Payment Made')),
            amount: r.amount,
            isInvoice: false,
            impact: impact 
        });
    });

    // Sort chronologically
    timeline.sort((a, b) => {
        if (a.id === 'open-bal') return -1;
        if (b.id === 'open-bal') return 1;
        return new Date(a.date) - new Date(b.date);
    });

    // Compute Running Balance
    let runningBalance = 0;
    timeline.forEach(t => {
        runningBalance += t.impact;
        t.runningBalance = runningBalance;
    });

    return {
        finalBalance: runningBalance,
        timeline: timeline
    };
};

const getGlobalTimeline = async (firmId) => {
    let timeline = [];
    const sales = (await getAllRecords('sales')).filter(r => r.firmId === firmId && r.status !== 'Open');
    const purchases = (await getAllRecords('purchases')).filter(r => r.firmId === firmId && r.status !== 'Open');
    const receipts = (await getAllRecords('receipts')).filter(r => r.firmId === firmId);
    const expenses = (await getAllRecords('expenses')).filter(r => r.firmId === firmId);

    sales.forEach(s => {
        const isReturn = s.documentType === 'return';
        timeline.push({ id: s.id, date: s.date, type: isReturn ? 'IN' : 'OUT', party: s.customerName, ref: s.invoiceNo, qty: `${isReturn ? '+' : ''}${(s.items || []).length} items`, amount: s.grandTotal, mode: 'Credit', desc: `${isReturn ? 'Return from' : 'Sale to'} ${s.customerName}` });
    });

    purchases.forEach(p => {
        const isReturn = p.documentType === 'return';
        timeline.push({ id: p.id, date: p.date, type: isReturn ? 'OUT' : 'IN', party: p.supplierName, ref: p.poNo || p.invoiceNo, qty: `${isReturn ? '-' : ''}${(p.items || []).length} items`, amount: p.grandTotal, mode: 'Credit', desc: `${isReturn ? 'Return to' : 'Purchase from'} ${p.supplierName}` });
    });

    receipts.forEach(r => {
        timeline.push({ id: r.id, date: r.date, type: r.type, amount: r.amount, mode: r.mode, desc: `Party: ${r.ledgerName}` });
    });

    expenses.forEach(e => {
        timeline.push({ id: e.id, date: e.date, type: 'out', amount: parseFloat(e.amount), mode: 'Cash', desc: `Expense: ${e.category}` });
    });

    return timeline.sort((a, b) => new Date(b.date) - new Date(a.date));
};

// ==========================================
// BULK EXPORT ENGINE
// ==========================================
const exportDatabase = async () => {
    const stores = ['firms', 'businessProfile', 'counters', 'items', 'ledgers', 'sales', 'purchases', 'receipts', 'expenses', 'accounts', 'adjustments', 'units', 'expenseCategories'];
    const backupData = {};
    for (const store of stores) {
        backupData[store] = await getAllRecords(store);
    }
    return backupData;
};

const importDatabase = async (parsedData) => {
    const stores = Object.keys(parsedData);
    return new Promise((resolve, reject) => {
        // Filter out obsolete stores to prevent transaction crashes
        const validStores = stores.filter(store => db.objectStoreNames.contains(store));
        
        if (validStores.length === 0) return reject(new Error("No valid data stores found in backup."));
        
        const transaction = db.transaction(validStores, 'readwrite');
        validStores.forEach(storeName => {
            const store = transaction.objectStore(storeName);
            store.clear(); 
            parsedData[storeName].forEach(record => {
                store.put(record);
            });
        });
        transaction.oncomplete = () => resolve();
        transaction.onerror = () => reject(transaction.error);
    });
};
// ==========================================
// 9. GST REPORTS AGGREGATION ENGINE
// ==========================================
async function generateGSTReport(yearMonth, firmId) {
    const sales = await getAllRecords('sales');
    const purchases = await getAllRecords('purchases');
    const ledgers = await getAllRecords('ledgers');

    // Filter by firm and specific month, and EXCLUDE DRAFTS
    const monthSales = sales.filter(s => s.firmId === firmId && s.date.startsWith(yearMonth) && s.status !== 'Open');
    const monthPurchases = purchases.filter(p => p.firmId === firmId && p.date.startsWith(yearMonth) && p.status !== 'Open');

    let gstr1 = { b2bTaxable: 0, b2bTax: 0, b2cTaxable: 0, b2cTax: 0, totalTaxable: 0, totalTax: 0 };
    let gstr2 = { totalTaxable: 0, totalTax: 0 }; 

    // Calculate GSTR-1 (Outward Supplies / Sales)
    monthSales.forEach(s => {
        // NEW: Omit "Bill of Supply" (Non-GST) from GSTR-1 and GSTR-3B
        if (s.invoiceType === 'Non-GST') return;

        let isReturn = s.documentType === 'return';
        let mult = isReturn ? -1 : 1; // Subtract returns from total sales
        
        // FIX: Reverted Double-Discount Bug! ui.js already applies the discount before saving s.subtotal.
        let taxable = parseFloat(s.subtotal) * mult;
        let tax = parseFloat(s.totalGst) * mult;

        gstr1.totalTaxable += taxable;
        gstr1.totalTax += tax;

        // Determine B2B vs B2C
        let customer = ledgers.find(l => l.id === s.customerId);
        if (customer && customer.gst && customer.gst.trim() !== '') {
            gstr1.b2bTaxable += taxable;
            gstr1.b2bTax += tax;
        } else {
            gstr1.b2cTaxable += taxable;
            gstr1.b2cTax += tax;
        }
    });

    // Calculate GSTR-2 (Inward Supplies / Purchases / ITC)
    monthPurchases.forEach(p => {
        // NEW: Omit Non-GST / Exempt purchases from GSTR-2 and GSTR-3B ITC
        if (p.invoiceType === 'Non-GST') return;

        let isReturn = p.documentType === 'return';
        let mult = isReturn ? -1 : 1;
        
        // FIX: Reverted Double-Discount Bug!
        gstr2.totalTaxable += parseFloat(p.subtotal) * mult;
        gstr2.totalTax += parseFloat(p.totalGst) * mult;
    });

    // Calculate GSTR-3B (Net Summary)
    let gstr3b = {
        outputTax: gstr1.totalTax,
        inputTax: gstr2.totalTax,
        netPayable: gstr1.totalTax - gstr2.totalTax
    };

    return { month: yearMonth, gstr1, gstr2, gstr3b, rawSales: monthSales, rawPurchases: monthPurchases };
}

// ==========================================
// NEW CODE: ES MODULE EXPORT & GLOBAL MAP
// ==========================================
// 1. Export all standalone functions for ES Modules
export {
    initDB, getAllRecords, getRecordById, saveRecord, deleteRecordById,
    getAllFirms, saveInvoiceTransaction, getNextDocumentNumber,
    getKhataStatement, getGlobalTimeline, exportDatabase, importDatabase, generateGSTReport
};

// 2. Map to window so inline HTML and older files don't break during the transition
window.initDB = initDB;
window.getAllRecords = getAllRecords;
window.getRecordById = getRecordById;
window.saveRecord = saveRecord;
window.deleteRecordById = deleteRecordById;
window.getAllFirms = getAllFirms;
window.saveInvoiceTransaction = saveInvoiceTransaction;
window.getNextDocumentNumber = getNextDocumentNumber;
window.getKhataStatement = getKhataStatement;
window.getGlobalTimeline = getGlobalTimeline;
window.exportDatabase = exportDatabase;
window.importDatabase = importDatabase;
window.generateGSTReport = generateGSTReport;
