// ==========================================
// SOLLO ERP - UTILITY, EXPORT & PDF ENGINE (v5.2 Enterprise)
// ==========================================

const Utils = {
    // ==========================================
    // 1. CORE UTILITIES & STRICT MATH
    // ==========================================
    generateId: () => typeof crypto !== 'undefined' && crypto.randomUUID ? crypto.randomUUID() : 'sollo-' + Math.random().toString(36).substr(2, 9) + '-' + Date.now(),

    getLocalDate: () => {
        const d = new Date();
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    },

    // --- NEW CODE: TOAST ENGINE ---
    showToast: (message) => {
        const container = document.getElementById('toast-container');
        if (!container) return;
        
        const toast = document.createElement('div');
        toast.className = 'toast';
        toast.innerText = message;
        container.appendChild(toast);
        
        setTimeout(() => toast.classList.add('show'), 10); // Trigger animation
        
        // Remove after 3 seconds
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    },
    // --- END OF NEW CODE ---

    calculateRowTotal: (qty, rate, gstPercent) => {
        let baseAmount = Math.round(((parseFloat(qty) || 0) * (parseFloat(rate) || 0)) * 100) / 100;
        let gstAmount = Math.round((baseAmount * ((parseFloat(gstPercent) || 0) / 100)) * 100) / 100;
        let finalTotal = Math.round((baseAmount + gstAmount) * 100) / 100;
        return { baseAmount, gstAmount, finalTotal };
    },

    downloadFile: (content, filename, contentType) => {
        const a = document.createElement("a");
        const file = new Blob([content], { type: contentType });
        a.href = URL.createObjectURL(file);
        a.download = filename;
        a.click();
        URL.revokeObjectURL(a.href);
    },
    // ==========================================
    // NEW CODE: UNIVERSAL SHARE ENGINE
    // ==========================================
    shareDocumentAsImage: async (elementId, documentTitle) => {
        try {
            const element = document.getElementById(elementId);
            if (!element) {
                alert("Error: Document area not found.");
                return;
            }

            // 1. Convert the HTML element (like an invoice) into an image
            const canvas = await html2canvas(element, { scale: 4, useCORS: true }); // UPGRADE: Changed 2 to 4
            
            canvas.toBlob(async (blob) => {
                const file = new File([blob], `${documentTitle}.png`, { type: 'image/png' });

                // 2. Check if the device's browser supports the Universal Share API with files
                if (navigator.canShare && navigator.canShare({ files: [file] })) {
                    await navigator.share({
                        title: documentTitle,
                        text: `Please find the attached document: ${documentTitle}`,
                        files: [file]
                    });
                } else {
                    // 3. Fallback: If sharing isn't supported (like on older desktop browsers), just download it
                    const url = URL.createObjectURL(blob);
                    const link = document.createElement("a");
                    link.href = url;
                    link.download = `${documentTitle}.png`;
                    link.click();
                    alert("Universal sharing is not supported on this device. The document has been downloaded instead.");
                }
            }, 'image/png');

        } catch (error) {
            console.error("Sharing failed:", error);
            alert("An error occurred while trying to share the document.");
        }
    },
    // ==========================================
    // END OF NEW CODE
    // ==========================================

    formatWhatsAppNumber: (phone) => {
        let clean = String(phone || '').replace(/\D/g, ''); 
        if (clean.length === 10) clean = '91' + clean; 
        return clean;
    },
    
    shareOverdueReminder: (phone, customerName, balanceAmount, invoiceNo) => {
        if (!phone) return alert("No phone number saved for this customer.");
        const cleanPhone = Utils.formatWhatsAppNumber(phone);
        const message = `🚨 *Payment Reminder* 🚨\n\nDear ${customerName},\nThis is a gentle reminder that a balance of *\u20B9${parseFloat(balanceAmount).toFixed(2)}* is currently overdue against Invoice *#${invoiceNo}*.\n\nPlease clear the dues at your earliest convenience to avoid any interruptions.\n\nThank you!`;
        const whatsappUrl = `https://wa.me/${cleanPhone}?text=${encodeURIComponent(message)}`;
        window.location.href = whatsappUrl;
    },

    // ==========================================
    // 3. DATABASE BACKUP (EXPORT/IMPORT) ENGINE
    // ==========================================
    exportData: async () => {
        try {
            if (typeof exportDatabase !== 'function') return alert("Database export not ready.");
            const data = await exportDatabase();
            const json = JSON.stringify(data); 

            const blob = new Blob([json], { type: "application/json" });
            const file = new File([blob], `SOLLO_Backup_${new Date().toISOString().split('T')[0]}.json`, { type: "application/json" });

            if (navigator.canShare && navigator.canShare({ files: [file] })) {
                await navigator.share({
                    title: "SOLLO ERP Backup",
                    text: "Here is your complete database backup.",
                    files: [file]
                });
            } else {
                const url = URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.style.display = "none";
                a.href = url;
                a.download = file.name;
                document.body.appendChild(a);
                a.click();
                URL.revokeObjectURL(url);
                document.body.removeChild(a);
                alert("✅ Backup downloaded successfully!");
            }
        } catch (e) {
            console.error(e);
            alert("Export failed. Please check your storage permissions.");
        }
    },

    importData: async (event) => {
        if (event && event.target && event.target.files && event.target.files.length > 0) {
            const file = event.target.files[0];
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    const data = JSON.parse(e.target.result);
                    if (typeof importDatabase === 'function') {
                        await importDatabase(data);
                        alert("Database imported successfully! Reloading app...");
                        window.location.reload();
                    }
                } catch (err) {
                    alert("Invalid backup file. Make sure it is a valid SOLLO JSON backup.");
                }
            };
            reader.readAsText(file);
            event.target.value = ''; 
            return;
        }

        // WEBINTOAPP CLIPBOARD FALLBACK
        const jsonStr = prompt("WEBINTOAPP RESTORE:\nSince your app blocks file downloads, you likely backed up your data as Text.\n\nPlease PASTE your backup text here to restore it:");
        if (jsonStr) {
            try {
                const data = JSON.parse(jsonStr);
                if (typeof importDatabase === 'function') {
                    await importDatabase(data);
                    alert("✅ Database restored successfully! Reloading app...");
                    window.location.reload();
                }
            } catch (err) {
                alert("❌ Invalid backup text. Please ensure you copied the exact JSON string.");
            }
        }
    },

    // ==========================================
    // 4. IN-APP INVOICE VIEWER (WEBINTOAPP SAFE)
    // ==========================================
    processPDFExport: async (elementId, filename) => {
        const element = document.getElementById(elementId);
        if (!element) return;

        if (typeof html2canvas === 'undefined') {
            alert("Installing Image Engine for the first time... Please wait 2 seconds and tap Print again.");
            const script = document.createElement('script');
            script.src = "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
            document.head.appendChild(script);
            return;
        }
        
        try {
            const canvas = await html2canvas(element, { 
                scale: 4, 
                useCORS: true,
                backgroundColor: '#ffffff',
                onclone: (clonedDoc) => {
                    const printArea = clonedDoc.getElementById('print-area');
                    if (printArea) {
                        printArea.className = ''; 
                        printArea.style.display = 'block';
                        printArea.style.position = 'relative';
                        printArea.style.visibility = 'visible';
                    }
                }
            });
            
            const imgSrc = canvas.toDataURL('image/jpeg', 1.0);

            const viewer = document.createElement('div');
            viewer.id = 'in-app-pdf-viewer';
            viewer.style.position = 'fixed';
            viewer.style.top = '0';
            viewer.style.left = '0';
            viewer.style.width = '100vw';
            viewer.style.height = '100vh';
            viewer.style.backgroundColor = '#e8eaed';
            viewer.style.zIndex = '999999';
            viewer.style.display = 'flex';
            viewer.style.flexDirection = 'column';

            viewer.innerHTML = `
                <div style="background:#0061a4; color:white; padding:16px; display:flex; justify-content:space-between; align-items:center; box-shadow:0 2px 4px rgba(0,0,0,0.2); flex-shrink:0;">
                    <div>
                        <div style="font-weight:bold; font-size:18px;">Document Preview</div>
                        <div style="font-size:12px; opacity:0.9; margin-top:2px;">Share directly or Print</div>
                    </div>
                    <div style="display: flex; gap: 20px; align-items: center;">
                        <span class="material-symbols-outlined tap-target" style="font-size:24px;" id="btn-print-preview">print</span>
                        <span class="material-symbols-outlined tap-target" style="font-size:24px;" id="btn-share-preview">share</span>
                        <span class="material-symbols-outlined tap-target" style="font-size:28px;" onclick="document.getElementById('in-app-pdf-viewer').remove()">close</span>
                    </div>
                </div>
                <div style="flex:1; overflow-y:auto; padding:16px; display:flex; justify-content:center; align-items:flex-start;">
                    <img src="${imgSrc}" style="max-width:100%; height:auto; box-shadow:0 4px 8px rgba(0,0,0,0.2); border-radius:4px; display:block;" />
                </div>
            `;
            document.body.appendChild(viewer);

            // NEW: Native OS Print Engine (PDF Export / Hardware Printing)
            document.getElementById('btn-print-preview').onclick = () => {
                document.getElementById('in-app-pdf-viewer').style.display = 'none'; // Temporarily hide viewer
                window.print();
                setTimeout(() => { document.getElementById('in-app-pdf-viewer').style.display = 'flex'; }, 500);
            };
            
            // NEW: Native OS Web Share API (WhatsApp, Telegram, Gmail, etc.)
            document.getElementById('btn-share-preview').onclick = async () => {
                try {
                    if (navigator.share) {
                        const response = await fetch(imgSrc);
                        const blob = await response.blob();
                        const file = new File([blob], filename.replace('.pdf', '.jpg'), { type: 'image/jpeg' });
                        await navigator.share({
                            title: filename,
                            files: [file]
                        });
                    } else {
                        // Fallback if testing on an outdated browser or locked-down wrapper
                        alert("Native sharing is blocked by your current browser. Please LONG-PRESS the image below to Share or Save it.");
                    }
                } catch (err) {
                    console.log("Share cancelled or failed", err);
                }
            };
        } catch (err) {
            console.error("Preview Generation Failed", err);
            alert("Failed to generate preview.");
        } finally {
            const printArea = document.getElementById('print-area');
            if (printArea) printArea.innerHTML = ''; 
        }
    },

    // ==========================================
    // 5. PRINT & TEMPLATE ENGINE
    // ==========================================
    generateInvoicePDF: (doc, biz, party, type) => {
        const isSales = type === 'sales';
        const partyName = (party && party.name) ? party.name : (isSales ? doc.customerName : doc.supplierName);
        const partyAddress = party ? (party.address || '') : '';
        
        const partyGst = party && party.gst ? party.gst.toUpperCase() : '';
        const bizGst = biz && biz.gst ? biz.gst.toUpperCase() : 'N/A';
        
        const isNonGST = doc.invoiceType === 'Non-GST';
        const isReturn = doc.documentType === 'return';
        
        let title = isSales ? 'TAX INVOICE' : 'PURCHASE RECORD';
        if (isNonGST && !isReturn) title = isSales ? 'BILL OF SUPPLY' : 'PURCHASE RECORD';
        if (isReturn) title = isSales ? 'CREDIT NOTE' : 'DEBIT NOTE';

        let itemsHtml = '';
        let rawSubtotal = 0;
        
        (doc.items || []).forEach((item, index) => {
            const rowTotal = Utils.calculateRowTotal(item.qty, item.rate, item.gstPercent).finalTotal;
            rawSubtotal += (parseFloat(item.qty) || 0) * (parseFloat(item.rate) || 0);
            
            itemsHtml += `
                <tr>
                    <td style="padding: 10px 8px; border-bottom: 1px solid #e0e0e0; text-align:center;">${index + 1}</td>
                    <td style="padding: 10px 8px; border-bottom: 1px solid #e0e0e0; font-weight: 500; color: #1a1c1e;">${item.name}</td>
                    ${!isNonGST ? `<td style="padding: 10px 8px; border-bottom: 1px solid #e0e0e0; text-align:center;">${item.hsn || '-'}</td>` : ''}
                    <td style="padding: 10px 8px; border-bottom: 1px solid #e0e0e0; text-align:center;">${item.qty} ${item.uom}</td>
                    <td style="padding: 10px 8px; border-bottom: 1px solid #e0e0e0; text-align:right;">${parseFloat(item.rate).toFixed(2)}</td>
                    ${!isNonGST ? `<td style="padding: 10px 8px; border-bottom: 1px solid #e0e0e0; text-align:center;">${item.gstPercent}%</td>` : ''}
                    <td style="padding: 10px 8px; border-bottom: 1px solid #e0e0e0; text-align:right; font-weight:bold; color: #1a1c1e;">${rowTotal.toFixed(2)}</td>
                </tr>
            `;
        });

        let discountAmt = doc.discountType === '%' ? (rawSubtotal * ((parseFloat(doc.discount) || 0) / 100)) : (parseFloat(doc.discount) || 0);
        if (discountAmt > rawSubtotal) discountAmt = rawSubtotal;
        const safeDocNo = doc.invoiceNo || doc.poNo || 'DRAFT';

        const html = `
            <div id="pdf-invoice-wrapper" class="a4-document" style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; box-sizing: border-box; position: relative; background: #ffffff; overflow: hidden;">
                
                ${biz.logo ? `<div style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); opacity: 0.04; z-index: 0; width: 70%; display: flex; justify-content: center; pointer-events: none;"><img src="${biz.logo}" style="width: 100%; height: auto; object-fit: contain; filter: grayscale(100%);" /></div>` : ''}

                <style>
                    #pdf-invoice-wrapper * { position: relative; z-index: 1; }
                    #pdf-invoice-wrapper table { width: 100%; border-collapse: collapse; margin-bottom: 25px; font-size: 12px; page-break-inside: auto; }
                    #pdf-invoice-wrapper th { background-color: #f8f9fa; color: #1a1c1e; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px; border-top: 2px solid #1a1c1e; border-bottom: 2px solid #1a1c1e; padding: 12px 8px; }
                    #pdf-invoice-wrapper td { border-bottom: 1px solid #e0e0e0; padding: 10px 8px; }
                    #pdf-invoice-wrapper tr { page-break-inside: avoid; page-break-after: auto; break-inside: avoid; }
                    #pdf-invoice-wrapper thead { display: table-header-group; }
                </style>

                <div style="display: flex; justify-content: space-between; align-items: flex-start; padding-bottom: 20px; margin-bottom: 20px; border-bottom: 3px solid #1a1c1e;">
                    <div style="max-width: 60%;">
                        <h1 style="margin: 0 0 8px 0; font-size: 26px; color: #1a1c1e; text-transform: uppercase; letter-spacing: 1px;">${biz.name || 'Company Name'}</h1>
                        <p style="margin: 4px 0; font-size: 12px; color: #43474e;">${biz.address || ''}</p>
                        <p style="margin: 4px 0; font-size: 12px; color: #43474e;">Phone: ${biz.phone || ''}</p>
                        <p style="margin: 4px 0; font-size: 12px; font-weight: bold; color: #1a1c1e;">GSTIN: ${bizGst}</p>
                    </div>
                    <div style="text-align: right; display: flex; flex-direction: column; align-items: flex-end;">
                        ${biz.logo ? `<img src="${biz.logo}" style="max-height: 70px; margin-bottom: 10px; border-radius: 4px;" />` : ''}
                        <div style="background: #f8f9fa; padding: 8px 12px; border: 1px solid #e0e0e0; border-radius: 4px; display: inline-block;">
                            <h2 style="margin: 0; font-size: 18px; color: #1a1c1e; letter-spacing: 2px; text-transform: uppercase;">${title}</h2>
                        </div>
                    </div>
                </div>
                
                <div style="display: flex; justify-content: space-between; margin-bottom: 30px;">
                    <div style="width: 48%;">
                        <p style="margin: 0 0 6px 0; font-size: 11px; color: #73777f; font-weight: bold; text-transform: uppercase; letter-spacing: 0.5px;">Billed To</p>
                        <div style="border-left: 3px solid #1a1c1e; padding-left: 12px;">
                            <p style="margin: 0 0 4px 0; font-size: 15px; font-weight: bold; color: #1a1c1e;">${partyName}</p>
                            ${partyAddress ? `<p style="margin: 0 0 4px 0; font-size: 12px; color: #43474e; white-space: pre-wrap; line-height: 1.4;">${partyAddress}</p>` : ''}
                            ${!isNonGST && partyGst ? `<p style="margin: 6px 0 0 0; font-size: 12px; font-weight: bold; color: #1a1c1e;">GSTIN: ${partyGst}</p>` : ''}
                        </div>
                    </div>
                    <div style="width: 45%;">
                        <table style="width: 100%; font-size: 12px; margin-bottom:0; border: none;">
                            <tr><td style="color: #73777f; padding: 4px 0; border:none; text-transform: uppercase; font-size: 10px; font-weight: bold;">Document No:</td><td style="font-weight: bold; text-align:right; padding: 4px 0; border:none; color: #1a1c1e;">${safeDocNo}</td></tr>
                            <tr><td style="color: #73777f; padding: 4px 0; border:none; text-transform: uppercase; font-size: 10px; font-weight: bold;">Invoice Date:</td><td style="font-weight: bold; text-align:right; padding: 4px 0; border:none; color: #1a1c1e;">${doc.date}</td></tr>
                            ${doc.orderNo ? `<tr><td style="color: #73777f; padding: 4px 0; border:none; text-transform: uppercase; font-size: 10px; font-weight: bold;">Order Ref:</td><td style="font-weight: bold; text-align:right; padding: 4px 0; border:none; color: #1a1c1e;">${doc.orderNo}</td></tr>` : ''}
                            ${doc.shippedDate ? `<tr><td style="color: #73777f; padding: 4px 0; border:none; text-transform: uppercase; font-size: 10px; font-weight: bold;">Shipped Date:</td><td style="font-weight: bold; text-align:right; padding: 4px 0; border:none; color: #1a1c1e;">${doc.shippedDate}</td></tr>` : ''}
                            ${doc.completedDate ? `<tr><td style="color: #73777f; padding: 4px 0; border:none; text-transform: uppercase; font-size: 10px; font-weight: bold;">Completed Date:</td><td style="font-weight: bold; text-align:right; padding: 4px 0; border:none; color: #1a1c1e;">${doc.completedDate}</td></tr>` : ''}
                            <tr><td style="color: #73777f; padding: 4px 0; border:none; text-transform: uppercase; font-size: 10px; font-weight: bold;">Status:</td><td style="font-weight: bold; text-align:right; padding: 4px 0; border:none; color: #0061a4;">${doc.status}</td></tr>
                        </table>
                    </div>
                </div>

                <table style="border-bottom: 2px solid #1a1c1e;">
                    <thead>
                        <tr>
                            <th style="text-align:center; width: 5%;">#</th>
                            <th style="text-align:left; width: 35%;">Item Description</th>
                            ${!isNonGST ? `<th style="text-align:center; width: 10%;">HSN</th>` : ''}
                            <th style="text-align:center; width: 10%;">Qty</th>
                            <th style="text-align:right; width: 15%;">Rate</th>
                            ${!isNonGST ? `<th style="text-align:center; width: 10%;">GST</th>` : ''}
                            <th style="text-align:right; width: 15%;">Total</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${itemsHtml}
                    </tbody>
                </table>

                <div class="avoid-break" style="display: flex; justify-content: flex-end; page-break-inside: avoid; margin-top: 15px;">
                    <table style="width: 55%; border: none; font-size: 13px;">
                        <tr><td style="padding: 6px 0; color: #73777f; border:none;">Subtotal:</td><td style="padding: 6px 0; text-align:right; font-weight:bold; color:#1a1c1e; border:none;">&#8377;${rawSubtotal.toFixed(2)}</td></tr>
                        ${discountAmt > 0 ? `<tr><td style="padding: 6px 0; color: #73777f; border:none;">Discount:</td><td style="padding: 6px 0; text-align:right; font-weight:bold; color:var(--md-error); border:none;">-&#8377;${discountAmt.toFixed(2)}</td></tr>
                        <tr><td style="padding: 6px 0; color: #73777f; border:none;">Taxable Value:</td><td style="padding: 6px 0; text-align:right; font-weight:bold; color:#1a1c1e; border:none;">&#8377;${(rawSubtotal - discountAmt).toFixed(2)}</td></tr>` : ''}
                        ${!isNonGST ? `<tr><td style="padding: 6px 0; color: #73777f; border:none;">Total GST:</td><td style="padding: 6px 0; text-align:right; font-weight:bold; color:#1a1c1e; border:none;">&#8377;${(parseFloat(doc.totalGst) || 0).toFixed(2)}</td></tr>` : ''}
                        ${(parseFloat(doc.freightAmount) || 0) > 0 ? `<tr><td style="padding: 6px 0; color: #73777f; border:none;">Freight / Shipping:</td><td style="padding: 6px 0; text-align:right; font-weight:bold; color:#1a1c1e; border:none;">&#8377;${(parseFloat(doc.freightAmount) || 0).toFixed(2)}</td></tr>` : ''}
                        
                        <tr style="background-color: #f8f9fa;"><td style="padding: 12px 10px; font-weight:bold; font-size: 14px; color: #1a1c1e; border-top: 2px solid #1a1c1e; border-bottom: 2px solid #1a1c1e; text-transform: uppercase;">Grand Total:</td><td style="padding: 12px 10px; text-align:right; font-size: 16px; font-weight:bold; color: #1a1c1e; border-top: 2px solid #1a1c1e; border-bottom: 2px solid #1a1c1e;">&#8377;${(parseFloat(doc.grandTotal) || 0).toFixed(2)}</td></tr>
                        
                        ${doc.linkedReceipts && doc.linkedReceipts.length > 0 ? doc.linkedReceipts.map(r => `
                            <tr>
                                <td style="padding: 6px 0; color: #146c2e; border:none; font-size: 11px;">
                                    Payment (${r.receiptNo || 'Auto'} | ${r.date}):
                                </td>
                                <td style="padding: 6px 0; text-align:right; font-weight:bold; color:#146c2e; border:none; font-size: 12px;">
                                    -&#8377;${parseFloat(r.amount).toFixed(2)}
                                </td>
                            </tr>
                        `).join('') : ''}
                        
                        ${((parseFloat(doc.grandTotal) || 0) - (doc.trueTotalPaid || 0)) > 0.01 ? `
                        <tr><td style="padding: 10px 0; font-weight:bold; font-size: 13px; color: #ba1a1a; border-top: 1px dashed #c3c7cf; border-bottom: none;">Balance Due:</td><td style="padding: 10px 0; text-align:right; font-size: 14px; font-weight:bold; color: #ba1a1a; border-top: 1px dashed #c3c7cf; border-bottom: none;">&#8377;${Math.max(0, (parseFloat(doc.grandTotal) || 0) - (doc.trueTotalPaid || 0)).toFixed(2)}</td></tr>
                        ` : `
                        <tr><td style="padding: 10px 0; font-weight:bold; font-size: 13px; color: #146c2e; border-top: 1px dashed #c3c7cf; border-bottom: none;">Balance Due:</td><td style="padding: 10px 0; text-align:right; font-size: 14px; font-weight:bold; color: #146c2e; border-top: 1px dashed #c3c7cf; border-bottom: none;">&#8377;0.00 (PAID)</td></tr>
                        `}
                    </table>
                </div>
                
                <div class="avoid-break" style="margin-top: 30px; display: flex; justify-content: space-between; align-items: stretch; page-break-inside: avoid;">
                    <div style="width: 55%; font-size: 11px; color: #43474e; display: flex; flex-direction: column; justify-content: space-between;">
                        ${biz.bankDetails ? `
                        <div style="border: 1px solid #e0e0e0; padding: 12px; border-radius: 4px; margin-bottom: 12px; background: #f8f9fa;">
                            <strong style="color: #1a1c1e; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Bank Details</strong><br>
                            <span style="white-space: pre-wrap; line-height: 1.5;">${biz.bankDetails}</span>
                        </div>` : '<div></div>'}
                        
                        <div style="padding-top: 10px;">
                            <strong style="color: #1a1c1e;">Terms & Conditions:</strong><br>
                            <span style="white-space: pre-wrap; line-height: 1.4;">${biz.terms ? biz.terms : '1. Subject to local jurisdiction.\\n2. This is a computer-generated document.'}</span>
                        </div>
                    </div>
                    
                    <div style="width: 38%; text-align: center; border: 1px solid #e0e0e0; border-radius: 4px; padding: 15px; display: flex; flex-direction: column; justify-content: flex-end; align-items: center; min-height: 130px; background: #ffffff;">
                        ${biz.signature ? `<img src="${biz.signature}" style="max-height: 50px; margin-bottom: 15px; object-fit: contain;" />` : '<div style="height: 50px; margin-bottom: 15px;"></div>'}
                        <div style="border-top: 1px solid #1a1c1e; width: 90%; padding-top: 8px; font-weight: bold; font-size: 12px; color: #1a1c1e;">Authorized Signatory</div>
                        <div style="font-size: 10px; color: #73777f; margin-top: 4px; font-weight: bold;">For ${biz.name || 'Company'}</div>
                    </div>
                </div>

                <div class="avoid-break" style="margin-top: 20px; font-size: 12px; text-align: center; color: #0061a4; font-weight: bold; page-break-inside: avoid;">
                    Thank you for your business!
                </div>
            </div>
        `;
        
        const printArea = document.getElementById('print-area');
        if (printArea) {
            printArea.innerHTML = html;
            setTimeout(() => {
                const safeFilenameDocNo = String(safeDocNo).replace(/[^a-zA-Z0-9_.-]/g, '-');
                Utils.processPDFExport('pdf-invoice-wrapper', `${title.replace(/ /g, '_')}_${safeFilenameDocNo}.pdf`);
            }, 100);
        }
    },

    printReceivablesReport: (reportData, grandTotal) => {
        let html = `
            <div id="pdf-rec-wrapper" style="padding: 20px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #1a1c1e; background: #ffffff;">
                <style>
                    #pdf-rec-wrapper table { width: 100%; border-collapse: collapse; margin-top: 20px; page-break-inside: auto; }
                    #pdf-rec-wrapper tr { page-break-inside: avoid; page-break-after: auto; break-inside: avoid; }
                    #pdf-rec-wrapper thead { display: table-header-group; }
                </style>
                <h1 style="text-align: center; color: #0061a4; text-transform: uppercase;">Receivables Report</h1>
                <p style="text-align: center; color: #73777f;">Generated on: ${Utils.getLocalDate()}</p>
                <table>
                    <thead>
                        <tr style="background-color: #f3f3f3;">
                            <th style="border: 1px solid #ccc; padding: 10px; text-align: left;">Customer Name</th>
                            <th style="border: 1px solid #ccc; padding: 10px; text-align: left;">Contact</th>
                            <th style="border: 1px solid #ccc; padding: 10px; text-align: right;">Outstanding Balance (&#8377;)</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        reportData.forEach(row => {
            html += `
                <tr>
                    <td style="border: 1px solid #ccc; padding: 10px; font-weight: bold;">${row.name}</td>
                    <td style="border: 1px solid #ccc; padding: 10px;">${row.phone || 'N/A'}</td>
                    <td style="border: 1px solid #ccc; padding: 10px; text-align: right; color: #ba1a1a; font-weight: bold;">${parseFloat(row.balance || 0).toFixed(2)}</td>
                </tr>
            `;
        });

        html += `
                    </tbody>
                </table>
                <div style="text-align: right; margin-top: 20px;">
                    <h2 style="color: #0061a4;">Total Market Receivables: &#8377;${parseFloat(grandTotal || 0).toFixed(2)}</h2>
                </div>
            </div>
        `;
        const printArea = document.getElementById('print-area');
        if (printArea) {
            printArea.innerHTML = html;
            setTimeout(() => {
                Utils.processPDFExport('pdf-rec-wrapper', `Receivables_${Utils.getLocalDate()}.pdf`);
            }, 100);
        }
    },

    downloadStatementPDF: async () => {
        const nameEl = document.getElementById('report-party-name');
        const balEl = document.getElementById('report-party-balance');
        if (!nameEl) return alert("Report data not found.");

        // FIX 1: Trim invisible spaces to prevent match failures
        const partyName = nameEl.innerText.trim();
        let isAccount = false;
        
        // SAFEGUARD: Case-insensitive search 
        let party = UI.state.rawData.ledgers.find(l => (l.name || '').trim().toLowerCase() === partyName.toLowerCase());
        
        // Check if this is a Bank Account/Cash Drawer/UPI instead of a Ledger Party
        if (!party) {
            party = UI.state.rawData.accounts.find(a => (a.name || '').trim().toLowerCase() === partyName.toLowerCase());
            // Make Cash Drawer detection bulletproof
            if (!party && partyName.toLowerCase().includes('cash')) {
                party = { id: 'cash', name: 'Cash Drawer', type: 'Account', firmId: typeof app !== 'undefined' && app.state ? app.state.firmId : 'firm1' };
            }
            if (party) isAccount = true;
        }
        
        if (!party) return alert("Could not identify details for: " + partyName);

        const biz = (party.firmId) ? await getRecordById('businessProfile', party.firmId) || {} : {};
        
        // Harness the timeline that is already populated in the UI to prevent bugs
        const timeline = UI.state.rawData.timeline || [];
        
        let finalBal = 0;
        if(balEl) {
            const balMatch = balEl.innerText.replace(/,/g, '').match(/[\d.]+/);
            if(balMatch) finalBal = parseFloat(balMatch[0]);
        }

        // FIX 2: Safely calculate Opening Balance from the timeline to prevent silent crashes!
        let openingBal = 0;
        const openEntry = timeline.find(t => t.id === 'open-bal');
        if (openEntry) openingBal = openEntry.impact || openEntry.amount || 0;
        
        let tableRows = '';
        timeline.forEach(t => {
            let debit = '';
            let credit = '';
            
            if (isAccount) {
                if (t.impact > 0) debit = Math.abs(t.impact).toFixed(2);
                else credit = Math.abs(t.impact).toFixed(2);
            } else {
                if (party.type === 'Customer') {
                    if (t.isInvoice) debit = t.amount.toFixed(2);
                    else credit = t.amount.toFixed(2);
                } else {
                    if (t.isInvoice) credit = t.amount.toFixed(2);
                    else debit = t.amount.toFixed(2);
                }
            }

            tableRows += `
                <tr>
                    <td style="text-align:center;">${t.date}</td>
                    <td style="font-weight: 500;">
                        ${t.desc || '-'} 
                        ${t.partyName && t.partyName !== 'Unknown' ? `<br><small style="color:#1a1c1e; font-weight:bold;">Party: ${t.partyName}</small>` : ''}
                        ${t.ref ? `<br><small style="color:#73777f; font-weight:normal;">Ref: ${t.ref}</small>` : ''}
                    </td>
                    <td style="text-align:right; color:#ba1a1a;">${debit ? '\u20B9' + debit : ''}</td>
                    <td style="text-align:right; color:#146c2e;">${credit ? '\u20B9' + credit : ''}</td>
                    <td style="text-align:right; font-weight:bold; color:#1a1c1e;">
                        \u20B9${Math.abs(t.runningBalance || 0).toFixed(2)} 
                        <span style="font-size:11px; color:#73777f;">${(t.runningBalance || 0) > 0 ? (party.type === 'Customer' ? 'Dr' : 'Cr') : (party.type === 'Customer' ? 'Cr' : 'Dr')}</span>
                    </td>
                </tr>
            `;
        });

        const safeDocNo = Utils.getLocalDate();
        const balSuffix = isAccount ? 'Available' : (finalBal > 0 ? (party.type === 'Customer' ? 'Dr (Due)' : 'Cr (To Pay)') : (party.type === 'Customer' ? 'Cr (Advance)' : 'Dr (Advance)'));

        const html = `
            <div id="pdf-statement-wrapper" class="a4-document" style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; box-sizing: border-box; position: relative; background: #ffffff; overflow: hidden;">
                
                ${biz.logo ? `<div style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); opacity: 0.04; z-index: 0; width: 70%; display: flex; justify-content: center; pointer-events: none;"><img src="${biz.logo}" style="width: 100%; height: auto; object-fit: contain; filter: grayscale(100%);" /></div>` : ''}

                <style>
                    #pdf-statement-wrapper * { position: relative; z-index: 1; }
                    #pdf-statement-wrapper table { width: 100%; border-collapse: collapse; margin-bottom: 25px; font-size: 12px; page-break-inside: auto; }
                    #pdf-statement-wrapper th { background-color: #f8f9fa; color: #1a1c1e; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px; border-top: 2px solid #1a1c1e; border-bottom: 2px solid #1a1c1e; padding: 12px 8px; }
                    #pdf-statement-wrapper td { border-bottom: 1px solid #e0e0e0; padding: 10px 8px; }
                    #pdf-statement-wrapper tr { page-break-inside: avoid; page-break-after: auto; break-inside: avoid; }
                    #pdf-statement-wrapper thead { display: table-header-group; }
                </style>

                <div style="display: flex; justify-content: space-between; align-items: flex-start; padding-bottom: 20px; margin-bottom: 20px; border-bottom: 3px solid #1a1c1e;">
                    <div style="max-width: 60%;">
                        <h1 style="margin: 0 0 8px 0; font-size: 26px; color: #1a1c1e; text-transform: uppercase; letter-spacing: 1px;">${biz.name || 'Company Name'}</h1>
                        <p style="margin: 4px 0; font-size: 12px; color: #43474e;">${biz.address || ''}</p>
                        <p style="margin: 4px 0; font-size: 12px; color: #43474e;">Phone: ${biz.phone || ''}</p>
                        <p style="margin: 4px 0; font-size: 12px; font-weight: bold; color: #1a1c1e;">GSTIN: ${biz.gst ? biz.gst.toUpperCase() : 'N/A'}</p>
                    </div>
                    <div style="text-align: right; display: flex; flex-direction: column; align-items: flex-end;">
                        ${biz.logo ? `<img src="${biz.logo}" style="max-height: 70px; margin-bottom: 10px; border-radius: 4px;" />` : ''}
                        <div style="background: #f8f9fa; padding: 8px 12px; border: 1px solid #e0e0e0; border-radius: 4px; display: inline-block;">
                            <h2 style="margin: 0; font-size: 18px; color: #1a1c1e; letter-spacing: 2px; text-transform: uppercase;">${isAccount ? 'Account Statement' : 'Ledger Statement'}</h2>
                        </div>
                    </div>
                </div>
                
                <div style="display: flex; justify-content: space-between; margin-bottom: 30px;">
                    <div style="width: 48%;">
                        <p style="margin: 0 0 6px 0; font-size: 11px; color: #73777f; font-weight: bold; text-transform: uppercase; letter-spacing: 0.5px;">${isAccount ? 'Account Details' : 'Party Details'}</p>
                        <div style="border-left: 3px solid #1a1c1e; padding-left: 12px;">
                            <p style="margin: 0 0 4px 0; font-size: 15px; font-weight: bold; color: #1a1c1e;">${party.name}</p>
                            ${party.phone ? `<p style="margin: 0 0 4px 0; font-size: 12px; color: #43474e;">Ph: ${party.phone}</p>` : ''}
                            ${party.gst ? `<p style="margin: 0 0 4px 0; font-size: 12px; font-weight: bold; color: #1a1c1e;">GSTIN: ${party.gst.toUpperCase()}</p>` : ''}
                            ${party.address ? `<p style="margin: 6px 0 0 0; font-size: 12px; color: #43474e; white-space: pre-wrap; line-height: 1.4;">${party.address}</p>` : ''}
                        </div>
                    </div>
                    <div style="width: 45%;">
                        <table style="width: 100%; font-size: 12px; margin-bottom:0; border: none;">
                            <tr><td style="color: #73777f; padding: 4px 0; border:none; text-transform: uppercase; font-size: 10px; font-weight: bold;">Generated On:</td><td style="font-weight: bold; text-align:right; padding: 4px 0; border:none; color: #1a1c1e;">${Utils.getLocalDate()}</td></tr>
                            <tr><td style="color: #73777f; padding: 4px 0; border:none; text-transform: uppercase; font-size: 10px; font-weight: bold;">Opening Balance:</td><td style="font-weight: bold; text-align:right; padding: 4px 0; border:none; color: #1a1c1e;">\u20B9${Math.abs(openingBal).toFixed(2)}</td></tr>
                            <tr><td style="color: #73777f; padding: 8px 0 4px 0; border:none; text-transform: uppercase; font-size: 11px; font-weight: bold;">Closing Balance:</td><td style="font-weight: bold; text-align:right; padding: 8px 0 4px 0; border:none; color: ${isAccount ? (finalBal >= 0 ? '#146c2e' : '#ba1a1a') : (finalBal > 0 ? '#ba1a1a' : '#146c2e')}; font-size: 16px;">\u20B9${Math.abs(finalBal).toFixed(2)} <span style="font-size:11px;">${balSuffix}</span></td></tr>
                        </table>
                    </div>
                </div>

                <table style="border-bottom: 2px solid #1a1c1e;">
                    <thead>
                        <tr>
                            <th style="text-align:center; width: 15%;">Date</th>
                            <th style="text-align:left; width: 40%;">Particulars / Voucher Type</th>
                            <th style="text-align:right; width: 15%;">Money Out (Dr)</th>
                            <th style="text-align:right; width: 15%;">Money In (Cr)</th>
                            <th style="text-align:right; width: 15%;">Balance</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${tableRows.length > 0 ? tableRows : '<tr><td colspan="5" style="padding:20px; text-align:center; color:#73777f; border:none;">No transactions found.</td></tr>'}
                    </tbody>
                </table>
                
                <div class="avoid-break" style="margin-top: 30px; display: flex; justify-content: space-between; align-items: stretch; page-break-inside: avoid;">
                    <div style="width: 55%; font-size: 11px; color: #43474e; display: flex; flex-direction: column; justify-content: flex-end;">
                        <p style="margin: 0 0 4px 0;">This is a computer-generated account statement.</p>
                        <p style="margin: 0;">Errors and omissions expected (E&OE).</p>
                    </div>
                    
                    <div style="width: 38%; text-align: center; border: 1px solid #e0e0e0; border-radius: 4px; padding: 15px; display: flex; flex-direction: column; justify-content: flex-end; align-items: center; min-height: 130px; background: #ffffff;">
                        ${biz.signature ? `<img src="${biz.signature}" style="max-height: 50px; margin-bottom: 15px; object-fit: contain;" />` : '<div style="height: 50px; margin-bottom: 15px;"></div>'}
                        <div style="border-top: 1px solid #1a1c1e; width: 90%; padding-top: 8px; font-weight: bold; font-size: 12px; color: #1a1c1e;">Authorized Signatory</div>
                        <div style="font-size: 10px; color: #73777f; margin-top: 4px; font-weight: bold;">For ${biz.name || 'Company'}</div>
                    </div>
                </div>
            </div>
        `;

        const printArea = document.getElementById('print-area');
        if (printArea) {
            printArea.innerHTML = html;
            setTimeout(() => {
                Utils.processPDFExport('pdf-statement-wrapper', `Account_Statement_${party.name.replace(/\s+/g, '_')}_${safeDocNo}.pdf`);
            }, 100);
        }
    },

    generateReceiptPDF: async (receiptId) => {
        const receipt = await getRecordById('receipts', receiptId);
        if (!receipt) return alert("Receipt not found. Please save it first.");
        
        const biz = await getRecordById('businessProfile', receipt.firmId) || {};
        
        const isMoneyIn = receipt.type === 'in';
        const title = isMoneyIn ? 'PAYMENT RECEIPT' : 'PAYMENT VOUCHER';
        const safeDocNo = receipt.receiptNo || String(receipt.id).substring(0, 12).toUpperCase();
        
        // NEW: Translate Multiple Invoices into readable names for the PDF
        let invoiceRefDisplay = '';
        if (receipt.invoiceRef) {
            const refs = String(receipt.invoiceRef).split(',').map(r => r.trim());
            const store = isMoneyIn ? 'sales' : 'purchases';
            let allDocs = [];
            if (typeof getAllRecords === 'function') allDocs = await getAllRecords(store);
            
            const displayNames = refs.map(ref => {
                // FIX: Check ALL cross-linked ID references (including orderNo) so the PDF can translate them!
                const doc = allDocs.find(d => d.id === ref || d.invoiceNo === ref || d.poNo === ref || d.orderNo === ref);
                if (doc) {
                    if (isMoneyIn) return doc.invoiceNo ? doc.invoiceNo : ('Bill of Supply' + (doc.orderNo ? ` (Ref: ${doc.orderNo})` : ''));
                    else return (doc.poNo || doc.invoiceNo) ? (doc.poNo || doc.invoiceNo) : ('Bill of Supply' + (doc.orderNo ? ` (Ref: ${doc.orderNo})` : ''));
                }
                return ref.startsWith('sollo-') ? 'Bill of Supply' : ref;
            });
            invoiceRefDisplay = displayNames.join(', ');
        }
        
        let balanceText = '';
        if (receipt.ledgerId && typeof getKhataStatement === 'function') {
            const party = await getRecordById('ledgers', receipt.ledgerId);
            if (party) {
                const statement = await getKhataStatement(party.id, party.type);
                balanceText = `<p style="margin: 8px 0 0 0; font-size: 14px; color: #43474e;">Current Party Balance: <strong>\u20B9${Math.abs(statement.finalBalance).toFixed(2)} ${statement.finalBalance > 0 ? (party.type === 'Customer' ? '(Due)' : '(To Pay)') : '(Advance)'}</strong></p>`;
            }
        }

        const html = `
            <div id="pdf-receipt-wrapper" class="a4-document" style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; box-sizing: border-box; position: relative; background: #ffffff; overflow: hidden;">
                
                ${biz.logo ? `<div style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); opacity: 0.04; z-index: 0; width: 70%; display: flex; justify-content: center; pointer-events: none;"><img src="${biz.logo}" style="width: 100%; height: auto; object-fit: contain; filter: grayscale(100%);" /></div>` : ''}

                <div style="display: flex; justify-content: space-between; align-items: flex-start; padding-bottom: 20px; margin-bottom: 20px; border-bottom: 3px solid #1a1c1e; position: relative; z-index: 1;">
                    <div style="max-width: 60%;">
                        <h1 style="margin: 0 0 8px 0; font-size: 26px; color: #1a1c1e; text-transform: uppercase; letter-spacing: 1px;">${biz.name || 'Company Name'}</h1>
                        <p style="margin: 4px 0; font-size: 12px; color: #43474e;">${biz.address || ''}</p>
                        <p style="margin: 4px 0; font-size: 12px; color: #43474e;">Phone: ${biz.phone || ''}</p>
                    </div>
                    <div style="text-align: right; display: flex; flex-direction: column; align-items: flex-end;">
                        ${biz.logo ? `<img src="${biz.logo}" style="max-height: 70px; margin-bottom: 10px; border-radius: 4px;" />` : ''}
                        <div style="background: #f8f9fa; padding: 8px 12px; border: 1px solid #e0e0e0; border-radius: 4px; display: inline-block;">
                            <h2 style="margin: 0; font-size: 18px; color: #1a1c1e; letter-spacing: 2px; text-transform: uppercase;">${title}</h2>
                        </div>
                    </div>
                </div>
                
                <div style="border: 2px solid #1a1c1e; border-radius: 4px; padding: 25px; background: rgba(255,255,255,0.8); margin-bottom: 30px; position: relative; z-index: 1;">
                    <table style="width: 100%; font-size: 14px; border-collapse: collapse; border: none;">
                        <tr><td style="padding: 12px 0; color: #73777f; border-bottom: 1px solid #e0e0e0; width: 35%; text-transform: uppercase; font-size: 11px; font-weight: bold;">Receipt No:</td><td style="padding: 12px 0; font-weight: bold; text-align: right; border-bottom: 1px solid #e0e0e0; color: #1a1c1e;">${safeDocNo}</td></tr>
                        <tr><td style="padding: 12px 0; color: #73777f; border-bottom: 1px solid #e0e0e0; text-transform: uppercase; font-size: 11px; font-weight: bold;">Date:</td><td style="padding: 12px 0; font-weight: bold; text-align: right; border-bottom: 1px solid #e0e0e0; color: #1a1c1e;">${receipt.date}</td></tr>
                        <tr><td style="padding: 12px 0; color: #73777f; border-bottom: 1px solid #e0e0e0; text-transform: uppercase; font-size: 11px; font-weight: bold;">${isMoneyIn ? 'Received From:' : 'Paid To:'}</td><td style="padding: 12px 0; font-weight: bold; text-align: right; border-bottom: 1px solid #e0e0e0; color: #1a1c1e; font-size: 16px;">${receipt.ledgerName}</td></tr>
                        <tr><td style="padding: 12px 0; color: #73777f; border-bottom: 1px solid #e0e0e0; text-transform: uppercase; font-size: 11px; font-weight: bold;">Amount:</td><td style="padding: 12px 0; font-weight: bold; text-align: right; border-bottom: 1px solid #e0e0e0; color: ${isMoneyIn ? '#146c2e' : '#ba1a1a'}; font-size: 22px;">\u20B9${parseFloat(receipt.amount).toFixed(2)}</td></tr>
                        <tr><td style="padding: 12px 0; color: #73777f; border-bottom: 1px solid #e0e0e0; text-transform: uppercase; font-size: 11px; font-weight: bold;">Mode:</td><td style="padding: 12px 0; font-weight: bold; text-align: right; border-bottom: 1px solid #e0e0e0; color: #1a1c1e;">${receipt.mode || 'Cash'} ${receipt.ref ? `(Ref: ${receipt.ref})` : ''}</td></tr>
                        ${invoiceRefDisplay ? `<tr><td style="padding: 12px 0; color: #73777f; border-bottom: none; text-transform: uppercase; font-size: 11px; font-weight: bold;">Settle Invoice(s):</td><td style="padding: 12px 0; font-weight: bold; text-align: right; border-bottom: none; color: #1a1c1e;">${invoiceRefDisplay}</td></tr>` : ''}
                    </table>
                </div>
                
                <div style="text-align: center; position: relative; z-index: 1;">${balanceText}</div>
                
                <div class="avoid-break" style="margin-top: 40px; display: flex; justify-content: flex-end; page-break-inside: avoid; position: relative; z-index: 1;">
                    <div style="width: 38%; text-align: center; border: 1px solid #e0e0e0; border-radius: 4px; padding: 15px; display: flex; flex-direction: column; justify-content: flex-end; align-items: center; min-height: 120px; background: #ffffff;">
                        ${biz.signature ? `<img src="${biz.signature}" style="max-height: 50px; margin-bottom: 15px; object-fit: contain;" />` : '<div style="height: 50px; margin-bottom: 15px;"></div>'}
                        <div style="border-top: 1px solid #1a1c1e; width: 90%; padding-top: 8px; font-weight: bold; font-size: 12px; color: #1a1c1e;">Authorized Signatory</div>
                        <div style="font-size: 10px; color: #73777f; margin-top: 4px; font-weight: bold;">For ${biz.name || 'Company'}</div>
                    </div>
                </div>
            </div>
        `;
        
        const printArea = document.getElementById('print-area');
        if (printArea) {
            printArea.innerHTML = html;
            setTimeout(() => {
                Utils.processPDFExport('pdf-receipt-wrapper', `${title.replace(/ /g, '_')}_${safeDocNo}.pdf`);
            }, 100);
        }
    },

    exportGSTCSV: async (reportData) => {
        try {
            let csvContent = "SOLLO ERP - MONTHLY GST REPORT\n";
            csvContent += `Month: ${reportData.month}\n\n`;

            csvContent += "GSTR-3B SUMMARY\n";
            csvContent += "Description,Taxable Value,Tax Amount\n";
            csvContent += `Total Sales (Output Tax),${reportData.gstr1.totalTaxable.toFixed(2)},${reportData.gstr1.totalTax.toFixed(2)}\n`;
            csvContent += `Total Purchases (Input Tax / ITC),${reportData.gstr2.totalTaxable.toFixed(2)},${reportData.gstr2.totalTax.toFixed(2)}\n`;
            csvContent += `NET GST PAYABLE TO GOVT,,${reportData.gstr3b.netPayable.toFixed(2)}\n\n`;

            csvContent += "GSTR-1 (SALES) BREAKDOWN\n";
            csvContent += "Category,Taxable Value,Tax Amount\n";
            csvContent += `B2B Sales (Registered),${reportData.gstr1.b2bTaxable.toFixed(2)},${reportData.gstr1.b2bTax.toFixed(2)}\n`;
            csvContent += `B2C Sales (Unregistered),${reportData.gstr1.b2cTaxable.toFixed(2)},${reportData.gstr1.b2cTax.toFixed(2)}\n\n`;

            csvContent += "DETAILED B2B SALES FOR CA PORTAL UPLOAD\n";
            csvContent += "Date,Invoice No,Customer Name,Customer GSTIN,Taxable Value,GST Amount,Total Invoice Value\n";
            
                        const ledgers = await window.getAllRecords('ledgers');
            
            reportData.rawSales.forEach(s => {
                // NEW: Skip Bill of Supply (Non-GST) from the detailed CA export
                if (s.invoiceType === 'Non-GST') return;

                let cust = ledgers.find(l => l.id === s.customerId);
                let gstin = cust ? cust.gst : '';
                if (gstin && gstin.trim() !== '') {
                    let taxable = parseFloat(s.subtotal) * (s.documentType === 'return' ? -1 : 1);
                    let tax = parseFloat(s.totalGst) * (s.documentType === 'return' ? -1 : 1);
                    let total = parseFloat(s.grandTotal) * (s.documentType === 'return' ? -1 : 1);
                    
                    // Fixed the CSV Quote Escaping Bug (replaces internal " with "")
                    const safeName = (s.customerName || '').replace(/"/g, '""');
                    csvContent += `${s.date},${s.invoiceNo},"${safeName}",${gstin.toUpperCase()},${taxable.toFixed(2)},${tax.toFixed(2)},${total.toFixed(2)}\n`;
                }
            });

            const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
            const file = new File([blob], `GST_Report_${new Date().toISOString().split('T')[0]}.csv`, { type: "text/csv" });

            if (navigator.canShare && navigator.canShare({ files: [file] })) {
                await navigator.share({
                    title: "SOLLO GST Report",
                    text: "Here is your GST Report for the accountant.",
                    files: [file]
                });
            } else {
                const url = URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.style.display = "none";
                a.href = url;
                a.download = file.name;
                document.body.appendChild(a);
                a.click();
                URL.revokeObjectURL(url);
                document.body.removeChild(a);
                alert("✅ GST Report downloaded successfully!");
            }

        } catch (e) {
            console.error(e);
            alert("Export failed.");
        }
    }
};

// ==========================================
// NEW CODE: ES MODULE EXPORT & GLOBAL MAP
// ==========================================
// 1. Export the module so app.js can import it safely
export default Utils; 

// 2. Attach to window so index.html onclick="Utils..." buttons don't break!
window.Utils = Utils; 
