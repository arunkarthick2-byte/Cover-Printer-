// ==========================================
// SOLLO ERP - UI & ANIMATION CONTROLLER (v5.2 Enterprise)
// ==========================================

const UI = {

    // --- NEW CODE: DARK MODE CONTROLLERS ---
    initTheme: () => {
        const isDark = localStorage.getItem('sollo_theme') === 'dark';
        if (isDark) {
            document.body.classList.add('dark-mode');
            const metaTheme = document.getElementById('meta-theme-color');
            if (metaTheme) metaTheme.setAttribute('content', '#111315'); // Matches dark background
        }
    },

    toggleTheme: () => {
        document.body.classList.toggle('dark-mode');
        const isDark = document.body.classList.contains('dark-mode');
        localStorage.setItem('sollo_theme', isDark ? 'dark' : 'light');
        UI.resetStatusBarColor();
        if (window.Utils) window.Utils.showToast(isDark ? "Dark Mode Enabled 🌙" : "Light Mode Enabled ☀️");
    },

    resetStatusBarColor: () => {
        const isDark = document.body.classList.contains('dark-mode');
        const metaTheme = document.getElementById('meta-theme-color');
        if (metaTheme) metaTheme.setAttribute('content', isDark ? '#111315' : '#0061a4');
    },

    setStatusBarColor: (color) => {
        const metaTheme = document.getElementById('meta-theme-color');
        if (metaTheme) metaTheme.setAttribute('content', color);
    },

    safeUpdateDOM: (container, htmlContent) => {
        if (!container) return;
        if (document.startViewTransition) {
            try { document.startViewTransition(() => { container.innerHTML = htmlContent; }); } 
            catch (e) { container.innerHTML = htmlContent; }
        } else {
            container.innerHTML = htmlContent;
        }
    },
    // --- END OF NEW CODE ---
    
    // ==========================================
    // LIVE BANK BALANCE CALCULATOR
    // ==========================================
    renderBankBalances: () => {
        // FIXED: Pull from state to respect firmId filtering, preventing data leaks across multiple companies
        const accounts = UI.state.rawData.accounts || [];
        const receipts = UI.state.rawData.cashbook || [];
        const container = document.getElementById('bank-balances-container');
        if (!container) return;

        // Use a shallow copy so we don't accidentally mutate the master state
        const displayAccounts = [...accounts];
        if (!displayAccounts.find(a => a.id === 'cash')) {
            displayAccounts.unshift({ id: 'cash', name: 'Cash Drawer', openingBalance: 0 });
        }

        let html = '';
        displayAccounts.forEach(acc => {
            let balance = parseFloat(acc.openingBalance) || 0;
            receipts.forEach(r => {
                // Included a fallback logic in case older receipts don't have an explicit accountId
                if (r.accountId === acc.id || (acc.id === 'cash' && !r.accountId)) {
                    if (r.type === 'in') balance += parseFloat(r.amount);
                    else if (r.type === 'out') balance -= parseFloat(r.amount);
                }
            });
            
            const color = balance >= 0 ? 'var(--md-success)' : 'var(--md-error)';
            
            // NEW: Clicking a bank balance card now opens its complete pin-to-pin passbook statement!
            const clickAction = `app.openAccountLedger('${acc.id}')`;

            // UPGRADE: Vertical premium card layout for accounts
            const icon = acc.id === 'cash' ? 'payments' : 'account_balance';
            
            html += `
                <div class="m3-card tap-target" onclick="${clickAction}" style="display: flex; align-items: center; gap: 16px; padding: 16px; margin-bottom: 0;">
                    <div class="icon-circle" style="background: var(--md-surface-variant); color: var(--md-on-surface-variant); width: 48px; height: 48px; flex-shrink: 0;">
                        <span class="material-symbols-outlined">${icon}</span>
                    </div>
                    <div style="flex: 1; min-width: 0; overflow: hidden;">
                        <strong class="large-text" style="display: block; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${acc.name}</strong>
                        <small style="color: var(--md-text-muted);">View Statement</small>
                    </div>
                    <div style="text-align: right; flex-shrink: 0;">
                        <strong style="font-size: 16px; color: ${color};">&#8377;${balance.toFixed(2)}</strong>
                    </div>
                </div>
            `;
        });
        container.innerHTML = html;
    },

    state: { 
        activeActivity: null, 
        selectedProducts: [], 
        activeMasterTab: 'products',
        currentMasterType: 'products',
        currentMasterTitle: 'Inventory Master',
        rawData: { sales: [], purchases: [], items: [], ledgers: [], expenses: [], cashbook: [], timeline: [], accounts: [] },
        activeFilters: { sales: 'All', purchases: 'All', masters: 'All', expenses: 'All', cashbook: 'All', timeline: 'All' }
    },

    // ==========================================
    // 1. SPLASH SCREEN & INSTANT NAVIGATION
    // ==========================================
    showSuccess: () => {
        const el = document.getElementById('success-animation');
        if(el) {
            el.classList.remove('hidden');
            const circle = el.querySelector('.success-circle');
            if(circle) { circle.style.animation = 'none'; circle.offsetHeight; circle.style.animation = null; } // Reset animation
            if (window.navigator && window.navigator.vibrate) window.navigator.vibrate([30, 50, 30]); // Haptic buzz
            setTimeout(() => el.classList.add('hidden'), 1200);
        }
    },

    hideSplash: () => {
        const splash = document.getElementById('splash-screen');
        if(splash) {
            splash.classList.add('fade-out');
            setTimeout(() => splash.classList.add('hidden'), 500); 
        }

        // UPGRADE 4: Read Home Screen Shortcuts on Boot
        setTimeout(() => {
            const urlParams = new URLSearchParams(window.location.search);
            const action = urlParams.get('action');
            if (action === 'new_sale' && typeof app !== 'undefined') {
                app.openForm('sales', null, 'invoice');
            } else if (action === 'cashbook') {
                UI.switchTab('tab-cashbook', 'Cashbook & Banking');
            }
        }, 600); // Wait for the splash screen to finish fading before opening
    },

    toggleDrawer: () => {
        const drawer = document.getElementById('side-drawer');
        const overlay = document.getElementById('side-drawer-overlay');
        const isOpening = !drawer.classList.contains('open');
        
        drawer.classList.toggle('open');
        
        if (isOpening) {
            overlay.classList.remove('hidden');
            requestAnimationFrame(() => overlay.classList.add('open'));
            history.pushState({ modalOpen: true }, '');
        } else {
            overlay.classList.remove('open');
            setTimeout(() => overlay.classList.add('hidden'), 300);
            
            if (history.state && history.state.modalOpen) {
                window._ignoreNextPop = true;
                history.back();
            }
        }
    },

    switchTab: (tabId, title, navElement) => {
        const doSwitch = () => {
            document.querySelectorAll('.screen-section').forEach(el => {
                el.classList.remove('active-screen');
                el.classList.add('hidden'); 
            });

            const titleEl = document.getElementById('screen-title');
            if (titleEl) titleEl.innerText = title;

            const target = document.getElementById(tabId);
            if(target) {
                target.classList.remove('hidden');
                target.classList.add('active-screen');
            }

            if (navElement) {
                document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
                navElement.classList.add('active');
            }

            if (tabId === 'tab-dashboard') UI.renderDashboard();
            else if (tabId === 'tab-documents') { UI.applyFilters('sales'); UI.applyFilters('purchases'); }
            else if (tabId === 'tab-cashbook') UI.applyFilters('cashbook');
            else if (tabId === 'tab-expenses') UI.applyFilters('expenses');
            else if (tabId === 'tab-masters') UI.applyFilters('masters');
        };

        // UPGRADE 4: Cinematic View Transitions
        if (document.startViewTransition) {
            document.startViewTransition(doSwitch);
        } else {
            doSwitch();
        }
    },

    handleFabClick: () => UI.openBottomSheet('sheet-fab-menu'),

    openActivity: (activityId) => {
        const a = document.getElementById(activityId);
        if(a) {
            history.pushState({ modalOpen: true }, ''); // Native Back Gesture hook
            a.classList.remove('hidden'); 
            requestAnimationFrame(() => { a.classList.add('open'); });
            
            if (activityId === 'activity-sales-form') {
                UI.state.activeActivity = 'sales';
            } else if (activityId === 'activity-purchase-form') {
                UI.state.activeActivity = 'purchase';
            }
        }
    },

    closeActivity: (activityId) => {
        const a = document.getElementById(activityId);
        if(a) {
            a.classList.remove('open'); 
            setTimeout(() => a.classList.add('hidden'), 300); 
            
            if (activityId === 'activity-sales-form' || activityId === 'activity-purchase-form') {
                UI.state.activeActivity = null;
            }
            
            // FIXED: Added an ignore flag to prevent double-closing
            if (history.state && history.state.modalOpen) { window._ignoreNextPop = true; history.back(); } 
        }
    },

    toggleDeleteButton: (type, show) => {
        const btnMap = {
            'sales': 'btn-delete-sales', 'purchase': 'btn-delete-purchase',
            'product': 'btn-delete-product', 'ledger': 'btn-delete-ledger',
            'expense': 'btn-delete-expense', 'receipt-in': 'btn-delete-receipt-in',
            'receipt-out': 'btn-delete-receipt-out', 'account': 'btn-delete-account'
        };
        const btnId = btnMap[type];
        if (!btnId) return;

        const btn = document.getElementById(btnId);
        if (btn) {
            if (show) btn.classList.remove('hidden');
            else btn.classList.add('hidden');
        }
    },

    // ==========================================
    // 2. MASTER VIEW NAVIGATION
    // ==========================================
    openMasterView: async (type, title) => {
        // CRITICAL FIX: Tracking the state correctly so search & filters remember what list you are looking at
        UI.state.activeMasterTab = type;
        UI.state.currentMasterType = type;
        UI.state.currentMasterTitle = title;
        
        document.getElementById('master-view-title').innerText = title;
        UI.openActivity('activity-master-view');
        
        const container = document.getElementById('master-list-container');
        // UPGRADE 3: Premium Shimmer Loaders instead of the old circular spinner
        container.innerHTML = `
            <div style="padding: 8px 0;">
                <div class="skeleton-card"></div>
                <div class="skeleton-card" style="animation-delay: 0.1s"></div>
                <div class="skeleton-card" style="animation-delay: 0.2s"></div>
                <div class="skeleton-card" style="animation-delay: 0.3s"></div>
                <div class="skeleton-card" style="animation-delay: 0.4s"></div>
            </div>
        `;
        
        // Handle Stock Adjustments Audit Trail
        if (type === 'adjustments') {
            // FIX: Hide the search bar so users don't trigger the crash
            const searchInput = document.getElementById('search-master-view');
            if (searchInput && searchInput.parentElement) searchInput.parentElement.style.display = 'none';

            const records = await getAllRecords('adjustments');
            const products = await getAllRecords('items');
            if (records.length === 0) {
                container.innerHTML = '<p class="empty-state">No stock adjustments logged yet.</p>';
                const actionBtn = document.getElementById('btn-master-action');
                if (actionBtn) actionBtn.classList.add('hidden');
                return;
            }
            let html = '<ul class="list-view">';
            // BULLETPROOF SORTING: Prevents WebView crash in Stock Adjustments!
            records.sort((a,b) => String(b.date || '').localeCompare(String(a.date || ''))).forEach(adj => {
                const prod = products.find(p => p.id === adj.itemId);
                const prodName = prod ? prod.name : 'Deleted Product';
                const sign = adj.type === 'add' ? '+' : '-';
                const color = adj.type === 'add' ? 'var(--md-success)' : 'var(--md-error)';
                html += `
                    <li style="flex-direction: column; align-items: flex-start; cursor: default;">
                        <div style="display: flex; justify-content: space-between; width: 100%;">
                            <strong>${prodName}</strong>
                            <strong style="color: ${color};">${sign}${parseFloat(adj.qty).toFixed(2)}</strong>
                        </div>
                        <div style="display: flex; justify-content: space-between; width: 100%; margin-top: 4px; font-size: 12px; color: var(--md-text-muted);">
                            <span>${adj.date}</span>
                            <span>${adj.notes || 'No Reason Provided'}</span>
                        </div>
                    </li>
                `;
            });
            container.innerHTML = html + '</ul>';
            const actionBtn = document.getElementById('btn-master-action');
            if (actionBtn) actionBtn.classList.add('hidden');
            return; 
        }

        const searchInput = document.getElementById('search-master-view');
        if (searchInput) {
            // FIX: Restore the search bar for Products, Customers, etc.
            if (searchInput.parentElement) searchInput.parentElement.style.display = 'flex'; 
            searchInput.value = '';
            searchInput.placeholder = `Search ${title}...`;
        }
        
        const actionBtn = document.getElementById('btn-master-action');
        if (actionBtn) {
            actionBtn.classList.add('hidden'); // Removed clunky Receivables button from UI
        }

        // UPGRADE: Show 3-Dot menu only for Products, Customers, and Suppliers
        const moreOptionsBtn = document.getElementById('master-more-options');
        if (moreOptionsBtn) {
            if (['products', 'customers', 'suppliers'].includes(type)) moreOptionsBtn.classList.remove('hidden');
            else moreOptionsBtn.classList.add('hidden');
        }

        const filterSelect = document.getElementById('filter-master-view');
        if (filterSelect) filterSelect.value = 'All';

        UI.applyFilters('masters'); 
    },

    renderRowWiseItem: (title, subtitle, rightText, rightSub, icon, iconColor, onClickAction) => {
        return `
        <div class="m3-card tap-target" style="padding: 12px; margin-bottom: 8px; border-radius: 8px; display: flex; align-items: center; gap: 12px; box-shadow: 0 1px 2px rgba(0,0,0,0.1);" onclick="${onClickAction}">
            <div class="icon-circle" style="width: 40px; height: 40px; background: var(--md-surface-variant); color: ${iconColor}; border-radius: 50%; display: flex; justify-content: center; align-items: center; flex-shrink: 0;">
                <span class="material-symbols-outlined" style="font-size: 20px;">${icon}</span>
            </div>
            <div style="flex: 1; min-width: 0; overflow: hidden;">
                <strong style="font-size: 14px; color: var(--md-on-surface); display: block; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${title}</strong>
                <small style="color: var(--md-text-muted); display: block; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${subtitle}</small>
            </div>
            <div style="text-align: right; flex-shrink: 0;">
                <strong style="font-size: 14px; color: var(--md-on-surface);">${rightText}</strong><br>
                <small style="color: var(--md-text-muted);">${rightSub}</small>
            </div>
        </div>`;
    },

    // ==========================================
    // 3. STRICT MATH, DATES & TAX DISTRIBUTION
    // ==========================================
    toggleDates: (type) => {
        const statusEl = document.getElementById(`${type}-order-status`);
        if(!statusEl) return;
        const status = statusEl.value;
        const shippedGroup = document.getElementById(`${type}-shipped-date-group`);
        const completedGroup = document.getElementById(`${type}-completed-date-group`);
        const shippedInput = document.getElementById(`${type}-shipped-date`);
        
        // Reset visibility and editability by default
        if (shippedGroup) shippedGroup.classList.add('hidden');
        if (completedGroup) completedGroup.classList.add('hidden');
        if (shippedInput) shippedInput.disabled = false;
        
        if (status === 'Shipped') {
            if (shippedGroup) shippedGroup.classList.remove('hidden');
        } else if (status === 'Completed') {
            // Keep the shipped date visible to retain history, but lock it
            if (shippedGroup) {
                shippedGroup.classList.remove('hidden');
                if (shippedInput) shippedInput.disabled = true;
            }
            if (completedGroup) completedGroup.classList.remove('hidden');
        }
    },

    calcSalesTotals: () => {
        let rawSubtotal = 0;
        const typeEl = document.getElementById('sales-invoice-type');
        const isGST = typeEl ? typeEl.value !== 'Non-GST' : true;

        const rows = document.querySelectorAll('#sales-items-body tr');
        
        rows.forEach(tr => {
            const qty = parseFloat(tr.querySelector('.row-qty').value) || 0;
            const rate = parseFloat(tr.querySelector('.row-rate').value) || 0;
            rawSubtotal += (qty * rate);
        });

        const discountInput = parseFloat(document.getElementById('sales-discount').value) || 0;
        const discountTypeEl = document.getElementById('sales-discount-type');
        const discountType = discountTypeEl ? discountTypeEl.value : '\u20B9';
        let discountAmt = discountType === '%' ? (rawSubtotal * (discountInput / 100)) : discountInput;
        
        // FIX: Prevent negative invoices by capping the discount at the subtotal maximum
        if (discountAmt > rawSubtotal) discountAmt = rawSubtotal;
        
        const discountRatio = rawSubtotal > 0 ? (discountAmt / rawSubtotal) : 0;
        
        let finalSubtotal = 0;
        let totalGst = 0;

        rows.forEach(tr => {
            const qty = parseFloat(tr.querySelector('.row-qty').value) || 0;
            const rate = parseFloat(tr.querySelector('.row-rate').value) || 0;
            const gstPercent = isGST ? (parseFloat(tr.querySelector('.row-gst').value) || 0) : 0;
            
            const baseAmount = qty * rate;
            const discountedBase = baseAmount - (baseAmount * discountRatio);
            const gstAmount = discountedBase * (gstPercent / 100);
            
            const roundedDiscountedBase = Math.round(discountedBase * 100) / 100;
            const roundedGst = Math.round(gstAmount * 100) / 100;
            const rowTotal = roundedDiscountedBase + roundedGst;
            
            tr.querySelector('.row-total').innerText = rowTotal.toFixed(2);
            
            finalSubtotal += roundedDiscountedBase;
            totalGst += roundedGst;

            const buyPrice = parseFloat(tr.querySelector('.row-item-buyprice').value) || 0;
            const marginSpan = tr.querySelector('.live-margin');
            if (marginSpan) {
                const effectiveRate = rate - (rate * discountRatio); 
                const profitAmt = effectiveRate - buyPrice;
                const marginPercent = effectiveRate > 0 ? ((profitAmt / effectiveRate) * 100).toFixed(1) : 0;
                // Upgrade: Shows the exact profit amount per unit based on the custom editable buy price
                marginSpan.innerText = `Margin: ${marginPercent}% (\u20B9${profitAmt.toFixed(2)}/unit)`;
                marginSpan.style.color = profitAmt < 0 ? 'var(--md-error)' : 'var(--md-success)';
            }
        });

        const freight = parseFloat(document.getElementById('sales-freight').value) || 0;
        const grandTotal = finalSubtotal + totalGst + freight;

        document.getElementById('sales-subtotal').innerText = `\u20B9${finalSubtotal.toFixed(2)}`;
        document.getElementById('sales-gst-total').innerText = `\u20B9${totalGst.toFixed(2)}`;
        document.getElementById('sales-grand-total').innerText = `\u20B9${grandTotal.toFixed(2)}`;
    },

    calcPurchaseTotals: () => {
        let rawSubtotal = 0;
        const typeEl = document.getElementById('purchase-invoice-type');
        const isGST = typeEl ? typeEl.value !== 'Non-GST' : true;

        const rows = document.querySelectorAll('#purchase-items-body tr');
        
        rows.forEach(tr => {
            const qty = parseFloat(tr.querySelector('.row-qty').value) || 0;
            const rate = parseFloat(tr.querySelector('.row-rate').value) || 0;
            rawSubtotal += (qty * rate);
        });

        const discountInput = parseFloat(document.getElementById('purchase-discount').value) || 0;
        const discountTypeEl = document.getElementById('purchase-discount-type');
        const discountType = discountTypeEl ? discountTypeEl.value : '\u20B9';
        let discountAmt = discountType === '%' ? (rawSubtotal * (discountInput / 100)) : discountInput;
        
        // FIX: Prevent negative invoices by capping the discount at the subtotal maximum
        if (discountAmt > rawSubtotal) discountAmt = rawSubtotal;
        
        const discountRatio = rawSubtotal > 0 ? (discountAmt / rawSubtotal) : 0;

        let finalSubtotal = 0;
        let totalGst = 0;

        rows.forEach(tr => {
            const qty = parseFloat(tr.querySelector('.row-qty').value) || 0;
            const rate = parseFloat(tr.querySelector('.row-rate').value) || 0;
            const gstPercent = isGST ? (parseFloat(tr.querySelector('.row-gst').value) || 0) : 0;
            
            const baseAmount = qty * rate;
            const discountedBase = baseAmount - (baseAmount * discountRatio);
            const gstAmount = discountedBase * (gstPercent / 100);
            
            const roundedDiscountedBase = Math.round(discountedBase * 100) / 100;
            const roundedGst = Math.round(gstAmount * 100) / 100;
            const rowTotal = roundedDiscountedBase + roundedGst;
            
            tr.querySelector('.row-total').innerText = rowTotal.toFixed(2);
            
            finalSubtotal += roundedDiscountedBase;
            totalGst += roundedGst;
        });

        const freight = parseFloat(document.getElementById('purchase-freight').value) || 0;
        const grandTotal = finalSubtotal + totalGst + freight;

        document.getElementById('purchase-subtotal').innerText = `\u20B9${finalSubtotal.toFixed(2)}`;
        document.getElementById('purchase-gst-total').innerText = `\u20B9${totalGst.toFixed(2)}`;
        document.getElementById('purchase-grand-total').innerText = `\u20B9${grandTotal.toFixed(2)}`;
    },

    // ==========================================
    // 4. UNIVERSAL SEARCH & DYNAMIC FILTERS
    // ==========================================
    highlightText: (text, term) => {
        if (!term || !text) return text;
        const regex = new RegExp(`(${term.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&')})`, 'gi');
        return String(text).replace(regex, '<span style="background: rgba(0,97,164,0.15); color: var(--md-primary); border-radius: 3px; font-weight: bold; padding: 0 2px;">$1</span>');
    },

    setFilter: (tab, filterValue, chipElement) => {
        // FIX: Safely ignore the color-changing logic if we are using the new Bottom Sheet menus (which send a null chipElement)
        if (chipElement) {
            const chips = chipElement.parentElement.querySelectorAll('.chip');
            chips.forEach(c => c.classList.remove('active'));
            chipElement.classList.add('active');
        }

        UI.state.activeFilters[tab] = filterValue;
        UI.applyFilters(tab);
    },

    applyFilters: (tab) => {
        const searchInput = tab === 'masters' ? document.getElementById('search-master-view') : document.getElementById(`search-${tab}`);
        const searchTerm = searchInput ? (searchInput.value || '').toLowerCase() : '';
        
        const sortSelect = tab === 'masters' ? document.getElementById('sort-master-view') : document.getElementById(`sort-${tab}`);
        const sortOption = sortSelect ? sortSelect.value : null;
        
        const activeFilter = UI.state.activeFilters[tab];
        
        let data = [];
        let containerId = '';
        let sumValue = 0;

        // Cashbook Map logic for True Balances (WITH REFUND & COLLISION FIX)
        const paymentMap = {};
        if (tab === 'sales' || tab === 'purchases') {
            UI.state.rawData.cashbook.forEach(c => {
                if (c.invoiceRef && c.ledgerId) {
                    let amt = parseFloat(c.amount) || 0;
                    if (tab === 'sales') amt = c.type === 'in' ? amt : -amt;
                    if (tab === 'purchases') amt = c.type === 'out' ? amt : -amt;
                    
                    // NEW: Split the payment map logic across all referenced invoices
                    const refs = String(c.invoiceRef).split(',').map(r => r.trim());
                    let splitAmt = amt / (refs.length || 1);
                    refs.forEach(r => {
                        const key = `${c.ledgerId}_${r}`;
                        paymentMap[key] = (paymentMap[key] || 0) + splitAmt;
                    });
                }
            });
        }

        // ------------------ SALES ------------------
        if (tab === 'sales') {
            containerId = 'sales-history-container';
            data = UI.state.rawData.sales.filter(s => {
                const matchSearch = (s.customerName || '').toLowerCase().includes(searchTerm) || (s.invoiceNo || '').toLowerCase().includes(searchTerm);
                let matchFilter = true;
                
                // FIX: Check ALL references to catch cross-linked payments, and respect FIFO completion!
                const uniqueRefs = [...new Set([s.orderNo, s.invoiceNo, s.id].filter(Boolean))];
                const paid = uniqueRefs.reduce((sum, ref) => sum + (paymentMap[`${s.customerId}_${ref}`] || 0), 0);
                const balance = Math.max(0, (parseFloat(s.grandTotal) || 0) - paid);
                
                if (activeFilter === 'Open') matchFilter = s.status === 'Open';
                else if (activeFilter === 'Completed') matchFilter = s.status === 'Completed'; // Removed balance restriction
                else if (activeFilter === 'Shipped') matchFilter = s.status === 'Shipped';
                else if (activeFilter === 'To Receive') matchFilter = balance > 0 && s.status !== 'Open' && s.documentType !== 'return';
                else if (activeFilter === 'GST') matchFilter = s.invoiceType !== 'Non-GST';
                else if (activeFilter === 'Non-GST') matchFilter = s.invoiceType === 'Non-GST';
                
                return matchSearch && matchFilter;
            });

            data.forEach(s => sumValue += (s.documentType === 'return' ? -(s.grandTotal || 0) : (s.grandTotal || 0)));
            if (document.getElementById('sum-sales')) document.getElementById('sum-sales').innerText = `\u20B9${sumValue.toFixed(2)}`;

            // BULLETPROOF SORTING: Prevents WebView crashes!
            if(sortOption === 'date-desc') data.sort((a,b) => new Date(b.date || 0) - new Date(a.date || 0));
            if(sortOption === 'date-asc') data.sort((a,b) => new Date(a.date || 0) - new Date(b.date || 0));
            if(sortOption === 'amt-desc') data.sort((a,b) => (parseFloat(b.grandTotal) || 0) - (parseFloat(a.grandTotal) || 0));
            if(sortOption === 'amt-asc') data.sort((a,b) => (parseFloat(a.grandTotal) || 0) - (parseFloat(b.grandTotal) || 0));

            const container = document.getElementById(containerId);
            if (container) {
                const newHTML = data.length ? data.map(s => {
                    const isReturn = s.documentType === 'return';
                    // FIX: Update the visual display to match the new cross-linked math!
                    const uniqueRefs = [...new Set([s.orderNo, s.invoiceNo, s.id].filter(Boolean))];
                    const paid = uniqueRefs.reduce((sum, ref) => sum + (paymentMap[`${s.customerId}_${ref}`] || 0), 0);
                    const balance = Math.max(0, (parseFloat(s.grandTotal) || 0) - paid);
                    const statusText = s.status === 'Open' ? 'Draft' : (balance > 0 && !isReturn ? `Due: \u20B9${balance.toFixed(2)}` : 'Paid');
                    const statusColor = s.status === 'Open' ? 'var(--md-text-muted)' : (balance > 0 && !isReturn ? 'var(--md-error)' : 'var(--md-success)');

                    return `
                    <div class="m3-card tap-target" style="${isReturn ? 'border-left: 4px solid var(--md-error);' : ''}" onclick="app.openForm('sales', '${s.id}', '${s.documentType}')">
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <div>
                                <div class="large-text">${s.customerName || 'Unknown Party'} ${isReturn ? '<span style="color:var(--md-error); font-size:12px;">(Credit Note)</span>' : ''}</div>
                                <small class="color-primary">${s.invoiceNo || s.orderNo || 'Draft'} | ${s.date || 'Unknown Date'}</small>
                            </div>
                            <div style="text-align:right;">
                                <strong style="font-size:16px; color:${isReturn ? 'var(--md-error)' : 'inherit'};">${isReturn ? '-' : ''}\u20B9${(s.grandTotal || 0).toFixed(2)}</strong><br>
                                <small style="background:var(--md-surface-variant); color:${statusColor}; font-weight:bold; padding:2px 6px; border-radius:4px;">${statusText}</small>
                            </div>
                        </div>
                    </div>`;
                }).join('') : `
                <div class="empty-state">
                    <svg width="100" height="100" viewBox="0 0 24 24" fill="none" stroke="var(--md-outline-variant)" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round" style="margin-bottom:16px;">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                        <polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline>
                    </svg>
                    <p style="margin: 0 0 16px 0; font-size: 16px;">No sales invoices match your filters.</p>
                    <button class="btn-primary" onclick="app.openForm('sales', null, 'invoice')">+ Create Sales Invoice</button>
                </div>`;
                UI.safeUpdateDOM(container, newHTML);
            }
        }

        // ------------------ PURCHASES ------------------
        else if (tab === 'purchases') {
            containerId = 'purchase-history-container';
            data = UI.state.rawData.purchases.filter(p => {
                const matchSearch = (p.supplierName || '').toLowerCase().includes(searchTerm) || (p.invoiceNo || p.poNo || '').toLowerCase().includes(searchTerm);
                let matchFilter = true;

                // FIX: Check ALL references to catch cross-linked payments, and respect FIFO completion!
                const uniqueRefs = [...new Set([p.orderNo, p.poNo, p.invoiceNo, p.id].filter(Boolean))];
                const paid = uniqueRefs.reduce((sum, ref) => sum + (paymentMap[`${p.supplierId}_${ref}`] || 0), 0);
                const balance = Math.max(0, (parseFloat(p.grandTotal) || 0) - paid);

                if (activeFilter === 'To Pay') matchFilter = balance > 0 && p.status !== 'Open' && p.documentType !== 'return';
                else if (activeFilter === 'Completed') matchFilter = p.status === 'Completed'; // Removed balance restriction
                else if (activeFilter === 'GST') matchFilter = p.invoiceType !== 'Non-GST';
                else if (activeFilter === 'Non-GST') matchFilter = p.invoiceType === 'Non-GST';

                return matchSearch && matchFilter;
            });

            data.forEach(p => sumValue += (p.documentType === 'return' ? -(p.grandTotal || 0) : (p.grandTotal || 0)));
            if (document.getElementById('sum-purchases')) document.getElementById('sum-purchases').innerText = `\u20B9${sumValue.toFixed(2)}`;

            // BULLETPROOF SORTING: Prevents WebView crashes!
            if(sortOption === 'date-desc') data.sort((a,b) => new Date(b.date || 0) - new Date(a.date || 0));
            if(sortOption === 'date-asc') data.sort((a,b) => new Date(a.date || 0) - new Date(b.date || 0));
            if(sortOption === 'amt-desc') data.sort((a,b) => (parseFloat(b.grandTotal) || 0) - (parseFloat(a.grandTotal) || 0));

            const container = document.getElementById(containerId);
            if (container) {
                const newHTML = data.length ? data.map(p => {
                    const isReturn = p.documentType === 'return';
                    // FIX: Update the visual display to match the new cross-linked math!
                    const uniqueRefs = [...new Set([p.orderNo, p.poNo, p.invoiceNo, p.id].filter(Boolean))];
                    const paid = uniqueRefs.reduce((sum, ref) => sum + (paymentMap[`${p.supplierId}_${ref}`] || 0), 0);
                    const balance = Math.max(0, (parseFloat(p.grandTotal) || 0) - paid);
                    const statusText = p.status === 'Open' ? 'Draft PO' : (balance > 0 && !isReturn ? `To Pay: \u20B9${balance.toFixed(2)}` : 'Paid');
                    const statusColor = p.status === 'Open' ? 'var(--md-text-muted)' : (balance > 0 && !isReturn ? 'var(--md-error)' : 'var(--md-success)');

                    return `
                    <div class="m3-card tap-target" style="${isReturn ? 'border-left: 4px solid var(--md-error);' : ''}" onclick="app.openForm('purchase', '${p.id}', '${p.documentType}')">
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <div>
                                <div class="large-text">${p.supplierName || 'Unknown Party'} ${isReturn ? '<span style="color:var(--md-error); font-size:12px;">(Debit Note)</span>' : ''}</div>
                                <small class="color-primary">${p.orderNo || p.poNo || p.invoiceNo || 'Draft'} | ${p.date || 'Unknown Date'}</small>
                            </div>
                            <div style="text-align:right;">
                                <strong style="font-size:16px; color:${isReturn ? 'var(--md-success)' : 'inherit'};">${isReturn ? '-' : ''}\u20B9${(p.grandTotal || 0).toFixed(2)}</strong><br>
                                ${p.status === 'Open' ? `<button class="btn-primary-small mt-2" onclick="event.stopPropagation(); app.convertPO('${p.id}')">Complete PI</button>` : `<small style="background:var(--md-surface-variant); color:${statusColor}; font-weight:bold; padding:2px 6px; border-radius:4px;">${statusText}</small>`}
                            </div>
                        </div>
                    </div>`;
                }).join('') : `
                <div class="empty-state">
                    <svg width="100" height="100" viewBox="0 0 24 24" fill="none" stroke="var(--md-outline-variant)" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round" style="margin-bottom:16px;">
                        <path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"></path>
                        <polyline points="3.27 6.96 12 12.01 20.73 6.96"></polyline><line x1="12" y1="22.08" x2="12" y2="12"></line>
                    </svg>
                    <p style="margin: 0 0 16px 0; font-size: 16px;">No purchase bills match your filters.</p>
                    <button class="btn-primary" onclick="app.openForm('purchase', null, 'invoice')">+ Create Purchase Bill</button>
                </div>`;
                UI.safeUpdateDOM(container, newHTML);
            }
        }

        // ------------------ ROW-WISE MASTERS ------------------
        else if (tab === 'masters') {
            containerId = 'master-list-container';
            const activeTab = UI.state.activeMasterTab || 'products';
            const filterSelect = document.getElementById('filter-master-view');
            const activeMasterFilter = filterSelect ? filterSelect.value : 'All';
            const html = [];

            const getBal = (id, type) => {
                let bal = 0;
                const isCustomer = String(type).toLowerCase() === 'customer';
                const ledger = UI.state.rawData.ledgers.find(l => l.id === id);
                
                if (ledger) {
                    let ob = parseFloat(ledger.openingBalance) || 0;
                    const balType = (ledger.balanceType || '').toLowerCase();
                    // FIX: Bulletproof string matching ignores spaces and case formatting
                    if (isCustomer) bal += (balType.includes('pay') || balType.includes('credit')) ? -ob : ob;
                    else bal += (balType.includes('receive') || balType.includes('debit')) ? -ob : ob;
                }
                
                if (isCustomer) {
                    UI.state.rawData.sales.forEach(s => { 
                        if (s.customerId === id && s.status !== 'Open') {
                            bal += (s.documentType === 'return' ? -parseFloat(s.grandTotal || 0) : parseFloat(s.grandTotal || 0)); 
                        }
                    });
                    UI.state.rawData.cashbook.forEach(c => { 
                        if (c.ledgerId === id) {
                            const amt = parseFloat(c.amount) || 0;
                            bal += (c.type === 'in' ? -amt : amt);
                        }
                    });
                } else {
                    UI.state.rawData.purchases.forEach(p => { 
                        if (p.supplierId === id && p.status !== 'Open') {
                            bal += (p.documentType === 'return' ? -parseFloat(p.grandTotal || 0) : parseFloat(p.grandTotal || 0)); 
                        }
                    });
                    UI.state.rawData.cashbook.forEach(c => { 
                        if (c.ledgerId === id) {
                            const amt = parseFloat(c.amount) || 0;
                            bal += (c.type === 'out' ? -amt : amt); 
                        }
                    });
                }
                return bal;
            };

            if (activeTab === 'products') {
                data = UI.state.rawData.items.filter(i => {
                    const matchSearch = (i.name || '').toLowerCase().includes(searchTerm) || (i.sku || '').toLowerCase().includes(searchTerm);
                    let matchFilter = true;
                    // NEW: Check if stock is greater than 0
                    if (activeMasterFilter === 'In Stock') {
                        matchFilter = (parseFloat(i.stock) || 0) > 0;
                    }
                    return matchSearch && matchFilter;
                });
                
                if(sortOption === 'name-asc') data.sort((a,b) => (a.name || '').localeCompare(b.name || ''));
                if(sortOption === 'stock-asc') data.sort((a,b) => (a.stock || 0) - (b.stock || 0));

                html.push(...data.map(i => {
                    const currentStock = parseFloat(i.stock) || 0;
                    const minStock = parseFloat(i.minStock) || 0;
                    // Check if stock is at or below the minimum threshold (and ensure minStock is actually set)
                    const isLowStock = minStock > 0 && currentStock <= minStock;
                    
                    const stockLabel = isLowStock 
                        ? `<span style="color:var(--md-error); font-weight:bold;">Stock: ${currentStock} ${i.uom || ''} ⚠️ Low</span>` 
                        : `Stock: ${currentStock} ${i.uom || ''}`;

                    return UI.renderRowWiseItem(
                        UI.highlightText(i.name || 'Unnamed Product', searchTerm), 
                        stockLabel, 
                        `\u20B9${(i.sellPrice || 0).toFixed(2)}`, 
                        `Buy: \u20B9${(i.buyPrice || 0).toFixed(2)}`, 
                        'inventory_2', 
                        isLowStock ? 'var(--md-error)' : 'var(--md-primary)', // Changes icon color to red if low
                        `app.openForm('product', '${i.id}')`
                    );
                }));
            } 
            else if (activeTab === 'customers' || activeTab === 'suppliers' || activeTab === 'contacts') {
                const typeFilter = activeTab === 'customers' ? 'Customer' : (activeTab === 'suppliers' ? 'Supplier' : 'All');
                const icon = activeTab === 'customers' ? 'person' : 'storefront';
                const color = activeTab === 'customers' ? '#0061a4' : '#ba1a1a';

                data = UI.state.rawData.ledgers.filter(l => {
                    // FIX: Pass actual ledger type for accurate math, allow 'All' to bypass type check
                    const matchSearch = (typeFilter === 'All' || l.type === typeFilter) && ((l.name || '').toLowerCase().includes(searchTerm) || (l.phone || '').toLowerCase().includes(searchTerm));
                    let matchFilter = true;
                    
                    // FIX: Pass the actual ledger type (l.type) instead of typeFilter to calculate correct math
                    const bal = getBal(l.id, l.type);

                    // NEW: Separating standard Dues from Advances perfectly!
                    if (activeMasterFilter === 'To Receive') {
                        matchFilter = l.type === 'Customer' && bal > 0.01;
                    } else if (activeMasterFilter === 'To Pay') {
                        matchFilter = l.type === 'Supplier' && bal > 0.01;
                    } else if (activeMasterFilter === 'Advance') {
                        matchFilter = bal < -0.01; // Any negative balance is an advance
                    } else if (activeMasterFilter === 'GST') {
                        matchFilter = l.gst && l.gst.trim() !== '';
                    } else if (activeMasterFilter === 'Non-GST') {
                        matchFilter = !l.gst || l.gst.trim() === '';
                    } else if (activeMasterFilter === 'Money In') {
                        matchFilter = UI.state.rawData.cashbook.some(c => c.ledgerId === l.id && c.type === 'in');
                    } else if (activeMasterFilter === 'Money Out') {
                        matchFilter = UI.state.rawData.cashbook.some(c => c.ledgerId === l.id && c.type === 'out');
                    }
                    
                    return matchSearch && matchFilter;
                });
                
                if(sortOption === 'name-asc') data.sort((a,b) => (a.name || '').localeCompare(b.name || ''));
                // NEW: Sort by Balance dynamically using your existing math engine
                if(sortOption === 'bal-desc') data.sort((a,b) => getBal(b.id, b.type) - getBal(a.id, a.type));
                if(sortOption === 'bal-asc') data.sort((a,b) => getBal(a.id, a.type) - getBal(b.id, b.type));

                html.push(...data.map(l => {
                    const isCustomer = String(l.type).toLowerCase() === 'customer';
                    let bal = getBal(l.id, l.type);
                    let balText = '';
                    let balColor = 'var(--md-text-muted)';
                    
                    // FIX: Case-insensitive evaluation!
                    if (isCustomer) {
                        if (bal > 0.01) { balText = `\u20B9${bal.toFixed(2)} (Receive)`; balColor = 'var(--md-error)'; }
                        else if (bal < -0.01) { balText = `\u20B9${Math.abs(bal).toFixed(2)} (Advance)`; balColor = 'var(--md-success)'; }
                        else { balText = `\u20B90.00`; }
                    } else {
                        if (bal > 0.01) { balText = `\u20B9${bal.toFixed(2)} (Pay)`; balColor = 'var(--md-error)'; }
                        else if (bal < -0.01) { balText = `\u20B9${Math.abs(bal).toFixed(2)} (Advance)`; balColor = 'var(--md-success)'; }
                        else { balText = `\u20B90.00`; }
                    }
                    
                    const rowIcon = isCustomer ? 'person' : 'storefront';
                    const rowColor = isCustomer ? '#0061a4' : '#ba1a1a';

                    return UI.renderRowWiseItem(
    UI.highlightText(l.name || 'Unnamed Party', searchTerm), UI.highlightText(l.phone || 'No Phone', searchTerm), `<span style="color:${balColor}; font-weight:500;">${balText}</span>`, 'View Ledger >', rowIcon, rowColor, `app.openPartyLedger('${l.id}', '${l.type}', '${(l.name || '').replace(/'/g, "\\'").replace(/"/g, "&quot;")}')`
);

                }));
            }
            else if (activeTab === 'pay-in' || activeTab === 'pay-out') {
                const targetType = activeTab === 'pay-in' ? 'in' : 'out';
                data = UI.state.rawData.cashbook.filter(c => {
                    const matchSearch = c.type === targetType && ((c.desc || '').toLowerCase().includes(searchTerm));
                    let matchFilter = true;
                    if (activeMasterFilter === 'Cash') matchFilter = (c.mode || '').toLowerCase() === 'cash';
                    if (activeMasterFilter === 'Bank') matchFilter = (c.mode || '').toLowerCase() !== 'cash';
                    return matchSearch && matchFilter && !(c.desc || '').startsWith('Expense');
                });

                html.push(...data.map(c => UI.renderRowWiseItem(
                    c.desc || 'Transaction', `${c.date} | ${c.mode || 'Cash'}`, `${targetType === 'in' ? '+' : '-'}\u20B9${(parseFloat(c.amount) || 0).toFixed(2)}`, targetType === 'in' ? 'Received' : 'Paid', targetType === 'in' ? 'arrow_downward' : 'arrow_upward', targetType === 'in' ? 'var(--md-success)' : 'var(--md-error)', `app.openReceipt('${c.id}', '${targetType}')`
                )));
            }
            
            // UPGRADE: Recycle Bin Render Logic
            else if (activeTab === 'trash') {
                const trashData = UI.state.rawData.trash || [];
                data = trashData.filter(t => (t.name || t.desc || t.invoiceNo || t.poNo || t.expenseNo || '').toLowerCase().includes(searchTerm));
                
                html.push(...data.map(t => UI.renderRowWiseItem(
                    t.name || t.desc || t.invoiceNo || t.poNo || t.expenseNo || 'Deleted Item', 
                    `Original Module: ${t._module}`, 
                    `<span style="color:var(--md-primary); font-weight:bold;">Restore</span>`, 
                    t.date || 'Unknown Date', 
                    'restore_from_trash', 
                    'var(--md-error)', 
                    `app.restoreRecord('${t.id}', '${t._module}')`
                )));
            }

            const container = document.getElementById(containerId);
            if (container) {
                const newHTML = html.length ? html.join('') : `
                <div class="empty-state" style="text-align: center; padding: 40px 20px;">
                    <svg width="100" height="100" viewBox="0 0 24 24" fill="none" stroke="var(--md-outline-variant)" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round" style="margin-bottom:16px;">
                        <circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line>
                    </svg>
                    <strong style="font-size: 18px; color: var(--md-on-surface); display:block;">No records found</strong>
                    <p style="margin: 8px 0 0 0; color: var(--md-text-muted);">Try adjusting your search or filters.</p>
                </div>`;
                UI.safeUpdateDOM(container, newHTML);
            }
        }

        // ------------------ EXPENSES ------------------
        else if (tab === 'expenses') {
            containerId = 'expense-history-container';
            data = UI.state.rawData.expenses.filter(e => {
                const matchSearch = (e.category || '').toLowerCase().includes(searchTerm) || (e.notes || '').toLowerCase().includes(searchTerm);
                let matchFilter = true;
                if (activeFilter !== 'All') matchFilter = e.category === activeFilter;
                return matchSearch && matchFilter;
            });

            data.forEach(e => sumValue += (parseFloat(e.amount) || 0));
            if (document.getElementById('sum-expenses')) document.getElementById('sum-expenses').innerText = `\u20B9${sumValue.toFixed(2)}`;

            // BULLETPROOF SORTING: Prevents WebView crashes!
            if(sortOption === 'date-desc') data.sort((a,b) => new Date(b.date || 0) - new Date(a.date || 0));
            if(sortOption === 'date-asc') data.sort((a,b) => new Date(a.date || 0) - new Date(b.date || 0));
            if(sortOption === 'amt-desc') data.sort((a,b) => (parseFloat(b.amount) || 0) - (parseFloat(a.amount) || 0));
            // NEW: Added Low to High sorting for expenses
            if(sortOption === 'amt-asc') data.sort((a,b) => (parseFloat(a.amount) || 0) - (parseFloat(b.amount) || 0));

            const container = document.getElementById(containerId);
            if (container) {
                container.innerHTML = data.length ? data.map((e, index) => {
                    // UPGRADE: Strictly prioritize and display the PO or Order Number!
                    let displayLink = e.linkedInvoice;
                    if (displayLink) {
                        const links = displayLink.split(',').map(x => x.trim()).filter(x => x);
                        const displayNames = links.map(linkId => {
                            // SELF-HEALING: Catch broken fragments like '8965' or '0778'
                            const sDoc = UI.state.rawData.sales.find(s => s.id === linkId || s.invoiceNo === linkId || s.orderNo === linkId || s.id.endsWith(linkId));
                            const pDoc = UI.state.rawData.purchases.find(p => p.id === linkId || p.poNo === linkId || p.invoiceNo === linkId || p.orderNo === linkId || p.id.endsWith(linkId));
                            
                            // Return STRICTLY the Order, PO, or Invoice No. NO prefixes!
                            if (sDoc) return sDoc.orderNo || sDoc.invoiceNo || sDoc.id.slice(-4).toUpperCase();
                            if (pDoc) return pDoc.orderNo || pDoc.poNo || pDoc.invoiceNo || pDoc.id.slice(-4).toUpperCase();
                            
                            // Absolute Fail-safe for missing documents
                            return linkId.startsWith('sollo-') ? linkId.slice(-4).toUpperCase() : linkId;
                        });
                        
                        // SELF-HEALING: Removes duplicates so "ORD-0003, ORD-0003" becomes just "ORD-0003"!
                        displayLink = [...new Set(displayNames)].join(', ');
                    }
                    
                    return `
                    <div class="m3-card tap-target" style="display:flex; justify-content:space-between; align-items:center;" onclick="app.openForm('expense', '${e.id}')">
                        <div>
                            <strong class="large-text">${e.expenseNo ? e.expenseNo + ' - ' : ''}${e.category || 'General Expense'}</strong><br>
                            <small>${e.date || ''} ${displayLink ? `| <span style="background:var(--md-primary-container); color:var(--md-primary); padding:2px 6px; border-radius:4px; font-weight:bold; font-size:10px;">🔗 ${displayLink}</span>` : ''} | ${e.notes || 'No notes'}</small>
                        </div>
                        <strong style="color:var(--md-error);">\u20B9${(parseFloat(e.amount) || 0).toFixed(2)}</strong>
                    </div>`;
                }).join('') : `
                <div class="empty-state">
                    <span class="material-symbols-outlined" style="font-size: 64px; color: var(--md-surface-variant);">account_balance_wallet</span>
                    <p style="margin: 12px 0;">No expenses match your filters.</p>
                    <button class="btn-primary" onclick="app.openForm('expense')">+ Log New Expense</button>
                </div>`;
            }
        }

        // ------------------ CASHBOOK ------------------
        else if (tab === 'cashbook') {
            // CRITICAL FIX: Ensure the Live Balance Bank calculator fires when tab is rendered
            UI.renderBankBalances(); 
            
            containerId = 'cashbook-container';
            data = UI.state.rawData.cashbook.filter(c => {
                const matchSearch = (c.desc || '').toLowerCase().includes(searchTerm) || 
                                    (c.mode || '').toLowerCase().includes(searchTerm) ||
                                    (c.ledgerName || '').toLowerCase().includes(searchTerm) ||
                                    (c.receiptNo || '').toLowerCase().includes(searchTerm) ||
                                    String(c.amount).includes(searchTerm);
                let matchFilter = true;
                if (activeFilter === 'In') matchFilter = c.type === 'in';
                if (activeFilter === 'Out') matchFilter = c.type === 'out';
                return matchSearch && matchFilter;
            });

            // BULLETPROOF SORTING: Prevents WebView crashes!
            if(sortOption === 'date-desc') data.sort((a,b) => new Date(b.date || 0) - new Date(a.date || 0));
            if(sortOption === 'date-asc') data.sort((a,b) => new Date(a.date || 0) - new Date(b.date || 0));

            const container = document.getElementById(containerId);
            if (container) {
                let lastDate = ''; // UPGRADE: Tracker for Google Pay style sticky headers
                container.innerHTML = data.length ? data.map((t, index) => {
                    
                    // UPGRADE: Extract and safely display the linked Invoice/PO Number!
                    let displayLink = '';
                    if (t.invoiceRef) {
                        const links = String(t.invoiceRef).split(',').map(x => x.trim()).filter(x => x);
                        const displayNames = links.map(linkId => {
                            const sDoc = UI.state.rawData.sales.find(s => s.id === linkId || s.invoiceNo === linkId || s.orderNo === linkId || s.id.endsWith(linkId));
                            const pDoc = UI.state.rawData.purchases.find(p => p.id === linkId || p.poNo === linkId || p.invoiceNo === linkId || p.orderNo === linkId || p.id.endsWith(linkId));
                            if (sDoc) return sDoc.orderNo || sDoc.invoiceNo || sDoc.id.slice(-4).toUpperCase();
                            if (pDoc) return pDoc.orderNo || pDoc.poNo || pDoc.invoiceNo || pDoc.id.slice(-4).toUpperCase();
                            return linkId.startsWith('sollo-') ? linkId.slice(-4).toUpperCase() : linkId;
                        });
                        displayLink = [...new Set(displayNames)].join(', ');
                    }

                    // UPGRADE: Inject Sticky Date Header if the date has changed
                    let dateHeader = '';
                    if (t.date !== lastDate) {
                        dateHeader = `<div style="position: sticky; top: 8px; z-index: 10; text-align: center; margin: 16px 0 8px 0; pointer-events: none;"><span style="background: var(--md-surface-variant); color: var(--md-on-surface-variant); padding: 4px 14px; border-radius: 20px; font-size: 11px; font-weight: bold; box-shadow: var(--elevation-1); border: 1px solid var(--md-outline-variant);">${t.date || 'Unknown Date'}</span></div>`;
                        lastDate = t.date;
                    }

                    return dateHeader + `
                    <div class="m3-card tap-target" style="display:flex; justify-content:space-between; align-items:center;" onclick="app.openReceipt('${t.id}', '${t.type}')">
                        <div>
                            <strong class="large-text">${t.receiptNo ? t.receiptNo + ' - ' : ''}${t.desc || 'Transaction'}</strong><br>
                            <small>${t.date || ''} ${displayLink ? `| <span style="background:var(--md-primary-container); color:var(--md-primary); padding:2px 6px; border-radius:4px; font-weight:bold; font-size:10px;">🔗 ${displayLink}</span>` : ''}</small><br>
                            <small>Party: <strong style="color:var(--md-primary)">${t.ledgerName || 'N/A'}</strong> | Mode: ${t.mode || 'Cash'}</small>
                        </div>
                        <strong style="font-size:16px; color:${t.type === 'in' ? 'var(--md-success)' : 'var(--md-error)'};">
                            ${t.type === 'in' ? '+' : '-'}\u20B9${(parseFloat(t.amount) || 0).toFixed(2)}
                        </strong>
                    </div>`;
                }).join('') : `
                <div class="empty-state">
                    <span class="material-symbols-outlined" style="font-size: 64px; color: var(--md-surface-variant);">account_balance</span>
                    <p style="margin: 12px 0;">No bank transactions match your filters.</p>
                </div>`;
            }
        }

        // ------------------ TIMELINE ------------------
        else if (tab === 'timeline') {
            containerId = 'timeline-list';
            
            // Get Custom Dates if set
            const startEl = document.getElementById('timeline-start-date');
            const endEl = document.getElementById('timeline-end-date');
            const startDate = startEl ? startEl.value : '';
            const endDate = endEl ? endEl.value : '';

            data = UI.state.rawData.timeline.filter(t => {
                const descStr = t.desc || (t.type === 'IN' ? 'Purchase' : (t.type === 'OUT' ? 'Sale' : ''));
                const matchSearch = descStr.toLowerCase().includes(searchTerm);
                let matchFilter = true;
                let matchDate = true;

                // 1. Check Date Range
                if (startDate && endDate && t.date) {
                    matchDate = (t.date >= startDate && t.date <= endDate);
                }
                
                // 2. Universal Type Check (Works for Banks, Parties, AND Expenses)
                if (activeFilter === 'Money In') {
                    if (t.hasOwnProperty('isInvoice')) matchFilter = t.isInvoice === false; 
                    else matchFilter = t.type === 'IN' || parseFloat(t.amount) > 0;
                } else if (activeFilter === 'Money Out') {
                    if (t.hasOwnProperty('isInvoice')) matchFilter = t.isInvoice === true; 
                    else matchFilter = t.type === 'OUT' || parseFloat(t.amount) < 0;
                } else if (activeFilter === 'Expenses') {
                    matchFilter = descStr.toLowerCase().includes('expense');
                }
                
                return matchSearch && matchFilter && matchDate;
            });

            const container = document.getElementById(containerId);
            if (container) {
                let lastDate = ''; // UPGRADE: Tracker for Google Pay style sticky headers
                container.innerHTML = data.length ? data.map(t => {
                    
                    // UPGRADE: Extract and safely display the linked Invoice/PO Number in the Statement Viewer!
                    let displayLink = '';
                    const refData = t.invoiceRef || t.linkedInvoice;
                    if (refData) {
                        const links = String(refData).split(',').map(x => x.trim()).filter(x => x);
                        const displayNames = links.map(linkId => {
                            const sDoc = UI.state.rawData.sales.find(s => s.id === linkId || s.invoiceNo === linkId || s.orderNo === linkId || s.id.endsWith(linkId));
                            const pDoc = UI.state.rawData.purchases.find(p => p.id === linkId || p.poNo === linkId || p.invoiceNo === linkId || p.orderNo === linkId || p.id.endsWith(linkId));
                            if (sDoc) return sDoc.orderNo || sDoc.invoiceNo || sDoc.id.slice(-4).toUpperCase();
                            if (pDoc) return pDoc.orderNo || pDoc.poNo || pDoc.invoiceNo || pDoc.id.slice(-4).toUpperCase();
                            return linkId.startsWith('sollo-') ? linkId.slice(-4).toUpperCase() : linkId;
                        });
                        displayLink = [...new Set(displayNames)].join(', ');
                    }

                    // UPGRADE: Inject Sticky Date Header if the date has changed
                    let dateHeader = '';
                    if (t.date !== lastDate) {
                        dateHeader = `<div style="position: sticky; top: 8px; z-index: 10; text-align: center; margin: 16px 0 8px 0; pointer-events: none;"><span style="background: var(--md-surface-variant); color: var(--md-on-surface-variant); padding: 4px 14px; border-radius: 20px; font-size: 11px; font-weight: bold; box-shadow: var(--elevation-1); border: 1px solid var(--md-outline-variant);">${t.date || 'Unknown Date'}</span></div>`;
                        lastDate = t.date;
                    }

                    if (t.hasOwnProperty('isInvoice')) {
                        return dateHeader + `
                        <div class="m3-card" style="display:flex; justify-content:space-between; align-items:center;">
                            <div>
                                <strong class="large-text">${t.desc || 'Document'}</strong><br>
                                <small>${t.date} ${displayLink ? `| <span style="background:var(--md-primary-container); color:var(--md-primary); padding:2px 6px; border-radius:4px; font-weight:bold; font-size:10px;">🔗 ${displayLink}</span>` : ''}</small>
                            </div>
                            <div style="text-align:right;">
                                <strong style="color:${t.isInvoice ? 'var(--md-error)' : 'var(--md-success)'};">\u20B9${(t.amount || 0).toFixed(2)}</strong><br>
                                <small>Bal: \u20B9${(t.runningBalance || 0).toFixed(2)}</small>
                            </div>
                        </div>`;
                    } else {
                        return `
                        <div class="m3-card" style="display:flex; justify-content:space-between; align-items:center;">
                            <div><strong class="large-text">${t.type === 'IN' ? 'Purchase' : 'Sale'} - ${t.party || 'Unknown'}</strong><br><small>${t.date} | Ref: ${t.ref}</small></div>
                            <strong style="font-size:16px; color:${t.type === 'IN' ? 'var(--md-success)' : 'var(--md-error)'};">${t.type === 'IN' ? '+' : '-'}${t.qty}</strong>
                        </div>`;
                    }
                }).join('') : '<p class="empty-state">No records match your filters.</p>';
            }
        }
    },

    // ==========================================
    // 5. TRUE PROFIT DASHBOARD & DATE FILTERS
    // ==========================================
    renderDashboard: () => {
        const filterEl = document.getElementById('dashboard-date-filter');
        const dateFilter = filterEl ? filterEl.value : 'all';
        const today = new Date();
        const todayStr = typeof Utils !== 'undefined' && Utils.getLocalDate ? Utils.getLocalDate() : today.toISOString().split('T')[0];
        
        const isDateInRange = (dateStr) => {
            if (dateFilter === 'all') return true;
            if (!dateStr) return false;
            if (dateFilter === 'today') return dateStr === todayStr;
            
            // FIXED: Manually split the string to bypass the UTC Timezone shift bug
            const [yearStr, monthStr, dayStr] = dateStr.split('-');
            const itemYear = parseInt(yearStr, 10);
            const itemMonth = parseInt(monthStr, 10) - 1; // JS months are 0-indexed
            
            if (dateFilter === 'month') return itemMonth === today.getMonth() && itemYear === today.getFullYear();
            if (dateFilter === 'year') return itemYear === today.getFullYear();
            return true;
        };

        const sales = UI.state.rawData.sales;
        const purchases = UI.state.rawData.purchases;
        const expenses = UI.state.rawData.expenses;

        let totalSales = 0, outputGst = 0, totalPurchases = 0, inputGst = 0, totalExpenses = 0;
        let cogs = 0; 

        // CRITICAL FIX: Dashboard Payment Map Math & Collision Guard
        const paymentMap = {};
        UI.state.rawData.cashbook.forEach(c => {
            if (c.invoiceRef && c.ledgerId) {
                let amt = c.type === 'in' ? parseFloat(c.amount) : -parseFloat(c.amount);
                // FIX: Split by comma to properly distribute payments covering multiple invoices!
                const refs = String(c.invoiceRef).split(',').map(r => r.trim());
                let splitAmt = amt / (refs.length || 1);
                refs.forEach(r => {
                    const key = `${c.ledgerId}_${r}`;
                    paymentMap[key] = (paymentMap[key] || 0) + splitAmt;
                });
            }
        });

        sales.forEach(s => { 
            if(s.status !== 'Open' && isDateInRange(s.date)) { 
                const isReturn = s.documentType === 'return';
                const modifier = isReturn ? -1 : 1;

                // UX FIX: Show Full Invoice Value (including GST) so it matches your Sales Tab
                totalSales += (parseFloat(s.grandTotal) || 0) * modifier; 
                outputGst += (parseFloat(s.totalGst) || 0) * modifier; 
                
                (s.items || []).forEach(item => {
                    const historicalCost = parseFloat(item.buyPrice) || 0;
                    cogs += ((parseFloat(item.qty) || 0) * historicalCost) * modifier;
                });
            }
        });        
        purchases.forEach(p => { 
            if (p.status !== 'Open' && isDateInRange(p.date)) { 
                const isReturn = p.documentType === 'return';
                const modifier = isReturn ? -1 : 1;
                const netPurchase = (parseFloat(p.grandTotal) || 0) - (parseFloat(p.totalGst) || 0);
                totalPurchases += netPurchase * modifier; 
                inputGst += (p.totalGst || 0) * modifier; 
            } 
        });
        
        expenses.forEach(e => { 
            if(isDateInRange(e.date)) totalExpenses += (parseFloat(e.amount) || 0); 
        });

        // TRUE ACCRUAL PROFIT MATH
        const netRevenue = totalSales - outputGst; // FIXED: Re-added the missing calculation!
        const grossMargin = netRevenue - cogs;
        const trueNetProfit = grossMargin - totalExpenses;

        // UPGRADE 1: Count-Up Animation Engine
        const animateValue = (id, start, end, duration) => {
            const obj = document.getElementById(id);
            if (!obj) return;
            let startTimestamp = null;
            const step = (timestamp) => {
                if (!startTimestamp) startTimestamp = timestamp;
                const progress = Math.min((timestamp - startTimestamp) / duration, 1);
                // Ease-out formula for smoother deceleration
                const easeOut = 1 - Math.pow(1 - progress, 3);
                obj.innerHTML = '\u20B9' + (easeOut * (end - start) + start).toFixed(2);
                if (progress < 1) window.requestAnimationFrame(step);
            };
            window.requestAnimationFrame(step);
        };

        // Update DOM with Animation (Duration: 800ms)
        animateValue('dash-total-sales', 0, totalSales, 800);
        
        const netProfitEl = document.getElementById('dash-net-profit');
        if(netProfitEl) {
            animateValue('dash-net-profit', 0, trueNetProfit, 800);
            netProfitEl.style.color = trueNetProfit >= 0 ? 'var(--md-success)' : 'var(--md-error)';
        }
        
        if(document.getElementById('profit-breakdown')) document.getElementById('profit-breakdown').innerText = `Gross: \u20B9${grossMargin.toFixed(0)} | Exp: \u20B9${totalExpenses.toFixed(0)}`;
        
        // NEW: Operational Order Volume Tracking (Open, Shipped, Completed)
        let openOrders = 0, shippedOrders = 0, completedOrders = 0;
        let openCount = 0, shippedCount = 0, completedCount = 0;

        sales.forEach(s => {
            // Apply the global dashboard date filter to the pipeline
            if (isDateInRange(s.date)) {
                const isReturn = s.documentType === 'return';
                const modifier = isReturn ? -1 : 1;
                const val = (parseFloat(s.grandTotal) || 0) * modifier;
                
                if (s.status === 'Open') {
                    openOrders += val;
                    if (!isReturn) openCount++;
                } else if (s.status === 'Shipped') {
                    shippedOrders += val;
                    if (!isReturn) shippedCount++;
                } else if (s.status === 'Completed') {
                    completedOrders += val;
                    if (!isReturn) completedCount++;
                }
            }
        });

        if(document.getElementById('dash-orders-open')) document.getElementById('dash-orders-open').innerHTML = `${openCount} Orders <strong style="color:var(--md-error); margin-left:8px;">\u20B9${openOrders.toFixed(2)}</strong>`;
        if(document.getElementById('dash-orders-shipped')) document.getElementById('dash-orders-shipped').innerHTML = `${shippedCount} Orders <strong style="color:#f57f17; margin-left:8px;">\u20B9${shippedOrders.toFixed(2)}</strong>`;
        if(document.getElementById('dash-orders-completed')) document.getElementById('dash-orders-completed').innerHTML = `${completedCount} Orders <strong style="color:var(--md-success); margin-left:8px;">\u20B9${completedOrders.toFixed(2)}</strong>`;
        // --- NEW CODE: Call the chart updater ---
        UI.updateChart(totalSales, totalPurchases, totalExpenses);
        // --- END OF NEW CODE ---

        // --- OVERDUE REMINDERS SAFELY INSIDE THE FUNCTION ---
        const overdueSales = sales.filter(s => {
            if (s.status === 'Open' || s.documentType === 'return') return false;
            
            // FIX: Check ALL references to catch cross-linked payments, and respect FIFO completion!
            const uniqueRefs = [...new Set([s.orderNo, s.invoiceNo, s.id].filter(Boolean))];
            const totalReceived = uniqueRefs.reduce((sum, ref) => sum + (paymentMap[`${s.customerId}_${ref}`] || 0), 0);
            const balance = Math.max(0, (parseFloat(s.grandTotal) || 0) - totalReceived);
            if (balance <= 0) return false;
            
            if (!s.date) return false;
            // BULLETPROOF DATE MATH: Manually parse YYYY-MM-DD so old WebViews don't panic!
            const parts = s.date.split('-'); 
            const invoiceDate = new Date(parts[0], parts[1] - 1, parts[2]); 
            
            const diffTime = today - invoiceDate;
            if (diffTime < 0) return false; // Prevent future/post-dated invoices from being flagged!
            
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
            return diffDays > 15;
        });

        // UPGRADE 2: Update Bottom Nav Notification Badge for Actionable Documents
        const pendingPurchases = purchases.filter(p => p.status === 'Open').length;
        const totalAlerts = overdueSales.length + pendingPurchases;
        const badgeEl = document.getElementById('badge-docs');
        if (badgeEl) {
            if (totalAlerts > 0) {
                badgeEl.innerText = totalAlerts;
                badgeEl.classList.remove('hidden');
            } else {
                badgeEl.classList.add('hidden');
            }
        }

        // UPGRADE 3: Native OS Home Screen App Badging
        if ('setAppBadge' in navigator) {
            if (totalAlerts > 0) {
                navigator.setAppBadge(totalAlerts).catch(e => console.error("Badging error", e));
            } else {
                navigator.clearAppBadge().catch(e => console.error("Badging clear error", e));
            }
        }

        const overdueContainer = document.getElementById('overdue-reminders-container');
        if (overdueContainer) {
            if (overdueSales.length > 0) {
                overdueContainer.classList.remove('hidden');
                document.getElementById('overdue-text').innerText = `${overdueSales.length} invoices pending over 15 days.`;
                
                document.getElementById('list-overdue').innerHTML = overdueSales.map(s => {
                    // FIX: Update the dashboard display to match the new cross-linked math!
                    const uniqueRefs = [...new Set([s.orderNo, s.invoiceNo, s.id].filter(Boolean))];
                    const totalReceived = uniqueRefs.reduce((sum, ref) => sum + (paymentMap[`${s.customerId}_${ref}`] || 0), 0);
                    const balance = Math.max(0, (parseFloat(s.grandTotal) || 0) - totalReceived);
                    return `
                    <li>
                        <div>
                            <strong class="large-text">${s.customerName || 'Unknown Party'}</strong><br>
                            <small class="color-primary">Inv: ${s.invoiceNo || 'Draft'} | Bal: <strong style="color:var(--md-error)">\u20B9${balance.toFixed(2)}</strong></small>
                        </div>
                        <span style="background:#fff0f2; color:#be123c; border:1px solid #be123c; padding:4px 8px; border-radius:4px; font-size:12px; font-weight:bold;">Overdue</span>
                    </li>`;
                }).join('');
            } else {
                overdueContainer.classList.add('hidden');
            }
        }
    }, // <-- ADDED THIS MISSING BRACKET AND COMMA!
    // ==========================================
    // NEW CODE: DYNAMIC CHART ENGINE
    // ==========================================
    chartInstance: null,

    updateChart: (salesAmt, purchaseAmt, expenseAmt) => {
        const canvas = document.getElementById('dashboard-chart');
        if (!canvas) return;
        const ctx = canvas.getContext('2d');

        // Destroy the old chart before drawing a new one
        if (UI.chartInstance) {
            UI.chartInstance.destroy();
        }

        // UPGRADE: Premium Vertical Gradients
        const salesGrad = ctx.createLinearGradient(0, 0, 0, 300);
        salesGrad.addColorStop(0, '#4ade80'); // Light Green
        salesGrad.addColorStop(1, '#146c2e'); // Deep Green

        const purchGrad = ctx.createLinearGradient(0, 0, 0, 300);
        purchGrad.addColorStop(0, '#facc15'); // Light Orange/Yellow
        purchGrad.addColorStop(1, '#f57f17'); // Deep Orange

        const expGrad = ctx.createLinearGradient(0, 0, 0, 300);
        expGrad.addColorStop(0, '#fb7185'); // Light Red
        expGrad.addColorStop(1, '#ba1a1a'); // Deep Red

        UI.chartInstance = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['Sales', 'Purchases', 'Expenses'],
                datasets: [{
                    label: 'Amount (₹)',
                    data: [salesAmt, purchaseAmt, expenseAmt],
                    backgroundColor: [salesGrad, purchGrad, expGrad],
                    borderRadius: 6, // Rounds the top of the bars
                    borderSkipped: false
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false } // Hides the legend since our labels are self-explanatory
                },
                scales: {
                    y: { 
                        beginAtZero: true,
                        ticks: {
                            callback: function(value) { return '₹' + value; }
                        }
                    }
                }
            }
        });
    },
    // ==========================================
    // END OF NEW CODE
    // ==========================================

    // ==========================================
    // 6. BOTTOM SHEETS & PRODUCT CRUD
    // ==========================================
        openBottomSheet: (sheetId) => {
        const sheet = document.getElementById(sheetId);
        const overlay = document.getElementById('sheet-overlay');
        
        history.pushState({ modalOpen: true }, ''); // Native Back Gesture hook

        if (overlay) {
            overlay.classList.remove('hidden');
            requestAnimationFrame(() => overlay.classList.add('open'));
        }

        if (sheet) {
            sheet.classList.remove('hidden'); 
            requestAnimationFrame(() => { sheet.classList.add('open'); });
        }
        
        // UPGRADE: Dynamic Status Bar Colors!
        if (sheetId === 'sheet-payment-in') { UI.toggleDeleteButton('receipt-in', false); UI.setStatusBarColor('#146c2e'); }
        if (sheetId === 'sheet-payment-out') { UI.toggleDeleteButton('receipt-out', false); UI.setStatusBarColor('#ba1a1a'); }
        if (sheetId === 'sheet-stock-adjustment') { UI.setStatusBarColor('#f57f17'); }

        if (sheetId === 'sheet-master-sort') {
            const filterSelect = document.getElementById('filter-master-view');
            const sortSelect = document.getElementById('sort-master-view');
            
            if (filterSelect) {
                const tab = UI.state.activeMasterTab;
                // NEW: Unified Filtering & Sorting for all Ledger types!
                if (tab === 'customers' || tab === 'suppliers' || tab === 'contacts') {
                    filterSelect.innerHTML = `
                        <option value="All">All Status</option>
                        <option value="To Receive">To Receive (Due)</option>
                        <option value="To Pay">To Pay (Due)</option>
                        <option value="Advance">Advance (Paid / Received)</option>
                        <option value="GST">GST Registered</option>
                        <option value="Non-GST">Non-GST</option>
                        <option value="Money In">Money In (Received)</option>
                        <option value="Money Out">Money Out (Paid)</option>
                    `;
                    if(sortSelect) sortSelect.innerHTML = `
                        <option value="name-asc">A to Z</option>
                        <option value="bal-desc">Balance: High to Low</option>
                        <option value="bal-asc">Balance: Low to High</option>
                    `;
                } else if (tab === 'pay-in' || tab === 'pay-out') {
                    filterSelect.innerHTML = `<option value="All">All Modes</option><option value="Cash">Cash Only</option><option value="Bank">Bank / Online Only</option>`;
                    if(sortSelect) sortSelect.innerHTML = `<option value="date-desc">Newest First</option><option value="date-asc">Oldest First</option>`;
                } else if (tab === 'trash') {
                    filterSelect.innerHTML = `<option value="All">All Trashed Items</option>`;
                    if(sortSelect) sortSelect.innerHTML = `<option value="date-desc">Recently Deleted</option>`;
                } else {
                    // NEW: Products gets the 'Stock Available' filter, but keeps its original sorting
                    filterSelect.innerHTML = `<option value="All">All Products</option><option value="In Stock">Stock Available</option>`;
                    if(sortSelect) sortSelect.innerHTML = `<option value="name-asc">A to Z</option><option value="stock-asc">Lowest Stock First</option>`;
                }
            }
        }

        if (sheetId === 'sheet-customers') {
            const searchBox = document.getElementById('search-customers');
            if (searchBox) { searchBox.value = ''; setTimeout(() => searchBox.focus(), 350); } // UPGRADE: Auto-focus keyboard!
            UI.renderLedgerList('list-customers', UI.state.rawData.ledgers.filter(l => l.type === 'Customer'), UI.state.currentPrefix || 'sales');
            document.querySelectorAll('#list-customers li').forEach(li => li.style.display = ''); // Force list to be visible
        }
        else if (sheetId === 'sheet-suppliers') {
            const searchBox = document.getElementById('search-suppliers');
            if (searchBox) { searchBox.value = ''; setTimeout(() => searchBox.focus(), 350); } // UPGRADE: Auto-focus keyboard!
            UI.renderLedgerList('list-suppliers', UI.state.rawData.ledgers.filter(l => l.type === 'Supplier'), UI.state.currentPrefix || 'purchase');
            document.querySelectorAll('#list-suppliers li').forEach(li => li.style.display = ''); // Force list to be visible
        }
        else if (sheetId === 'sheet-products') {
            const searchBox = document.getElementById('search-products');
            if (searchBox) { searchBox.value = ''; setTimeout(() => searchBox.focus(), 350); } // UPGRADE: Auto-focus keyboard!
            UI.renderProductList(UI.state.rawData.items);
            document.querySelectorAll('#list-products li').forEach(li => li.style.display = ''); // Force list to be visible
        }
        else if (sheetId === 'sheet-stock-adjustment') {
            const options = UI.state.rawData.items.map(i => 
                `<option value="${i.id}">${i.name || 'Unnamed'} (Cur: ${i.stock || 0} ${i.uom || ''})</option>`
            ).join('');
            
            const prodSelect = document.getElementById('adj-product-id');
            if(prodSelect) prodSelect.innerHTML = options ? `<option value="">Select Product...</option>` + options : `<option value="">No products found...</option>`;
            
            const dateInput = document.getElementById('adj-date');
            if(dateInput) {
                const today = typeof Utils !== 'undefined' && Utils.getLocalDate ? Utils.getLocalDate() : new Date().toISOString().split('T')[0];
                dateInput.value = today;
                if (dateInput._flatpickr) dateInput._flatpickr.setDate(today); // FIX: Sync Flatpickr
            }
        }
        else if (sheetId === 'sheet-payment-in' || sheetId === 'sheet-payment-out') {
            // UPGRADE: Reset the display card when opening a new payment form
            const prefix = sheetId === 'sheet-payment-in' ? 'pay-in' : 'pay-out';
            const display = document.getElementById(`${prefix}-customer-display`) || document.getElementById(`${prefix}-supplier-display`);
            const input = document.getElementById(`${prefix}-customer`) || document.getElementById(`${prefix}-supplier`);
            
            if (display) {
                display.innerText = sheetId === 'sheet-payment-in' ? 'Select Customer...' : 'Select Supplier...';
                display.style.color = 'var(--md-text-muted)';
            }
            if (input) input.value = '';
        }
    },

        closeBottomSheet: (sheetId) => {
        const sheet = document.getElementById(sheetId);
        const overlay = document.getElementById('sheet-overlay');

        if (sheet) sheet.classList.remove('open');
        if (overlay) overlay.classList.remove('open');

        setTimeout(() => { 
            if (sheet) sheet.classList.add('hidden');
            if (overlay) overlay.classList.add('hidden'); 
        }, 300);

        UI.resetStatusBarColor(); // Reset phone status bar back to normal
        if(typeof app !== 'undefined' && app.state) app.state.currentReceiptId = null;
        // FIXED: Added an ignore flag to prevent double-closing
        if (history.state && history.state.modalOpen) { window._ignoreNextPop = true; history.back(); }
    },

    closeAllBottomSheets: () => {
        let closedSomething = false;
        document.querySelectorAll('.bottom-sheet.open').forEach(sheet => {
            sheet.classList.remove('open');
            setTimeout(() => sheet.classList.add('hidden'), 300);
            closedSomething = true;
        });
        
        const overlay = document.getElementById('sheet-overlay');
        if (overlay && overlay.classList.contains('open')) {
            overlay.classList.remove('open');
            setTimeout(() => overlay.classList.add('hidden'), 300);
            closedSomething = true;
        }
        
        // FIXED: Restored the missing closing logic!
        if (closedSomething && history.state && history.state.modalOpen) { window._ignoreNextPop = true; history.back(); }
    },
        
    // UPGRADE: iOS-Style Haptic Context Menu
    showContextMenu: (clickAction) => {
        let overlay = document.getElementById('haptic-overlay');
        if (!overlay) {
            overlay = document.createElement('div');
            overlay.id = 'haptic-overlay';
            overlay.className = 'drawer-overlay hidden'; // Reuses your premium blur effect!
            overlay.style.zIndex = '4500';
            
            overlay.innerHTML = `
                <div id="haptic-menu" class="bottom-sheet" style="z-index: 4501; padding: 0 0 calc(24px + env(safe-area-inset-bottom, 0px)) 0; background: transparent !important; box-shadow: none;">
                    
                    <div style="margin: 0 16px 16px 16px; background: var(--md-surface); border-radius: 20px; box-shadow: 0 8px 32px rgba(0,0,0,0.15); overflow: hidden;">
                        <div style="padding: 14px; text-align: center; font-size: 12px; font-weight: bold; color: var(--md-text-muted); border-bottom: 1px solid var(--md-surface-variant); background: rgba(0,0,0,0.02);">
                            DOCUMENT OPTIONS
                        </div>
                        
                        <div class="tap-target" onclick="document.getElementById('haptic-overlay').click(); setTimeout(() => { ${clickAction} }, 300);" style="padding: 16px 20px; display: flex; align-items: center; gap: 16px; font-size: 16px; font-weight: 500; color: var(--md-on-surface); border-bottom: 1px solid var(--md-surface-variant); background: var(--md-surface);">
                            <div class="icon-circle" style="width:36px; height:36px; background: var(--md-primary-container); color: var(--md-primary);"><span class="material-symbols-outlined" style="font-size:20px;">edit_document</span></div>
                            Open & Edit Document
                        </div>
                        
                        <div class="tap-target" onclick="document.getElementById('haptic-overlay').click(); setTimeout(() => { if(window.Utils) window.Utils.showToast('Open the document to generate a PDF'); }, 300);" style="padding: 16px 20px; display: flex; align-items: center; gap: 16px; font-size: 16px; font-weight: 500; color: var(--md-on-surface); background: var(--md-surface);">
                            <div class="icon-circle" style="width:36px; height:36px; background: #fff8e1; color: #f57f17;"><span class="material-symbols-outlined" style="font-size:20px;">picture_as_pdf</span></div>
                            Generate PDF
                        </div>
                    </div>

                    <div class="tap-target" onclick="document.getElementById('haptic-overlay').click();" style="margin: 0 16px; padding: 16px; text-align: center; background: var(--md-surface); border-radius: 20px; font-size: 16px; font-weight: bold; color: var(--md-error); box-shadow: 0 4px 16px rgba(0,0,0,0.1);">
                        Cancel
                    </div>

                </div>
            `;
            document.body.appendChild(overlay);
            
            // Close logic
            overlay.addEventListener('click', (e) => {
                if (e.target === overlay) {
                    document.getElementById('haptic-menu').classList.remove('open');
                    overlay.classList.remove('open');
                    setTimeout(() => overlay.classList.add('hidden'), 300);
                }
            });
        }
        
        // Show the menu
        overlay.classList.remove('hidden');
        requestAnimationFrame(() => {
            overlay.classList.add('open');
            document.getElementById('haptic-menu').classList.add('open');
        });
        
        // Micro-vibration when the menu pops up
        if (window.navigator && window.navigator.vibrate) window.navigator.vibrate(15);
    },
    renderLedgerList: (containerId, ledgers, prefix) => {
        const container = document.getElementById(containerId);
        if(!container) return;
        container.innerHTML = ledgers.map(l => `
            <li onclick="UI.selectLedger('${l.id}', '${(l.name || '').replace(/'/g, "\\'").replace(/"/g, "&quot;")}', '${prefix}')">
                <div><div class="large-text">${l.name || 'Unnamed'}</div><small class="color-primary">${l.phone || 'No Phone'}</small></div>
            </li>
        `).join('');
    },

    selectLedger: (id, name, prefix) => {
        const inputId = document.getElementById(`${prefix}-customer-id`) || document.getElementById(`${prefix}-supplier-id`) || document.getElementById(`${prefix}-customer`) || document.getElementById(`${prefix}-supplier`);
        if (inputId) inputId.value = id;
        
        const display = document.getElementById(`${prefix}-customer-display`) || document.getElementById(`${prefix}-supplier-display`);
        if(display) {
            display.innerText = name; 
            display.style.color = 'var(--md-on-surface)';
        }
        
        if (typeof app !== 'undefined') {
            if ((prefix === 'sales' || prefix === 'purchase') && typeof app.loadOriginalDocuments === 'function') {
                app.loadOriginalDocuments(id, prefix);
            } else if (prefix === 'pay-in' && typeof app.loadPendingInvoices === 'function') {
                app.loadPendingInvoices(id, 'in');
            } else if (prefix === 'pay-out' && typeof app.loadPendingInvoices === 'function') {
                app.loadPendingInvoices(id, 'out');
            }
        }

        UI.closeBottomSheet(prefix === 'sales' || prefix === 'pay-in' ? 'sheet-customers' : 'sheet-suppliers');
    },

    renderProductList: (items) => {
        UI.state.selectedProducts = [];
        const isPurchase = UI.state.activeActivity === 'purchase';
        const container = document.getElementById('list-products');
        if(!container) return;
        
        container.innerHTML = items.map(item => {
            const price = isPurchase ? (item.buyPrice || 0) : (item.sellPrice || 0);
            
            const currentStock = parseFloat(item.stock) || 0;
            const minStock = parseFloat(item.minStock) || 0;
            const isLowStock = minStock > 0 && currentStock <= minStock;
            
            return `
            <li onclick="UI.toggleProductSelection(this, '${item.id}', '${(item.name || '').replace(/'/g, "\\'").replace(/"/g, "&quot;")}', ${price}, ${item.gst || 0}, '${(item.uom || '').replace(/'/g, "\\'")}', '${(item.hsn || '').replace(/'/g, "\\'")}', ${item.buyPrice || 0})">
                <div>
                    <div class="large-text">${item.name || 'Unnamed Product'}</div>
                    <small>
                        <span style="${isLowStock ? 'color:var(--md-error); font-weight:bold;' : ''}">Stock: ${currentStock} ${item.uom || ''} ${isLowStock ? '⚠️' : ''}</span> 
                        | Rate: \u20B9${price.toFixed(2)}
                    </small>
                </div>
                <input type="checkbox" style="width: 20px; height: 20px; pointer-events: none;">
            </li>
        `}).join('');
    },

    toggleProductSelection: (li, id, name, price, gst, uom, hsn, buyPrice) => {
        const cb = li.querySelector('input'); cb.checked = !cb.checked;
        if (cb.checked) { 
            UI.state.selectedProducts.push({ id, name, price, gst, uom, hsn, buyPrice }); 
            li.style.background = 'var(--md-surface-variant)'; 
        } else { 
            UI.state.selectedProducts = UI.state.selectedProducts.filter(p => p.id !== id); 
            li.style.background = 'transparent'; 
        }
    },

    confirmProducts: () => {
        const prefix = UI.state.activeActivity;
        const tbody = document.getElementById(`${prefix}-items-body`);
        if(!tbody) return;
        
        UI.state.selectedProducts.forEach(p => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>
                    <div style="font-weight:500;">${p.name}</div>
                    ${prefix === 'sales' ? `
                    <div style="display:flex; align-items:center; gap:4px; margin-top:4px;">
                        <span style="font-size:11px; color:var(--md-text-muted);">Buy: ₹</span>
                        <input type="number" class="row-item-buyprice" value="${p.buyPrice || 0}" step="any" oninput="UI.calcSalesTotals()" style="width:60px; padding:2px 4px; font-size:11px; border:1px solid var(--md-outline-variant); border-radius:4px; background:var(--md-surface);">
                    </div>
                    <small class="live-margin" style="font-size:10px; display:block; margin-top:4px;"></small>
                    ` : `<input type="hidden" class="row-item-buyprice" value="${p.buyPrice || 0}">`}
                    <input type="hidden" class="row-item-id" value="${p.id}">
                    <input type="hidden" class="row-item-name" value="${(p.name || '').replace(/"/g, '&quot;')}">
                </td>
                <td><input type="text" class="row-hsn" value="${p.hsn}" readonly style="width:60px; text-align:center; padding:4px;"></td>
                <td><input type="number" class="row-qty" value="1" min="0.01" step="any" oninput="UI.calc${prefix.charAt(0).toUpperCase() + prefix.slice(1)}Totals()" style="width:60px; padding:4px;"></td>
                <td><input type="text" class="row-uom" value="${p.uom}" readonly style="width:50px; padding:4px;"></td>
                <td><input type="number" class="row-rate" value="${p.price}" step="any" oninput="UI.calc${prefix.charAt(0).toUpperCase() + prefix.slice(1)}Totals()" style="width:80px; padding:4px;"></td>
                <td><input type="number" class="row-gst" value="${p.gst}" step="any" oninput="UI.calc${prefix.charAt(0).toUpperCase() + prefix.slice(1)}Totals()" style="width:50px; padding:4px;"></td>
                <td class="row-total" style="font-weight:bold; text-align:right;">0.00</td>
                <td style="text-align:center;">
                    <span class="material-symbols-outlined tap-target" style="color:var(--md-error); font-size:20px;" onclick="this.closest('tr').remove(); UI.calc${prefix.charAt(0).toUpperCase() + prefix.slice(1)}Totals()">cancel</span>
                </td>
            `;
            tbody.appendChild(tr);
        });
        prefix === 'sales' ? UI.calcSalesTotals() : UI.calcPurchaseTotals();
        UI.closeBottomSheet('sheet-products');
    },

    // ==========================================
    // 7. RECEIVABLES REPORT ENGINE
    // ==========================================
    downloadReceivablesReport: async () => {
        if (typeof Utils === 'undefined' || typeof Utils.printReceivablesReport !== 'function') return alert("Print engine unavailable.");
        
        const customerLedgers = UI.state.rawData.ledgers.filter(l => l.type === 'Customer');
        const reportData = [];
        let grandTotal = 0;

        for (const ledger of customerLedgers) {
            if (typeof getKhataStatement !== 'undefined') {
                const statement = await getKhataStatement(ledger.id, 'Customer');
                if (statement.finalBalance > 0) {
                    reportData.push({
                        name: ledger.name || 'Unknown',
                        phone: ledger.phone || 'N/A',
                        balance: statement.finalBalance
                    });
                    grandTotal += statement.finalBalance;
                }
            }
        }
        
        if (reportData.length === 0) return alert("No pending receivables found.");
        Utils.printReceivablesReport(reportData, grandTotal);
    },

    // ==========================================
    // ADVANCED REPORTING SUITE
    // ==========================================
    switchReportTab: (tab) => {
        document.getElementById('view-daybook').classList.add('hidden');
        document.getElementById('view-pnl').classList.add('hidden');
        document.getElementById('btn-tab-daybook').style.background = 'transparent';
        document.getElementById('btn-tab-daybook').style.boxShadow = 'none';
        document.getElementById('btn-tab-pnl').style.background = 'transparent';
        document.getElementById('btn-tab-pnl').style.boxShadow = 'none';

        if (tab === 'daybook') {
            document.getElementById('view-daybook').classList.remove('hidden');
            document.getElementById('btn-tab-daybook').style.background = 'var(--md-surface)';
            document.getElementById('btn-tab-daybook').style.boxShadow = 'var(--elevation-1)';
            UI.renderDayBook();
        } else {
            document.getElementById('view-pnl').classList.remove('hidden');
            document.getElementById('btn-tab-pnl').style.background = 'var(--md-surface)';
            document.getElementById('btn-tab-pnl').style.boxShadow = 'var(--elevation-1)';
            UI.renderPnL();
        }
    },

    renderDayBook: () => {
        const dateInput = document.getElementById('report-daybook-date');
        if (!dateInput.value) {
            dateInput.value = typeof Utils !== 'undefined' ? Utils.getLocalDate() : new Date().toISOString().split('T')[0];
        }
        const targetDate = dateInput.value;
        const container = document.getElementById('daybook-container');
        let html = '';

        // Combine all relevant activity into one chronological stream
        const dailyActivity = [];
        
        UI.state.rawData.sales.filter(s => s.date === targetDate && s.status !== 'Open').forEach(s => {
            const isRet = s.documentType === 'return';
            const isNonGST = s.invoiceType === 'Non-GST';
            const docLabel = isRet ? 'Credit Note' : (isNonGST ? 'Bill of Supply' : 'Sales Invoice');
            dailyActivity.push({ time: s.id, type: docLabel, desc: s.customerName, amount: s.grandTotal, icon: isRet ? 'assignment_return' : 'receipt_long', color: isRet ? 'var(--md-error)' : 'var(--md-success)', sign: isRet ? '-' : '+' });
        });
        
        UI.state.rawData.purchases.filter(p => p.date === targetDate && p.status !== 'Open').forEach(p => {
            const isRet = p.documentType === 'return';
            const isNonGST = p.invoiceType === 'Non-GST';
            const docLabel = isRet ? 'Debit Note' : (isNonGST ? 'Bill of Supply' : 'Purchase Bill');
            dailyActivity.push({ time: p.id, type: docLabel, desc: p.supplierName, amount: p.grandTotal, icon: isRet ? 'assignment_returned' : 'local_shipping', color: isRet ? 'var(--md-success)' : 'var(--md-error)', sign: isRet ? '+' : '-' });
        });
        
        UI.state.rawData.cashbook.filter(c => c.date === targetDate && !c.isAutoGenerated).forEach(c => {
            const isIn = c.type === 'in';
            dailyActivity.push({ time: c.id, type: isIn ? 'Money In' : 'Money Out', desc: c.ledgerName, amount: parseFloat(c.amount), icon: isIn ? 'arrow_downward' : 'arrow_upward', color: isIn ? 'var(--md-success)' : 'var(--md-error)', sign: isIn ? '+' : '-' });
        });
        
        UI.state.rawData.expenses.filter(e => e.date === targetDate).forEach(e => {
            dailyActivity.push({ time: e.id, type: 'Expense', desc: e.category, amount: parseFloat(e.amount), icon: 'account_balance_wallet', color: 'var(--md-error)', sign: '-' });
        });

        if (dailyActivity.length === 0) {
            container.innerHTML = '<p class="empty-state">No financial activity recorded on this date.</p>';
            return;
        }

        // FIX: Extract the actual timestamp chunk from the end of the ID string to sort chronologically!
        dailyActivity.sort((a, b) => {
            const timeA = parseInt(a.time.split('-').pop()) || 0;
            const timeB = parseInt(b.time.split('-').pop()) || 0;
            return timeB - timeA;
        });

        // UPGRADE: Added Export CSV Button dynamically
        html += `
            <div style="display:flex; justify-content:flex-end; margin-bottom: 12px; padding: 0 4px;">
                <button class="btn-primary-small tap-target" onclick="UI.exportDaybookCSV()" style="display:flex; align-items:center; gap:4px; background: #e8f5e9; color: #146c2e; border: 1px solid rgba(20, 108, 46, 0.3);">
                    <span class="material-symbols-outlined" style="font-size: 16px;">download</span> Export CSV
                </button>
            </div>
        `;

        html += dailyActivity.map(t => `
            <div class="m3-card" style="display:flex; align-items:center; gap: 12px; padding: 12px; margin-bottom: 8px;">
                <div class="icon-circle" style="background: var(--md-surface-variant); color: ${t.color}; width: 40px; height: 40px; flex-shrink: 0;"><span class="material-symbols-outlined" style="font-size:20px;">${t.icon}</span></div>
                <div style="flex: 1;"><strong class="large-text">${t.type}</strong><br><small style="color: var(--md-text-muted);">${t.desc}</small></div>
                <strong style="font-size: 16px; color: ${t.color};">${t.sign}&#8377;${(t.amount || 0).toFixed(2)}</strong>
            </div>
        `).join('');

        container.innerHTML = html;
    },

    renderPnL: () => {
        const startEl = document.getElementById('report-pnl-start');
        const endEl = document.getElementById('report-pnl-end');
        
        if (!startEl.value || !endEl.value) {
            const d = new Date();
            endEl.value = typeof Utils !== 'undefined' ? Utils.getLocalDate() : d.toISOString().split('T')[0];
            d.setDate(1); // Set to first of month
            startEl.value = d.toISOString().split('T')[0];
        }

        const startDate = startEl.value;
        const endDate = endEl.value;
        const container = document.getElementById('pnl-container');

        let totalRevenue = 0, totalCOGS = 0, totalExpenses = 0;

        // 1. Calculate Accrual Revenue & COGS
        UI.state.rawData.sales.forEach(s => {
            if (s.date >= startDate && s.date <= endDate && s.status !== 'Open') {
                const modifier = s.documentType === 'return' ? -1 : 1;
                const netSales = (parseFloat(s.grandTotal) || 0) - (parseFloat(s.totalGst) || 0); // Profit excludes tax
                totalRevenue += netSales * modifier;

                (s.items || []).forEach(item => {
                    const cost = parseFloat(item.buyPrice) || 0;
                    totalCOGS += ((parseFloat(item.qty) || 0) * cost) * modifier;
                });
            }
        });

        // 2. Calculate Expenses
        UI.state.rawData.expenses.forEach(e => {
            if (e.date >= startDate && e.date <= endDate) {
                totalExpenses += parseFloat(e.amount) || 0;
            }
        });

        const grossProfit = totalRevenue - totalCOGS;
        const netProfit = grossProfit - totalExpenses;
        const isProfitable = netProfit >= 0;

        container.innerHTML = `
            <div style="display:flex; justify-content:flex-end; margin-bottom: 12px; padding: 0 4px;">
                <button class="btn-primary-small tap-target" onclick="UI.exportPnLCSV()" style="display:flex; align-items:center; gap:4px; background: #e8f5e9; color: #146c2e; border: 1px solid rgba(20, 108, 46, 0.3);">
                    <span class="material-symbols-outlined" style="font-size: 16px;">download</span> Export CSV
                </button>
            </div>

            <div class="m3-card" style="padding: 16px; margin-bottom: 12px; border-left: 4px solid var(--md-success);">
                <small style="color: var(--md-text-muted);">Trading Account</small>
                <div style="display:flex; justify-content:space-between; margin-top:8px;"><span>Total Net Revenue</span><strong>&#8377;${totalRevenue.toFixed(2)}</strong></div>
                <div style="display:flex; justify-content:space-between; margin-top:4px;"><span>Cost of Goods Sold (COGS)</span><strong style="color:var(--md-error);">-&#8377;${totalCOGS.toFixed(2)}</strong></div>
                <hr style="border:0; border-top:1px dashed var(--md-outline-variant); margin: 8px 0;">
                <div style="display:flex; justify-content:space-between; font-size:16px;"><strong>Gross Profit</strong><strong style="color:var(--md-success);">&#8377;${grossProfit.toFixed(2)}</strong></div>
            </div>

            <div class="m3-card" style="padding: 16px; margin-bottom: 12px; border-left: 4px solid var(--md-error);">
                <small style="color: var(--md-text-muted);">Operating Account</small>
                <div style="display:flex; justify-content:space-between; margin-top:8px;"><span>Total Operating Expenses</span><strong>&#8377;${totalExpenses.toFixed(2)}</strong></div>
            </div>

            <div class="m3-card" style="padding: 16px; background: ${isProfitable ? '#e8f5e9' : '#fff0f2'}; border: 1px solid ${isProfitable ? 'var(--md-success)' : 'var(--md-error)'};">
                <small style="color: ${isProfitable ? 'var(--md-success)' : 'var(--md-error)'}; font-weight:bold; text-transform:uppercase;">Bottom Line Summary</small>
                <div style="display:flex; justify-content:space-between; align-items:center; margin-top:8px;">
                    <strong style="font-size:18px;">Net ${isProfitable ? 'Profit' : 'Loss'}</strong>
                    <strong style="font-size:24px; color:${isProfitable ? 'var(--md-success)' : 'var(--md-error)'};">&#8377;${Math.abs(netProfit).toFixed(2)}</strong>
                </div>
            </div>
        `;
    },

    // ==========================================
    // UPGRADE: EXPORT ENGINES (CSV)
    // ==========================================
    exportDaybookCSV: () => {
        const dateInput = document.getElementById('report-daybook-date').value;
        const dailyActivity = [];
        
        UI.state.rawData.sales.filter(s => s.date === dateInput && s.status !== 'Open').forEach(s => {
            const isRet = s.documentType === 'return';
            const isNonGST = s.invoiceType === 'Non-GST';
            const docLabel = isRet ? 'Credit Note' : (isNonGST ? 'Bill of Supply' : 'Sales Invoice');
            dailyActivity.push({ time: s.id, type: docLabel, desc: s.customerName, amount: s.grandTotal, sign: isRet ? '-' : '+' });
        });
        UI.state.rawData.purchases.filter(p => p.date === dateInput && p.status !== 'Open').forEach(p => {
            const isRet = p.documentType === 'return';
            const isNonGST = p.invoiceType === 'Non-GST';
            const docLabel = isRet ? 'Debit Note' : (isNonGST ? 'Bill of Supply' : 'Purchase Bill');
            dailyActivity.push({ time: p.id, type: docLabel, desc: p.supplierName, amount: p.grandTotal, sign: isRet ? '+' : '-' });
        });
        UI.state.rawData.cashbook.filter(c => c.date === dateInput && !c.isAutoGenerated).forEach(c => {
            const isIn = c.type === 'in';
            dailyActivity.push({ time: c.id, type: isIn ? 'Money In' : 'Money Out', desc: c.ledgerName, amount: parseFloat(c.amount), sign: isIn ? '+' : '-' });
        });
        UI.state.rawData.expenses.filter(e => e.date === dateInput).forEach(e => {
            dailyActivity.push({ time: e.id, type: 'Expense', desc: e.category, amount: parseFloat(e.amount), sign: '-' });
        });

        if (dailyActivity.length === 0) return alert("No data to export for this date.");

        // FIX: Extract the actual timestamp chunk from the end of the ID string to sort chronologically!
        dailyActivity.sort((a, b) => {
            const timeA = parseInt(a.time.split('-').pop()) || 0;
            const timeB = parseInt(b.time.split('-').pop()) || 0;
            return timeB - timeA;
        });

        let csv = "Type,Description,Amount (INR)\n";
        dailyActivity.forEach(t => {
            // Replace internal double quotes with two double quotes to safely escape them in CSV formatting
            const safeDesc = (t.desc || '').replace(/"/g, '""'); 
            csv += `"${t.type}","${safeDesc}","${t.sign}${(t.amount || 0).toFixed(2)}"\n`;
        });

        // FIX: Directly generate the CSV Blob because downloadFile does not exist in utils.js
        const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `Daybook_${dateInput}.csv`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    },

    exportPnLCSV: () => {
        const start = document.getElementById('report-pnl-start').value;
        const end = document.getElementById('report-pnl-end').value;
        
        let totalRevenue = 0, totalCOGS = 0, totalExpenses = 0;

        UI.state.rawData.sales.forEach(s => {
            if (s.date >= start && s.date <= end && s.status !== 'Open') {
                const modifier = s.documentType === 'return' ? -1 : 1;
                totalRevenue += ((parseFloat(s.grandTotal) || 0) - (parseFloat(s.totalGst) || 0)) * modifier;
                (s.items || []).forEach(item => totalCOGS += ((parseFloat(item.qty) || 0) * (parseFloat(item.buyPrice) || 0)) * modifier);
            }
        });
        UI.state.rawData.expenses.forEach(e => {
            if (e.date >= start && e.date <= end) totalExpenses += parseFloat(e.amount) || 0;
        });

        const grossProfit = totalRevenue - totalCOGS;
        const netProfit = grossProfit - totalExpenses;

        let csv = `Profit & Loss Statement (${start} to ${end})\n\n`;
        csv += `Account,Amount (INR)\n`;
        csv += `"Total Net Revenue","${totalRevenue.toFixed(2)}"\n`;
        csv += `"Cost of Goods Sold (COGS)","-${totalCOGS.toFixed(2)}"\n`;
        csv += `"Gross Profit","${grossProfit.toFixed(2)}"\n`;
        csv += `"Operating Expenses","-${totalExpenses.toFixed(2)}"\n`;
        csv += `"Net ${netProfit >= 0 ? 'Profit' : 'Loss'}","${netProfit.toFixed(2)}"\n`;

        // FIX: Directly generate the CSV Blob
        const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `PnL_${start}_to_${end}.csv`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    },

    // ==========================================
    // DATABASE OPTIMIZATION ENGINE
    // ==========================================
    optimizeMemory: () => {
        // FIX: Memory optimization disabled! Pruning arrays from RAM breaks the Address Book's running balance math. 
        // Modern devices have plenty of RAM for text arrays, so we keep all data loaded for perfect accuracy.
        return;
    }

}; // <--- MAKE SURE YOU HAVE THIS CLOSING BRACKET AND SEMICOLON!

// Bind Listeners for Live Search Filtering
document.addEventListener('DOMContentLoaded', () => {
    
    // ==========================================
    // FINAL POLISH: MATERIAL RIPPLES & HAPTICS
    // ==========================================
    document.addEventListener('pointerdown', (e) => {
        const target = e.target.closest('.tap-target, .btn-primary, .btn-primary-small, .nav-item, .list-view li, .chip');
        if (target) {
            // 1. Trigger Micro-Haptic Vibration (15ms tick for physical weight)
            if (window.navigator && window.navigator.vibrate) {
                try { window.navigator.vibrate(15); } catch(err){}
            }

            // 2. True Material Touch Ripple (Calculates exact X/Y finger coordinates)
            const rect = target.getBoundingClientRect();
            const x = e.clientX - rect.left;
            const y = e.clientY - rect.top;
            
            const ripple = document.createElement('span');
            ripple.classList.add('ripple');
            
            // Center the ripple exactly under the finger
            ripple.style.left = `${x}px`;
            ripple.style.top = `${y}px`;
            
            // FIXED: Restored the missing ripple cleanup!
            target.appendChild(ripple);
            setTimeout(() => { if(ripple.parentElement) ripple.remove(); }, 600);
        }
    });

    // UPGRADE: iOS-Style Haptic Touch Engine (Long Press)
    let pressTimer;
    let isLongPress = false;
    let startY = 0;

    document.addEventListener('touchstart', (e) => {
        const card = e.target.closest('.m3-card.tap-target');
        if (!card) return;
        
        // Only trigger on items that open forms or receipts
        const clickAction = card.getAttribute('onclick');
        if (!clickAction || (!clickAction.includes('app.openForm') && !clickAction.includes('app.openReceipt'))) return;

        isLongPress = false;
        startY = e.touches[0].clientY;

        pressTimer = setTimeout(() => {
            isLongPress = true;
            if (window.navigator && window.navigator.vibrate) window.navigator.vibrate([40, 50, 40]); // Heavy haptic bump
            UI.showContextMenu(clickAction); // Open the menu!
        }, 550); // 550ms hold time
    }, { passive: true });

    // Cancel the long-press if the user starts scrolling the screen
    document.addEventListener('touchmove', (e) => {
        if (!pressTimer) return;
        if (Math.abs(e.touches[0].clientY - startY) > 15) {
            clearTimeout(pressTimer); 
        }
    }, { passive: true });

    document.addEventListener('touchend', () => clearTimeout(pressTimer), { passive: true });

    // Intercept and stop the normal click if a long press just happened
    document.addEventListener('click', (e) => {
        if (isLongPress) {
            e.preventDefault();
            e.stopPropagation();
            isLongPress = false;
        }
    }, true);

    // UPGRADE 3: Smart Keyboard Scroll-into-View Engine (Now supports Dropdowns)
    document.addEventListener('focusin', (e) => {
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.tagName === 'SELECT') {
            setTimeout(() => {
                e.target.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }, 350); // 350ms delay ensures the physical keyboard fully slides up first
        }
    });

    // UPGRADE 4: Enter-to-Next Data Entry Engine
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && e.target.tagName === 'INPUT' && e.target.type !== 'submit') {
            e.preventDefault(); // Prevent accidental form submission
            
            // Find all focusable inputs inside the current active screen/modal
            const activeContainer = e.target.closest('.activity-screen.open') || e.target.closest('.bottom-sheet.open') || document;
            const focusable = Array.from(activeContainer.querySelectorAll('input:not([type="hidden"]):not([disabled]), select:not([disabled])'));
            
            const index = focusable.indexOf(e.target);
            if (index > -1 && index + 1 < focusable.length) {
                focusable[index + 1].focus(); // Jump to next input
                if (focusable[index + 1].select) focusable[index + 1].select(); // Highlight text for fast replacing
            } else {
                e.target.blur(); // If it's the last input, close the keyboard
            }
        }
    });

    ['sales', 'purchases', 'expenses', 'cashbook', 'timeline'].forEach(tab => {
        const searchInput = document.getElementById(`search-${tab}`);
        if(searchInput) searchInput.addEventListener('input', () => UI.applyFilters(tab));
    });
    
    // Master View Hook
    const searchMasterView = document.getElementById('search-master-view');
    if(searchMasterView) searchMasterView.addEventListener('input', () => UI.applyFilters('masters'));

    // UPGRADE 5: Smart Search Clear Buttons (Flagship UI)
    // Automatically injects a clear 'X' into every search bar in the entire app
    document.querySelectorAll('.search-bar').forEach(bar => {
        const input = bar.querySelector('input');
        if (!input) return;
        
        const clearBtn = document.createElement('span');
        clearBtn.className = 'material-symbols-outlined tap-target hidden';
        clearBtn.innerText = 'close';
        clearBtn.style.cssText = 'font-size: 16px; color: var(--md-on-surface-variant); padding: 4px; margin-left: 4px; margin-right: 4px; border-radius: 50%; background: var(--md-surface-variant); cursor: pointer;';
        
        input.parentNode.insertBefore(clearBtn, input.nextSibling);

        input.addEventListener('input', () => {
            if (input.value.length > 0) clearBtn.classList.remove('hidden');
            else clearBtn.classList.add('hidden');
        });

        clearBtn.addEventListener('click', () => {
            input.value = '';
            clearBtn.classList.add('hidden');
            input.focus(); 
            input.dispatchEvent(new Event('input')); // Instantly updates the list below it
        });
    });

    // UPGRADE 6: Swipe-to-Switch Main Tabs (Gesture Navigation)
    let touchStartX = 0;
    let touchStartY = 0;
    const mainAppShell = document.getElementById('app-shell');

    if (mainAppShell) {
        mainAppShell.addEventListener('touchstart', e => {
            touchStartX = e.changedTouches[0].screenX;
            touchStartY = e.changedTouches[0].screenY;
        }, { passive: true });

        mainAppShell.addEventListener('touchend', e => {
            const touchEndX = e.changedTouches[0].screenX;
            const touchEndY = e.changedTouches[0].screenY;
            
            // Block 1: Don't swipe if a menu, drawer, or modal is open
            if (document.querySelector('.bottom-sheet.open') || document.querySelector('.activity-screen.open') || document.getElementById('side-drawer').classList.contains('open')) return;

            // Block 2: Don't swipe if the user is scrolling horizontally on a table
            if (e.target.closest('table') || e.target.closest('.filter-chips')) return;

            const diffX = touchEndX - touchStartX;
            const diffY = touchEndY - touchStartY;
            
            // Block 3: If they are scrolling vertically (up/down), ignore the swipe
            if (Math.abs(diffY) > Math.abs(diffX)) return;

            const swipeThreshold = 80; // Minimum pixel drag to trigger a screen change
            if (Math.abs(diffX) < swipeThreshold) return;

            const tabs = [
                { id: 'tab-dashboard', title: 'Dashboard', navId: 'nav-dash' },
                { id: 'tab-masters', title: 'Master Data', navId: 'nav-masters' },
                { id: 'tab-documents', title: 'Documents', navId: 'nav-docs' },
                { id: 'tab-cashbook', title: 'Cashbook', navId: 'nav-cashbook' }
            ];

            const currentIndex = tabs.findIndex(t => document.getElementById(t.id) && document.getElementById(t.id).classList.contains('active-screen'));
            if (currentIndex === -1) return;

            if (diffX < 0 && currentIndex < tabs.length - 1) {
                // Swipe Left -> Go to Next Tab
                const nextTab = tabs[currentIndex + 1];
                UI.switchTab(nextTab.id, nextTab.title, document.getElementById(nextTab.navId));
                if (nextTab.id === 'tab-masters') UI.renderBankBalances();
            } else if (diffX > 0 && currentIndex > 0) {
                // Swipe Right -> Go to Previous Tab
                const prevTab = tabs[currentIndex - 1];
                UI.switchTab(prevTab.id, prevTab.title, document.getElementById(prevTab.navId));
                if (prevTab.id === 'tab-masters') UI.renderBankBalances();
            }
        }, { passive: true });
    }

    // UPGRADE: Auto-Hiding FAB on Scroll
    let lastScrollY = 0;
    const mainContent = document.querySelector('.main-content');
    const fab = document.querySelector('.floating-action-button');
    if (mainContent && fab) {
        mainContent.addEventListener('scroll', () => {
            const currentScrollY = mainContent.scrollTop;
            if (currentScrollY > lastScrollY && currentScrollY > 50) {
                fab.classList.add('fab-hidden'); // Hide when scrolling down
            } else {
                fab.classList.remove('fab-hidden'); // Show when scrolling up
            }
            lastScrollY = currentScrollY;
        }, {passive: true});
    }
    
    // Bottom Sheet Searches
    ['customers', 'suppliers', 'products'].forEach(type => {
        const input = document.getElementById(`search-${type}`);
        if(input) input.addEventListener('input', (e) => {
            const term = (e.target.value || '').toLowerCase();
            document.querySelectorAll(`#list-${type} li`).forEach(li => {
                li.style.display = (li.innerText || '').toLowerCase().includes(term) ? '' : 'none';
            });
        });
    });

    // (Swipe Gesture Engine Removed to prevent accidental deletions)

    // ==========================================
    // HARDWARE BACK BUTTON GESTURE INTERCEPTOR
    // ==========================================
    window.addEventListener('popstate', () => {
        // FIXED: If this was a manual button click, ignore the back gesture completely
        if (window._ignoreNextPop) {
            window._ignoreNextPop = false;
            return;
        }

        let handled = false;
        
        // 1. Check if Side Drawer is open
        const drawer = document.getElementById('side-drawer');
        if (drawer && drawer.classList.contains('open')) {
            drawer.classList.remove('open');
            const overlay = document.getElementById('side-drawer-overlay');
            if (overlay) {
                overlay.classList.remove('open');
                setTimeout(() => overlay.classList.add('hidden'), 300);
            }
            handled = true;
        }
        
        // 2. Check if Bottom Sheets are open
        if (!handled) {
            const openSheets = document.querySelectorAll('.bottom-sheet.open');
            if (openSheets.length > 0) {
                openSheets.forEach(sheet => {
                    sheet.classList.remove('open');
                    setTimeout(() => sheet.classList.add('hidden'), 300);
                });
                const overlay = document.getElementById('sheet-overlay');
                if (overlay) {
                    overlay.classList.remove('open');
                    setTimeout(() => overlay.classList.add('hidden'), 300);
                }
                if(typeof app !== 'undefined' && app.state) app.state.currentReceiptId = null;
                handled = true;
            }
        }
        
        // 3. Check if Full-Screen Activities are open
        if (!handled) {
            // FIX: Convert NodeList to an Array so we can pop the topmost screen
            const openActivities = Array.from(document.querySelectorAll('.activity-screen.open'));
            if (openActivities.length > 0) {
                // Only close the single topmost activity to prevent nuking the entire back-stack
                const topActivity = openActivities.pop(); 
                topActivity.classList.remove('open');
                setTimeout(() => topActivity.classList.add('hidden'), 300);
                
                // If we just closed a main form, wipe the active form tracker
                if (topActivity.id === 'activity-sales-form' || topActivity.id === 'activity-purchase-form') {
                    UI.state.activeActivity = null;
                }
            }
        }
    });
});

// ==========================================
// NEW CODE: ES MODULE EXPORT & GLOBAL MAP
// ==========================================
// 1. Export the module so app.js can import it
export default UI;

// 2. Attach to window so index.html inline scripts don't break
window.UI = UI;
