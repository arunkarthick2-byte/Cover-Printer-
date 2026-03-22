// ==========================================
// SOLLO ERP - GOOGLE DRIVE CLOUD SYNC
// ==========================================

const CLIENT_ID = '965559909950-a9g92ipjvdof9s38r4g4c9avkj86lphf.apps.googleusercontent.com';
const SCOPES = 'https://www.googleapis.com/auth/drive.file';
const DISCOVERY_DOC = 'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest';

let tokenClient;
let gapiInited = false;
let gisInited = false;

// 1. Initialize Google API Client
window.gapiLoaded = () => {
    gapi.load('client', async () => {
        await gapi.client.init({ discoveryDocs: [DISCOVERY_DOC] });
        gapiInited = true;
        maybeEnableButtons();
    });
};

// 2. Initialize Google Identity Services
window.gisLoaded = () => {
    tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CLIENT_ID,
        scope: SCOPES,
        callback: '', // Defined dynamically at request time
    });
    gisInited = true;
    maybeEnableButtons();
};

const maybeEnableButtons = () => {
    if (gapiInited && gisInited) {
        const statusEl = document.getElementById('cloud-status');
        if(statusEl) {
            statusEl.innerText = 'Google Drive Sync Ready';
            statusEl.style.color = 'var(--md-success)';
        }
    }
};

const Cloud = {
    authenticate: (callback) => {
        if (!gapiInited || !gisInited) {
            alert("Google Cloud Services are still loading. Please check your internet connection.");
            return;
        }
        if (gapi.client.getToken() === null) {
            // Prompt the user to log in
            tokenClient.callback = async (resp) => {
                if (resp.error !== undefined) throw (resp);
                callback();
            };
            tokenClient.requestAccessToken({ prompt: 'consent' });
        } else {
            // Already logged in
            callback();
            // Trigger auto-backup since we know we have a fresh, valid token right now
            if (Cloud.autoBackup) Cloud.autoBackup();
        }
    },

    autoBackup: async () => {
        // Only run if Google API is loaded and token is actively authorized
        if (!gapiInited || !gisInited || gapi.client.getToken() === null) return;
        
        const lastBackup = localStorage.getItem('sollo_last_backup');
        const now = Date.now();
        
        // 86400000 ms = 24 hours. Trigger if no backup in the last 24hrs
        if (!lastBackup || (now - parseInt(lastBackup)) > 86400000) {
            console.log("Triggering silent background auto-backup...");
            try {
                const data = await window.exportDatabase(); 
                const fileContent = JSON.stringify(data);
                const file = new Blob([fileContent], { type: 'application/json' });
                const metadata = { 'name': 'SOLLO_ERP_Backup.json', 'mimeType': 'application/json' };

                let response = await gapi.client.drive.files.list({
                    q: "name='SOLLO_ERP_Backup.json' and trashed=false",
                    spaces: 'drive', fields: 'files(id)'
                });

                let fileId = response.result.files.length > 0 ? response.result.files[0].id : null;
                const form = new FormData();
                form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
                form.append('file', file);

                const url = fileId ? 
                    `https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=multipart` : 
                    `https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart`;
                
                let uploadRes = await fetch(url, {
                    method: fileId ? 'PATCH' : 'POST',
                    headers: new Headers({ 'Authorization': 'Bearer ' + gapi.client.getToken().access_token }),
                    body: form,
                });

                if (uploadRes.ok) {
                    localStorage.setItem('sollo_last_backup', now.toString());
                    if (window.Utils) window.Utils.showToast("☁️ Auto-Backup Completed in Background!");
                } else if (uploadRes.status === 401) {
                    gapi.client.setToken(null);
                    if (window.Utils) window.Utils.showToast("⚠️ Cloud Auto-Sync Paused (Session Expired). Please open Data Management to reconnect.");
                }
            } catch (e) {
                console.error("Auto backup failed quietly", e);
            }
        }
    },

    backup: async () => {
        Cloud.authenticate(async () => {
            window.Utils.showToast("Preparing Backup...");
            try {
                // Get the database dump using the globally mapped function from db.js
                const data = await window.exportDatabase(); 
                const fileContent = JSON.stringify(data);
                const file = new Blob([fileContent], { type: 'application/json' });
                const metadata = {
                    'name': 'SOLLO_ERP_Backup.json',
                    'mimeType': 'application/json'
                };

                // Search for existing backup file to overwrite
                let response = await gapi.client.drive.files.list({
                    q: "name='SOLLO_ERP_Backup.json' and trashed=false",
                    spaces: 'drive',
                    fields: 'files(id, name)'
                });

                let fileId = response.result.files.length > 0 ? response.result.files[0].id : null;
                window.Utils.showToast("Uploading to Google Drive...");

                const form = new FormData();
                form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
                form.append('file', file);

                const url = fileId ? 
                    `https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=multipart` : 
                    `https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart`;
                const method = fileId ? 'PATCH' : 'POST';

                let uploadRes = await fetch(url, {
                    method: method,
                    headers: new Headers({ 'Authorization': 'Bearer ' + gapi.client.getToken().access_token }),
                    body: form,
                });

                if (uploadRes.ok) {
                    // NEW: Save the timestamp of this successful backup
                    localStorage.setItem('sollo_last_backup', Date.now().toString());
                    window.Utils.showToast("✅ Cloud Backup Successful!");
                } else {
                    if (uploadRes.status === 401) {
                        gapi.client.setToken(null); // Clear expired token to force re-auth
                        alert("Cloud session expired. Please click Backup again to re-authenticate.");
                        return;
                    }
                    throw new Error('Upload failed');
                }
            } catch (e) {
                console.error(e);
                alert("Cloud Backup Failed: " + (e.message || "Connection Error"));
            }
        });
    },

    restore: async () => {
        Cloud.authenticate(async () => {
            window.Utils.showToast("Searching Google Drive...");
            try {
                let response = await gapi.client.drive.files.list({
                    q: "name='SOLLO_ERP_Backup.json' and trashed=false",
                    spaces: 'drive',
                    fields: 'files(id, name, modifiedTime)'
                });

                if (response.result.files.length === 0) {
                    alert("No backup found on your Google Drive. Please create a backup first.");
                    return;
                }

                const fileId = response.result.files[0].id;
                const modDate = new Date(response.result.files[0].modifiedTime).toLocaleString();

                if (!confirm(`Found backup from: ${modDate}\n\nDo you want to restore this? WARNING: This will overwrite your current data on this device!`)) {
                    return;
                }

                window.Utils.showToast("Downloading Backup...");

                let fileRes = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`, {
                    headers: new Headers({ 'Authorization': 'Bearer ' + gapi.client.getToken().access_token }),
                });

                if (fileRes.ok) {
                    let jsonData = await fileRes.json();
                    window.Utils.showToast("Installing Data...");
                    
                    // Route the downloaded JSON directly into your database engine
                    await window.importDatabase(jsonData); 
                    
                    window.Utils.showToast("✅ Restore Successful! Reloading...");
                    setTimeout(() => location.reload(), 1500);
                } else {
                    // FIX: Catch expired tokens during Restore to prevent silent failures
                    if (fileRes.status === 401) {
                        gapi.client.setToken(null);
                        alert("Cloud session expired. Please click Restore again to re-authenticate.");
                        return;
                    }
                    throw new Error('Download failed');
                }
            } catch (e) {
                console.error(e);
                alert("Cloud Restore Failed: " + (e.message || "Connection Error"));
            }
        });
    }
};

window.Cloud = Cloud;
