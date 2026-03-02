// ==UserScript==
// @name         Multi-Env Chat Scraper v8.1 (Secured)
// @version      8.1.2
// @description  chat scraper and upload the data to sharepoint after assessment
// @author       bsv
// @match        https://pre-prod.amazon.com/*
// @match        https://www.amazon.com/*
// @grant        GM_xmlhttpRequest
// @connect      amazon.sharepoint.com
// @require      https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js
// ==/UserScript==

(function () {
    'use strict';

    // ════════════════════════════════════════════════════════
    //  SECURITY UTILITIES
    // ════════════════════════════════════════════════════════

    function escapeHTML(str) {
        if (!str) return '';
        const div = document.createElement('div');
        div.textContent = String(str);
        return div.innerHTML;
    }

    function safeSetText(el, text) {
        if (el) el.textContent = String(text || '');
    }

    const CryptoStore = {
        _keyCache: null,
        _salt: 'chatScraper_v8.1_salt_2024',

        async _deriveKey(passphrase) {
            if (this._keyCache) return this._keyCache;
            const enc = new TextEncoder();
            const keyMaterial = await crypto.subtle.importKey(
                'raw', enc.encode(passphrase), 'PBKDF2', false, ['deriveKey']
            );
            this._keyCache = await crypto.subtle.deriveKey(
                { name: 'PBKDF2', salt: enc.encode(this._salt), iterations: 100000, hash: 'SHA-256' },
                keyMaterial,
                { name: 'AES-GCM', length: 256 },
                false,
                ['encrypt', 'decrypt']
            );
            return this._keyCache;
        },

        async encrypt(data, passphrase) {
            try {
                const key = await this._deriveKey(passphrase);
                const enc = new TextEncoder();
                const iv = crypto.getRandomValues(new Uint8Array(12));
                const encrypted = await crypto.subtle.encrypt(
                    { name: 'AES-GCM', iv }, key, enc.encode(JSON.stringify(data))
                );
                return JSON.stringify({
                    iv: Array.from(iv),
                    data: Array.from(new Uint8Array(encrypted)),
                    _enc: true
                });
            } catch { return JSON.stringify(data); }
        },

        async decrypt(encryptedStr, passphrase) {
            try {
                const parsed = JSON.parse(encryptedStr);
                if (!parsed._enc) return parsed;
                const key = await this._deriveKey(passphrase);
                const dec = new TextDecoder();
                const decrypted = await crypto.subtle.decrypt(
                    { name: 'AES-GCM', iv: new Uint8Array(parsed.iv) },
                    key, new Uint8Array(parsed.data)
                );
                return JSON.parse(dec.decode(decrypted));
            } catch {
                try { return JSON.parse(encryptedStr); } catch { return null; }
            }
        },

        clearKeyCache() { this._keyCache = null; }
    };

    const SecureLog = {
        info(msg) { console.log(`[ChatScraper] ${msg}`); },
        warn(msg) { console.warn(`[ChatScraper] ${msg}`); },
        error(msg, error) {
            const safeMsg = error?.message
                ? error.message.replace(/password|token|digest|secret|key/gi, '[REDACTED]')
                : 'Unknown error';
            console.error(`[ChatScraper] ${msg}: ${safeMsg}`);
        }
    };

    function isValidSPUrl(url) {
        try {
            const parsed = new URL(url);
            return parsed.protocol === 'https:' &&
                (parsed.hostname.endsWith('.sharepoint.com') ||
                 parsed.hostname.endsWith('.microsoft.com'));
        } catch { return false; }
    }

    const CHAT_URL_PATTERNS = [
        /\/api\/chat\//i, /\/chat\/message/i, /\/chat\/send/i,
        /\/bot\/respond/i, /\/conversation\//i
    ];

    function isChatEndpoint(url) {
        if (!url || typeof url !== 'string') return false;
        return CHAT_URL_PATTERNS.some(p => p.test(url));
    }

    const LIMITS = {
        MAX_CHAT_ENTRIES: 500,
        MAX_QUEUE_SIZE: 50,
        MAX_STRING_LENGTH: 10000,
        MAX_IMAGE_SIZE_KB: 8192,
        DATA_RETENTION_HOURS: 24
    };

    function truncateString(str, max = LIMITS.MAX_STRING_LENGTH) {
        if (!str) return '';
        return String(str).substring(0, max);
    }

    // ════════════════════════════════════════════════════════
    //  ENVIRONMENT DEFINITIONS
    // ════════════════════════════════════════════════════════
    const ENVIRONMENTS = {
        'pre-prod': {
            id: 'pre-prod', label: 'Pre-Prod', shortLabel: 'PP', color: '#FF9900',
            gradient: 'linear-gradient(135deg, #FF9900, #e88b00)',
            headerGradient: 'linear-gradient(135deg, #232F3E, #37475A)',
            borderColor: '#FF9900',
            listPrefix: 'ABC_Pre-Prod_Environment_Tool_Feedback_Data',
            siteUrl: 'https://amazon.sharepoint.com/sites/Pre-ProdDataDump',
            hasAreaOfImprovement: false, autoUpload: true
        },
        'regression': {
            id: 'regression', label: 'Regression Test', shortLabel: 'RT', color: '#6f42c1',
            gradient: 'linear-gradient(135deg, #6f42c1, #5a32a3)',
            headerGradient: 'linear-gradient(135deg, #2d1b69, #4a2d8a)',
            borderColor: '#6f42c1',
            listPrefix: 'ABC_Regression_Test_Tool_Feedback_Data',
            siteUrl: 'https://amazon.sharepoint.com/sites/Pre-ProdDataDump',
            hasAreaOfImprovement: true, autoUpload: true
        },
        'prod': {
            id: 'prod', label: 'Prod Environment', shortLabel: 'PR', color: '#dc3545',
            gradient: 'linear-gradient(135deg, #dc3545, #c82333)',
            headerGradient: 'linear-gradient(135deg, #4a0e0e, #721c24)',
            borderColor: '#dc3545',
            listPrefix: 'ABC_Prod_Environment_Tool_Feedback_Data',
            siteUrl: 'https://amazon.sharepoint.com/sites/Pre-ProdDataDump',
            hasAreaOfImprovement: false, autoUpload: true
        }
    };

    Object.values(ENVIRONMENTS).forEach(env => {
        if (!isValidSPUrl(env.siteUrl)) {
            SecureLog.error(`Invalid SP URL for ${env.id}`, new Error('Bad URL'));
        }
    });

    // ════════════════════════════════════════════════════════
    //  STATE
    // ════════════════════════════════════════════════════════
    let activeEnvId = 'pre-prod';
    let chatData = [];
    let pendingRequest = null;
    let currentUsername = null;
    let cachedElements = {};
    let savedTestAccount = {};
    let responseQueue = [];
    let dailyResetTimer = null;

    let cachedFieldMapping = {};
    let cachedEntityType = {};
    let cachedDigest = null;
    let cachedDigestTime = 0;
    const DIGEST_LIFETIME = 600000;
    let siteAssetsFolderVerified = {};

    const TAB_ID = `tab_${Date.now()}_${Math.random().toString(36).substr(2, 8)}`;

    const CONFIG = {
        PDF_BATCH_SIZE: 200,
        JPEG_QUALITY: 0.92,
        MAX_IMG_DIM: 1920,
        IMAGE_PROCESS_DELAY: 50,
        SP_ATTACH_MAX_BYTES: 4 * 1024 * 1024,
        SP_SCREENSHOTS_FOLDER: 'PreProdScreenshots',
    };

    // ════════════════════════════════════════════════════════
    //  ENCRYPTION PASSPHRASE
    // ════════════════════════════════════════════════════════
    function getEncryptionKey() {
        return `scraper_${currentUsername || 'default'}_${window.location.hostname}`;
    }

    // ════════════════════════════════════════════════════════
    //  ENVIRONMENT ACCESS RULES
    // ════════════════════════════════════════════════════════
    function getAllowedEnvironments() {
        const url = window.location.href.toLowerCase();
        if (url.includes('pre-prod')) return { 'pre-prod': true, 'regression': true, 'prod': false };
        if (url.includes('regression')) return { 'pre-prod': true, 'regression': true, 'prod': false };
        return { 'pre-prod': false, 'regression': false, 'prod': true };
    }

    function isEnvAllowed(envId) { return getAllowedEnvironments()[envId] === true; }

    // ════════════════════════════════════════════════════════
    //  ENVIRONMENT HELPERS
    // ════════════════════════════════════════════════════════
    function getActiveEnv() { return ENVIRONMENTS[activeEnvId] || ENVIRONMENTS['pre-prod']; }

    function getListNameForEnv(envId) {
        const env = ENVIRONMENTS[envId] || ENVIRONMENTS['pre-prod'];
        const now = new Date();
        return `${env.listPrefix}_${now.getFullYear()}_${String(now.getMonth() + 1).padStart(2, '0')}`;
    }

    function getSPListUrl() { return getActiveEnv().siteUrl; }
    function getSPListName() { return getListNameForEnv(activeEnvId); }
    function getAutoUpload() { return getActiveEnv().autoUpload; }

    // ════════════════════════════════════════════════════════
    //  STORAGE KEYS
    // ════════════════════════════════════════════════════════
    const GLOBAL_KEYS = {
        username: 'chatScraperUsername',
        testAccount: 'chatScraperTestAccount',
        lastResetDate: 'chatScraper_lastResetDate',
        activeEnv: 'chatScraper_activeEnv'
    };

    function envKey(base) { return `chatScraper_${activeEnvId}_${base}`; }

    // ════════════════════════════════════════════════════════
    //  SP CACHE
    // ════════════════════════════════════════════════════════
    function clearSPCache(envId) {
        const eid = envId || activeEnvId;
        delete cachedFieldMapping[eid];
        delete cachedEntityType[eid];
        delete siteAssetsFolderVerified[eid];
        cachedDigest = null; cachedDigestTime = 0;
        localStorage.removeItem(`chatScraper_${eid}_spFieldMap`);
    }

    function clearAllSPCaches() {
        cachedFieldMapping = {}; cachedEntityType = {};
        cachedDigest = null; cachedDigestTime = 0;
        siteAssetsFolderVerified = {};
        Object.keys(ENVIRONMENTS).forEach(eid => {
            localStorage.removeItem(`chatScraper_${eid}_spFieldMap`);
        });
    }

    function saveCachedFieldMap(mapping) {
        cachedFieldMapping[activeEnvId] = mapping;
        try { localStorage.setItem(envKey('spFieldMap'), JSON.stringify(mapping)); } catch {}
    }

    function loadCachedFieldMap() {
        if (cachedFieldMapping[activeEnvId]) return cachedFieldMapping[activeEnvId];
        try {
            const s = localStorage.getItem(envKey('spFieldMap'));
            if (s) { cachedFieldMapping[activeEnvId] = JSON.parse(s); return cachedFieldMapping[activeEnvId]; }
        } catch {}
        return null;
    }

    // ════════════════════════════════════════════════════════
    //  ENCRYPTED PERSISTENCE
    // ════════════════════════════════════════════════════════
    async function saveQueueToStorage() {
        try {
            const data = responseQueue.map(i => ({ ...i, _savedAt: i._savedAt || Date.now() }));
            const encrypted = await CryptoStore.encrypt(data, getEncryptionKey());
            localStorage.setItem(envKey('responseQueue'), encrypted);
        } catch {}
    }

    async function loadQueueFromStorage() {
        try {
            const s = localStorage.getItem(envKey('responseQueue'));
            if (s) {
                const p = await CryptoStore.decrypt(s, getEncryptionKey());
                if (Array.isArray(p)) {
                    responseQueue = p.filter(i =>
                        (Date.now() - (i._savedAt || 0)) < (LIMITS.DATA_RETENTION_HOURS * 3600000)
                    );
                } else { responseQueue = []; }
            } else { responseQueue = []; }
        } catch { responseQueue = []; }
    }

    async function saveDataToStorage() {
        try {
            const data = chatData.map(i => { const o = { ...i }; delete o.attachment; return o; });
            const encrypted = await CryptoStore.encrypt(data, getEncryptionKey());
            localStorage.setItem(envKey('savedData'), encrypted);
        } catch {}
    }

    async function loadDataFromStorage() {
        try {
            const s = localStorage.getItem(envKey('savedData'));
            if (s) {
                const parsed = await CryptoStore.decrypt(s, getEncryptionKey());
                if (Array.isArray(parsed)) {
                    chatData = parsed;
                    chatData.forEach(i => delete i.attachment);
                } else { chatData = []; }
            } else { chatData = []; }
        } catch { chatData = []; }
    }

    // ════════════════════════════════════════════════════════
    //  SECURE TEST ACCOUNT
    // ════════════════════════════════════════════════════════
    function loadSavedTestAccount() {
        try {
            const s = localStorage.getItem(GLOBAL_KEYS.testAccount);
            if (s) {
                const parsed = JSON.parse(s);
                delete parsed.password;
                savedTestAccount = parsed;
            }
        } catch { savedTestAccount = {}; }
    }

    function saveTestAccount(d) {
        const safe = {
            email: truncateString(d.email || '', 255),
            customerId: truncateString(d.customerId || '', 255),
            gcsLink: truncateString(d.gcsLink || '', 500),
            gcsLinkType: truncateString(d.gcsLinkType || '', 255)
        };
        savedTestAccount = safe;
        localStorage.setItem(GLOBAL_KEYS.testAccount, JSON.stringify(safe));
    }

    // ════════════════════════════════════════════════════════
    //  ENVIRONMENT SWITCHING
    // ════════════════════════════════════════════════════════
    async function switchEnvironment(newEnvId) {
        if (!ENVIRONMENTS[newEnvId] || newEnvId === activeEnvId) return;
        if (!isEnvAllowed(newEnvId)) {
            showNotification(`${ENVIRONMENTS[newEnvId].label} is not available on this site`, 'error');
            return;
        }

        if (chatData.length > 0) {
            const oldEnv = getActiveEnv();
            showNotification(`[${oldEnv.shortLabel}] Auto-downloading ${chatData.length} entries...`, 'info');
            try {
                const buffer = await generateExcelBuffer(chatData);
                const excelName = `${oldEnv.id}_Results_${currentUsername}_${new Date().toISOString().slice(0, 10)}.xlsx`;
                const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url; link.download = excelName;
                document.body.appendChild(link); link.click(); document.body.removeChild(link);
                setTimeout(() => URL.revokeObjectURL(url), 5000);
            } catch (e) { SecureLog.error('AutoDownload Excel failed', e); }

            await new Promise(r => setTimeout(r, 1000));

            try {
                const imgEntries = chatData.filter(d => d._hasImage && d.entryId);
                if (imgEntries.length > 0) {
                    let hasImages = false;
                    for (const item of imgEntries) { if (await ImageStore.has(item.entryId)) { hasImages = true; break; } }
                    if (hasImages) await downloadPDFSilent(chatData, oldEnv);
                }
            } catch (e) { SecureLog.error('AutoDownload PDF failed', e); }

            showNotification(`[${oldEnv.shortLabel}] Files downloaded`, 'success');
            await new Promise(r => setTimeout(r, 500));
        }

        await saveQueueToStorage();
        await saveDataToStorage();

        activeEnvId = newEnvId;
        localStorage.setItem(GLOBAL_KEYS.activeEnv, newEnvId);

        await loadQueueFromStorage();
        await loadDataFromStorage();

        cachedDigest = null; cachedDigestTime = 0;

        clearElementCache();
        rebuildPanelForEnv();
        refreshPanelState();
        updateStatus();
        updateBubble();

        const env = getActiveEnv();
        showNotification(`Switched to ${env.label} (${chatData.length} entries, ${responseQueue.length} pending)`, 'info');
    }

    // ════════════════════════════════════════════════════════
    //  IMAGE UTILITIES
    // ════════════════════════════════════════════════════════
    function base64ToArrayBuffer(base64DataUrl) {
        const base64 = base64DataUrl.split(',')[1];
        if (!base64) throw new Error('Invalid base64 data');
        const bin = atob(base64);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        return bytes.buffer;
    }

    function validateImageDataUrl(dataUrl) {
        if (!dataUrl || typeof dataUrl !== 'string') return false;
        if (!dataUrl.startsWith('data:image/')) return false;
        const sizeKB = Math.round(dataUrl.length * 0.75 / 1024);
        return sizeKB <= LIMITS.MAX_IMAGE_SIZE_KB;
    }

    function compressImageForAttachment(dataUrl, maxBytes = CONFIG.SP_ATTACH_MAX_BYTES) {
        if (!validateImageDataUrl(dataUrl)) return Promise.reject(new Error('Invalid image data'));
        return new Promise(resolve => {
            const img = new Image();
            img.onload = () => {
                const c = document.createElement('canvas');
                const ctx = c.getContext('2d');
                let w = img.naturalWidth, h = img.naturalHeight;
                if (w > 1600 || h > 1600) { const s = Math.min(1600 / w, 1600 / h); w = Math.round(w * s); h = Math.round(h * s); }
                c.width = w; c.height = h;
                ctx.fillStyle = '#FFF'; ctx.fillRect(0, 0, w, h);
                ctx.drawImage(img, 0, 0, w, h);
                let q = 0.85, r = c.toDataURL('image/jpeg', q), est = Math.round(r.length * 0.75);
                while (est > maxBytes && q > 0.3) { q -= 0.1; r = c.toDataURL('image/jpeg', q); est = Math.round(r.length * 0.75); }
                c.width = 0; c.height = 0;
                resolve({ dataUrl: r, width: w, height: h, quality: q, sizeKB: Math.round(est / 1024), format: 'jpeg' });
            };
            img.onerror = () => resolve({ dataUrl, width: 800, height: 600, quality: 0, sizeKB: 0, format: 'unknown' });
            img.src = dataUrl;
        });
    }

    function processImageForStorage(dataUrl) {
        if (!validateImageDataUrl(dataUrl)) return Promise.reject(new Error('Invalid image data'));
        return new Promise(res => {
            const img = new Image();
            img.onload = () => {
                const ow = img.naturalWidth, oh = img.naturalHeight;
                const c = document.createElement('canvas');
                const ctx = c.getContext('2d');
                let tw = ow, th = oh;
                if (tw > 3840 || th > 3840) { const s = Math.min(3840 / tw, 3840 / th); tw = Math.round(tw * s); th = Math.round(th * s); }
                c.width = tw; c.height = th;
                ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = 'high';
                ctx.drawImage(img, 0, 0, tw, th);
                const d = c.toDataURL('image/png');
                c.width = 0; c.height = 0;
                res({ dataUrl: d, format: 'png', originalWidth: ow, originalHeight: oh, processedWidth: tw, processedHeight: th, fileSize: Math.round(d.length * 0.75 / 1024) });
            };
            img.onerror = () => res({ dataUrl, format: 'png', originalWidth: 800, originalHeight: 600, processedWidth: 800, processedHeight: 600, fileSize: 0 });
            img.src = dataUrl;
        });
    }

    function compressImageForPDF(dataUrl) {
        return new Promise(res => {
            const img = new Image();
            img.onload = () => {
                const c = document.createElement('canvas');
                const ctx = c.getContext('2d');
                let tw = img.naturalWidth, th = img.naturalHeight;
                if (tw > CONFIG.MAX_IMG_DIM || th > CONFIG.MAX_IMG_DIM) { const s = Math.min(CONFIG.MAX_IMG_DIM / tw, CONFIG.MAX_IMG_DIM / th); tw = Math.round(tw * s); th = Math.round(th * s); }
                c.width = tw; c.height = th;
                ctx.fillStyle = '#FFF'; ctx.fillRect(0, 0, tw, th);
                ctx.drawImage(img, 0, 0, tw, th);
                const d = c.toDataURL('image/jpeg', CONFIG.JPEG_QUALITY);
                c.width = 0; c.height = 0;
                res({ dataUrl: d, width: tw, height: th, compressedKB: Math.round(d.length * 0.75 / 1024) });
            };
            img.onerror = () => res({ dataUrl, width: 800, height: 600, compressedKB: 0 });
            img.src = dataUrl;
        });
    }

    // ════════════════════════════════════════════════════════
    //  SHAREPOINT LIST ENGINE
    // ════════════════════════════════════════════════════════
    const SharePointList = {

        async getFormDigest(siteUrl) {
            if (!isValidSPUrl(siteUrl)) throw new Error('Invalid SharePoint URL');
            if (cachedDigest && (Date.now() - cachedDigestTime) < DIGEST_LIFETIME) return cachedDigest;
            const errors = [];
            try {
                const digest = await new Promise((resolve, reject) => {
                    GM_xmlhttpRequest({
                        method: 'POST', url: `${siteUrl}/_api/contextinfo`,
                        headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'X-Requested-With': 'XMLHttpRequest' },
                        anonymous: false, withCredentials: true, timeout: 15000,
                        onload: r => { if (r.status === 200) { try { const v = JSON.parse(r.responseText).d?.GetContextWebInformation?.FormDigestValue; v ? resolve(v) : reject(new Error('No digest value')); } catch (e) { reject(e); } } else { reject(new Error(`HTTP ${r.status}`)); } },
                        onerror: () => reject(new Error('Network error')),
                        ontimeout: () => reject(new Error('Timeout'))
                    });
                });
                cachedDigest = digest; cachedDigestTime = Date.now(); return digest;
            } catch (e) { errors.push('GM: ' + e.message); }
            try {
                const resp = await fetch(`${siteUrl}/_api/contextinfo`, { method: 'POST', credentials: 'include', headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' } });
                if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
                const d = await resp.json();
                const digest = d.d?.GetContextWebInformation?.FormDigestValue;
                if (!digest) throw new Error('No digest');
                cachedDigest = digest; cachedDigestTime = Date.now(); return digest;
            } catch (e) { errors.push('Fetch: ' + e.message); }
            throw new Error('Auth failed');
        },

        async getListEntityType(siteUrl, listName, digest) {
            if (!isValidSPUrl(siteUrl)) throw new Error('Invalid SP URL');
            if (cachedEntityType[activeEnvId]) return cachedEntityType[activeEnvId];
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'GET', url: `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')?$select=ListItemEntityTypeFullName`,
                    headers: { 'Accept': 'application/json;odata=verbose', 'X-RequestDigest': digest, 'X-Requested-With': 'XMLHttpRequest' },
                    anonymous: false, withCredentials: true, timeout: 10000,
                    onload: r => { if (r.status === 200) { try { const et = JSON.parse(r.responseText).d.ListItemEntityTypeFullName; cachedEntityType[activeEnvId] = et; resolve(et); } catch { reject(new Error('Parse error')); } } else { reject(new Error(`HTTP ${r.status}`)); } },
                    onerror: () => reject(new Error('Network')),
                    ontimeout: () => reject(new Error('Timeout'))
                });
            });
        },

        async getListFields(siteUrl, listName, digest) {
            if (!isValidSPUrl(siteUrl)) throw new Error('Invalid SP URL');
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'GET', url: `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/fields?$filter=Hidden eq false and ReadOnlyField eq false&$select=Title,InternalName,TypeAsString,Required&$top=200`,
                    headers: { 'Accept': 'application/json;odata=verbose', 'X-RequestDigest': digest, 'X-Requested-With': 'XMLHttpRequest' },
                    anonymous: false, withCredentials: true, timeout: 10000,
                    onload: r => { if (r.status === 200) { try { resolve(JSON.parse(r.responseText).d.results.map(f => ({ title: f.Title, internal: f.InternalName, type: f.TypeAsString, required: f.Required }))); } catch (e) { reject(e); } } else { reject(new Error(`HTTP ${r.status}`)); } },
                    onerror: () => reject(new Error('Network')),
                    ontimeout: () => reject(new Error('Timeout'))
                });
            });
        },

        async buildFieldMapping(siteUrl, listName, digest) {
            const cached = loadCachedFieldMap();
            if (cached) return cached;
            const fields = await this.getListFields(siteUrl, listName, digest);
            const lookup = {}, fieldTypeMap = {};
            fields.forEach(f => {
                lookup[f.title.toLowerCase()] = f.internal;
                lookup[f.internal.toLowerCase()] = f.internal;
                lookup[f.title.replace(/[_\s]/g, '').toLowerCase()] = f.internal;
                lookup[f.internal.replace(/[_\s]/g, '').toLowerCase()] = f.internal;
                fieldTypeMap[f.internal] = f.type;
            });
            const fieldDefs = {
                hva: ['HVA', 'hva'],
                messageId: ['Message_Id', 'MessageId', 'Message Id'],
                conversationId: ['Conversation_Id', 'ConversationId', 'Conversation Id'],
                query: ['Query', 'query', 'UserQuery'],
                botResponse: ['BotResponse', 'Bot_Response', 'Bot Response'],
                responseTime: ['Response_Time', 'ResponseTime', 'Response Time'],
                status: ['Status', 'status', 'Result'],
                groundTruth: ['GroundTruth', 'Ground_Truth', 'Ground Truth'],
                observations: ['Observations', 'observations', 'Notes'],
                areaOfImprovement: ['Area_of_Improvement', 'AreaOfImprovement', 'Area of Improvement'],
                gcsLinkType: ['GCS_MCS_Type', 'GCSMCSType', 'GCS MCS Type', 'Link_Type', 'LinkType'],
                gcsLink: ['GCS_MCS_Link', 'GCSLink', 'GCS Link'],
                testEmail: ['Test_Email', 'TestEmail', 'Test Email'],
                testCustomerId: ['Test_Customer_Id', 'TestCustomerId', 'Test Customer Id'],
                testingDate: ['Testing_Date', 'TestingDate', 'Testing Date'],
                testerLogin: ['Tester_Login', 'TesterLogin', 'Tester Login', 'Tester'],
                queryTimestamp: ['Query_Timestamp', 'QueryTimestamp', 'Query Timestamp'],
                savedTimestamp: ['Saved_Timestamp', 'SavedTimestamp', 'Saved Timestamp'],
                screenshot: ['Screenshot', 'screenshot', 'ScreenshotUrl', 'Image', 'Picture'],
                environment: ['Environment', 'environment', 'Env', 'TestEnvironment']
            };
            const mapping = {};
            for (const [key, names] of Object.entries(fieldDefs)) {
                let found = false;
                for (const name of names) {
                    const internal = lookup[name.toLowerCase()] || lookup[name.replace(/[_\s]/g, '').toLowerCase()];
                    if (internal) { mapping[key] = internal; found = true; break; }
                }
                if (!found) mapping[key] = names[0];
            }
            mapping._screenshotType = fieldTypeMap[mapping.screenshot] || 'URL';
            mapping._screenshotFound = !!fieldTypeMap[mapping.screenshot];
            saveCachedFieldMap(mapping);
            return mapping;
        },

        async ensureScreenshotFolder(siteUrl, digest) {
            if (!isValidSPUrl(siteUrl)) throw new Error('Invalid SP URL');
            if (siteAssetsFolderVerified[activeEnvId]) return true;
            const fp = `SiteAssets/${CONFIG.SP_SCREENSHOTS_FOLDER}`;
            const exists = await new Promise(res => {
                GM_xmlhttpRequest({
                    method: 'GET', url: `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(fp)}')`,
                    headers: { 'Accept': 'application/json;odata=verbose', 'X-RequestDigest': digest, 'X-Requested-With': 'XMLHttpRequest' },
                    anonymous: false, withCredentials: true, timeout: 10000,
                    onload: r => res(r.status === 200), onerror: () => res(false), ontimeout: () => res(false)
                });
            });
            if (exists) { siteAssetsFolderVerified[activeEnvId] = true; return true; }
            await new Promise(res => {
                GM_xmlhttpRequest({
                    method: 'POST', url: `${siteUrl}/_api/web/folders`,
                    headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'X-RequestDigest': digest, 'X-Requested-With': 'XMLHttpRequest' },
                    data: JSON.stringify({ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': 'SiteAssets' }),
                    anonymous: false, withCredentials: true, timeout: 10000,
                    onload: () => res(), onerror: () => res(), ontimeout: () => res()
                });
            });
            const created = await new Promise(res => {
                GM_xmlhttpRequest({
                    method: 'POST', url: `${siteUrl}/_api/web/folders`,
                    headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'X-RequestDigest': digest, 'X-Requested-With': 'XMLHttpRequest' },
                    data: JSON.stringify({ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': fp }),
                    anonymous: false, withCredentials: true, timeout: 10000,
                    onload: r => res(r.status === 200 || r.status === 201), onerror: () => res(false), ontimeout: () => res(false)
                });
            });
            if (created) siteAssetsFolderVerified[activeEnvId] = true;
            return created;
        },

        async uploadImageToSiteAssets(siteUrl, imageDataUrl, fileName, digest) {
            if (!isValidSPUrl(siteUrl)) throw new Error('Invalid SP URL');
            if (!validateImageDataUrl(imageDataUrl)) throw new Error('Invalid image');
            const comp = await compressImageForAttachment(imageDataUrl);
            const buf = base64ToArrayBuffer(comp.dataUrl);
            await this.ensureScreenshotFolder(siteUrl, digest);
            const fp = `SiteAssets/${CONFIG.SP_SCREENSHOTS_FOLDER}`;
            const safeName = fileName.replace(/[^a-zA-Z0-9._-]/g, '_');
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'POST', url: `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(fp)}')/Files/add(url='${encodeURIComponent(safeName)}',overwrite=true)`,
                    headers: { 'Accept': 'application/json;odata=verbose', 'X-RequestDigest': digest, 'Content-Type': 'application/octet-stream', 'X-Requested-With': 'XMLHttpRequest' },
                    data: buf, anonymous: false, withCredentials: true, timeout: 30000,
                    onload: r => {
                        if (r.status === 200 || r.status === 201) {
                            try { const d = JSON.parse(r.responseText); resolve({ success: true, serverRelativeUrl: d.d.ServerRelativeUrl, absoluteUrl: `${new URL(siteUrl).origin}${d.d.ServerRelativeUrl}`, fileName: d.d.Name, sizeKB: comp.sizeKB }); }
                            catch { resolve({ success: true, absoluteUrl: '', sizeKB: comp.sizeKB }); }
                        } else { let m = `HTTP ${r.status}`; try { m = JSON.parse(r.responseText).error?.message?.value || m; } catch {} reject(new Error(m)); }
                    },
                    onerror: () => reject(new Error('Network error')),
                    ontimeout: () => reject(new Error('Timeout'))
                });
            });
        },

        async updateItemScreenshot(siteUrl, listName, itemId, imageUrl, mapping, entityType, digest) {
            if (!isValidSPUrl(siteUrl)) throw new Error('Invalid SP URL');
            const col = mapping.screenshot;
            const colType = mapping._screenshotType || 'URL';
            if (!mapping._screenshotFound) return { success: false, error: 'No screenshot column' };
            const body = { '__metadata': { 'type': entityType } };
            if (colType === 'URL') {
                body[col] = { '__metadata': { 'type': 'SP.FieldUrlValue' }, 'Url': imageUrl, 'Description': `Screenshot_${itemId}` };
            } else if (colType === 'Thumbnail' || colType === 'Image') {
                body[col] = JSON.stringify({ type: 'thumbnail', fileName: `Screenshot_${itemId}.jpg`, nativeFile: {}, fieldName: col, serverUrl: new URL(siteUrl).origin, serverRelativeUrl: imageUrl.replace(new URL(siteUrl).origin, '') });
            } else { body[col] = imageUrl; }
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'POST', url: `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${itemId})`,
                    headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'X-RequestDigest': digest, 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE', 'X-Requested-With': 'XMLHttpRequest' },
                    data: JSON.stringify(body), anonymous: false, withCredentials: true, timeout: 15000,
                    onload: r => { if (r.status === 204 || r.status === 200) { resolve({ success: true }); } else { let m = `HTTP ${r.status}`; try { m = JSON.parse(r.responseText).error?.message?.value || m; } catch {} if (m.includes('does not exist')) clearSPCache(); reject(new Error(m)); } },
                    onerror: () => reject(new Error('Network')),
                    ontimeout: () => reject(new Error('Timeout'))
                });
            });
        },

        async findItemByTitle(siteUrl, listName, title, digest) {
            if (!title || !isValidSPUrl(siteUrl)) return null;
            return new Promise(resolve => {
                GM_xmlhttpRequest({
                    method: 'GET', url: `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items?$filter=Title eq '${encodeURIComponent(title)}'&$select=Id,Title&$top=1`,
                    headers: { 'Accept': 'application/json;odata=verbose', 'X-RequestDigest': digest, 'X-Requested-With': 'XMLHttpRequest' },
                    anonymous: false, withCredentials: true, timeout: 10000,
                    onload: r => { if (r.status === 200) { try { const res = JSON.parse(r.responseText).d.results; resolve(res?.length > 0 ? res[0].Id : null); } catch { resolve(null); } } else { resolve(null); } },
                    onerror: () => resolve(null), ontimeout: () => resolve(null)
                });
            });
        },

        async _uploadImageForItem(siteUrl, listName, itemId, imageDataUrl, entryData, mapping, entityType, digest) {
            if (!imageDataUrl || !validateImageDataUrl(imageDataUrl)) return null;
            try {
                const fn = `Screenshot_${activeEnvId}_${(entryData.entryId || itemId).replace(/[^a-zA-Z0-9_-]/g, '_')}_${Date.now()}.jpg`;
                const up = await this.uploadImageToSiteAssets(siteUrl, imageDataUrl, fn, digest);
                if (up.success && up.absoluteUrl) {
                    try { await this.updateItemScreenshot(siteUrl, listName, itemId, up.absoluteUrl, mapping, entityType, digest); return { success: true, url: up.absoluteUrl, fileName: up.fileName, sizeKB: up.sizeKB }; }
                    catch { return { success: false, url: up.absoluteUrl, error: 'Column update failed', uploaded: true }; }
                }
            } catch { return { success: false, error: 'Upload failed' }; }
            return null;
        },

        async checkListExists(siteUrl, listName, digest) {
            if (!isValidSPUrl(siteUrl)) return false;
            return new Promise(resolve => {
                GM_xmlhttpRequest({
                    method: 'GET', url: `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')`,
                    headers: { 'Accept': 'application/json;odata=verbose', 'X-RequestDigest': digest, 'X-Requested-With': 'XMLHttpRequest' },
                    anonymous: false, withCredentials: true, timeout: 10000,
                    onload: r => resolve(r.status === 200), onerror: () => resolve(false), ontimeout: () => resolve(false)
                });
            });
        },

        async createList(siteUrl, listName, digest) {
            if (!isValidSPUrl(siteUrl)) throw new Error('Invalid SP URL');
            const env = getActiveEnv();
            await new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'POST', url: `${siteUrl}/_api/web/lists`,
                    headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'X-RequestDigest': digest, 'X-Requested-With': 'XMLHttpRequest' },
                    data: JSON.stringify({ '__metadata': { 'type': 'SP.List' }, 'Title': listName, 'BaseTemplate': 100, 'AllowContentTypes': true, 'ContentTypesEnabled': false, 'Description': `Auto-created by Chat Scraper v8.1 for ${env.label}` }),
                    anonymous: false, withCredentials: true, timeout: 15000,
                    onload: r => { if (r.status === 201 || r.status === 200) resolve(); else { let msg = `HTTP ${r.status}`; try { msg = JSON.parse(r.responseText).error?.message?.value || msg; } catch {} reject(new Error(msg)); } },
                    onerror: () => reject(new Error('Network error')),
                    ontimeout: () => reject(new Error('Timeout'))
                });
            });
            const columns = [
                { name: 'HVA', type: 2 }, { name: 'Message_Id', type: 2 }, { name: 'Conversation_Id', type: 2 },
                { name: 'Query', type: 3 }, { name: 'BotResponse', type: 3 }, { name: 'Response_Time', type: 2 },
                { name: 'Status', type: 2 }, { name: 'GroundTruth', type: 3 }, { name: 'Observations', type: 3 },
                { name: 'GCS_MCS_Type', type: 2 },
                { name: 'GCS_MCS_Link', type: 2 }, { name: 'Test_Email', type: 2 }, { name: 'Test_Customer_Id', type: 2 },
                { name: 'Testing_Date', type: 2 }, { name: 'Tester_Login', type: 2 }, { name: 'Query_Timestamp', type: 2 },
                { name: 'Saved_Timestamp', type: 2 }, { name: 'Screenshot', type: 11 }
            ];
            if (env.hasAreaOfImprovement) columns.push({ name: 'Area_of_Improvement', type: 3 });

            for (const col of columns) {
                try {
                    await new Promise((resolve) => {
                        GM_xmlhttpRequest({
                            method: 'POST', url: `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/fields`,
                            headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'X-RequestDigest': digest, 'X-Requested-With': 'XMLHttpRequest' },
                            data: JSON.stringify(col.type === 11 ? { '__metadata': { 'type': 'SP.FieldUrl' }, 'Title': col.name, 'FieldTypeKind': col.type, 'Required': false, 'EnforceUniqueValues': false, 'StaticName': col.name, 'InternalName': col.name, 'DisplayFormat': 1 } : { '__metadata': { 'type': 'SP.Field' }, 'Title': col.name, 'FieldTypeKind': col.type, 'Required': false, 'EnforceUniqueValues': false, 'StaticName': col.name, 'InternalName': col.name }),
                            anonymous: false, withCredentials: true, timeout: 10000,
                            onload: () => resolve(), onerror: () => resolve(), ontimeout: () => resolve()
                        });
                    });
                    await new Promise(r => setTimeout(r, 300));
                } catch {}
            }
            try {
                for (const col of columns) {
                    await new Promise(resolve => {
                        GM_xmlhttpRequest({
                            method: 'POST', url: `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/views/getbytitle('All Items')/viewfields/addviewfield('${col.name}')`,
                            headers: { 'Accept': 'application/json;odata=verbose', 'X-RequestDigest': digest, 'X-Requested-With': 'XMLHttpRequest' },
                            anonymous: false, withCredentials: true, timeout: 5000,
                            onload: () => resolve(), onerror: () => resolve(), ontimeout: () => resolve()
                        });
                    });
                }
            } catch {}
            clearSPCache();
        },

        async ensureListExists(siteUrl, listName, digest) {
            const cacheKey = `listVerified_${listName}`;
            if (this[cacheKey]) return true;
            const exists = await this.checkListExists(siteUrl, listName, digest);
            if (exists) { this[cacheKey] = true; return true; }
            showAutoUploadStatus(`Creating new list: ${listName}...`, 'uploading');
            try {
                await this.createList(siteUrl, listName, digest);
                this[cacheKey] = true;
                showAutoUploadStatus(`List created: ${listName}`, 'success');
                showNotification(`Auto-created SP list: ${listName}`, 'success');
                return true;
            } catch (e) {
                showAutoUploadStatus('Failed to create list', 'error');
                throw new Error(`Cannot create list: ${e.message}`);
            }
        },

        async addItem(siteUrl, listName, entryData, imageDataUrl = null) {
            if (!isValidSPUrl(siteUrl)) throw new Error('Invalid SP URL');
            const digest = await this.getFormDigest(siteUrl);
            await this.ensureListExists(siteUrl, listName, digest);
            const entityType = await this.getListEntityType(siteUrl, listName, digest);
            const mapping = await this.buildFieldMapping(siteUrl, listName, digest);
            const env = getActiveEnv();

            if (entryData._spItemId && entryData._spItemId !== 'unknown') {
                let imgRes = null;
                if (imageDataUrl && !entryData._spImageAttached) {
                    imgRes = await this._uploadImageForItem(siteUrl, listName, entryData._spItemId, imageDataUrl, entryData, mapping, entityType, digest);
                } else { imgRes = entryData._spImageAttached ? { success: true, alreadyDone: true } : null; }
                return { success: true, id: entryData._spItemId, alreadyExisted: true, image: imgRes };
            }

            let existId = null;
            try { existId = await this.findItemByTitle(siteUrl, listName, entryData.entryId || '', digest); } catch {}
            if (existId) {
                const imgRes = imageDataUrl ? await this._uploadImageForItem(siteUrl, listName, existId, imageDataUrl, entryData, mapping, entityType, digest) : null;
                return { success: true, id: existId, alreadyExisted: true, image: imgRes };
            }

            const body = { '__metadata': { 'type': entityType }, 'Title': truncateString(entryData.entryId || `Entry_${Date.now()}`, 255) };
            const af = (k, v) => { const n = mapping[k]; if (n && n !== 'Title' && k !== 'screenshot' && !k.startsWith('_')) body[n] = v; };

            af('hva', truncateString(entryData.hva, 255));
            af('messageId', truncateString(entryData.messageId, 255));
            af('conversationId', truncateString(entryData.conversationId, 255));
            af('responseTime', truncateString(entryData.responseTimeFormatted, 255));
            af('status', truncateString(entryData.responseCorrectOrIncorrect, 255));
            af('gcsLinkType', truncateString(entryData.gcsLinkType, 255));
            af('gcsLink', truncateString(entryData.gcsLink, 255));
            af('testEmail', truncateString(entryData.testAccountEmail, 255));
            af('testCustomerId', truncateString(entryData.testAccountCustomerId, 255));
            af('testingDate', truncateString(entryData.testingDate, 255));
            af('testerLogin', truncateString(entryData.testerLogin, 255));
            af('queryTimestamp', truncateString(entryData.queryLocalTime, 255));
            af('savedTimestamp', truncateString(entryData.savedLocalTime, 255));
            af('query', truncateString(entryData.query));
            af('botResponse', truncateString(entryData.preProdBotResponse));
            af('groundTruth', truncateString(entryData.groundTruthResponse));
            af('observations', truncateString(entryData.observations));
            if (env.hasAreaOfImprovement) af('areaOfImprovement', truncateString(entryData.areaOfImprovement));

            const result = await new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'POST', url: `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items`,
                    headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'X-RequestDigest': digest, 'X-Requested-With': 'XMLHttpRequest' },
                    data: JSON.stringify(body), anonymous: false, withCredentials: true, timeout: 20000,
                    onload: r => {
                        if (r.status === 201) { try { const d = JSON.parse(r.responseText); resolve({ success: true, id: d.d.Id, title: d.d.Title }); } catch { resolve({ success: true, id: 'unknown' }); } }
                        else if (r.status === 400) { let m = 'Bad request'; try { m = JSON.parse(r.responseText).error?.message?.value || m; } catch {} if (m.includes('does not exist')) clearSPCache(); reject(new Error(m)); }
                        else if (r.status === 401 || r.status === 403) { cachedDigest = null; reject(new Error(`Permission denied (${r.status})`)); }
                        else { let m = `HTTP ${r.status}`; try { m = JSON.parse(r.responseText).error?.message?.value || m; } catch {} reject(new Error(m)); }
                    },
                    onerror: () => reject(new Error('Network error')),
                    ontimeout: () => reject(new Error('Timeout'))
                });
            });

            const imgRes = (imageDataUrl && result.id && result.id !== 'unknown') ? await this._uploadImageForItem(siteUrl, listName, result.id, imageDataUrl, entryData, mapping, entityType, digest) : null;
            return { ...result, image: imgRes };
        }
    };

    // ════════════════════════════════════════════════════════
    //  AUTO-UPLOAD ENGINE
    // ════════════════════════════════════════════════════════
    let autoUploadInProgress = false;
    let autoUploadRetryQueue = [];
    let uploadLockKey = 'chatScraper_uploadLock';

    function acquireUploadLock() {
        const existing = localStorage.getItem(uploadLockKey);
        if (existing) {
            try {
                const lock = JSON.parse(existing);
                if (Date.now() - lock.time < 60000 && lock.tabId !== TAB_ID) return false;
            } catch {}
        }
        localStorage.setItem(uploadLockKey, JSON.stringify({ tabId: TAB_ID, time: Date.now() }));
        return true;
    }

    function releaseUploadLock() {
        try {
            const existing = localStorage.getItem(uploadLockKey);
            if (existing) { const lock = JSON.parse(existing); if (lock.tabId === TAB_ID) localStorage.removeItem(uploadLockKey); }
        } catch { localStorage.removeItem(uploadLockKey); }
    }

    function showAutoUploadStatus(message, type = 'info') {
        let ind = document.getElementById('autoUploadIndicator');
        if (!ind) {
            ind = document.createElement('div');
            ind.id = 'autoUploadIndicator';
            ind.style.cssText = `position:fixed;bottom:70px;right:20px;padding:10px 18px;border-radius:8px;font-family:Arial;font-size:12px;z-index:200001;box-shadow:0 4px 15px rgba(0,0,0,0.3);transition:all 0.3s ease;max-width:420px;display:flex;align-items:center;gap:8px;`;
            document.body.appendChild(ind);
        }
        const styles = {
            uploading: { bg: '#0078d4', c: '#fff', i: '⏳' },
            success: { bg: '#28a745', c: '#fff', i: '✅' },
            error: { bg: '#dc3545', c: '#fff', i: '❌' },
            retry: { bg: '#fd7e14', c: '#fff', i: '🔄' },
            info: { bg: '#17a2b8', c: '#fff', i: 'ℹ️' }
        };
        const s = styles[type] || styles.info;
        ind.style.background = s.bg;
        ind.style.color = s.c;
        ind.textContent = '';
        const iconSpan = document.createElement('span');
        iconSpan.textContent = s.i;
        const msgSpan = document.createElement('span');
        msgSpan.textContent = `[${getActiveEnv().shortLabel}] ${message}`;
        ind.appendChild(iconSpan);
        ind.appendChild(msgSpan);
               ind.style.opacity = '1';
        if (type !== 'uploading' && type !== 'retry') {
            setTimeout(() => {
                if (ind) { ind.style.opacity = '0'; setTimeout(() => { try { ind.remove(); } catch {} }, 300); }
            }, type === 'success' ? 3000 : 5000);
        }
    }

    async function autoUploadEntryToList(entry) {
        if (!getAutoUpload()) return;
        const siteUrl = getSPListUrl();
        const listName = getSPListName();
        if (!siteUrl || !listName || !isValidSPUrl(siteUrl)) return;
        if (entry._spUploaded === true && entry._spImageAttached === true) return;
        if (entry._spUploaded === true && !entry._hasImage) return;

        if (autoUploadInProgress) {
            if (!autoUploadRetryQueue.some(q => q.entryId === entry.entryId)) {
                if (autoUploadRetryQueue.length < LIMITS.MAX_QUEUE_SIZE) autoUploadRetryQueue.push(entry);
            }
            return;
        }

        if (!acquireUploadLock()) {
            SecureLog.info('Another tab is uploading, queuing...');
            if (!autoUploadRetryQueue.some(q => q.entryId === entry.entryId)) autoUploadRetryQueue.push(entry);
            return;
        }

        autoUploadInProgress = true;
        const t0 = Date.now();

        try {
            showAutoUploadStatus(`Pushing #${entry.sNo} to SP...`, 'uploading');
            let img = null;
            if (entry._hasImage && entry.entryId) {
                try { img = await ImageStore.getImageData(entry.entryId); } catch {}
            }
            showAutoUploadStatus(`#${entry.sNo} → SP${img ? ' + image' : ''}...`, 'uploading');
            const r = await SharePointList.addItem(siteUrl, listName, entry, img);
            const el = ((Date.now() - t0) / 1000).toFixed(1);
            const iN = r.image?.success ? ' + image' : r.image?.error ? ' (no img)' : '';
            showAutoUploadStatus(`#${entry.sNo} → SP (ID: ${r.id})${iN} — ${el}s`, 'success');
            entry._spUploaded = true;
            entry._spItemId = r.id;
            entry._spImageAttached = r.image?.success || entry._spImageAttached || false;
            entry._spImageUrl = r.image?.url || entry._spImageUrl || '';
            delete entry._spError;
            await saveDataToStorage();
            updateStatus();
        } catch (e) {
            showAutoUploadStatus(`#${entry.sNo} failed`, 'error');
            entry._spUploaded = false;
            entry._spError = 'Upload failed';
            await saveDataToStorage();
            updateStatus();
            SecureLog.error('Auto-upload failed', e);
        } finally {
            autoUploadInProgress = false;
            releaseUploadLock();
            if (autoUploadRetryQueue.length > 0) {
                const next = autoUploadRetryQueue.shift();
                setTimeout(() => autoUploadEntryToList(next), 1500);
            }
        }
    }

    async function retryFailedUploads() {
        const siteUrl = getSPListUrl();
        const listName = getSPListName();
        if (!siteUrl || !listName || !isValidSPUrl(siteUrl)) return;

        const needsUpload = chatData.filter(d => d._spUploaded === false);
        const needsImage = chatData.filter(d => d._spUploaded === true && d._spImageAttached === false && d._hasImage);
        const total = needsUpload.length + needsImage.length;

        if (!total) { showNotification('Everything synced!', 'info'); return; }
        showNotification(`Retrying ${needsUpload.length} failed + ${needsImage.length} images...`, 'info');

        let okR = 0, okI = 0, fail = 0;

        for (const entry of needsUpload) {
            if (!acquireUploadLock()) { showNotification('Another tab is uploading', 'warning'); break; }
            try {
                showAutoUploadStatus(`Retrying #${entry.sNo}...`, 'retry');
                let img = null;
                if (entry._hasImage && entry.entryId) { try { img = await ImageStore.getImageData(entry.entryId); } catch {} }
                const r = await SharePointList.addItem(siteUrl, listName, entry, img);
                entry._spUploaded = true; entry._spItemId = r.id;
                entry._spImageAttached = r.image?.success || false;
                entry._spImageUrl = r.image?.url || '';
                delete entry._spError; okR++;
                if (r.image?.success) okI++;
            } catch (e) { entry._spError = 'Retry failed'; fail++; SecureLog.error('Retry failed', e); }
            finally { releaseUploadLock(); }
            await new Promise(r => setTimeout(r, 1000));
        }

        for (const entry of needsImage) {
            if (!acquireUploadLock()) break;
            try {
                showAutoUploadStatus(`Adding image #${entry.sNo}...`, 'retry');
                let img = null;
                if (entry.entryId) { try { img = await ImageStore.getImageData(entry.entryId); } catch {} }
                if (!img) continue;
                const r = await SharePointList.addItem(siteUrl, listName, entry, img);
                if (r.image?.success) { entry._spImageAttached = true; entry._spImageUrl = r.image.url || ''; okI++; }
            } catch {} finally { releaseUploadLock(); }
            await new Promise(r => setTimeout(r, 500));
        }

        await saveDataToStorage();
        updateStatus();
        const parts = [];
        if (okR) parts.push(`${okR} rows`);
        if (okI) parts.push(`${okI} images`);
        if (fail) parts.push(`${fail} failed`);
        showNotification(`Retry: ${parts.join(', ')}`, fail > 0 ? 'warning' : 'success');
    }

    // ════════════════════════════════════════════════════════
    //  IndexedDB Image Store
    // ════════════════════════════════════════════════════════
    const ImageStore = {
        dbName: 'chatScraperImagesDB', storeName: 'images', db: null, _initPromise: null,

        async init() {
            if (this.db) return;
            if (this._initPromise) return this._initPromise;
            this._initPromise = new Promise((res, rej) => {
                const r = indexedDB.open(this.dbName, 1);
                r.onupgradeneeded = e => { const db = e.target.result; if (!db.objectStoreNames.contains(this.storeName)) db.createObjectStore(this.storeName, { keyPath: 'id' }); };
                r.onsuccess = e => { this.db = e.target.result; this.db.onclose = () => { this.db = null; this._initPromise = null; }; res(); };
                r.onerror = () => { this._initPromise = null; rej(r.error); };
            });
            return this._initPromise;
        },

        async ensureDB() {
            if (!this.db) { this._initPromise = null; await this.init(); }
            try { this.db.transaction(this.storeName, 'readonly'); }
            catch { this.db = null; this._initPromise = null; await this.init(); }
            return this.db;
        },

        async save(id, data, meta = {}) {
            if (!validateImageDataUrl(data)) throw new Error('Invalid image data for storage');
            const db = await this.ensureDB();
            return new Promise((res, rej) => {
                const tx = db.transaction(this.storeName, 'readwrite');
                tx.objectStore(this.storeName).put({ id, data, width: meta.width || 0, height: meta.height || 0, sizeKB: Math.round((data?.length || 0) * 0.75 / 1024), savedAt: Date.now() });
                tx.oncomplete = () => res(true); tx.onerror = () => rej(tx.error);
            });
        },

        async get(id) {
            const db = await this.ensureDB();
            return new Promise((res, rej) => {
                const tx = db.transaction(this.storeName, 'readonly');
                const r = tx.objectStore(this.storeName).get(id);
                r.onsuccess = () => res(r.result || null); r.onerror = () => rej(r.error);
            });
        },

        async getImageData(id) { try { const r = await this.get(id); return r?.data?.startsWith('data:image') ? r.data : null; } catch { return null; } },
        async has(id) { try { return (await this.getImageData(id)) !== null; } catch { return false; } },
        async verify(id) { try { const r = await this.get(id); return { exists: !!r, hasData: !!r?.data, isImage: !!r?.data?.startsWith('data:image'), sizeKB: r?.sizeKB || 0 }; } catch { return { exists: false, hasData: false, isImage: false, sizeKB: 0 }; } },
        async clear() { const db = await this.ensureDB(); return new Promise((res, rej) => { const tx = db.transaction(this.storeName, 'readwrite'); tx.objectStore(this.storeName).clear(); tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error); }); },
        async count() { const db = await this.ensureDB(); return new Promise(res => { const tx = db.transaction(this.storeName, 'readonly'); const r = tx.objectStore(this.storeName).count(); r.onsuccess = () => res(r.result); r.onerror = () => res(0); }); }
    };

    // ════════════════════════════════════════════════════════
    //  DAILY RESET
    // ════════════════════════════════════════════════════════
    function checkDailyReset() {
        const today = new Date().toLocaleDateString('en-US');
        const last = localStorage.getItem(GLOBAL_KEYS.lastResetDate);
        if (last && last !== today) performDailyReset();
        localStorage.setItem(GLOBAL_KEYS.lastResetDate, today);
    }

    function performDailyReset() {
        Object.keys(ENVIRONMENTS).forEach(eid => {
            localStorage.removeItem(`chatScraper_${eid}_savedData`);
            localStorage.removeItem(`chatScraper_${eid}_responseQueue`);
        });
        chatData = []; responseQueue = [];
        ImageStore.clear().catch(() => {});
        clearAllSPCaches(); CryptoStore.clearKeyCache();
        localStorage.setItem(GLOBAL_KEYS.lastResetDate, new Date().toLocaleDateString('en-US'));
        updateStatus(); refreshPanelState();
        showNotification('Daily reset: all environments cleared', 'info');
    }

    function scheduleDailyReset() {
        if (dailyResetTimer) clearTimeout(dailyResetTimer);
        const now = new Date(); const rt = new Date(now);
        rt.setHours(23, 59, 0, 0);
        if (now >= rt) rt.setDate(rt.getDate() + 1);
        dailyResetTimer = setTimeout(() => { performDailyReset(); scheduleDailyReset(); }, rt - now);
    }

    // ════════════════════════════════════════════════════════
    //  SAFE UTILITIES
    // ════════════════════════════════════════════════════════
    const getElement = id => { if (!cachedElements[id]) cachedElements[id] = document.getElementById(id); return cachedElements[id]; };
    const clearElementCache = () => { cachedElements = {}; };

    const formatLocalTime = iso => new Date(iso).toLocaleString('en-US', { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true });
    const formatTimestamp = iso => { const d = new Date(iso); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}:${String(d.getSeconds()).padStart(2, '0')}`; };
    const formatDate = iso => new Date(iso).toLocaleDateString('en-US', { year: 'numeric', month: '2-digit', day: '2-digit' });
    const formatResponseTime = ms => ms < 1000 ? `${Math.round(ms)} ms` : `${(ms / 1000).toFixed(2)} s`;
    const debounce = (fn, d) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), d); }; };

    const HVA_OPTIONS = [
        'Select HVA...',
        '3WM (3-way match)', 'ATEP (Amazon Tax Exemption Program)', 'Account authority',
        'Add User', 'Business Lists', 'Business Order Information', 'Custom Quotes',
        'Business Prime', 'Guided Buying', 'PBI', 'Quantity Discount', 'Recurring Delivery',
        'SSO', 'Shared Settings', 'Services', 'UNSPSC', 'Amazon Business Payments',
        'Your orders', 'Budget Management', 'Business Giving', 'Business Pricing',
        'Business Pricing - Lightning Deals', 'Business Pricing - Promotions',
        'Business Pricing - Pre order price', 'Business Pricing - Custom Quotes',
        'Business Pricing - Bulk Buying', 'Business Pricing - Industry Vertical Pricing',
        'Business Pricing - Negotiated Prices', 'Spend Anomaly Monitoring', 'Approvals',
        'AB Registration', 'AB Search', 'AB Cart', 'Socially Responsible Purchasing',
        'AB App Center', 'None of the above', 'Custom'
    ];

    const GCS_MCS_OPTIONS = [
        { value: '', label: 'Select Link Type...' },
        { value: 'gcs', label: 'GCS' },
        { value: 'mcs', label: 'MCS' },
    ];

    // ════════════════════════════════════════════════════════
    //  GCS/MCS HELPERS — SEPARATED TYPE AND URL
    // ════════════════════════════════════════════════════════

    /** Returns ONLY the URL (no type prefix) */
    function getGcsLinkFromSelector(containerId, inputId, selectId) {
        const input = document.getElementById(inputId);
        return input ? input.value.trim() : '';
    }

    /** Returns ONLY the type label: "GCS", "MCS", or custom text */
    function getGcsLinkTypeLabel(selectId, containerId) {
        const select = document.getElementById(selectId);
        if (!select || !select.value) return '';
        const val = select.value;
        if (val === 'gcs') return 'GCS';
        if (val === 'mcs') return 'MCS';
        if (val === 'custom') {
            const customInput = document.getElementById(`${containerId}_customType`);
            return customInput && customInput.value.trim() ? customInput.value.trim() : 'Custom';
        }
        const opt = GCS_MCS_OPTIONS.find(o => o.value === val);
        return opt ? opt.label.trim() : val;
    }

    // ════════════════════════════════════════════════════════
    //  QUEUE
    // ════════════════════════════════════════════════════════
    function getCurrentResponse() { return responseQueue[0] || null; }

    function isValidMessage(t) { return t && !['N/A', 'NA', '', 'UNDEFINED', 'NULL'].includes(String(t).trim().toUpperCase()); }

    function addToQueue(data) {
        if (!isValidMessage(data.userMessage) || !isValidMessage(data.botMessage)) return;
        if (responseQueue.some(r => r.messageId === data.messageId && data.messageId !== 'N/A')) return;
        if (responseQueue.length >= LIMITS.MAX_QUEUE_SIZE) { SecureLog.warn('Queue full, dropping oldest'); responseQueue.shift(); }
        data._savedAt = Date.now();
        data._queueId = `q_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`;
        data.userMessage = truncateString(data.userMessage);
        data.botMessage = truncateString(data.botMessage);
        responseQueue.push(data);
        saveQueueToStorage();
        refreshPanelState();
    }

    function removeCurrentFromQueue() { if (responseQueue.length > 0) responseQueue.shift(); saveQueueToStorage(); refreshPanelState(); }

    function skipCurrentResponse() {
        if (!responseQueue.length) return;
        removeCurrentFromQueue();
        showNotification(responseQueue.length > 0 ? `Skipped. ${responseQueue.length} left.` : 'Skipped.', 'info');
    }

    function skipAllResponses() { const c = responseQueue.length; responseQueue = []; saveQueueToStorage(); refreshPanelState(); showNotification(`Skipped ${c}.`, 'info'); }

    function refreshPanelState() { updateCaptureButtonState(responseQueue.length > 0); updateQueueDisplay(); updateBubble(); }

    function updateQueueDisplay() {
        const qi = getElement('queueInfo'), qd = getElement('queueDetail'), sb = getElement('skipBtn');
        const sa = getElement('skipAllBtn'), rt = getElement('responseTimeDisplay'), bi = getElement('botResponseIndicator');
        const wi = getElement('waitingIndicator');
        const c = responseQueue.length, cur = getCurrentResponse();

        if (c > 0 && cur) {
            if (bi) bi.style.display = 'block'; if (wi) wi.style.display = 'none'; if (sb) sb.style.display = 'block';
            if (rt) safeSetText(rt, `Response time: ${cur.responseTimeFormatted}`);
            if (qi) { safeSetText(qi, c === 1 ? '1 pending' : `${c} pending`); qi.style.display = 'block'; }
            if (qd) {
                qd.textContent = '';
                const strong = document.createElement('strong'); strong.textContent = 'Next: '; qd.appendChild(strong);
                const msgText = (cur.userMessage || '').substring(0, 50);
                qd.appendChild(document.createTextNode(`"${msgText}${cur.userMessage?.length > 50 ? '...' : ''}"`));
                qd.style.display = 'block';
            }
            if (sa) sa.style.display = c > 1 ? 'block' : 'none';
        } else {
            if (bi) bi.style.display = 'none'; if (wi) wi.style.display = 'block';
            [sb, sa, qi, qd].forEach(el => { if (el) el.style.display = 'none'; });
        }
    }

    // ════════════════════════════════════════════════════════
    //  USERNAME PROMPT
    // ════════════════════════════════════════════════════════
    function showUsernamePrompt() {
        return new Promise(resolve => {
            const saved = localStorage.getItem(GLOBAL_KEYS.username);
            if (saved) { currentUsername = saved; resolve(saved); return; }

            const modal = document.createElement('div');
            modal.style.cssText = `position:fixed;inset:0;background:rgba(0,0,0,0.9);z-index:999999;display:flex;align-items:center;justify-content:center;font-family:Arial;`;
            const content = document.createElement('div');
            content.style.cssText = `background:#fff;width:380px;max-width:95vw;border-radius:12px;overflow:hidden`;
            const header = document.createElement('div');
            header.style.cssText = `background:linear-gradient(135deg,#FF9900,#e88b00);color:#fff;padding:20px;text-align:center`;
            const h2 = document.createElement('h2'); h2.style.margin = '0'; h2.textContent = 'Multi-Env v8.1 (Secured)'; header.appendChild(h2);
            const body = document.createElement('div'); body.style.cssText = `padding:20px`;
            const label = document.createElement('label'); label.style.cssText = `font-weight:bold;font-size:13px`; label.textContent = 'Tester Login ';
            const req = document.createElement('span'); req.style.color = 'red'; req.textContent = '*'; label.appendChild(req);
            const input = document.createElement('input');
            input.type = 'text'; input.id = 'usernameInput'; input.placeholder = 'Enter login';
            input.style.cssText = `width:100%;padding:12px;border:2px solid #ddd;border-radius:6px;font-size:14px;box-sizing:border-box;margin-top:8px`;
            input.setAttribute('maxlength', '50'); input.setAttribute('autocomplete', 'off');
            const errDiv = document.createElement('div'); errDiv.style.cssText = `color:red;font-size:12px;margin-top:6px;display:none`; errDiv.textContent = 'Required';
            const submitBtn = document.createElement('button');
            submitBtn.style.cssText = `width:100%;padding:12px;margin-top:15px;background:#28a745;color:#fff;border:none;border-radius:6px;font-size:14px;font-weight:bold;cursor:pointer`;
            submitBtn.textContent = 'Continue';
            body.appendChild(label); body.appendChild(input); body.appendChild(errDiv); body.appendChild(submitBtn);
            content.appendChild(header); content.appendChild(body); modal.appendChild(content); document.body.appendChild(modal);
            setTimeout(() => input.focus(), 100);

            const submit = () => {
                const u = input.value.trim().replace(/[^a-zA-Z0-9._@-]/g, '');
                if (!u) { errDiv.style.display = 'block'; return; }
                currentUsername = u; localStorage.setItem(GLOBAL_KEYS.username, u); modal.remove(); resolve(u);
            };
            submitBtn.onclick = submit;
            input.onkeypress = e => { if (e.key === 'Enter') submit(); };
        });
    }

    // ════════════════════════════════════════════════════════
    //  BUBBLE
    // ════════════════════════════════════════════════════════
    function createFloatingBubble() {
        const b = document.createElement('div');
        b.id = 'chatScraperBubble';
        b.style.cssText = `position:fixed;top:15px;left:15px;z-index:10001;width:50px;height:50px;border-radius:50%;background:${getActiveEnv().gradient};box-shadow:0 4px 15px rgba(0,0,0,0.4);cursor:pointer;display:none;align-items:center;justify-content:center;font-family:Arial;font-size:11px;font-weight:bold;color:#fff;text-align:center;line-height:1.1;user-select:none;`;
        const span = document.createElement('span'); span.id = 'bubbleCount'; span.style.pointerEvents = 'none'; span.textContent = `${getActiveEnv().shortLabel}`;
        b.appendChild(span); document.body.appendChild(b);

        let bx = 0, by = 0, drag = false;
        b.onmousedown = e => {
            e.preventDefault(); bx = e.clientX; by = e.clientY; drag = false;
            document.onmouseup = () => { document.onmouseup = null; document.onmousemove = null; if (!drag) toggleMinimize(); };
            document.onmousemove = e => { drag = true; b.style.top = (b.offsetTop - (by - e.clientY)) + "px"; b.style.left = (b.offsetLeft - (bx - e.clientX)) + "px"; bx = e.clientX; by = e.clientY; };
        };
    }

    function updateBubble() {
        const b = getElement('chatScraperBubble'); if (!b) return;
        const env = getActiveEnv(); b.style.background = env.gradient;
        const q = responseQueue.length, c = chatData.length;
        const el = b.querySelector('#bubbleCount'); if (!el) return;
        if (q > 0) safeSetText(el, `${q} ${env.shortLabel} pend`);
        else if (c > 0) safeSetText(el, `${c} ${env.shortLabel} saved`);
        else safeSetText(el, `${env.shortLabel} 8.1`);
    }

    let isMinimized = false;
    function toggleMinimize() {
        const p = getElement('chatScraperPanel'), b = getElement('chatScraperBubble');
        if (!p || !b) return;
        isMinimized = !isMinimized;
        if (isMinimized) { p.style.display = 'none'; b.style.display = 'flex'; updateBubble(); }
        else { p.style.display = 'block'; b.style.display = 'none'; }
    }

    // ════════════════════════════════════════════════════════
    //  ENVIRONMENT SWITCHER UI
    // ════════════════════════════════════════════════════════
    function buildEnvSwitcherHTML() {
        const container = document.createElement('div');
        container.style.cssText = `display:flex;gap:4px;margin-bottom:10px;border-radius:8px;overflow:hidden;border:2px solid #dee2e6`;
        Object.values(ENVIRONMENTS).forEach(env => {
            const isActive = env.id === activeEnvId;
            const btn = document.createElement('button');
            btn.className = 'env-switch-btn'; btn.setAttribute('data-env', env.id);
            btn.style.cssText = `flex:1;padding:8px 4px;border:none;cursor:pointer;font-size:10px;font-weight:bold;transition:all 0.2s;background:${isActive ? env.color : '#f8f9fa'};color:${isActive ? '#fff' : '#495057'};${isActive ? 'box-shadow:0 2px 8px rgba(0,0,0,0.25);' : ''}`;
            const label = document.createElement('span'); label.textContent = env.shortLabel; btn.appendChild(label);
            btn.appendChild(document.createElement('br'));
            const sub = document.createElement('span'); sub.style.cssText = `font-size:8px;font-weight:normal;opacity:0.85`; sub.textContent = env.label.split(' ')[0];
            btn.appendChild(sub); container.appendChild(btn);
        });
        return container;
    }

    function attachEnvSwitcherListeners() {
        document.querySelectorAll('.env-switch-btn').forEach(btn => {
            btn.onclick = () => { const envId = btn.getAttribute('data-env'); if (envId && envId !== activeEnvId) switchEnvironment(envId); };
        });
    }

    // ════════════════════════════════════════════════════════
    //  GCS/MCS LINK SELECTOR BUILDER
    // ════════════════════════════════════════════════════════
    function buildGcsLinkSelector(containerId, inputId, selectId, currentValue, currentType) {
        const wrapper = document.createElement('div');
        wrapper.id = containerId;
        wrapper.style.marginBottom = '15px';

        const label = document.createElement('label');
        label.style.cssText = `font-weight:bold;font-size:12px`;
        label.textContent = 'GCS/MCS Link:';
        wrapper.appendChild(label);

        const select = document.createElement('select');
        select.id = selectId;
        select.style.cssText = `width:100%;padding:8px;border:2px solid #17a2b8;border-radius:6px;margin-top:5px;font-size:11px;background:#f0fbff;cursor:pointer;box-sizing:border-box`;

        GCS_MCS_OPTIONS.forEach(opt => {
            const option = document.createElement('option');
            option.value = opt.value;
            option.textContent = opt.label;
            if (opt.value === (currentType || '')) option.selected = true;
            select.appendChild(option);
        });
        wrapper.appendChild(select);

        const urlInput = document.createElement('input');
        urlInput.type = 'text';
        urlInput.id = inputId;
        urlInput.placeholder = 'Paste GCS/MCS link here...';
        urlInput.value = currentValue || '';
        urlInput.setAttribute('maxlength', '500');
        urlInput.setAttribute('autocomplete', 'off');
        urlInput.style.cssText = `width:100%;padding:8px;border:1px solid #ddd;border-radius:6px;font-size:11px;box-sizing:border-box;margin-top:6px`;
        wrapper.appendChild(urlInput);

        const badge = document.createElement('div');
        badge.id = `${containerId}_badge`;
        badge.style.cssText = `display:none;margin-top:4px;padding:3px 8px;border-radius:4px;font-size:9px;font-weight:bold`;
        wrapper.appendChild(badge);

        select.addEventListener('change', () => {
            const val = select.value;
            if (val === 'gcs') {
                badge.textContent = 'GCS'; badge.style.display = 'inline-block';
                badge.style.background = '#d4edda'; badge.style.color = '#155724';
            } else if (val === 'mcs') {
                badge.textContent = 'MCS'; badge.style.display = 'inline-block';
                badge.style.background = '#cce5ff'; badge.style.color = '#004085';
            } else {
                badge.style.display = 'none';
            }
        });

        // Trigger initial state
        if (currentType === 'gcs') {
            badge.textContent = 'GCS'; badge.style.display = 'inline-block';
            badge.style.background = '#d4edda'; badge.style.color = '#155724';
        } else if (currentType === 'mcs') {
            badge.textContent = 'MCS'; badge.style.display = 'inline-block';
            badge.style.background = '#cce5ff'; badge.style.color = '#004085';
        }

        return wrapper;
    }

    // ════════════════════════════════════════════════════════
    //  CUSTOM HVA INPUT BUILDER
    // ════════════════════════════════════════════════════════
    function buildHvaSelector(selectId, customInputId) {
        const wrapper = document.createElement('div');
        wrapper.style.marginBottom = '15px';

        const label = document.createElement('label');
        label.style.fontWeight = 'bold';
        label.textContent = 'HVA ';
        const req = document.createElement('span');
        req.style.color = 'red'; req.textContent = '*';
        label.appendChild(req);
        wrapper.appendChild(label);

        const select = document.createElement('select');
        select.id = selectId;
        select.style.cssText = `width:100%;padding:10px;border:2px solid #ddd;border-radius:6px;margin-top:5px;font-size:12px`;
        HVA_OPTIONS.forEach(o => {
            const opt = document.createElement('option');
            opt.value = o === 'Select HVA...' ? '' : o;
            opt.textContent = o;
            select.appendChild(opt);
        });
        wrapper.appendChild(select);

        const customWrap = document.createElement('div');
        customWrap.id = `${customInputId}_wrap`;
        customWrap.style.cssText = `display:none;margin-top:8px`;

        const customLabel = document.createElement('label');
        customLabel.style.cssText = `font-size:11px;color:#856404;font-weight:bold`;
        customLabel.textContent = '✏️ Enter Custom HVA:';
        customWrap.appendChild(customLabel);

        const customInput = document.createElement('input');
        customInput.type = 'text';
        customInput.id = customInputId;
        customInput.placeholder = 'Type your custom HVA here...';
        customInput.setAttribute('maxlength', '255');
        customInput.setAttribute('autocomplete', 'off');
        customInput.style.cssText = `width:100%;padding:10px;border:2px solid #fd7e14;border-radius:6px;font-size:12px;box-sizing:border-box;margin-top:4px;background:#fff8f0;transition:border-color 0.2s`;
        customWrap.appendChild(customInput);

        const hint = document.createElement('div');
        hint.style.cssText = `font-size:9px;color:#856404;margin-top:3px`;
        hint.textContent = 'This will be used as the HVA value for this entry';
        customWrap.appendChild(hint);
        wrapper.appendChild(customWrap);

        select.addEventListener('change', () => {
            if (select.value === 'Custom') {
                customWrap.style.display = 'block';
                customInput.focus();
                select.style.borderColor = '#fd7e14';
                select.style.background = '#fff8f0';
            } else {
                customWrap.style.display = 'none';
                customInput.value = '';
                select.style.borderColor = '#ddd';
                select.style.background = '';
            }
        });

        return wrapper;
    }

    function getHvaValue(selectId, customInputId) {
        const select = document.getElementById(selectId);
        const customInput = document.getElementById(customInputId);
        if (select && select.value === 'Custom' && customInput && customInput.value.trim()) {
            return customInput.value.trim();
        }
        return select ? select.value : '';
    }

    // ════════════════════════════════════════════════════════
    //  CONTROL PANEL
    // ════════════════════════════════════════════════════════
    function createControlPanel() {
        const env = getActiveEnv();
        const panel = document.createElement('div');
        panel.id = 'chatScraperPanel';
        buildPanelDOM(panel, env);
        document.body.appendChild(panel);
        clearElementCache();
        createFloatingBubble();
        makeDraggable(panel, panel.querySelector('#panelHeader'));
        attachPanelListeners();
        attachEnvSwitcherListeners();
        refreshPanelState();
        updateStatus();
    }

    function rebuildPanelForEnv() {
        const env = getActiveEnv();
        const panel = getElement('chatScraperPanel') || document.getElementById('chatScraperPanel');
        if (!panel) return;
        const top = panel.style.top, left = panel.style.left;
        while (panel.firstChild) panel.removeChild(panel.firstChild);
        buildPanelDOM(panel, env);
        panel.style.borderColor = env.borderColor;
        if (top) panel.style.top = top; if (left) panel.style.left = left;
        clearElementCache();
        makeDraggable(panel, panel.querySelector('#panelHeader'));
        attachPanelListeners();
        attachEnvSwitcherListeners();
        const bubble = document.getElementById('chatScraperBubble');
        if (bubble) bubble.style.background = env.gradient;
    }

    function buildPanelDOM(panel, env) {
        const style = document.createElement('style');
        style.textContent = `
            #chatScraperPanel{position:fixed;top:10px;left:10px;z-index:10000;background:#fff;
                border:2px solid ${env.borderColor};border-radius:10px;
                box-shadow:0 4px 20px rgba(0,0,0,0.2);font-family:Arial;width:280px;font-size:12px}
            .csp-btn{width:100%;padding:10px;border:none;border-radius:6px;cursor:pointer;
                font-weight:bold;font-size:12px;margin-bottom:8px;transition:opacity 0.15s}
            .csp-btn:hover{opacity:0.9}
            .env-switch-btn:hover{filter:brightness(1.1)}
        `;
        panel.appendChild(style);

        const header = document.createElement('div');
        header.id = 'panelHeader';
        header.style.cssText = `background:${env.headerGradient};color:#fff;padding:10px 12px;border-radius:8px 8px 0 0;display:flex;align-items:center;justify-content:space-between;font-weight:bold;font-size:13px;cursor:move`;
        const headerLeft = document.createElement('span'); headerLeft.style.cssText = `display:flex;align-items:center;gap:6px`;
        const envBadge = document.createElement('span'); envBadge.style.cssText = `background:${env.color};padding:2px 6px;border-radius:4px;font-size:10px`; envBadge.textContent = env.shortLabel;
        headerLeft.appendChild(envBadge); headerLeft.appendChild(document.createTextNode(' v8.1'));
        const minBtn = document.createElement('button'); minBtn.id = 'minimizeBtn';
        minBtn.style.cssText = `background:rgba(255,255,255,0.3);border:none;color:#fff;width:22px;height:22px;border-radius:4px;cursor:pointer;font-size:14px`; minBtn.textContent = '−';
        header.appendChild(headerLeft); header.appendChild(minBtn); panel.appendChild(header);

        const content = document.createElement('div'); content.id = 'panelContent'; content.style.padding = '12px';
        content.appendChild(buildEnvSwitcherHTML());

        const envInd = document.createElement('div');
        envInd.style.cssText = `padding:6px 10px;margin-bottom:10px;background:${env.color}15;color:${env.color};border:1px solid ${env.color}40;border-radius:6px;font-size:11px;text-align:center;font-weight:bold`;
        envInd.textContent = `📍 ${env.label}`;
        content.appendChild(envInd);

        const userDiv = document.createElement('div');
        userDiv.style.cssText = `padding:6px 10px;margin-bottom:10px;background:#e3f2fd;color:#1565c0;border-radius:6px;font-size:11px;display:flex;align-items:center;justify-content:space-between`;
        const userSpan = document.createElement('span'); userSpan.textContent = 'User: ';
        const userStrong = document.createElement('strong'); userStrong.id = 'usernameDisplay'; userStrong.textContent = currentUsername || '';
        userSpan.appendChild(userStrong);
        const changeBtn = document.createElement('button'); changeBtn.id = 'changeUserBtn';
        changeBtn.style.cssText = `background:none;border:none;color:#1565c0;cursor:pointer;font-size:10px;text-decoration:underline`; changeBtn.textContent = 'Change';
        userDiv.appendChild(userSpan); userDiv.appendChild(changeBtn); content.appendChild(userDiv);

        const botInd = document.createElement('div'); botInd.id = 'botResponseIndicator';
        botInd.style.cssText = `display:none;padding:8px;margin-bottom:10px;background:#d4edda;color:#155724;border-radius:6px;font-size:11px;text-align:center`;
        const botLabel = document.createElement('div'); botLabel.style.fontWeight = 'bold'; botLabel.textContent = 'Bot response detected!';
        const rtDisp = document.createElement('div'); rtDisp.id = 'responseTimeDisplay'; rtDisp.style.cssText = `margin-top:4px;font-size:10px`; rtDisp.textContent = '--';
        botInd.appendChild(botLabel); botInd.appendChild(rtDisp); content.appendChild(botInd);

        const qInfo = document.createElement('div'); qInfo.id = 'queueInfo';
        qInfo.style.cssText = `display:none;padding:6px 10px;margin-bottom:6px;background:#e8f4fd;color:#0c5460;border-radius:6px;font-size:11px;text-align:center;font-weight:bold;border:1px solid #bee5eb`;
        content.appendChild(qInfo);
        const qDetail = document.createElement('div'); qDetail.id = 'queueDetail';
        qDetail.style.cssText = `display:none;padding:6px 10px;margin-bottom:10px;background:#f8f9fa;color:#495057;border-radius:6px;font-size:10px;max-height:40px;overflow:hidden`;
        content.appendChild(qDetail);
        const waitInd = document.createElement('div'); waitInd.id = 'waitingIndicator';
        waitInd.style.cssText = `padding:8px;margin-bottom:10px;background:#fff3cd;color:#856404;border-radius:6px;font-size:11px;text-align:center`;
        waitInd.textContent = 'Waiting for bot response...';
        content.appendChild(waitInd);

        const createButton = (id, text, bg, extra = '') => {
            const btn = document.createElement('button'); btn.id = id; btn.className = 'csp-btn';
            btn.style.cssText = `background:${bg};color:#fff;${extra}`; btn.textContent = text; return btn;
        };

        content.appendChild(createButton('captureBtn', 'Capture & Review', 'linear-gradient(135deg,#6c757d,#5a6268)', 'opacity:0.6;cursor:not-allowed'));
        const skipBtn = createButton('skipBtn', 'Skip This', '#fd7e14'); skipBtn.style.display = 'none'; content.appendChild(skipBtn);
        const skipAllBtn = createButton('skipAllBtn', 'Skip All', '#dc3545', 'font-size:10px;padding:6px'); skipAllBtn.style.display = 'none'; content.appendChild(skipAllBtn);
        content.appendChild(createButton('manageTestAccBtn', 'Test Account Settings', '#17a2b8'));

        const dlSection = document.createElement('div');
        dlSection.style.cssText = `border:1px solid #dee2e6;border-radius:8px;padding:8px;margin-bottom:8px;background:#f8f9fa`;
        dlSection.appendChild(createButton('downloadExcelBtn', '📊 Excel (0)', 'linear-gradient(135deg,#28a745,#218838)', 'margin-bottom:6px'));
        dlSection.appendChild(createButton('downloadPdfBtn', '📄 PDF (0)', 'linear-gradient(135deg,#dc3545,#c82333)', 'margin-bottom:6px'));
        dlSection.appendChild(createButton('downloadBothBtn', 'Export Both', env.gradient, 'font-size:10px;padding:7px;margin:0'));
        content.appendChild(dlSection);

        content.appendChild(createButton('retryUploadsBtn', '🔄 Retry Failed SP Uploads', '#6610f2', 'font-size:10px;padding:7px'));
        content.appendChild(createButton('envDashboardBtn', '📊 Environment Dashboard (Alt+D)', 'linear-gradient(135deg,#232F3E,#37475A)', 'font-size:10px;padding:7px'));

        const statusDiv = document.createElement('div'); statusDiv.id = 'statusText';
        statusDiv.style.cssText = `margin-top:8px;padding:8px;background:#f8f9fa;border-radius:6px;font-size:11px;color:#495057;text-align:center`;
        statusDiv.textContent = 'Entries: 0';
        content.appendChild(statusDiv);
        panel.appendChild(content);
    }

    function attachPanelListeners() {
        const $ = id => document.getElementById(id);
        $('minimizeBtn').onclick = e => { e.stopPropagation(); toggleMinimize(); };
        $('changeUserBtn').onclick = async () => {
            localStorage.removeItem(GLOBAL_KEYS.username); CryptoStore.clearKeyCache();
            const panel = $('chatScraperPanel'); panel.style.display = 'none';
            await showUsernamePrompt(); safeSetText($('usernameDisplay'), currentUsername); panel.style.display = 'block';
        };
        $('captureBtn').onclick = () => { const c = getCurrentResponse(); if (c) openReviewModal(c); else showNotification('No pending response', 'warning'); };
        $('skipBtn').onclick = () => { if (responseQueue.length > 0) skipCurrentResponse(); };
        $('skipAllBtn').onclick = () => { if (responseQueue.length > 1 && confirm(`Skip all ${responseQueue.length}?`)) skipAllResponses(); };
        $('manageTestAccBtn').onclick = openTestAccountSettings;
        $('downloadExcelBtn').onclick = downloadExcel;
        $('downloadPdfBtn').onclick = downloadPDF;
        $('downloadBothBtn').onclick = async () => { if (!chatData.length) return alert('No data'); await downloadExcel(); await downloadPDF(); };
        $('retryUploadsBtn').onclick = retryFailedUploads;
        $('envDashboardBtn').onclick = openEnvDashboard;
    }

    // ════════════════════════════════════════════════════════
    //  TEST ACCOUNT SETTINGS
    // ════════════════════════════════════════════════════════
    function openTestAccountSettings() {
        const modal = document.createElement('div');
        modal.style.cssText = `position:fixed;inset:0;background:rgba(0,0,0,0.85);z-index:100000;display:flex;align-items:center;justify-content:center;font-family:Arial`;

        const box = document.createElement('div');
        box.style.cssText = `background:#fff;width:500px;max-width:95vw;border-radius:12px;overflow:hidden`;

        const header = document.createElement('div');
        header.style.cssText = `background:#17a2b8;color:#fff;padding:15px 20px`;
        const h3 = document.createElement('h3'); h3.style.margin = '0'; h3.textContent = 'Test Account'; header.appendChild(h3);

        const body = document.createElement('div'); body.style.padding = '20px';

        const createField = (label, id, value, type = 'text') => {
            const wrap = document.createElement('div'); wrap.style.marginBottom = '15px';
            const lbl = document.createElement('label'); lbl.style.cssText = `font-weight:bold;font-size:12px`; lbl.textContent = label;
            const inp = document.createElement('input'); inp.type = type; inp.id = id; inp.value = value;
            inp.setAttribute('maxlength', '255'); inp.setAttribute('autocomplete', 'off');
            inp.style.cssText = `width:100%;padding:10px;border:1px solid #ddd;border-radius:6px;box-sizing:border-box;margin-top:4px`;
            wrap.appendChild(lbl); wrap.appendChild(inp); return wrap;
        };

        body.appendChild(createField('Email:', 'sEmail', savedTestAccount.email || '', 'email'));
        body.appendChild(createField('Customer ID:', 'sCid', savedTestAccount.customerId || ''));

        const gcsSelector = buildGcsLinkSelector(
            'settingsGcsContainer', 'sGcs', 'sGcsType',
            savedTestAccount.gcsLink || '', savedTestAccount.gcsLinkType || ''
        );
        body.appendChild(gcsSelector);

        const notice = document.createElement('div');
        notice.style.cssText = `padding:8px;background:#d4edda;border-radius:6px;margin-bottom:15px;font-size:10px;color:#155724;border:1px solid #c3e6cb`;
        notice.textContent = '🔒 Passwords are NOT stored or transmitted for security.';
        body.appendChild(notice);

        const btnRow = document.createElement('div'); btnRow.style.cssText = `display:flex;gap:10px`;
        const cancelBtn = document.createElement('button');
        cancelBtn.style.cssText = `flex:1;padding:10px;background:#6c757d;color:#fff;border:none;border-radius:6px;cursor:pointer`; cancelBtn.textContent = 'Cancel';
        const saveBtn = document.createElement('button');
        saveBtn.style.cssText = `flex:1;padding:10px;background:#28a745;color:#fff;border:none;border-radius:6px;cursor:pointer;font-weight:bold`; saveBtn.textContent = 'Save';
        btnRow.appendChild(cancelBtn); btnRow.appendChild(saveBtn);
        body.appendChild(btnRow);

        box.appendChild(header); box.appendChild(body); modal.appendChild(box); document.body.appendChild(modal);

        cancelBtn.onclick = () => modal.remove();
        saveBtn.onclick = () => {
            const gcsTypeSelect = modal.querySelector('#sGcsType');
            const gcsUrl = modal.querySelector('#sGcs')?.value.trim() || '';
            const gcsTypeVal = gcsTypeSelect ? gcsTypeSelect.value : '';
            saveTestAccount({
                email: modal.querySelector('#sEmail').value.trim(),
                customerId: modal.querySelector('#sCid').value.trim(),
                gcsLink: gcsUrl,
                gcsLinkType: gcsTypeVal
            });
            modal.remove();
            showNotification('Saved!', 'success');
        };
        modal.onclick = e => { if (e.target === modal) modal.remove(); };
    }

    function makeDraggable(el, handle) {
        let p1 = 0, p2 = 0, p3 = 0, p4 = 0;
        handle.onmousedown = e => {
            if (e.target.tagName === 'BUTTON') return;
            e.preventDefault(); p3 = e.clientX; p4 = e.clientY;
            document.onmouseup = () => { document.onmouseup = null; document.onmousemove = null; };
            document.onmousemove = e => { p1 = p3 - e.clientX; p2 = p4 - e.clientY; p3 = e.clientX; p4 = e.clientY; el.style.top = (el.offsetTop - p2) + "px"; el.style.left = (el.offsetLeft - p1) + "px"; };
        };
    }

    // ════════════════════════════════════════════════════════
    //  REVIEW MODAL
    // ════════════════════════════════════════════════════════
    function openReviewModal(responseData) {
        if (!responseData) return;
        const env = getActiveEnv();
        const sNo = chatData.length + 1;

        if (chatData.length >= LIMITS.MAX_CHAT_ENTRIES) {
            showNotification(`Entry limit (${LIMITS.MAX_CHAT_ENTRIES}) reached. Download and reset.`, 'error');
            return;
        }

        const queryTimestamp = responseData.timestamp || new Date().toISOString();
        const chatEntry = {
            sNo, queryTimestamp,
            queryLocalTime: formatTimestamp(queryTimestamp),
            testingDate: formatDate(queryTimestamp),
            testerLogin: currentUsername,
            conversationId: truncateString(responseData.conversationId || 'N/A', 255),
            messageId: truncateString(responseData.messageId || 'N/A', 255),
            query: truncateString(responseData.userMessage || 'N/A'),
            preProdBotResponse: truncateString(responseData.botMessage || 'N/A'),
            responseTime: responseData.responseTime || 0,
            responseTimeFormatted: responseData.responseTimeFormatted || 'N/A'
        };

        let pastedImageData = null;
        let pastedImageInfo = null;

        const modal = document.createElement('div');
        modal.id = 'reviewModal';
        modal.style.cssText = `position:fixed;inset:0;background:rgba(0,0,0,0.85);z-index:100000;display:flex;align-items:center;justify-content:center;font-family:Arial`;

        const box = document.createElement('div');
        box.style.cssText = `background:#fff;width:900px;max-width:95vw;max-height:95vh;border-radius:12px;overflow:hidden;display:flex;flex-direction:column`;

        // Header
        const headerDiv = document.createElement('div');
        headerDiv.style.cssText = `background:${env.headerGradient};color:#fff;padding:15px 20px;border-bottom:3px solid ${env.color}`;
        const headerRow = document.createElement('div'); headerRow.style.cssText = `display:flex;justify-content:space-between;align-items:center`;
        const headerInfo = document.createElement('div');
        const h3 = document.createElement('h3'); h3.style.cssText = `margin:0;display:flex;align-items:center;gap:8px`;
        const envSpan = document.createElement('span'); envSpan.style.cssText = `background:${env.color};padding:3px 8px;border-radius:4px;font-size:11px`; envSpan.textContent = env.label;
        h3.appendChild(envSpan);
        h3.appendChild(document.createTextNode(` #${sNo}`));
        headerInfo.appendChild(h3);
        const subInfo = document.createElement('p');
        subInfo.style.cssText = `margin:5px 0 0;font-size:11px;opacity:0.85`;
        subInfo.textContent = `${chatEntry.testerLogin} | ${chatEntry.testingDate} | ${chatEntry.responseTimeFormatted}`;
        headerInfo.appendChild(subInfo);
        headerRow.appendChild(headerInfo);

        if (responseQueue.length > 1) {
            const badge = document.createElement('div');
            badge.style.cssText = `background:#fd7e14;padding:4px 8px;border-radius:4px;font-size:10px;font-weight:bold`;
            badge.textContent = `${responseQueue.length - 1} more`;
            headerRow.appendChild(badge);
        }

        headerDiv.appendChild(headerRow);
        box.appendChild(headerDiv);

        // Scrollable body
        const bodyDiv = document.createElement('div');
        bodyDiv.style.cssText = `overflow-y:auto;padding:20px;flex:1`;

        // Screenshot section
        const screenshotSection = document.createElement('div');
        screenshotSection.style.marginBottom = '20px';
        const ssH4 = document.createElement('h4');
        ssH4.style.cssText = `margin:0 0 10px;border-bottom:2px solid ${env.color};padding-bottom:5px`;
        ssH4.textContent = 'Screenshot';
        screenshotSection.appendChild(ssH4);

        const pasteZone = document.createElement('div');
        pasteZone.id = 'pasteZone';
        pasteZone.tabIndex = 0;
        pasteZone.style.cssText = `border:3px dashed ${env.color};border-radius:12px;padding:25px;text-align:center;background:linear-gradient(135deg,#fff9f0,#fff3e0);cursor:pointer;min-height:120px;display:flex;flex-direction:column;align-items:center;justify-content:center;outline:none`;

        const pastePrompt = document.createElement('div');
        pastePrompt.id = 'pastePrompt';
        const promptH3 = document.createElement('h3');
        promptH3.style.cssText = `margin:0 0 6px;font-size:14px`;
        promptH3.textContent = 'Click here & Ctrl+V';
        const promptP = document.createElement('p');
        promptP.style.cssText = `margin:0;color:#666;font-size:11px`;
        promptP.textContent = 'Win+Shift+S then paste';
        pastePrompt.appendChild(promptH3);
        pastePrompt.appendChild(promptP);

        const pastePreview = document.createElement('div');
        pastePreview.id = 'pastePreview';
        pastePreview.style.cssText = `display:none;width:100%`;
        const previewImg = document.createElement('img');
        previewImg.id = 'previewImage';
        previewImg.style.cssText = `max-width:100%;max-height:250px;border-radius:8px`;
        const imageInfo = document.createElement('div');
        imageInfo.id = 'imageInfo';
        imageInfo.style.cssText = `margin-top:8px;font-size:10px;color:#666`;
        const clearImgBtn = document.createElement('button');
        clearImgBtn.id = 'clearImage';
        clearImgBtn.style.cssText = `padding:5px 14px;background:#dc3545;color:#fff;border:none;border-radius:4px;cursor:pointer;margin-top:8px`;
        clearImgBtn.textContent = '✕ Clear';
        pastePreview.appendChild(previewImg);
        pastePreview.appendChild(imageInfo);
        pastePreview.appendChild(document.createElement('br'));
        pastePreview.appendChild(clearImgBtn);

        pasteZone.appendChild(pastePrompt);
        pasteZone.appendChild(pastePreview);
        screenshotSection.appendChild(pasteZone);
        bodyDiv.appendChild(screenshotSection);

        // Auto-captured section
        const autoSection = document.createElement('div');
        autoSection.style.marginBottom = '20px';
        const autoH4 = document.createElement('h4');
        autoH4.style.cssText = `margin:0 0 10px;border-bottom:2px solid #17a2b8;padding-bottom:5px`;
        autoH4.textContent = 'Auto-Captured';
        autoSection.appendChild(autoH4);

        const autoGrid = document.createElement('div');
        autoGrid.style.cssText = `background:#f8f9fa;padding:12px;border-radius:8px;font-size:11px`;

        const idRow = document.createElement('div');
        idRow.style.cssText = `display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px`;

        const createInfoBox = (label, value) => {
            const d = document.createElement('div');
            const s = document.createElement('strong'); s.textContent = label; d.appendChild(s);
            const v = document.createElement('div');
            v.style.cssText = `background:#fff;padding:6px;border-radius:4px;margin-top:3px;word-break:break-all;font-size:10px;border:1px solid #ddd`;
            v.textContent = value; d.appendChild(v); return d;
        };

        idRow.appendChild(createInfoBox('Msg ID:', chatEntry.messageId));
        idRow.appendChild(createInfoBox('Conv ID:', chatEntry.conversationId));
        autoGrid.appendChild(idRow);

        const queryBox = document.createElement('div'); queryBox.style.marginBottom = '10px';
        const qLabel = document.createElement('strong'); qLabel.textContent = 'Query:'; queryBox.appendChild(qLabel);
        const qValue = document.createElement('div');
        qValue.style.cssText = `background:#e3f2fd;padding:8px;border-radius:4px;margin-top:3px;max-height:60px;overflow-y:auto`;
        qValue.textContent = chatEntry.query; queryBox.appendChild(qValue); autoGrid.appendChild(queryBox);

        const botBox = document.createElement('div');
        const bLabel = document.createElement('strong'); bLabel.textContent = 'Bot:'; botBox.appendChild(bLabel);
        const bValue = document.createElement('div');        bValue.style.cssText = `background:#e8f5e9;padding:8px;border-radius:4px;margin-top:3px;max-height:80px;overflow-y:auto`;
        bValue.textContent = chatEntry.preProdBotResponse; botBox.appendChild(bValue); autoGrid.appendChild(botBox);

        autoSection.appendChild(autoGrid);
        bodyDiv.appendChild(autoSection);

        // Assessment section
        const assessSection = document.createElement('div');
        assessSection.style.marginBottom = '20px';
        const assessH4 = document.createElement('h4');
        assessH4.style.cssText = `margin:0 0 10px;border-bottom:2px solid #28a745;padding-bottom:5px`;
        assessH4.textContent = 'Assessment';
        assessSection.appendChild(assessH4);

        // HVA Select with Custom Input
        const hvaWidget = buildHvaSelector('hvaSelect', 'customHvaInput');
        assessSection.appendChild(hvaWidget);

        // Status radio buttons
        const statusWrap = document.createElement('div'); statusWrap.style.marginBottom = '15px';
        const statusLabel = document.createElement('label'); statusLabel.style.fontWeight = 'bold'; statusLabel.textContent = 'Status ';
        const statusReq = document.createElement('span'); statusReq.style.color = 'red'; statusReq.textContent = '*';
        statusLabel.appendChild(statusReq); statusWrap.appendChild(statusLabel);

        const radioRow = document.createElement('div'); radioRow.style.cssText = `display:flex;gap:15px;margin-top:8px`;
        const createRadio = (value, bg, border, color) => {
            const lbl = document.createElement('label');
            lbl.style.cssText = `display:flex;align-items:center;gap:8px;padding:12px 30px;background:${bg};border-radius:8px;border:2px solid ${border};font-weight:bold;color:${color};cursor:pointer;flex:1;justify-content:center`;
            const inp = document.createElement('input'); inp.type = 'radio'; inp.name = 'responseStatus'; inp.value = value;
            inp.style.cssText = `width:18px;height:18px`;
            lbl.appendChild(inp); lbl.appendChild(document.createTextNode(` ${value}`)); return lbl;
        };

        radioRow.appendChild(createRadio('Correct', '#d4edda', '#28a745', '#155724'));
        radioRow.appendChild(createRadio('Incorrect', '#f8d7da', '#dc3545', '#721c24'));
        statusWrap.appendChild(radioRow);
        assessSection.appendChild(statusWrap);

        // Ground Truth
        const gtWrap = document.createElement('div'); gtWrap.style.marginBottom = '15px';
        const gtLabel = document.createElement('label'); gtLabel.style.fontWeight = 'bold'; gtLabel.textContent = 'Ground Truth';
        gtWrap.appendChild(gtLabel);
        const gtTextarea = document.createElement('textarea'); gtTextarea.id = 'groundTruth';
        gtTextarea.placeholder = 'Expected response...'; gtTextarea.setAttribute('maxlength', String(LIMITS.MAX_STRING_LENGTH));
        gtTextarea.style.cssText = `width:100%;height:80px;padding:10px;border:2px solid #ddd;border-radius:6px;margin-top:5px;box-sizing:border-box;resize:vertical`;
        gtWrap.appendChild(gtTextarea); assessSection.appendChild(gtWrap);

        // Observations
        const obsWrap = document.createElement('div'); obsWrap.style.marginBottom = '15px';
        const obsLabel = document.createElement('label'); obsLabel.style.fontWeight = 'bold'; obsLabel.textContent = 'Observations';
        obsWrap.appendChild(obsLabel);
        const obsTextarea = document.createElement('textarea'); obsTextarea.id = 'observations';
        obsTextarea.placeholder = 'Notes...'; obsTextarea.setAttribute('maxlength', String(LIMITS.MAX_STRING_LENGTH));
        obsTextarea.style.cssText = `width:100%;height:70px;padding:10px;border:2px solid #ddd;border-radius:6px;margin-top:5px;box-sizing:border-box;resize:vertical`;
        obsWrap.appendChild(obsTextarea); assessSection.appendChild(obsWrap);

        // Area of Improvement (regression only)
        let aoiTextarea = null;
        if (env.hasAreaOfImprovement) {
            const aoiWrap = document.createElement('div'); aoiWrap.style.marginBottom = '15px';
            const aoiLabel = document.createElement('label'); aoiLabel.style.fontWeight = 'bold'; aoiLabel.textContent = 'Area of Improvement';
            aoiWrap.appendChild(aoiLabel);
            aoiTextarea = document.createElement('textarea'); aoiTextarea.id = 'areaOfImprovement';
            aoiTextarea.placeholder = 'Describe areas where the bot response could be improved...';
            aoiTextarea.setAttribute('maxlength', String(LIMITS.MAX_STRING_LENGTH));
            aoiTextarea.style.cssText = `width:100%;height:80px;padding:10px;border:2px solid #6f42c1;border-radius:6px;margin-top:5px;box-sizing:border-box;resize:vertical;background:#f8f5ff`;
            aoiWrap.appendChild(aoiTextarea); assessSection.appendChild(aoiWrap);
        }

        bodyDiv.appendChild(assessSection);

        // Test Account section
        const testSection = document.createElement('div'); testSection.style.marginBottom = '10px';
        const testH4 = document.createElement('h4');
        testH4.style.cssText = `margin:0 0 10px;border-bottom:2px solid #6c757d;padding-bottom:5px`;
        testH4.textContent = 'Test Account';
        testSection.appendChild(testH4);

        const testGrid = document.createElement('div');
        testGrid.style.cssText = `display:grid;grid-template-columns:1fr 1fr;gap:15px`;

        const createTestField = (label, id, value, type = 'text') => {
            const d = document.createElement('div');
            const l = document.createElement('label'); l.style.cssText = `font-size:11px;font-weight:bold`; l.textContent = label; d.appendChild(l);
            const inp = document.createElement('input'); inp.type = type; inp.id = id; inp.value = value;
            inp.setAttribute('maxlength', '255'); inp.setAttribute('autocomplete', 'off');
            inp.style.cssText = `width:100%;padding:8px;border:1px solid #ddd;border-radius:6px;font-size:11px;box-sizing:border-box`;
            d.appendChild(inp); return d;
        };

        testGrid.appendChild(createTestField('Email:', 'testEmail', savedTestAccount.email || '', 'email'));
        testGrid.appendChild(createTestField('Customer ID:', 'testCustomerId', savedTestAccount.customerId || ''));
        testSection.appendChild(testGrid);

        // GCS/MCS Link Selector in Review Modal
        const gcsReviewSelector = buildGcsLinkSelector(
            'reviewGcsContainer', 'gcsLink', 'gcsLinkType',
            savedTestAccount.gcsLink || '', savedTestAccount.gcsLinkType || ''
        );
        gcsReviewSelector.style.marginTop = '12px';
        testSection.appendChild(gcsReviewSelector);

        const testNotice = document.createElement('div');
        testNotice.style.cssText = `padding:6px;background:#fff3cd;border-radius:4px;font-size:9px;color:#856404;text-align:center;border:1px solid #ffeeba;margin-top:10px`;
        testNotice.textContent = '🔒 Passwords are never stored or sent.';
        testSection.appendChild(testNotice);

        bodyDiv.appendChild(testSection);
        box.appendChild(bodyDiv);

        // Footer
        const footer = document.createElement('div');
        footer.style.cssText = `display:flex;justify-content:space-between;align-items:center;padding:12px 20px;background:#f8f9fa;border-top:1px solid #ddd`;

        const footerLeft = document.createElement('div'); footerLeft.style.cssText = `display:flex;align-items:center;gap:10px`;
        const imageStatus = document.createElement('span'); imageStatus.id = 'imageStatus';
        imageStatus.style.cssText = `font-size:11px;color:#dc3545;font-weight:bold`; imageStatus.textContent = 'Screenshot required';
        const saveAccLabel = document.createElement('label'); saveAccLabel.style.cssText = `font-size:11px;color:#666;cursor:pointer`;
        const saveAccCheck = document.createElement('input'); saveAccCheck.type = 'checkbox'; saveAccCheck.id = 'saveTestAccCheck'; saveAccCheck.checked = true;
        saveAccLabel.appendChild(saveAccCheck); saveAccLabel.appendChild(document.createTextNode(' Save acct'));
        footerLeft.appendChild(imageStatus); footerLeft.appendChild(saveAccLabel);

        const footerRight = document.createElement('div'); footerRight.style.cssText = `display:flex;gap:10px`;
        const skipFromModalBtn = document.createElement('button'); skipFromModalBtn.id = 'skipFromModalBtn';
        skipFromModalBtn.style.cssText = `padding:10px 15px;background:#fd7e14;color:#fff;border:none;border-radius:6px;cursor:pointer`; skipFromModalBtn.textContent = 'Skip';
        const cancelBtn = document.createElement('button'); cancelBtn.id = 'cancelBtn';
        cancelBtn.style.cssText = `padding:10px 20px;background:#6c757d;color:#fff;border:none;border-radius:6px;cursor:pointer`; cancelBtn.textContent = 'Cancel';
        const saveBtn = document.createElement('button'); saveBtn.id = 'saveBtn'; saveBtn.disabled = true;
        saveBtn.style.cssText = `padding:10px 30px;background:${env.color};color:#fff;border:none;border-radius:6px;cursor:pointer;font-weight:bold;opacity:0.5`; saveBtn.textContent = 'Save';
        footerRight.appendChild(skipFromModalBtn); footerRight.appendChild(cancelBtn); footerRight.appendChild(saveBtn);

        footer.appendChild(footerLeft); footer.appendChild(footerRight);
        box.appendChild(footer);
        modal.appendChild(box);
        document.body.appendChild(modal);

        setTimeout(() => pasteZone.focus(), 100);

        // Validation helper
        const updateSaveBtn = () => {
            const hvaVal = getHvaValue('hvaSelect', 'customHvaInput');
            const ok = pastedImageData && hvaVal && modal.querySelector('input[name="responseStatus"]:checked');
            saveBtn.disabled = !ok;
            saveBtn.style.opacity = ok ? '1' : '0.5';
        };

        // Paste handler
        const handlePaste = async (e) => {
            const items = e.clipboardData?.items;
            if (!items) return;
            for (const item of items) {
                if (item.type.startsWith('image/')) {
                    e.preventDefault();
                    const blob = item.getAsFile();
                    if (!blob) return;
                    const reader = new FileReader();
                    reader.onload = async (ev) => {
                        try {
                            safeSetText(promptH3, 'Processing...'); promptP.textContent = '';
                            if (!validateImageDataUrl(ev.target.result)) { safeSetText(promptH3, 'Image too large or invalid'); return; }
                            const proc = await processImageForStorage(ev.target.result);
                            pastedImageData = proc.dataUrl; pastedImageInfo = proc;
                            previewImg.src = pastedImageData;
                            pastePrompt.style.display = 'none'; pastePreview.style.display = 'block';
                            pasteZone.style.borderStyle = 'solid'; pasteZone.style.borderColor = '#28a745'; pasteZone.style.background = '#f0fff0';
                            safeSetText(imageInfo, `${proc.originalWidth}x${proc.originalHeight}px | ~${proc.fileSize}KB`);
                            updateSaveBtn();
                            safeSetText(imageStatus, `✓ ${proc.originalWidth}x${proc.originalHeight}`);
                            imageStatus.style.color = '#28a745';
                        } catch (err) {
                            SecureLog.error('Paste processing error', err);
                            safeSetText(promptH3, 'Error processing image'); promptH3.style.color = 'red';
                            pastePrompt.style.display = 'flex'; pastePreview.style.display = 'none';
                        }
                    };
                    reader.onerror = () => { safeSetText(promptH3, 'Failed to read image'); promptH3.style.color = 'red'; };
                    reader.readAsDataURL(blob);
                    break;
                }
            }
        };

        pasteZone.addEventListener('paste', handlePaste);
        modal.addEventListener('paste', handlePaste);
        pasteZone.addEventListener('click', () => pasteZone.focus());

        // Clear image
        clearImgBtn.onclick = e => {
            e.stopPropagation();
            pastedImageData = null; pastedImageInfo = null; previewImg.src = '';
            safeSetText(promptH3, 'Click here & Ctrl+V'); promptH3.style.color = '';
            safeSetText(promptP, 'Win+Shift+S then paste');
            pastePrompt.style.display = 'flex'; pastePreview.style.display = 'none';
            pasteZone.style.borderStyle = 'dashed'; pasteZone.style.borderColor = env.color;
            pasteZone.style.background = 'linear-gradient(135deg,#fff9f0,#fff3e0)';
            updateSaveBtn();
            safeSetText(imageStatus, 'Screenshot required'); imageStatus.style.color = '#dc3545';
            pasteZone.focus();
        };

        // Validation listeners
        const hvaSelect = modal.querySelector('#hvaSelect');
        const customHvaInput = modal.querySelector('#customHvaInput');
        if (hvaSelect) hvaSelect.onchange = updateSaveBtn;
        if (customHvaInput) customHvaInput.oninput = updateSaveBtn;
        modal.querySelectorAll('input[name="responseStatus"]').forEach(r => { r.onchange = updateSaveBtn; });

        // Cancel
        cancelBtn.onclick = () => { pastedImageData = null; pastedImageInfo = null; modal.remove(); };

        // Skip from modal
        skipFromModalBtn.onclick = () => { pastedImageData = null; pastedImageInfo = null; modal.remove(); skipCurrentResponse(); };

        // ★ SAVE HANDLER — separated GCS type and URL
        saveBtn.onclick = async () => {
            const hva = getHvaValue('hvaSelect', 'customHvaInput');
            const status = modal.querySelector('input[name="responseStatus"]:checked')?.value;
            if (!pastedImageData || !hva || !status) return;

            saveBtn.disabled = true; saveBtn.textContent = 'Saving...';

            // Get GCS/MCS type and URL SEPARATELY
            const gcsLinkUrl = getGcsLinkFromSelector('reviewGcsContainer', 'gcsLink', 'gcsLinkType');
            const gcsLinkTypeLabel = getGcsLinkTypeLabel('gcsLinkType', 'reviewGcsContainer');
            const gcsTypeSelectVal = modal.querySelector('#gcsLinkType')?.value || '';

            // Save test account if checked
            if (saveAccCheck.checked) {
                saveTestAccount({
                    email: modal.querySelector('#testEmail').value.trim(),
                    customerId: modal.querySelector('#testCustomerId').value.trim(),
                    gcsLink: gcsLinkUrl,
                    gcsLinkType: gcsTypeSelectVal
                });
            }

            const entryId = `${activeEnvId}_entry_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`;
            let idbOk = false;

            for (let a = 1; a <= 3; a++) {
                try {
                    await ImageStore.save(entryId, pastedImageData, {
                        width: pastedImageInfo?.processedWidth || 800,
                        height: pastedImageInfo?.processedHeight || 600
                    });
                    const v = await ImageStore.verify(entryId);
                    if (v.isImage) { idbOk = true; break; }
                } catch (e) { SecureLog.error(`Image save attempt ${a}`, e); }
                if (a < 3) await new Promise(r => setTimeout(r, 500));
            }

            if (!idbOk) {
                saveBtn.disabled = false; saveBtn.textContent = 'Save';
                showNotification('Image save failed!', 'error'); return;
            }

            const savedTimestamp = new Date().toISOString();
            const entry = {
                sNo: chatEntry.sNo,
                entryId,
                environment: env.label,
                environmentId: env.id,
                hva: truncateString(hva, 255),
                messageId: chatEntry.messageId,
                conversationId: chatEntry.conversationId,
                query: chatEntry.query,
                preProdBotResponse: chatEntry.preProdBotResponse,
                responseCorrectOrIncorrect: status,
                groundTruthResponse: truncateString(gtTextarea.value.trim()),
                observations: truncateString(obsTextarea.value.trim()),
                _hasImage: true,
                imageWidth: pastedImageInfo?.processedWidth || 800,
                imageHeight: pastedImageInfo?.processedHeight || 600,
                imageOriginalWidth: pastedImageInfo?.originalWidth || 800,
                imageOriginalHeight: pastedImageInfo?.originalHeight || 600,
                imageFileSize: pastedImageInfo?.fileSize || 0,
                gcsLinkType: truncateString(gcsLinkTypeLabel, 255),
                gcsLink: truncateString(gcsLinkUrl, 500),
                testAccountEmail: truncateString(modal.querySelector('#testEmail').value.trim(), 255),
                testAccountCustomerId: truncateString(modal.querySelector('#testCustomerId').value.trim(), 255),
                testingDate: chatEntry.testingDate,
                testerLogin: chatEntry.testerLogin,
                responseTime: chatEntry.responseTime,
                responseTimeFormatted: chatEntry.responseTimeFormatted,
                queryTimestamp: chatEntry.queryTimestamp,
                queryLocalTime: chatEntry.queryLocalTime,
                savedTimestamp,
                savedLocalTime: formatTimestamp(savedTimestamp),
                _spImageAttached: false
            };

            if (env.hasAreaOfImprovement && aoiTextarea) {
                entry.areaOfImprovement = truncateString(aoiTextarea.value.trim());
            }

            chatData.push(entry);
            await saveDataToStorage();
            removeCurrentFromQueue();
            updateStatus();

            pastedImageData = null; pastedImageInfo = null;
            modal.remove();

            showNotification(`[${env.shortLabel}] #${entry.sNo} saved! ${status}`, 'success');
            setTimeout(() => autoUploadEntryToList(entry), 500);
        };

        // Close on backdrop
        modal.onclick = e => { if (e.target === modal) { pastedImageData = null; pastedImageInfo = null; modal.remove(); } };

        // Escape key
        const escKey = e => {
            if (e.key === 'Escape') { pastedImageData = null; pastedImageInfo = null; modal.remove(); document.removeEventListener('keydown', escKey); }
        };
        document.addEventListener('keydown', escKey);
    }

    // ════════════════════════════════════════════════════════
    //  STATE UPDATES
    // ════════════════════════════════════════════════════════
    function updateCaptureButtonState(ready) {
        const btn = getElement('captureBtn'); if (!btn) return;
        const env = getActiveEnv();
        if (ready) {
            btn.style.background = env.gradient; btn.style.cursor = 'pointer'; btn.style.opacity = '1';
            btn.textContent = `Capture & Review (${responseQueue.length})`;
        } else {
            btn.style.background = 'linear-gradient(135deg,#6c757d,#5a6268)'; btn.style.cursor = 'not-allowed'; btn.style.opacity = '0.6';
            btn.textContent = 'Capture & Review';
        }
    }

    function updateStatus() {
        const el = getElement('statusText'), de = getElement('downloadExcelBtn'), dp = getElement('downloadPdfBtn'), db = getElement('downloadBothBtn');
        const env = getActiveEnv();
        const imgs = chatData.filter(d => d._hasImage).length;
        const spOk = chatData.filter(d => d._spUploaded === true).length;
        const spFail = chatData.filter(d => d._spUploaded === false).length;
        const correct = chatData.filter(d => d.responseCorrectOrIncorrect === 'Correct').length;
        const incorrect = chatData.filter(d => d.responseCorrectOrIncorrect === 'Incorrect').length;

        if (el) {
            el.textContent = '';
            const wrapper = document.createElement('div');
            wrapper.style.cssText = `display:flex;align-items:center;justify-content:center;gap:6px;flex-wrap:wrap`;
            const envBadge = document.createElement('span');
            envBadge.style.cssText = `background:${env.color};color:#fff;padding:1px 5px;border-radius:3px;font-size:9px`;
            envBadge.textContent = env.shortLabel; wrapper.appendChild(envBadge);
            const countSpan = document.createElement('span');
            const countStrong = document.createElement('strong'); countStrong.textContent = chatData.length;
            countSpan.textContent = 'Entries: '; countSpan.appendChild(countStrong); wrapper.appendChild(countSpan);
            if (chatData.length > 0) {
                const cSpan = document.createElement('span'); cSpan.style.color = '#155724'; cSpan.textContent = `✓${correct}`; wrapper.appendChild(cSpan);
                const iSpan = document.createElement('span'); iSpan.style.color = '#721c24'; iSpan.textContent = `✗${incorrect}`; wrapper.appendChild(iSpan);
            }
            if (spOk > 0) { const spSpan = document.createElement('span'); spSpan.style.cssText = `color:#0078d4;font-size:10px`; spSpan.textContent = `☁${spOk}`; wrapper.appendChild(spSpan); }
            if (spFail > 0) { const sfSpan = document.createElement('span'); sfSpan.style.cssText = `color:#dc3545;font-size:10px`; sfSpan.textContent = `⚠${spFail}`; wrapper.appendChild(sfSpan); }
            el.appendChild(wrapper);
        }
        if (de) de.textContent = `📊 Excel (${chatData.length})`;
        if (dp) dp.textContent = `📄 PDF (${imgs})`;
        if (db) db.textContent = `Export Both (${chatData.length})`;
        updateBubble();
    }

    // ════════════════════════════════════════════════════════
    //  NOTIFICATION
    // ════════════════════════════════════════════════════════
    function showNotification(msg, type = 'info') {
        const colors = { success: '#28a745', error: '#dc3545', info: '#17a2b8', warning: '#ffc107' };
        const n = document.createElement('div');
        n.style.cssText = `position:fixed;bottom:20px;right:20px;padding:12px 20px;border-radius:8px;color:${type === 'warning' ? '#212529' : '#fff'};font-family:Arial;font-size:13px;z-index:200000;background:${colors[type]};box-shadow:0 4px 15px rgba(0,0,0,0.3);transform:translateX(100%);transition:transform 0.25s;max-width:400px`;
        n.textContent = msg; document.body.appendChild(n);
        requestAnimationFrame(() => n.style.transform = 'translateX(0)');
        setTimeout(() => { n.style.transform = 'translateX(100%)'; setTimeout(() => n.remove(), 250); }, 4000);
    }

    // ════════════════════════════════════════════════════════
    //  EXCEL GENERATOR
    // ════════════════════════════════════════════════════════
    async function generateExcelBuffer(entries) {
        const env = getActiveEnv();
        const wb = new ExcelJS.Workbook(); wb.creator = currentUsername; wb.created = new Date();
        const sh = wb.addWorksheet(`${env.label} Data`, { views: [{ state: 'frozen', ySplit: 1, xSplit: 1 }] });

        const cols = [
            { header: 'S.No', key: 'sNo', width: 6 },
            { header: 'Environment', key: 'environment', width: 14 },
            { header: 'HVA', key: 'hva', width: 20 },
            { header: 'Message ID', key: 'messageId', width: 30 },
            { header: 'Conversation ID', key: 'conversationId', width: 30 },
            { header: 'Query', key: 'query', width: 45 },
            { header: 'Bot Response', key: 'botResponse', width: 50 },
            { header: 'Response Time', key: 'responseTime', width: 14 },
            { header: 'Status', key: 'status', width: 18 },
            { header: 'Ground Truth', key: 'groundTruth', width: 45 },
            { header: 'Observations', key: 'observations', width: 35 }
        ];
        if (env.hasAreaOfImprovement) cols.push({ header: 'Area of Improvement', key: 'areaOfImprovement', width: 45 });
        cols.push(
            { header: 'Screenshot Ref', key: 'screenshotRef', width: 25 },
            { header: 'Link Type', key: 'linkType', width: 12 },
            { header: 'GCS/MCS Link', key: 'gcsLink', width: 35 },
            { header: 'Test Email', key: 'testEmail', width: 25 },
            { header: 'Test Customer ID', key: 'testCustId', width: 22 },
            { header: 'Testing Date', key: 'testDate', width: 14 },
            { header: 'Tester', key: 'tester', width: 14 },
            { header: 'Query Time', key: 'queryTime', width: 16 },
            { header: 'Saved Time', key: 'savedTime', width: 16 }
        );
        sh.columns = cols;

        const hr = sh.getRow(1);
        hr.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
        hr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF232F3E' } };
        hr.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }; hr.height = 30;

        let imgIdx = 0;
        const totalImg = entries.filter(d => d._hasImage).length;
        const batches = Math.ceil(totalImg / CONFIG.PDF_BATCH_SIZE);

        for (let i = 0; i < entries.length; i++) {
            const item = entries[i];
            let ref = 'No image';
            if (item._hasImage) {
                imgIdx++;
                const pg = ((imgIdx - 1) % CONFIG.PDF_BATCH_SIZE) + 2;
                ref = batches > 1 ? `PDF Part ${Math.ceil(imgIdx / CONFIG.PDF_BATCH_SIZE)}, Pg ${pg}` : `PDF Pg ${pg}`;
            }
            const rowData = {
                sNo: item.sNo, environment: item.environment || env.label, hva: item.hva,
                messageId: item.messageId, conversationId: item.conversationId, query: item.query,
                botResponse: item.preProdBotResponse, responseTime: item.responseTimeFormatted,
                status: item.responseCorrectOrIncorrect, groundTruth: item.groundTruthResponse || '',
                observations: item.observations || '', screenshotRef: ref,
                linkType: item.gcsLinkType || '',
                gcsLink: item.gcsLink || '',
                testEmail: item.testAccountEmail || '',
                testCustId: item.testAccountCustomerId || '', testDate: item.testingDate,
                tester: item.testerLogin, queryTime: item.queryLocalTime || '', savedTime: item.savedLocalTime || ''
            };
            if (env.hasAreaOfImprovement) rowData.areaOfImprovement = item.areaOfImprovement || '';

            const row = sh.addRow(rowData);
            if (i % 2 === 0) row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8F9FA' } };
            const statusCol = cols.findIndex(c => c.key === 'status') + 1;
            const sc = { 'Correct': { bg: 'FFD4EDDA', t: 'FF155724' }, 'Incorrect': { bg: 'FFF8D7DA', t: 'FF721C24' } }[item.responseCorrectOrIncorrect];
            if (sc && statusCol > 0) {
                row.getCell(statusCol).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: sc.bg } };
                row.getCell(statusCol).font = { bold: true, color: { argb: sc.t } };
            }
            const totalCols = cols.length;
            for (let c = 1; c <= totalCols; c++) {
                row.getCell(c).alignment = { wrapText: true, vertical: 'top' };
                row.getCell(c).border = { top: { style: 'thin', color: { argb: 'FFE0E0E0' } }, bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } } };
            }
            row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' }; row.height = 45;
        }
        return await wb.xlsx.writeBuffer();
    }

    // ════════════════════════════════════════════════════════
    //  NETWORK INTERCEPTION
    // ════════════════════════════════════════════════════════
    const EXCLUDED_URLS = [
        /login/i, /auth/i, /oauth/i, /token/i, /signin/i, /password/i, /credential/i, /saml/i,
        /\.css/i, /\.js(\?|$)/i, /\.png/i, /\.jpg/i, /\.gif/i, /\.svg/i, /\.woff/i,
        /analytics/i, /tracking/i, /telemetry/i, /sharepoint\.com/i, /microsoftonline/i, /contextinfo/i
    ];

    function shouldIntercept(url) {
        if (!url || typeof url !== 'string') return false;
        if (EXCLUDED_URLS.some(p => p.test(url))) return false;
        const lower = url.toLowerCase();
        return lower.includes('chat') || lower.includes('message') || lower.includes('conversation') ||
               lower.includes('bot') || lower.includes('assist') || lower.includes('query') || lower.includes('copilot');
    }

    const onBotResponse = debounce(data => {
        if (!isValidMessage(data.userMessage) || !isValidMessage(data.botMessage)) return;
        addToQueue(data);
        showNotification(`[${getActiveEnv().shortLabel}] Bot response! (${data.responseTimeFormatted}) | Queue: ${responseQueue.length}`, 'success');
    }, 150);

    const origXHROpen = XMLHttpRequest.prototype.open;
    const origXHRSend = XMLHttpRequest.prototype.send;

    XMLHttpRequest.prototype.open = function (m, u, ...r) {
        Object.defineProperty(this, '_scraperUrl', { value: u, writable: true, enumerable: false, configurable: true });
        return origXHROpen.apply(this, [m, u, ...r]);
    };

    XMLHttpRequest.prototype.send = function (body) {
        if (this._scraperUrl && shouldIntercept(this._scraperUrl)) {
            try {
                const r = JSON.parse(body);
                pendingRequest = {
                    conversationId: truncateString(r.conversationId || r.conversation_id || r.sessionId || 'N/A', 255),
                    userMessage: truncateString(r.message || r.userMessage || r.query || r.text || r.input || r.prompt || 'N/A'),
                    timestamp: new Date().toISOString(),
                    requestStartTime: performance.now()
                };
            } catch {}

            this.addEventListener('load', function () {
                if (this.status >= 200 && this.status < 300 && pendingRequest) {
                    try {
                        const t = performance.now(); const rt = t - pendingRequest.requestStartTime;
                        const res = JSON.parse(this.responseText);
                        onBotResponse({
                            ...pendingRequest,
                            messageId: truncateString(res.messageId || res.message_id || res.id || 'N/A', 255),
                            botMessage: truncateString(res.response || res.message || res.botMessage || res.botResponse || res.answer || res.text || res.output || res.reply || 'N/A'),
                            responseTime: rt, responseTimeFormatted: formatResponseTime(rt)
                        });
                        pendingRequest = null;
                    } catch {}
                }
            });
        }
        return origXHRSend.apply(this, arguments);
    };

    const origFetch = window.fetch;
    window.fetch = function (...args) {
        const url = typeof args[0] === 'string' ? args[0] : args[0]?.url;
        if (url && shouldIntercept(url)) {
            const opts = args[1] || {}; const st = performance.now();
            try {
                const r = JSON.parse(opts.body);
                pendingRequest = {
                    conversationId: truncateString(r.conversationId || r.conversation_id || r.sessionId || 'N/A', 255),
                    userMessage: truncateString(r.message || r.userMessage || r.query || r.text || r.input || r.prompt || 'N/A'),
                    timestamp: new Date().toISOString(),
                    requestStartTime: st
                };
            } catch {}
            return origFetch.apply(this, args).then(response => {
                const et = performance.now();
                response.clone().json().then(res => {
                    if (pendingRequest) {
                        const rt = et - pendingRequest.requestStartTime;
                        onBotResponse({
                            ...pendingRequest,
                            messageId: truncateString(res.messageId || res.message_id || res.id || 'N/A', 255),
                            botMessage: truncateString(res.response || res.message || res.botMessage || res.botResponse || res.answer || res.text || res.output || res.reply || 'N/A'),
                            responseTime: rt, responseTimeFormatted: formatResponseTime(rt)
                        });
                        pendingRequest = null;
                    }
                }).catch(() => {});
                return response;
            });
        }
        return origFetch.apply(this, args);
    };

    // ════════════════════════════════════════════════════════
    //  EXCEL DOWNLOAD
    // ════════════════════════════════════════════════════════
    async function downloadExcel() {
        if (!chatData.length) return alert('No data');
        const env = getActiveEnv();
        const loader = document.createElement('div');
        loader.style.cssText = `position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);background:#232F3E;color:#fff;padding:30px 50px;border-radius:12px;z-index:150000;font-family:Arial;border:2px solid #28a745;text-align:center;min-width:300px`;
        const loaderText = document.createElement('div'); loaderText.style.cssText = `font-size:16px;font-weight:bold`;
        loaderText.textContent = `Generating ${env.label} Excel...`; loader.appendChild(loaderText); document.body.appendChild(loader);
        try {
            const buffer = await generateExcelBuffer(chatData);
            const fileName = `${env.id}_Results_${currentUsername}_${new Date().toISOString().slice(0, 10)}.xlsx`;
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a'); link.href = url; link.download = fileName;
            document.body.appendChild(link); link.click(); document.body.removeChild(link);
            setTimeout(() => URL.revokeObjectURL(url), 5000);
            loader.remove();
            showNotification(`[${env.shortLabel}] Excel: ${chatData.length} entries`, 'success');
        } catch (e) { loader.remove(); SecureLog.error('Excel generation failed', e); showNotification('Excel failed', 'error'); }
    }

    // ════════════════════════════════════════════════════════
    //  COLOR HELPER
    // ════════════════════════════════════════════════════════
    function hexToRgb(hex) {
        const h = hex.replace('#', '');
        return [parseInt(h.substring(0, 2), 16), parseInt(h.substring(2, 4), 16), parseInt(h.substring(4, 6), 16)];
    }

    // ════════════════════════════════════════════════════════
    //  PDF DOWNLOAD
    // ════════════════════════════════════════════════════════
    async function downloadPDF() {
        if (!chatData.length) return alert('No data');
        const env = getActiveEnv();
        const entries = chatData.filter(d => d._hasImage && d.entryId);
        if (!entries.length) return alert('No images');

        let okCount = 0, missCount = 0;
        for (const item of entries) { if (await ImageStore.has(item.entryId)) okCount++; else missCount++; }
        if (okCount === 0) return alert('All images missing!');
        if (missCount > 0 && !confirm(`${missCount} missing. ${okCount} available. Continue?`)) return;
        await generatePDFFromEntries(entries, env, true);
    }

    async function downloadPDFSilent(data, env) {
        const entries = data.filter(d => d._hasImage && d.entryId);
        if (!entries.length) return;
        let okCount = 0;
        for (const item of entries) { if (await ImageStore.has(item.entryId)) okCount++; }
        if (okCount === 0) return;
        await generatePDFFromEntries(entries, env, false);
    }

    async function generatePDFFromEntries(entries, env, showLoader = true) {
        const totalBatches = Math.ceil(entries.length / CONFIG.PDF_BATCH_SIZE);
        const correct = chatData.filter(d => d.responseCorrectOrIncorrect === 'Correct').length;
        const incorrect = chatData.filter(d => d.responseCorrectOrIncorrect === 'Incorrect').length;
        const accuracy = chatData.length > 0 ? ((correct / chatData.length) * 100).toFixed(1) : '0.0';

        let loader = null, pdfBar = null, pdfProg = null;
        if (showLoader) {
            loader = document.createElement('div');
            loader.style.cssText = `position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);background:#232F3E;color:#fff;padding:30px 50px;border-radius:12px;z-index:150000;font-family:Arial;border:2px solid ${env.color};text-align:center;min-width:350px`;
            const title = document.createElement('div'); title.style.cssText = `font-size:16px;font-weight:bold`; title.textContent = `Generating ${env.label} PDF`; loader.appendChild(title);
            pdfProg = document.createElement('div'); pdfProg.style.cssText = `margin-top:10px;font-size:12px;color:${env.color}`; pdfProg.textContent = 'Preparing...'; loader.appendChild(pdfProg);
            const barContainer = document.createElement('div'); barContainer.style.cssText = `margin-top:12px;background:#37475A;border-radius:10px;overflow:hidden;height:10px`;
            pdfBar = document.createElement('div'); pdfBar.style.cssText = `width:0%;height:100%;background:${env.color};transition:width 0.3s;border-radius:10px`;
            barContainer.appendChild(pdfBar); loader.appendChild(barContainer); document.body.appendChild(loader);
        }

        const setP = (p, t) => { if (pdfBar) pdfBar.style.width = p + '%'; if (pdfProg) pdfProg.textContent = t; };

        try {
            const dateStr = new Date().toISOString().slice(0, 10);
            let gIdx = 0, totalKB = 0, success = 0, fail = 0;

            for (let bn = 0; bn < totalBatches; bn++) {
                const bStart = bn * CONFIG.PDF_BATCH_SIZE;
                const bEnd = Math.min(bStart + CONFIG.PDF_BATCH_SIZE, entries.length);
                const batch = entries.slice(bStart, bEnd);
                const bLabel = totalBatches > 1 ? ` (Part ${bn + 1}/${totalBatches})` : '';
                const { jsPDF } = window.jspdf;
                const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
                const pw = pdf.internal.pageSize.getWidth(), ph = pdf.internal.pageSize.getHeight(), m = 10, uw = pw - m * 2;

                // Cover page
                pdf.setFillColor(35, 47, 62); pdf.rect(0, 0, pw, ph, 'F');
                pdf.setFillColor(...hexToRgb(env.color)); pdf.rect(0, 55, pw, 4, 'F');
                pdf.setTextColor(255, 255, 255); pdf.setFontSize(26); pdf.setFont('helvetica', 'bold');
                pdf.text(`${env.label} Screenshots${bLabel}`, pw / 2, 40, { align: 'center' });
                pdf.setFontSize(13); pdf.setFont('helvetica', 'normal'); pdf.setTextColor(...hexToRgb(env.color));
                pdf.text(`Tester: ${currentUsername}`, pw / 2, 72, { align: 'center' });
                pdf.setTextColor(200, 200, 200); pdf.setFontSize(11);
                pdf.text(formatLocalTime(new Date().toISOString()), pw / 2, 82, { align: 'center' });
                pdf.text(`Entries ${bStart + 1}-${bEnd} of ${entries.length}`, pw / 2, 90, { align: 'center' });
                pdf.setTextColor(40, 167, 69); pdf.setFontSize(12);
                pdf.text(`Correct: ${correct}`, pw / 2 - 40, 105, { align: 'center' });
                pdf.setTextColor(220, 53, 69); pdf.text(`Incorrect: ${incorrect}`, pw / 2 + 40, 105, { align: 'center' });
                pdf.setTextColor(255, 255, 255); pdf.setFontSize(16);
                pdf.text(`Accuracy: ${accuracy}%`, pw / 2, 120, { align: 'center' });

                for (let i = 0; i < batch.length; i++) {
                    const item = batch[i]; gIdx++;
                    setP(Math.round(((bn * CONFIG.PDF_BATCH_SIZE + i + 1) / entries.length) * 85), `Image ${gIdx}/${entries.length}${bLabel}`);
                    pdf.addPage('a4', 'landscape');
                    let imgData = await ImageStore.getImageData(item.entryId);

                    // Top bar
                    pdf.setFillColor(35, 47, 62); pdf.rect(0, 0, pw, 10, 'F');
                    pdf.setTextColor(255, 255, 255); pdf.setFontSize(11); pdf.setFont('helvetica', 'bold');
                    pdf.text(`#${item.sNo} [${env.shortLabel}]`, m, 7);
                    if (item.responseCorrectOrIncorrect === 'Correct') pdf.setFillColor(40, 167, 69); else pdf.setFillColor(220, 53, 69);
                    const stxt = item.responseCorrectOrIncorrect || 'N/A';
                    pdf.roundedRect(m + 35, 2, pdf.getTextWidth(stxt) + 8, 7, 1.5, 1.5, 'F');
                    pdf.setTextColor(255, 255, 255); pdf.setFontSize(7); pdf.text(stxt, m + 39, 7);
                    pdf.setTextColor(255, 200, 100); pdf.setFontSize(8); pdf.setFont('helvetica', 'normal');
                    pdf.text(`${item.hva || ''} | ${item.responseTimeFormatted || ''}`, pw - m, 7, { align: 'right' });

                    // ID bar
                    pdf.setFillColor(44, 62, 80); pdf.rect(0, 10, pw, 7, 'F');
                    pdf.setFontSize(6); pdf.setTextColor(100, 180, 255); pdf.setFont('helvetica', 'bold');
                    pdf.text('Msg:', m, 14.5); pdf.setFont('helvetica', 'normal'); pdf.setTextColor(220, 230, 240);
                    pdf.text((item.messageId || 'N/A').substring(0, 55), m + 10, 14.5);
                    pdf.setFont('helvetica', 'bold'); pdf.setTextColor(100, 180, 255);
                    pdf.text('Conv:', pw / 2, 14.5); pdf.setFont('helvetica', 'normal'); pdf.setTextColor(220, 230, 240);
                    pdf.text((item.conversationId || 'N/A').substring(0, 55), pw / 2 + 12, 14.5);
                    pdf.setFillColor(...hexToRgb(env.color)); pdf.rect(0, 17, pw, 1.5, 'F');

                    // Message box
                    const my = 20, mh = 26;
                    pdf.setFillColor(248, 249, 250); pdf.rect(m, my, uw, mh, 'F');
                    pdf.setDrawColor(222, 226, 230); pdf.rect(m, my, uw, mh, 'S');
                    pdf.setFontSize(7); pdf.setFont('helvetica', 'bold'); pdf.setTextColor(21, 101, 192);
                    pdf.text('Query:', m + 3, my + 5); pdf.setFont('helvetica', 'normal'); pdf.setTextColor(33, 37, 41);
                    pdf.text(pdf.splitTextToSize((item.query || '').substring(0, 300), uw - 22).slice(0, 2), m + 18, my + 5);

                    const sy = my + 12;
                    pdf.setDrawColor(230, 230, 230); pdf.setLineWidth(0.15); pdf.line(m + 3, sy, m + uw - 3, sy);
                    pdf.setFont('helvetica', 'bold'); pdf.setTextColor(21, 128, 61);
                    pdf.text('Bot:', m + 3, sy + 4); pdf.setFont('helvetica', 'normal'); pdf.setTextColor(33, 37, 41);
                    pdf.text(pdf.splitTextToSize((item.preProdBotResponse || '').substring(0, 350), uw - 18).slice(0, 3), m + 14, sy + 4);

                    // Area of Improvement for regression
                    let extraOffset = 0;
                    if (env.hasAreaOfImprovement && item.areaOfImprovement) {
                        extraOffset = 8;
                        const aoiY = my + mh + 1;
                        pdf.setFillColor(248, 245, 255); pdf.rect(m, aoiY, uw, 7, 'F');
                        pdf.setDrawColor(111, 66, 193); pdf.rect(m, aoiY, uw, 7, 'S');
                        pdf.setFontSize(6); pdf.setFont('helvetica', 'bold'); pdf.setTextColor(111, 66, 193);
                        pdf.text('Area of Improvement:', m + 3, aoiY + 4); pdf.setFont('helvetica', 'normal'); pdf.setTextColor(80, 60, 120);
                        pdf.text(pdf.splitTextToSize((item.areaOfImprovement || '').substring(0, 200), uw - 45).slice(0, 1), m + 40, aoiY + 4);
                    }

                    // Image area
                    const iy = my + mh + 2 + extraOffset;
                    const ia = ph - iy - 10;

                    if (imgData) {
                        try {
                            const comp = await compressImageForPDF(imgData); imgData = null;
                            const ar = comp.width / comp.height;
                            let iw = uw, ih = iw / ar;
                            if (ih > ia) { ih = ia; iw = ih * ar; }
                            pdf.addImage(comp.dataUrl, 'JPEG', m + (uw - iw) / 2, iy, iw, ih);
                            totalKB += comp.compressedKB; success++; comp.dataUrl = null;
                        } catch (err) {
                            imgData = null;
                            pdf.setFillColor(255, 243, 205); pdf.rect(m, iy, uw, 30, 'F');
                            pdf.setTextColor(133, 100, 4); pdf.setFontSize(10);
                            pdf.text('Image error', pw / 2, iy + 15, { align: 'center' }); fail++;
                        }
                    } else {
                        pdf.setFillColor(248, 215, 218); pdf.rect(m, iy, uw, 30, 'F');
                        pdf.setTextColor(114, 28, 36); pdf.setFontSize(10);
                        pdf.text('Image missing', pw / 2, iy + 15, { align: 'center' }); fail++;
                    }

                    // Footer
                    pdf.setTextColor(150, 150, 150); pdf.setFontSize(7);
                    pdf.text(`${currentUsername} | ${env.shortLabel} | ${item.testingDate} | Pg ${i + 2}/${batch.length + 1} | #${item.sNo}`, pw / 2, ph - 5, { align: 'center' });
                    await new Promise(r => setTimeout(r, CONFIG.IMAGE_PROCESS_DELAY));
                }

                const suffix = totalBatches > 1 ? `_Part${bn + 1}of${totalBatches}` : '';
                pdf.save(`${env.id}_Images_${currentUsername}_${dateStr}${suffix}.pdf`);
                if (bn < totalBatches - 1) await new Promise(r => setTimeout(r, 2000));
            }

            if (showLoader) {
                setP(100, 'Done!');
                setTimeout(() => {
                    if (loader) loader.remove();
                    showNotification(`[${env.shortLabel}] ${success} images in ${totalBatches} PDF(s) (~${(totalKB / 1024).toFixed(1)}MB)${fail > 0 ? ` | ${fail} missing` : ''}`, fail > 0 ? 'warning' : 'success');
                }, 800);
            }
        } catch (e) {
            if (loader) loader.remove();
            SecureLog.error('PDF generation failed', e);
            if (showLoader) showNotification('PDF failed', 'error');
        }
    }

    // ════════════════════════════════════════════════════════
    //  ENVIRONMENT DASHBOARD
    // ════════════════════════════════════════════════════════
    function getEnvStats() {
        const stats = {};
        Object.keys(ENVIRONMENTS).forEach(eid => {
            try {
                let data = [];
                if (eid === activeEnvId) { data = chatData; }
                else {
                    const key = `chatScraper_${eid}_savedData`;
                    const raw = localStorage.getItem(key);
                    if (raw) { try { const parsed = JSON.parse(raw); if (!parsed._enc && Array.isArray(parsed)) data = parsed; } catch {} }
                }
                const correct = data.filter(d => d.responseCorrectOrIncorrect === 'Correct').length;
                const incorrect = data.filter(d => d.responseCorrectOrIncorrect === 'Incorrect').length;
                stats[eid] = {
                    total: data.length, correct, incorrect,
                    accuracy: data.length > 0 ? ((correct / data.length) * 100).toFixed(1) : '0.0',
                    spOk: data.filter(d => d._spUploaded === true).length,
                    spFail: data.filter(d => d._spUploaded === false).length,
                    images: data.filter(d => d._hasImage).length,
                    isActive: eid === activeEnvId
                };
            } catch {
                stats[eid] = { total: 0, correct: 0, incorrect: 0, accuracy: '0.0', spOk: 0, spFail: 0, images: 0, isActive: eid === activeEnvId };
            }
        });
        return stats;
    }

    function openEnvDashboard() {
        const stats = getEnvStats();
        const modal = document.createElement('div');
        modal.style.cssText = `position:fixed;inset:0;background:rgba(0,0,0,0.9);z-index:100000;display:flex;align-items:center;justify-content:center;font-family:Arial;`;

        const container = document.createElement('div');
        container.style.cssText = `background:#fff;width:800px;max-width:95vw;border-radius:12px;overflow:hidden`;

        const header = document.createElement('div');
        header.style.cssText = `background:linear-gradient(135deg,#232F3E,#37475A);color:#fff;padding:15px 20px;display:flex;justify-content:space-between;align-items:center`;
        const headerInfo = document.createElement('div');
        const h3 = document.createElement('h3'); h3.style.margin = '0'; h3.textContent = '📊 Environment Dashboard'; headerInfo.appendChild(h3);
        const subP = document.createElement('p'); subP.style.cssText = `margin:4px 0 0;font-size:11px;opacity:0.8`; subP.textContent = 'All environments | Shortcuts: Alt+1, Alt+2, Alt+3';
        headerInfo.appendChild(subP);
        const closeBtn = document.createElement('button');
        closeBtn.style.cssText = `background:rgba(255,255,255,0.2);border:none;color:#fff;width:30px;height:30px;border-radius:6px;cursor:pointer;font-size:18px`;
        closeBtn.textContent = '✕'; closeBtn.onclick = () => modal.remove();
        header.appendChild(headerInfo); header.appendChild(closeBtn); container.appendChild(header);

        const cardsRow = document.createElement('div');
        cardsRow.style.cssText = `padding:20px;display:flex;gap:15px;flex-wrap:wrap`;

        Object.entries(ENVIRONMENTS).forEach(([eid, env]) => {
            const s = stats[eid];
            const isActive = eid === activeEnvId;
            const allowed = isEnvAllowed(eid);

            const card = document.createElement('div');
            card.style.cssText = `flex:1;min-width:220px;background:${isActive ? '#fff' : '#f8f9fa'};border:3px solid ${isActive ? env.color : '#dee2e6'};border-radius:12px;overflow:hidden;${isActive ? 'box-shadow:0 4px 20px rgba(0,0,0,0.15);' : ''}`;

            const cardHeader = document.createElement('div');
            cardHeader.style.cssText = `background:${env.gradient};color:#fff;padding:12px 15px;text-align:center`;
            const cardTitle = document.createElement('div'); cardTitle.style.cssText = `font-size:16px;font-weight:bold`; cardTitle.textContent = env.label;
            cardHeader.appendChild(cardTitle);
            const cardSub = document.createElement('div'); cardSub.style.cssText = `font-size:10px;opacity:0.85;margin-top:2px`;
            cardSub.textContent = isActive ? '● ACTIVE' : `Alt+${Object.keys(ENVIRONMENTS).indexOf(eid) + 1}`;
            cardHeader.appendChild(cardSub); card.appendChild(cardHeader);

            const cardBody = document.createElement('div'); cardBody.style.cssText = `padding:15px;font-size:12px`;

            const createStatRow = (label, value, color = '') => {
                const row = document.createElement('div'); row.style.cssText = `display:flex;justify-content:space-between;margin-bottom:8px`;
                const lbl = document.createElement('span'); if (color) lbl.style.color = color; lbl.textContent = label;
                const val = document.createElement('strong'); if (color) val.style.color = color; val.textContent = value;
                row.appendChild(lbl); row.appendChild(val); return row;
            };

            cardBody.appendChild(createStatRow('Total Entries:', s.total));
            cardBody.appendChild(createStatRow('✓ Correct:', s.correct, '#155724'));
            cardBody.appendChild(createStatRow('✗ Incorrect:', s.incorrect, '#721c24'));
            cardBody.appendChild(createStatRow('Accuracy:', s.accuracy + '%'));
            cardBody.appendChild(createStatRow('🖼️ Images:', s.images));
            cardBody.appendChild(createStatRow('☁ SP Synced:', s.spOk, '#0078d4'));
            if (s.spFail > 0) cardBody.appendChild(createStatRow('⚠ SP Failed:', s.spFail, '#dc3545'));

            const listInfo = document.createElement('div');
            listInfo.style.cssText = `border-top:1px solid #dee2e6;padding-top:10px;margin-top:10px;text-align:center`;
            const listLabel = document.createElement('div'); listLabel.style.cssText = `font-size:9px;color:#888;word-break:break-all`;
            listLabel.textContent = `List: ${getListNameForEnv(eid)}`; listInfo.appendChild(listLabel);
            cardBody.appendChild(listInfo);

            if (!isActive && allowed) {
                const switchBtn = document.createElement('button');
                switchBtn.setAttribute('data-env', eid);
                switchBtn.style.cssText = `width:100%;margin-top:10px;padding:8px;background:${env.gradient};color:#fff;border:none;border-radius:6px;cursor:pointer;font-weight:bold;font-size:11px`;
                switchBtn.textContent = `Switch to ${env.shortLabel}`;
                switchBtn.onclick = () => { modal.remove(); switchEnvironment(eid); };
                cardBody.appendChild(switchBtn);
            } else if (!isActive && !allowed) {
                const blockedDiv = document.createElement('div');
                blockedDiv.style.cssText = `text-align:center;margin-top:10px;padding:8px;background:#f8d7da;color:#721c24;border-radius:6px;font-weight:bold;font-size:11px`;
                blockedDiv.textContent = '🔒 Blocked on this site'; cardBody.appendChild(blockedDiv);
            } else {
                const activeDiv = document.createElement('div');
                activeDiv.style.cssText = `text-align:center;margin-top:10px;padding:8px;background:${env.color}20;color:${env.color};border-radius:6px;font-weight:bold;font-size:11px`;
                activeDiv.textContent = 'Currently Active'; cardBody.appendChild(activeDiv);
            }

            card.appendChild(cardBody); cardsRow.appendChild(card);
        });

        container.appendChild(cardsRow);

        const footerDiv = document.createElement('div');
        footerDiv.style.cssText = `padding:10px 20px 15px;text-align:center;border-top:1px solid #dee2e6`;
        const secInfo = document.createElement('div'); secInfo.style.cssText = `font-size:9px;color:#666;margin-bottom:8px`;
        secInfo.textContent = '🔒 v8.1 Secured: AES-256 encrypted storage | No passwords stored or transmitted';
        footerDiv.appendChild(secInfo);
        const closeBottom = document.createElement('button');
        closeBottom.style.cssText = `padding:8px 30px;background:#6c757d;color:#fff;border:none;border-radius:6px;cursor:pointer;font-weight:bold`;
        closeBottom.textContent = 'Close'; closeBottom.onclick = () => modal.remove();
        footerDiv.appendChild(closeBottom); container.appendChild(footerDiv);

        modal.appendChild(container); document.body.appendChild(modal);
        modal.onclick = e => { if (e.target === modal) modal.remove(); };
        const escDash = e => { if (e.key === 'Escape') { modal.remove(); document.removeEventListener('keydown', escDash); } };
        document.addEventListener('keydown', escDash);
    }

    // ════════════════════════════════════════════════════════
    //  KEYBOARD SHORTCUTS
    // ════════════════════════════════════════════════════════
    document.addEventListener('keydown', e => {
        if (e.altKey && !e.ctrlKey && !e.shiftKey) {
            const envMap = { '1': 'pre-prod', '2': 'regression', '3': 'prod' };
            const target = envMap[e.key];
            if (target && target !== activeEnvId) {
                e.preventDefault();
                if (!isEnvAllowed(target)) {
                    showNotification(`${ENVIRONMENTS[target].label} is blocked on this site`, 'error');
                    return;
                }
                switchEnvironment(target);
            }
            if (e.key.toLowerCase() === 'd') {
                e.preventDefault();
                openEnvDashboard();
            }
        }
    });

    // ════════════════════════════════════════════════════════
    //  AUTO-DETECT ENVIRONMENT FROM URL
    // ════════════════════════════════════════════════════════
    function autoDetectEnvironment() {
        const url = window.location.href.toLowerCase();
        if (url.includes('pre-prod')) return 'pre-prod';
        if (url.includes('regression')) return 'regression';
        if (url.includes('www.amazon.com') || url.includes('prod')) return 'prod';
        return null;
    }

    // ════════════════════════════════════════════════════════
    //  CLEANUP ON TAB CLOSE
    // ════════════════════════════════════════════════════════
    function cleanupOnExit() {
        releaseUploadLock();
        saveQueueToStorage();
        saveDataToStorage();
        cachedDigest = null; cachedDigestTime = 0;
        CryptoStore.clearKeyCache();
    }

    // ════════════════════════════════════════════════════════
    //  CSP VIOLATION LISTENER
    // ════════════════════════════════════════════════════════
    document.addEventListener('securitypolicyviolation', (e) => {
        SecureLog.warn(`CSP violation: ${e.violatedDirective} - ${e.blockedURI}`);
    });

    // ════════════════════════════════════════════════════════
    //  STORAGE EVENT LISTENER (Cross-tab sync)
    // ════════════════════════════════════════════════════════
    window.addEventListener('storage', (e) => {
        if (e.key === GLOBAL_KEYS.activeEnv && e.newValue && e.newValue !== activeEnvId) {
            SecureLog.info(`Another tab switched to ${e.newValue}`);
            showNotification(`Another tab switched to ${ENVIRONMENTS[e.newValue]?.label || e.newValue}`, 'info');
        }
    });

    // ════════════════════════════════════════════════════════
    //  INIT
    // ════════════════════════════════════════════════════════
    window.addEventListener('load', async () => {
        try { await ImageStore.init(); } catch (e) { SecureLog.error('IndexedDB init failed', e); }

        loadSavedTestAccount();

        const savedEnv = localStorage.getItem(GLOBAL_KEYS.activeEnv);
        const detectedEnv = autoDetectEnvironment();

        if (detectedEnv && ENVIRONMENTS[detectedEnv]) {
            activeEnvId = detectedEnv;
        } else if (savedEnv && ENVIRONMENTS[savedEnv]) {
            activeEnvId = savedEnv;
        }
        localStorage.setItem(GLOBAL_KEYS.activeEnv, activeEnvId);

        await showUsernamePrompt();

        await loadQueueFromStorage();
        await loadDataFromStorage();

        checkDailyReset();
        scheduleDailyReset();
        createControlPanel();

        const dc = chatData.length;
        const qc = responseQueue.length;
        const fc = chatData.filter(d => d._spUploaded === false).length;
        const env = getActiveEnv();

        if (fc > 0 && getAutoUpload()) {
            setTimeout(async () => {
                showNotification(`[${env.shortLabel}] Retrying ${fc} failed uploads...`, 'info');
                await retryFailedUploads();
            }, 5000);
        }

        if (qc > 0) {
            showNotification(`Welcome ${currentUsername}! [${env.label}] ${qc} pending.`, 'success');
        } else if (dc > 0) {
            showNotification(`Welcome ${currentUsername}! [${env.label}] ${dc} entries.`, 'success');
        } else {
            showNotification(`Welcome ${currentUsername}! [${env.label}] Ready.`, 'success');
        }

        if (detectedEnv) {
            setTimeout(() => {
                showNotification(`Auto-detected: ${env.label} environment`, 'info');
            }, 2000);
        }

        const secNoticeKey = 'chatScraper_secNoticeShown_v81';
        if (!localStorage.getItem(secNoticeKey)) {
            setTimeout(() => {
                showNotification('🔒 v8.1 Secured: Encrypted storage, no passwords stored/sent', 'info');
                localStorage.setItem(secNoticeKey, 'true');
            }, 3000);
        }
    });

    window.addEventListener('beforeunload', () => {
        cleanupOnExit();
    });

    document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'hidden') {
            saveQueueToStorage();
            saveDataToStorage();
        }
    });

    setInterval(() => {
        if (chatData.length > 0 || responseQueue.length > 0) {
            saveQueueToStorage();
            saveDataToStorage();
        }
        const existing = localStorage.getItem(uploadLockKey);
        if (existing) {
            try {
                const lock = JSON.parse(existing);
                if (lock.tabId === TAB_ID) {
                    lock.time = Date.now();
                    localStorage.setItem(uploadLockKey, JSON.stringify(lock));
                }
            } catch {}
        }
    }, 30000);

    try {
        const existing = localStorage.getItem(uploadLockKey);
        if (existing) {
            const lock = JSON.parse(existing);
            if (Date.now() - lock.time > 120000) {
                localStorage.removeItem(uploadLockKey);
                SecureLog.info('Cleared stale upload lock');
            }
        }
    } catch {
        localStorage.removeItem(uploadLockKey);
    }

})();
