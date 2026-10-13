'use strict';
/* =====================================================================
   CloudNest — Google Drive + OneDrive manager
   100% client-side: you sign in with the official Google / Microsoft
   OAuth popups and the files travel directly between this page and
   Google / Microsoft. There is no CloudNest server or database.
   ===================================================================== */

/* When MSAL opens its sign-in popup it redirects back to this page and reads
   the URL hash from the main window. In that popup we must not start the app. */
const IS_AUTH_POPUP = !!(window.opener && window.opener !== window && /[#?&](code|error)=/.test(location.hash + location.search));

/* ---------------------------------------------------------------- helpers */
const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const sleep = ms => new Promise(r => setTimeout(r, ms));
const el = (tag, cls, html) => { const e = document.createElement(tag); if (cls) e.className = cls; if (html != null) e.innerHTML = html; return e; };
function readStore(store, key, fallback) { try { const v = store.getItem(key); return v ? JSON.parse(v) : fallback; } catch (e) { return fallback; } }
function writeStore(store, key, val) { try { store.setItem(key, JSON.stringify(val)); } catch (e) { /* storage full / blocked */ } }
function formatBytes(bytes, d = 1) {
    if (!bytes) return '0 B';
    const k = 1024, u = ['B', 'KB', 'MB', 'GB', 'TB', 'PB'];
    const i = Math.min(u.length - 1, Math.floor(Math.log(bytes) / Math.log(k)));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(d)) + ' ' + u[i];
}

/* ----------------------------------------------------------------- config */
const BASE_CFG = window.CLOUDNEST_CONFIG || {};
const SAVED_CFG = readStore(localStorage, 'cloudnest.config', {});
const CFG = {
    googleClientId: String(SAVED_CFG.googleClientId || BASE_CFG.GOOGLE_CLIENT_ID || '').trim(),
    msClientId: String(SAVED_CFG.msClientId || BASE_CFG.MICROSOFT_CLIENT_ID || '').trim(),
    googleScope: BASE_CFG.GOOGLE_SCOPES || 'https://www.googleapis.com/auth/drive',
    msScopes: BASE_CFG.MICROSOFT_SCOPES || ['User.Read', 'Files.ReadWrite.All'],
    msAuthority: BASE_CFG.MICROSOFT_AUTHORITY || 'https://login.microsoftonline.com/common'
};
const APP_ORIGIN = location.origin;
const APP_URL = location.origin + location.pathname.replace(/index\.html$/, '');

const GAPI = 'https://www.googleapis.com/drive/v3';
const GUPLOAD = 'https://www.googleapis.com/upload/drive/v3';
const GRAPH = 'https://graph.microsoft.com/v1.0';

const LIB = {
    pdf: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
    pdfWorker: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js',
    mammoth: 'https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js',
    xlsx: 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
    jszip: 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js'
};

/* ------------------------------------------------------- file kinds/icons */
const EXT = {
    image: ['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp', 'svg', 'avif', 'ico'],
    pdf: ['pdf'],
    doc: ['doc', 'docx', 'odt', 'rtf'],
    sheet: ['xls', 'xlsx', 'xlsm', 'ods', 'csv'],
    ppt: ['ppt', 'pptx', 'odp'],
    video: ['mp4', 'mov', 'mkv', 'webm', 'avi', 'm4v', '3gp'],
    audio: ['mp3', 'wav', 'm4a', 'aac', 'ogg', 'flac', 'opus'],
    archive: ['zip', 'rar', '7z', 'tar', 'gz', 'tgz'],
    text: ['txt', 'md', 'json', 'xml', 'html', 'htm', 'css', 'js', 'ts', 'jsx', 'tsx', 'log', 'yml', 'yaml', 'py', 'java', 'c', 'cpp', 'h', 'sh', 'ini', 'conf', 'sql', 'toml']
};
function extOf(name) { return name.includes('.') ? name.split('.').pop().toLowerCase() : ''; }
function kindOf(name, mime, isFolder) {
    if (isFolder) return 'folder';
    mime = mime || '';
    if (mime === 'application/vnd.google-apps.document') return 'doc';
    if (mime === 'application/vnd.google-apps.spreadsheet') return 'sheet';
    if (mime === 'application/vnd.google-apps.presentation') return 'ppt';
    const ext = extOf(name);
    for (const k in EXT) if (EXT[k].includes(ext)) return k;
    if (mime.startsWith('image/')) return 'image';
    if (mime.startsWith('video/')) return 'video';
    if (mime.startsWith('audio/')) return 'audio';
    if (mime === 'application/pdf') return 'pdf';
    if (mime.startsWith('text/')) return 'text';
    return 'file';
}
const KIND_ICON = {
    folder: 'fa-folder', pdf: 'fa-file-pdf', image: 'fa-file-image', doc: 'fa-file-word', sheet: 'fa-file-excel', ppt: 'fa-file-powerpoint',
    video: 'fa-file-video', audio: 'fa-file-audio', archive: 'fa-file-zipper', text: 'fa-file-lines', file: 'fa-file'
};
const KIND_STYLE = {
    folder: 'bg-amber-100 dark:bg-amber-950/40 text-amber-600',
    pdf: 'bg-red-100 dark:bg-red-950/40 text-red-600',
    image: 'bg-emerald-100 dark:bg-emerald-950/40 text-emerald-600',
    doc: 'bg-blue-100 dark:bg-blue-950/40 text-blue-600',
    sheet: 'bg-green-100 dark:bg-green-950/40 text-green-600',
    ppt: 'bg-orange-100 dark:bg-orange-950/40 text-orange-600',
    video: 'bg-purple-100 dark:bg-purple-950/40 text-purple-600',
    audio: 'bg-pink-100 dark:bg-pink-950/40 text-pink-600',
    archive: 'bg-yellow-100 dark:bg-yellow-950/40 text-yellow-700',
    text: 'bg-slate-200 dark:bg-slate-800 text-slate-600',
    file: 'bg-slate-100 dark:bg-slate-800 text-slate-600'
};
const kindIcon = k => KIND_ICON[k] || KIND_ICON.file;
const kindStyle = k => KIND_STYLE[k] || KIND_STYLE.file;

/* ------------------------------------------------------------------ state */
const state = {
    accounts: (readStore(localStorage, 'cloudnest.accounts.v1', []) || []).filter(a => a && a.id && a.provider),
    view: 'accounts',
    activeId: null,
    path: [],
    items: [],
    next: null,
    loading: false,
    error: null,
    needReconnect: false,
    viewMode: readStore(localStorage, 'cloudnest.viewMode', 'grid') === 'list' ? 'list' : 'grid',
    uploads: [],
    uploading: false,
    trayCollapsed: false,
    preview: null
};
const getAcc = id => state.accounts.find(a => a.id === id);
const prov = acc => PROVIDERS[acc.provider];
const currentFolder = () => state.path[state.path.length - 1];
function saveAccounts() { writeStore(localStorage, 'cloudnest.accounts.v1', state.accounts); }

/* ---------------------------------------------------------------- toasts */
function toast(msg, type = 'info') {
    const c = $('toast-container');
    if (!c) return;
    const t = el('div', 'pointer-events-auto max-w-sm flex items-start space-x-2 px-4 py-3 rounded-xl shadow-lg text-xs font-semibold transition-all duration-300 transform translate-y-2 opacity-0 ' +
        (type === 'error' ? 'bg-red-600 text-white' : 'bg-slate-900 dark:bg-slate-100 text-white dark:text-slate-900'),
        `<i class="fa-solid ${type === 'error' ? 'fa-circle-exclamation' : 'fa-circle-check'} mt-0.5"></i><span></span>`);
    t.querySelector('span').textContent = msg;
    c.appendChild(t);
    setTimeout(() => t.classList.remove('translate-y-2', 'opacity-0'), 10);
    setTimeout(() => { t.classList.add('translate-y-2', 'opacity-0'); setTimeout(() => t.remove(), 300); }, type === 'error' ? 6000 : 3000);
}

/* ---------------------------------------------------------------- modals */
function openModal(id) { $(id).classList.remove('hidden'); }
function closeModal(id) { $(id).classList.add('hidden'); }
const anyModalOpen = () => !!document.querySelector('[data-modal]:not(.hidden)');

let confirmResolve = null;
function confirmDialog({ title, text, okLabel = 'Delete' }) {
    return new Promise(resolve => {
        if (confirmResolve) confirmResolve(false);
        confirmResolve = resolve;
        $('confirm-title').textContent = title;
        $('confirm-text').textContent = text;
        $('confirm-yes').textContent = okLabel;
        openModal('confirm-modal');
    });
}
function settleConfirm(v) { closeModal('confirm-modal'); if (confirmResolve) { const r = confirmResolve; confirmResolve = null; r(v); } }

/* ---------------------------------------------------------- token handling */
const tokens = readStore(sessionStorage, 'cloudnest.tokens', {});
const pendingTok = {};
class NeedReconnect extends Error { constructor(msg) { super(msg || 'Sign-in expired. Reconnect this account to continue.'); this.needReconnect = true; } }
function setToken(id, token, exp) { tokens[id] = { token, exp }; writeStore(sessionStorage, 'cloudnest.tokens', tokens); }
function dropToken(id) { delete tokens[id]; writeStore(sessionStorage, 'cloudnest.tokens', tokens); }
const tokenValid = acc => !!tokens[acc.id] && tokens[acc.id].exp > Date.now();

function getToken(acc, { interactive = false } = {}) {
    if (tokenValid(acc)) return Promise.resolve(tokens[acc.id].token);
    if (pendingTok[acc.id]) return pendingTok[acc.id];
    // refreshToken() starts synchronously, so the sign-in popup stays tied to the user's tap
    const p = (async () => prov(acc).refreshToken(acc, interactive))().finally(() => { delete pendingTok[acc.id]; });
    pendingTok[acc.id] = p;
    return p;
}

async function apiError(res) {
    let msg = `${res.status} ${res.statusText}`.trim();
    try { const j = await res.clone().json(); msg = (j.error && (j.error.message || j.error)) || j.error_description || msg; } catch (e) { /* not JSON */ }
    const e = new Error(typeof msg === 'string' ? msg : JSON.stringify(msg));
    e.status = res.status;
    return e;
}

async function authFetch(acc, url, opts = {}, retry = true) {
    const token = await getToken(acc);
    const res = await fetch(url, { ...opts, headers: { ...(opts.headers || {}), Authorization: 'Bearer ' + token } });
    if (res.status === 401 && retry) { dropToken(acc.id); return authFetch(acc, url, opts, false); }
    if ((res.status === 429 || res.status === 503) && retry) {
        await sleep(Math.min(10, parseInt(res.headers.get('Retry-After'), 10) || 2) * 1000);
        return authFetch(acc, url, opts, false);
    }
    if (!res.ok) throw await apiError(res);
    return res;
}

/* XHR is used for uploads because fetch() cannot report upload progress */
function xhrSend(method, url, body, headers, onProgress, ctl) {
    return new Promise((resolve, reject) => {
        const x = new XMLHttpRequest();
        x.open(method, url);
        for (const k in headers || {}) x.setRequestHeader(k, headers[k]);
        x.upload.onprogress = e => { if (e.lengthComputable && onProgress) onProgress(e.loaded); };
        x.onload = () => resolve({ status: x.status, text: x.responseText });
        x.onerror = () => reject(new Error('Network error during upload. Check your connection and retry.'));
        x.onabort = () => reject(Object.assign(new Error('Cancelled'), { cancelled: true }));
        if (ctl) ctl.abort = () => x.abort();
        x.send(body);
    });
}
function xhrError(r) {
    try { const j = JSON.parse(r.text); return (j.error && (j.error.message || j.error)) || `Upload failed (${r.status})`; } catch (e) { return `Upload failed (${r.status})`; }
}

/* -------------------------------------------------------- Google provider */
function gisReady() {
    if (window.google && google.accounts && google.accounts.oauth2) return Promise.resolve();
    return new Promise((resolve, reject) => {
        let n = 0;
        const t = setInterval(() => {
            if (window.google && google.accounts && google.accounts.oauth2) { clearInterval(t); resolve(); }
            else if (++n > 150) { clearInterval(t); reject(new Error('Google sign-in script could not load. Check your internet or ad-blocker.')); }
        }, 100);
    });
}
function authErr(code, msg) {
    const map = {
        popup_closed: 'Sign-in window was closed before finishing.',
        popup_failed_to_open: 'Your browser blocked the sign-in popup. Allow popups for this site and try again.',
        access_denied: 'Access was denied.'
    };
    const e = new Error(map[code] || msg || String(code));
    e.code = code;
    return e;
}
function gRequestToken({ hint, prompt }) {
    if (!CFG.googleClientId) return Promise.reject(new Error('Add your Google Client ID in Settings first.'));
    return gisReady().then(() => new Promise((resolve, reject) => {
        const client = google.accounts.oauth2.initTokenClient({
            client_id: CFG.googleClientId,
            scope: CFG.googleScope,
            callback: r => {
                if (r.error) reject(authErr(r.error, r.error_description));
                else if (!google.accounts.oauth2.hasGrantedAllScopes(r, CFG.googleScope)) reject(new Error('Google Drive permission was not granted. Tick the Drive access box on the consent screen.'));
                else resolve(r);
            },
            error_callback: e => reject(authErr(e.type, e.message))
        });
        const o = { prompt };
        if (hint) o.hint = hint;
        client.requestAccessToken(o);
    }));
}
function gNorm(f) {
    const isFolder = f.mimeType === 'application/vnd.google-apps.folder';
    return {
        id: f.id, name: f.name, mime: f.mimeType || '', isFolder,
        type: kindOf(f.name, f.mimeType, isFolder),
        size: +f.size || 0,
        modified: (f.modifiedTime || '').slice(0, 10),
        webUrl: f.webViewLink || '',
        native: !isFolder && (f.mimeType || '').startsWith('application/vnd.google-apps.')
    };
}
const G_EXPORT = {
    doc: ['application/vnd.openxmlformats-officedocument.wordprocessingml.document', '.docx'],
    sheet: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.xlsx'],
    ppt: ['application/vnd.openxmlformats-officedocument.presentationml.presentation', '.pptx']
};

const GoogleP = {
    id: 'google', name: 'Google Drive', icon: 'fa-brands fa-google-drive', rootName: 'My Drive', trashName: 'Google Drive trash',
    badge: 'bg-blue-50 dark:bg-blue-950/40 text-blue-600',
    configured: () => !!CFG.googleClientId,

    async refreshToken(acc, interactive) {
        if (!interactive) throw new NeedReconnect();
        const r = await gRequestToken({ hint: acc.email, prompt: '' });
        setToken(acc.id, r.access_token, Date.now() + (Number(r.expires_in) - 60) * 1000);
        return r.access_token;
    },
    async about(token) {
        const res = await fetch(`${GAPI}/about?fields=user(displayName,emailAddress,photoLink),storageQuota(limit,usage)`, { headers: { Authorization: 'Bearer ' + token } });
        if (!res.ok) throw await apiError(res);
        return res.json();
    },
    async connect() {
        const r = await gRequestToken({ prompt: 'select_account' });
        const about = await this.about(r.access_token);
        const email = String(about.user.emailAddress || '').toLowerCase();
        const acc = {
            id: 'g:' + email, provider: 'google', email,
            name: about.user.displayName || email,
            avatar: about.user.photoLink || '',
            used: Number(about.storageQuota.usage) || 0,
            limit: Number(about.storageQuota.limit) || 0
        };
        setToken(acc.id, r.access_token, Date.now() + (Number(r.expires_in) - 60) * 1000);
        return acc;
    },
    disconnect(acc) {
        const t = tokens[acc.id];
        try { if (t && window.google && google.accounts.oauth2.revoke) google.accounts.oauth2.revoke(t.token, () => { }); } catch (e) { /* ignore */ }
        dropToken(acc.id);
    },
    async quota(acc) {
        const res = await authFetch(acc, `${GAPI}/about?fields=storageQuota(limit,usage)`);
        const j = await res.json();
        return { used: Number(j.storageQuota.usage) || 0, limit: Number(j.storageQuota.limit) || 0 };
    },
    async list(acc, folderId, pageToken) {
        const params = new URLSearchParams({
            q: `'${String(folderId).replace(/'/g, "\\'")}' in parents and trashed=false`,
            pageSize: '100',
            fields: 'nextPageToken,files(id,name,mimeType,size,modifiedTime,webViewLink)',
            orderBy: 'folder,name'
        });
        if (pageToken) params.set('pageToken', pageToken);
        const j = await (await authFetch(acc, `${GAPI}/files?${params}`)).json();
        return { items: (j.files || []).map(gNorm), next: j.nextPageToken || null };
    },
    async createFolder(acc, parentId, name) {
        const res = await authFetch(acc, `${GAPI}/files?fields=id`, {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name, mimeType: 'application/vnd.google-apps.folder', parents: [parentId] })
        });
        return res.json();
    },
    async remove(acc, item) {
        await authFetch(acc, `${GAPI}/files/${encodeURIComponent(item.id)}`, {
            method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ trashed: true })
        });
    },
    async upload(acc, parentId, file, onProgress, ctl) {
        const mime = file.type || 'application/octet-stream';
        if (file.size === 0) {   // resumable sessions cannot carry an empty body
            await (await authFetch(acc, `${GAPI}/files?fields=id`, {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ name: file.name, parents: [parentId], mimeType: mime })
            })).json();
            onProgress(1);
            return;
        }
        const init = await authFetch(acc, `${GUPLOAD}/files?uploadType=resumable&fields=id,name`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json; charset=UTF-8', 'X-Upload-Content-Type': mime, 'X-Upload-Content-Length': String(file.size) },
            body: JSON.stringify({ name: file.name, parents: [parentId] })
        });
        const loc = init.headers.get('Location');
        if (!loc) throw new Error('Google did not return an upload address. Try again.');
        const r = await xhrSend('PUT', loc, file, { 'Content-Type': mime }, loaded => onProgress(loaded / file.size), ctl);
        if (r.status < 200 || r.status >= 300) throw new Error(xhrError(r));
    },
    async blob(acc, item, { download = false } = {}) {
        let url;
        if (item.native) {
            const mt = download && G_EXPORT[item.type] ? G_EXPORT[item.type][0] : 'application/pdf';
            url = `${GAPI}/files/${encodeURIComponent(item.id)}/export?mimeType=${encodeURIComponent(mt)}`;
        } else {
            url = `${GAPI}/files/${encodeURIComponent(item.id)}?alt=media`;
        }
        return (await authFetch(acc, url)).blob();
    },
    async directUrl() { return null; },
    async download(acc, item) {
        if (item.size > 1.5e9) throw new Error('This file is too large to download inside the browser. Use "Open in Google Drive".');
        const blob = await this.blob(acc, item, { download: true });
        const suffix = item.native ? (G_EXPORT[item.type] ? G_EXPORT[item.type][1] : '.pdf') : '';
        saveBlob(blob, item.name + suffix);
    }
};

/* ------------------------------------------------------ Microsoft provider */
let msalApp = null, msalReady = null;
function initMsal() {
    if (msalApp || !CFG.msClientId || !window.msal) return;
    msalApp = new msal.PublicClientApplication({
        auth: { clientId: CFG.msClientId, authority: CFG.msAuthority, redirectUri: APP_URL },
        cache: { cacheLocation: 'localStorage' }
    });
    msalReady = Promise.resolve(typeof msalApp.initialize === 'function' ? msalApp.initialize() : null);
}
function getMsal() {
    initMsal();
    if (!msalApp) return Promise.reject(new Error(window.msal ? 'Add your Microsoft Client ID in Settings first.' : 'Microsoft sign-in script could not load. Check your internet or ad-blocker.'));
    return msalReady.then(() => msalApp);
}
function msErr(e) {
    const code = e && (e.errorCode || e.code);
    const map = {
        user_cancelled: 'Sign-in window was closed before finishing.',
        popup_window_error: 'Your browser blocked the sign-in popup. Allow popups for this site and try again.',
        empty_window_error: 'Your browser blocked the sign-in popup. Allow popups for this site and try again.',
        interaction_in_progress: 'Another Microsoft sign-in is still open. Finish or close it, then try again.'
    };
    const err = new Error(map[code] || (e && (e.errorMessage || e.message)) || 'Microsoft sign-in failed.');
    err.code = code;
    return err;
}
const msEnc = encodeURIComponent;
const msItemPath = id => id === 'root' ? '/me/drive/root' : '/me/drive/items/' + msEnc(id);
const graph = (acc, path, opts) => authFetch(acc, path.startsWith('http') ? path : GRAPH + path, opts);

function mNorm(f) {
    const isFolder = !!f.folder;
    const mime = f.file && f.file.mimeType || '';
    return {
        id: f.id, name: f.name, mime, isFolder,
        type: kindOf(f.name, mime, isFolder),
        size: +f.size || 0,
        modified: (f.lastModifiedDateTime || '').slice(0, 10),
        webUrl: f.webUrl || '', native: false
    };
}
const MS_SMALL_LIMIT = 4 * 1024 * 1024;          // simple PUT works up to 4 MB
const MS_CHUNK = 320 * 1024 * 16;                // 5 MiB, must be a multiple of 320 KiB

const MicrosoftP = {
    id: 'microsoft', name: 'OneDrive', icon: 'fa-brands fa-microsoft', rootName: 'OneDrive', trashName: 'OneDrive recycle bin',
    badge: 'bg-sky-50 dark:bg-sky-950/40 text-sky-700',
    configured: () => !!CFG.msClientId,

    msAccount(acc) { return msalApp && msalApp.getAllAccounts().find(a => a.homeAccountId === acc.msId); },
    async refreshToken(acc, interactive) {
        const app = await getMsal();
        const account = this.msAccount(acc);
        if (!account) throw new NeedReconnect();
        try {
            const r = await app.acquireTokenSilent({ scopes: CFG.msScopes, account });
            setToken(acc.id, r.accessToken, r.expiresOn.getTime() - 60000);
            return r.accessToken;
        } catch (e) {
            if (!interactive) throw new NeedReconnect();
            try {
                const r = await app.acquireTokenPopup({ scopes: CFG.msScopes, account });
                setToken(acc.id, r.accessToken, r.expiresOn.getTime() - 60000);
                return r.accessToken;
            } catch (e2) { throw msErr(e2); }
        }
    },
    async connect() {
        const app = await getMsal();
        let r;
        try { r = await app.loginPopup({ scopes: CFG.msScopes, prompt: 'select_account' }); } catch (e) { throw msErr(e); }
        const get = async p => { const res = await fetch(GRAPH + p, { headers: { Authorization: 'Bearer ' + r.accessToken } }); if (!res.ok) throw await apiError(res); return res; };
        const me = await (await get('/me?$select=displayName,mail,userPrincipalName')).json();
        const drive = await (await get('/me/drive?$select=id,quota')).json();
        const email = String(me.mail || me.userPrincipalName || r.account.username || '').toLowerCase();
        let avatar = '';
        try {
            const pr = await fetch(GRAPH + '/me/photo/$value', { headers: { Authorization: 'Bearer ' + r.accessToken } });
            if (pr.ok) {
                const b = await pr.blob();
                if (b.size < 60000) avatar = await new Promise(res => { const fr = new FileReader(); fr.onload = () => res(fr.result); fr.onerror = () => res(''); fr.readAsDataURL(b); });
            }
        } catch (e) { /* no profile photo */ }
        const acc = {
            id: 'm:' + r.account.homeAccountId, msId: r.account.homeAccountId, provider: 'microsoft', email,
            name: me.displayName || email, avatar,
            used: Number(drive.quota && drive.quota.used) || 0, limit: Number(drive.quota && drive.quota.total) || 0
        };
        setToken(acc.id, r.accessToken, r.expiresOn.getTime() - 60000);
        return acc;
    },
    disconnect(acc) {
        dropToken(acc.id);
        try {   // clear this account's cached MSAL tokens
            const needle = String(acc.msId || '').toLowerCase();
            if (needle) Object.keys(localStorage).forEach(k => { if (k.toLowerCase().includes(needle)) localStorage.removeItem(k); });
        } catch (e) { /* ignore */ }
    },
    async quota(acc) {
        const j = await (await graph(acc, '/me/drive?$select=quota')).json();
        return { used: Number(j.quota && j.quota.used) || 0, limit: Number(j.quota && j.quota.total) || 0 };
    },
    async list(acc, folderId, pageToken) {
        const url = pageToken || `${msItemPath(folderId)}/children?$top=100`;
        const j = await (await graph(acc, url)).json();
        return { items: (j.value || []).map(mNorm), next: j['@odata.nextLink'] || null };
    },
    async createFolder(acc, parentId, name) {
        const res = await graph(acc, `${msItemPath(parentId)}/children`, {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' })
        });
        return res.json();
    },
    async remove(acc, item) { await graph(acc, `/me/drive/items/${msEnc(item.id)}`, { method: 'DELETE' }); },
    async upload(acc, parentId, file, onProgress, ctl) {
        const base = `${msItemPath(parentId)}:/${msEnc(file.name)}:`;
        if (file.size < MS_SMALL_LIMIT) {
            const token = await getToken(acc);
            const r = await xhrSend('PUT', `${GRAPH}${base}/content?@microsoft.graph.conflictBehavior=rename`, file,
                { Authorization: 'Bearer ' + token, 'Content-Type': file.type || 'application/octet-stream' },
                loaded => onProgress(file.size ? loaded / file.size : 1), ctl);
            if (r.status < 200 || r.status >= 300) throw new Error(xhrError(r));
            onProgress(1);
            return;
        }
        const sess = await (await graph(acc, `${base}/createUploadSession`, {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'rename', name: file.name } })
        })).json();
        const url = sess.uploadUrl;
        if (!url) throw new Error('OneDrive did not return an upload address. Try again.');
        try {
            for (let start = 0; start < file.size; start += MS_CHUNK) {
                if (ctl && ctl.cancelled) throw Object.assign(new Error('Cancelled'), { cancelled: true });
                const end = Math.min(start + MS_CHUNK, file.size) - 1;
                let attempt = 0, r;
                for (;;) {
                    try {
                        r = await xhrSend('PUT', url, file.slice(start, end + 1),
                            { 'Content-Range': `bytes ${start}-${end}/${file.size}` },
                            loaded => onProgress((start + loaded) / file.size), ctl);
                    } catch (e) { if (e.cancelled || ++attempt >= 3) throw e; await sleep(1500 * attempt); continue; }
                    if ((r.status >= 500 || r.status === 429) && ++attempt < 3) { await sleep(1500 * attempt); continue; }
                    break;
                }
                if (r.status < 200 || r.status >= 300) throw new Error(xhrError(r));
            }
        } catch (e) {
            if (e.cancelled) { try { await fetch(url, { method: 'DELETE' }); } catch (x) { /* best effort */ } }
            throw e;
        }
    },
    async directUrl(acc, item) {
        const j = await (await graph(acc, `/me/drive/items/${msEnc(item.id)}`)).json();
        return j['@microsoft.graph.downloadUrl'] || null;
    },
    async blob(acc, item) { return (await graph(acc, `/me/drive/items/${msEnc(item.id)}/content`)).blob(); },
    async download(acc, item) {
        const url = await this.directUrl(acc, item);
        if (!url) throw new Error('OneDrive did not return a download link.');
        const a = el('a'); a.href = url; a.download = item.name; a.rel = 'noopener';
        document.body.appendChild(a); a.click(); a.remove();
    }
};
const PROVIDERS = { google: GoogleP, microsoft: MicrosoftP };

function saveBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = el('a'); a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
}

/* ----------------------------------------------------------- rendering: shell */
function avatarHtml(acc, cls) {
    const initial = esc((acc.name || acc.email || '?').trim().charAt(0).toUpperCase());
    const img = acc.avatar ? `<img src="${esc(acc.avatar)}" referrerpolicy="no-referrer" alt="" class="absolute inset-0 w-full h-full object-cover" onerror="this.remove()">` : '';
    return `<div class="relative shrink-0 overflow-hidden bg-cloudprimary/10 text-cloudprimary font-bold flex items-center justify-center ${cls}">${initial}${img}</div>`;
}

function renderUI() {
    $('view-accounts').classList.toggle('hidden', state.view !== 'accounts');
    $('view-files').classList.toggle('hidden', state.view !== 'files');
    if (state.view === 'accounts') renderAccounts(); else renderFiles();
}

function renderAccounts() {
    const has = state.accounts.length > 0;
    $('accounts-empty').classList.toggle('hidden', has);
    $('storage-cards-grid').classList.toggle('hidden', !has);
    $('storage-summary').classList.toggle('hidden', !has);
    $('refresh-quota-btn').classList.toggle('hidden', !has);
    if (!has) return;

    const used = state.accounts.reduce((s, a) => s + (a.used || 0), 0);
    const allLimited = state.accounts.every(a => a.limit > 0);
    const total = state.accounts.reduce((s, a) => s + (a.limit || 0), 0);
    $('storage-summary').textContent = `${formatBytes(used)} used across ${state.accounts.length} account${state.accounts.length > 1 ? 's' : ''}` + (allLimited ? ` of ${formatBytes(total)} total` : '');

    $('storage-cards-grid').innerHTML = state.accounts.map(acc => {
        const P = prov(acc);
        const pct = acc.limit ? Math.min(100, Math.round((acc.used / acc.limit) * 100)) : 0;
        const bar = pct > 90 ? 'bg-red-500' : pct > 75 ? 'bg-amber-500' : 'bg-cloudprimary';
        const needsSignIn = acc.provider === 'google' && !tokenValid(acc);
        return `
        <div data-act="select-account" data-id="${esc(acc.id)}" role="button" tabindex="0" class="cute-card bg-cloudcard-light dark:bg-cloudcard-dark p-6 rounded-3xl border border-cloudborder-light dark:border-cloudborder-dark shadow-sm hover:shadow-md cursor-pointer space-y-4">
            <div class="flex items-start justify-between gap-2">
                <div class="flex items-center space-x-3 overflow-hidden">
                    ${avatarHtml(acc, 'w-10 h-10 rounded-2xl text-sm')}
                    <div class="overflow-hidden">
                        <h3 class="font-bold text-sm text-slate-900 dark:text-white truncate">${esc(acc.name)}</h3>
                        <p class="text-xs text-slate-500 dark:text-slate-400 truncate">${esc(acc.email)}</p>
                    </div>
                </div>
                <button data-act="disconnect" data-id="${esc(acc.id)}" title="Remove account" aria-label="Remove account" class="text-slate-400 hover:text-red-500 text-xs p-1.5 shrink-0"><i class="fa-solid fa-link-slash"></i></button>
            </div>
            <div class="flex items-center gap-2">
                <span class="inline-flex items-center gap-1.5 text-[11px] font-semibold px-2.5 py-1 rounded-full ${P.badge}"><i class="${P.icon}"></i>${esc(P.name)}</span>
                ${needsSignIn ? '<span class="text-[11px] font-semibold px-2 py-1 rounded-full bg-amber-50 dark:bg-amber-950/40 text-amber-700">Sign in again</span>' : ''}
            </div>
            <div class="space-y-1.5">
                <div class="flex justify-between text-xs font-medium">
                    <span class="text-slate-500">Storage used</span>
                    <span class="text-slate-900 dark:text-white font-semibold">${formatBytes(acc.used)} / ${acc.limit ? formatBytes(acc.limit) : 'Unlimited'}</span>
                </div>
                <div class="w-full bg-slate-100 dark:bg-slate-800 rounded-full h-2 overflow-hidden"><div class="${bar} h-full rounded-full transition-all duration-500" style="width:${pct}%"></div></div>
            </div>
            <div class="flex items-center justify-between pt-2 border-t border-slate-100 dark:border-slate-800 text-xs text-slate-500">
                <span>${acc.limit ? pct + '% full' : 'No storage cap'}</span>
                <span class="text-cloudprimary font-semibold flex items-center space-x-1"><span>Open files</span><i class="fa-solid fa-chevron-right text-[10px]"></i></span>
            </div>
        </div>`;
    }).join('');
}

function renderModeButtons() {
    const on = 'px-3 py-1.5 rounded-xl text-xs font-semibold transition bg-white dark:bg-slate-700 shadow-sm text-cloudprimary';
    const off = 'px-3 py-1.5 rounded-xl text-xs font-semibold transition text-slate-500 hover:text-slate-900 dark:hover:text-white';
    $('view-mode-grid').className = state.viewMode === 'grid' ? on : off;
    $('view-mode-list').className = state.viewMode === 'list' ? on : off;
}

function renderFiles() {
    const acc = getAcc(state.activeId);
    if (!acc) return;
    const P = prov(acc);
    $('active-account-badge').innerHTML = `${avatarHtml(acc, 'w-6 h-6 rounded-full text-[10px]')}<span class="font-bold text-xs text-slate-900 dark:text-white truncate">${esc(acc.name)}</span><i class="${P.icon} text-[11px] text-slate-400"></i>`;

    $('breadcrumb-container').innerHTML = state.path.map((p, i) => {
        const last = i === state.path.length - 1;
        return `<button data-act="crumb" data-i="${i}" class="font-semibold text-xs transition ${last ? 'text-cloudprimary' : 'text-slate-500 hover:text-slate-900 dark:hover:text-white'}">${esc(p.name)}</button>` +
            (last ? '' : '<i class="fa-solid fa-chevron-right text-[10px] text-slate-400"></i>');
    }).join('');
    renderModeButtons();

    // notice (sign-in needed / error)
    const notice = $('files-notice');
    if (state.error || state.needReconnect) {
        notice.classList.remove('hidden');
        notice.innerHTML = `<div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3 p-4 rounded-2xl bg-amber-50 dark:bg-amber-950/40 border border-amber-200 dark:border-amber-900 text-xs text-amber-800 dark:text-amber-300">
            <span><i class="fa-solid fa-circle-exclamation mr-2"></i>${esc(state.error || 'Sign-in expired. Reconnect this account to continue.')}</span>
            <button data-act="reconnect" class="shrink-0 px-3.5 py-2 bg-cloudprimary hover:bg-cloudprimaryhover text-white font-semibold rounded-xl">${state.needReconnect ? 'Reconnect' : 'Try again'}</button></div>`;
    } else { notice.classList.add('hidden'); notice.innerHTML = ''; }

    const container = $('files-container');
    $('load-more-wrap').classList.toggle('hidden', !state.next || state.loading);

    if (state.loading && !state.items.length) {
        container.className = '';
        container.innerHTML = `<div class="text-center py-16 text-xs text-slate-400"><i class="fa-solid fa-circle-notch spin text-2xl text-cloudprimary mb-3 block"></i>Loading files…</div>`;
        return;
    }
    if (!state.items.length) {
        container.className = '';
        container.innerHTML = (state.error || state.needReconnect) ? '' : `
            <div class="text-center py-16 bg-cloudcard-light dark:bg-cloudcard-dark rounded-3xl border border-cloudborder-light dark:border-cloudborder-dark p-8">
                <i class="fa-solid fa-folder-open text-4xl text-slate-300 dark:text-slate-600 mb-3"></i>
                <h3 class="font-bold text-sm text-slate-900 dark:text-white">This folder is empty</h3>
                <p class="text-xs text-slate-400 mt-1">Tap Upload, drop files here, or create a folder to get started.</p>
            </div>`;
        return;
    }

    const sortBy = $('sort-by').value;
    const items = state.items.slice().sort((a, b) => {
        if (a.isFolder !== b.isFolder) return a.isFolder ? -1 : 1;
        if (sortBy === 'date') return (b.modified || '').localeCompare(a.modified || '');
        if (sortBy === 'size') return b.size - a.size;
        return a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' });
    });
    const open = f => f.isFolder ? `data-act="open-folder" data-id="${esc(f.id)}" data-name="${esc(f.name)}"` : `data-act="preview" data-id="${esc(f.id)}"`;

    if (state.viewMode === 'grid') {
        container.className = 'grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4';
        container.innerHTML = items.map(f => `
            <div class="cute-card bg-cloudcard-light dark:bg-cloudcard-dark p-4 rounded-2xl border border-cloudborder-light dark:border-cloudborder-dark shadow-sm hover:shadow-md flex flex-col justify-between space-y-3">
                <div class="flex items-start justify-between space-x-2">
                    <div ${open(f)} class="flex items-center space-x-3 overflow-hidden cursor-pointer flex-1">
                        <div class="w-10 h-10 rounded-xl flex items-center justify-center shrink-0 ${kindStyle(f.type)}"><i class="fa-solid ${kindIcon(f.type)} text-base"></i></div>
                        <div class="overflow-hidden">
                            <h4 class="font-bold text-xs text-slate-900 dark:text-white truncate hover:text-cloudprimary transition">${esc(f.name)}</h4>
                            <p class="text-[11px] text-slate-400">${f.isFolder ? 'Folder' : formatBytes(f.size)}</p>
                        </div>
                    </div>
                    <button data-act="delete" data-id="${esc(f.id)}" title="Move to trash" aria-label="Move to trash" class="text-slate-400 hover:text-red-500 text-xs p-1"><i class="fa-solid fa-trash-can"></i></button>
                </div>
                <div class="flex items-center justify-between pt-2 border-t border-slate-100 dark:border-slate-800 text-[10px] text-slate-400">
                    <span>${esc(f.modified)}</span>
                    ${f.isFolder ? '' : `<button data-act="preview" data-id="${esc(f.id)}" class="text-cloudprimary font-semibold hover:underline">Preview</button>`}
                </div>
            </div>`).join('');
    } else {
        container.className = 'bg-cloudcard-light dark:bg-cloudcard-dark rounded-3xl border border-cloudborder-light dark:border-cloudborder-dark divide-y divide-slate-100 dark:divide-slate-800 overflow-hidden shadow-sm';
        container.innerHTML = items.map(f => `
            <div class="p-4 flex items-center justify-between hover:bg-slate-50 dark:hover:bg-slate-800/50 transition">
                <div ${open(f)} class="flex items-center space-x-3 cursor-pointer overflow-hidden flex-1">
                    <div class="w-8 h-8 rounded-lg flex items-center justify-center shrink-0 ${kindStyle(f.type)}"><i class="fa-solid ${kindIcon(f.type)} text-xs"></i></div>
                    <span class="font-semibold text-xs text-slate-900 dark:text-white truncate">${esc(f.name)}</span>
                </div>
                <div class="flex items-center space-x-4 text-xs text-slate-400 shrink-0">
                    <span class="hidden sm:inline">${esc(f.modified)}</span>
                    <span class="w-16 text-right">${f.isFolder ? '—' : formatBytes(f.size)}</span>
                    <button data-act="delete" data-id="${esc(f.id)}" title="Move to trash" aria-label="Move to trash" class="text-slate-400 hover:text-red-500 p-1"><i class="fa-solid fa-trash-can"></i></button>
                </div>
            </div>`).join('');
    }
}

/* ----------------------------------------------------------- navigation / data */
let loadSeq = 0;
const RECONNECT_CODES = ['popup_closed', 'popup_failed_to_open', 'access_denied', 'user_cancelled', 'popup_window_error', 'empty_window_error', 'interaction_required', 'login_required', 'consent_required'];

async function loadFolder({ more = false, silent = false, interactive = false } = {}) {
    const acc = getAcc(state.activeId);
    if (!acc) return;
    const my = ++loadSeq;
    const folder = currentFolder();
    if (!more && !silent) { state.items = []; state.next = null; }
    state.loading = true; state.error = null; state.needReconnect = false;
    renderFiles();
    try {
        await getToken(acc, { interactive });
        const r = await prov(acc).list(acc, folder.id, more ? state.next : null);
        if (my !== loadSeq) return;
        state.items = more ? state.items.concat(r.items) : r.items;
        state.next = r.next;
        if (!more && !silent) refreshQuota(acc);
    } catch (e) {
        if (my !== loadSeq) return;
        if (e.needReconnect || RECONNECT_CODES.includes(e.code)) { state.needReconnect = true; state.error = e.needReconnect ? null : e.message; }
        else state.error = e.message || 'Could not load files.';
    } finally {
        if (my === loadSeq) { state.loading = false; renderFiles(); }
    }
}

function selectAccount(id) {
    const acc = getAcc(id);
    if (!acc) return;
    state.activeId = id; state.view = 'files';
    state.path = [{ id: 'root', name: prov(acc).rootName }];
    state.items = []; state.next = null; state.error = null; state.needReconnect = false;
    renderUI();
    loadFolder({ interactive: true });
}
function backToAccounts() { loadSeq++; state.view = 'accounts'; state.activeId = null; renderUI(); refreshAllQuotas(); }
function openFolder(id, name) { state.path.push({ id, name }); loadFolder(); }
function goCrumb(i) { if (i === state.path.length - 1) return; state.path = state.path.slice(0, i + 1); loadFolder(); }

async function refreshQuota(acc) {
    try {
        const q = await prov(acc).quota(acc);
        acc.used = q.used; acc.limit = q.limit;
        saveAccounts();
        if (state.view === 'accounts') renderAccounts();
    } catch (e) { /* quota is a nicety, never block on it */ }
}
async function refreshAllQuotas() {
    await Promise.all(state.accounts.map(async acc => {
        try { await getToken(acc, { interactive: false }); await refreshQuota(acc); } catch (e) { /* needs sign-in; shown on the card */ }
    }));
    if (state.view === 'accounts') renderAccounts();
}

/* --------------------------------------------------------------- accounts */
function upsertAccount(acc) {
    const i = state.accounts.findIndex(a => a.id === acc.id);
    if (i >= 0) state.accounts[i] = { ...state.accounts[i], ...acc }; else state.accounts.push(acc);
    saveAccounts();
}
async function connectProvider(kind) {
    const P = PROVIDERS[kind];
    if (!P.configured()) {
        closeModal('connect-modal');
        openSettings();
        toast(`Add your ${P.name} Client ID first (steps are in README.md)`, 'error');
        return;
    }
    try {
        const acc = await P.connect();
        const existed = !!getAcc(acc.id);
        upsertAccount(acc);
        closeModal('connect-modal');
        toast(existed ? `${acc.email} reconnected` : `${acc.email} connected`);
        if (state.view === 'files' && state.activeId === acc.id) loadFolder({ silent: true }); else renderUI();
    } catch (e) {
        toast(e.message || 'Could not connect the account', 'error');
    }
}
async function disconnectAccount(id) {
    const acc = getAcc(id);
    if (!acc) return;
    const ok = await confirmDialog({ title: 'Remove this account?', text: `${acc.email} will be removed from CloudNest on this device. Your files are not touched.`, okLabel: 'Remove' });
    if (!ok) return;
    prov(acc).disconnect(acc);
    state.accounts = state.accounts.filter(a => a.id !== id);
    saveAccounts();
    renderUI();
    toast('Account removed');
}

/* ----------------------------------------------------------- folder / delete */
function openFolderModal() { $('new-folder-name').value = ''; openModal('folder-modal'); setTimeout(() => $('new-folder-name').focus(), 50); }
async function createFolder() {
    const name = $('new-folder-name').value.trim();
    if (!name) { toast('Enter a folder name', 'error'); return; }
    const acc = getAcc(state.activeId);
    if (acc.provider === 'microsoft' && /[\\/:*?"<>|]/.test(name)) { toast('OneDrive folder names cannot contain \\ / : * ? " < > |', 'error'); return; }
    try {
        await prov(acc).createFolder(acc, currentFolder().id, name);
        closeModal('folder-modal');
        toast(`Folder "${name}" created`);
        loadFolder({ silent: true });
    } catch (e) { toast(e.message || 'Could not create the folder', 'error'); }
}
async function deleteItem(id) {
    const item = state.items.find(i => i.id === id);
    const acc = getAcc(state.activeId);
    if (!item || !acc) return;
    const P = prov(acc);
    const ok = await confirmDialog({
        title: item.isFolder ? 'Delete this folder?' : 'Delete this file?',
        text: `"${item.name}" will move to the ${P.trashName}.` + (item.isFolder ? ' Everything inside it moves too.' : ''),
        okLabel: 'Move to trash'
    });
    if (!ok) return;
    try {
        await P.remove(acc, item);
        state.items = state.items.filter(i => i.id !== id);
        renderFiles();
        toast('Moved to trash');
        refreshQuota(acc);
    } catch (e) { toast(e.message || 'Could not delete', 'error'); }
}

/* ------------------------------------------------------------------ uploads */
function openUpload() {
    const sel = $('upload-account-select');
    sel.innerHTML = state.accounts.map(a => `<option value="${esc(a.id)}">${esc(a.name)} — ${esc(a.email)} (${esc(prov(a).name)})</option>`).join('');
    if (state.view === 'files' && state.activeId) sel.value = state.activeId;
    $('upload-file-input').value = '';
    updateUploadSummary();
    updateUploadDest();
    openModal('upload-modal');
}
function updateUploadDest() {
    const id = $('upload-account-select').value;
    const acc = getAcc(id);
    if (!acc) { $('upload-dest').textContent = ''; return; }
    $('upload-dest').textContent = 'Destination: ' + (state.view === 'files' && state.activeId === id ? state.path.map(p => p.name).join(' / ') : prov(acc).rootName);
}
function updateUploadSummary() {
    const files = [...$('upload-file-input').files];
    $('upload-summary').textContent = files.length
        ? `${files.length} file${files.length > 1 ? 's' : ''} selected • ${formatBytes(files.reduce((s, f) => s + f.size, 0))}`
        : 'Any file type, any size your account allows';
}
function startUpload() {
    const files = [...$('upload-file-input').files];
    if (!files.length) { toast('Choose at least one file', 'error'); return; }
    const acc = getAcc($('upload-account-select').value);
    if (!acc) { toast('Choose an account', 'error'); return; }
    const inFolder = state.view === 'files' && state.activeId === acc.id;
    closeModal('upload-modal');
    enqueueFiles(files, acc, inFolder ? currentFolder().id : 'root', inFolder ? currentFolder().name : prov(acc).rootName);
}

let jobSeq = 0;
async function enqueueFiles(files, acc, parentId, parentLabel) {
    try { await getToken(acc, { interactive: true }); }
    catch (e) { toast(e.message || 'Sign in to this account first', 'error'); return; }
    files.forEach(file => state.uploads.push({ id: ++jobSeq, accId: acc.id, parentId, parentLabel, file, pct: 0, status: 'queued', error: '', ctl: {} }));
    state.trayCollapsed = false;
    renderTray();
    processUploads();
}
async function processUploads() {
    if (state.uploading) return;
    state.uploading = true;
    const touched = new Set();
    for (;;) {
        const job = state.uploads.find(u => u.status === 'queued');
        if (!job) break;
        job.status = 'uploading'; job.pct = 0; job.error = '';
        renderTray();
        const acc = getAcc(job.accId);
        try {
            if (!acc) throw new Error('Account was removed');
            await getToken(acc, { interactive: false });
            await prov(acc).upload(acc, job.parentId, job.file, p => { job.pct = Math.max(0, Math.min(100, Math.round(p * 100))); updateTrayRow(job); }, job.ctl);
            job.status = 'done'; job.pct = 100;
            touched.add(job.accId);
            if (state.view === 'files' && state.activeId === job.accId && currentFolder().id === job.parentId) loadFolder({ silent: true });
        } catch (e) {
            if (job.ctl.cancelled || e.cancelled) job.status = 'cancelled';
            else { job.status = 'error'; job.error = e.needReconnect ? 'Sign-in expired. Tap retry to sign in again.' : (e.message || 'Upload failed'); }
        }
        renderTray();
    }
    state.uploading = false;
    touched.forEach(id => { const a = getAcc(id); if (a) refreshQuota(a); });
}
function trayRowHtml(u) {
    const icon = { queued: 'fa-clock text-slate-400', uploading: 'fa-circle-notch spin text-cloudprimary', done: 'fa-circle-check text-emerald-500', error: 'fa-circle-exclamation text-red-500', cancelled: 'fa-ban text-slate-400' }[u.status];
    const sub = u.status === 'error' ? esc(u.error) : u.status === 'done' ? `Uploaded to ${esc(u.parentLabel)}` : u.status === 'cancelled' ? 'Cancelled' : u.status === 'queued' ? 'Waiting…' : `<span data-pct>${u.pct}%</span> • ${formatBytes(u.file.size)}`;
    const action = (u.status === 'queued' || u.status === 'uploading') ? `<button data-act="upload-cancel" data-job="${u.id}" aria-label="Cancel upload" class="text-slate-400 hover:text-red-500 p-1"><i class="fa-solid fa-xmark text-xs"></i></button>`
        : u.status === 'error' ? `<button data-act="upload-retry" data-job="${u.id}" class="text-[11px] font-semibold text-cloudprimary px-2 py-1">Retry</button>` : '';
    return `<div data-job-row="${u.id}" class="px-4 py-3 space-y-1.5">
        <div class="flex items-center gap-3">
            <i class="fa-solid ${icon} text-sm shrink-0"></i>
            <div class="flex-1 overflow-hidden"><p class="text-xs font-semibold text-slate-900 dark:text-white truncate">${esc(u.file.name)}</p><p class="text-[11px] text-slate-400 truncate">${sub}</p></div>
            ${action}
        </div>
        ${u.status === 'uploading' ? `<div class="h-1.5 bg-slate-100 dark:bg-slate-800 rounded-full overflow-hidden"><div data-bar class="h-full bg-cloudprimary rounded-full transition-all" style="width:${u.pct}%"></div></div>` : ''}
    </div>`;
}
function renderTray() {
    const tray = $('upload-tray');
    if (!state.uploads.length) { tray.classList.add('hidden'); return; }
    tray.classList.remove('hidden');
    const active = state.uploads.filter(u => u.status === 'queued' || u.status === 'uploading').length;
    const failed = state.uploads.filter(u => u.status === 'error').length;
    $('tray-title').textContent = active ? `Uploading ${active} file${active > 1 ? 's' : ''}…` : failed ? 'Some uploads failed' : 'Uploads complete';
    $('tray-list').classList.toggle('hidden', state.trayCollapsed);
    $('tray-toggle-icon').className = 'fa-solid text-xs ' + (state.trayCollapsed ? 'fa-chevron-up' : 'fa-chevron-down');
    $('tray-list').innerHTML = state.uploads.map(trayRowHtml).join('');
}
function updateTrayRow(job) {
    const row = document.querySelector(`[data-job-row="${job.id}"]`);
    if (!row) return;
    const bar = row.querySelector('[data-bar]'); if (bar) bar.style.width = job.pct + '%';
    const pct = row.querySelector('[data-pct]'); if (pct) pct.textContent = job.pct + '%';
}
function cancelUpload(id) {
    const job = state.uploads.find(u => u.id === id);
    if (!job) return;
    job.ctl.cancelled = true;
    if (job.status === 'queued') job.status = 'cancelled';
    else if (job.ctl.abort) job.ctl.abort();
    renderTray();
}
async function retryUpload(id) {
    const job = state.uploads.find(u => u.id === id);
    const acc = job && getAcc(job.accId);
    if (!job || !acc) return;
    try { await getToken(acc, { interactive: true }); } catch (e) { toast(e.message, 'error'); return; }
    job.ctl = {}; job.status = 'queued'; job.error = '';
    renderTray(); processUploads();
}
function closeTray() {
    state.uploads = state.uploads.filter(u => u.status === 'queued' || u.status === 'uploading');
    if (state.uploads.length) state.trayCollapsed = true;
    renderTray();
}

/* drag & drop */
let dragDepth = 0;
const hasFiles = e => !!(e.dataTransfer && [...(e.dataTransfer.types || [])].includes('Files'));
function showDrop(on) { const o = $('drop-overlay'); o.classList.toggle('hidden', !on); o.classList.toggle('flex', on); }
function filesFromDrop(dt) {
    const out = []; let dirs = 0;
    if (dt.items && dt.items.length && dt.items[0].webkitGetAsEntry) {
        for (const it of dt.items) {
            if (it.kind !== 'file') continue;
            const en = it.webkitGetAsEntry();
            if (en && en.isDirectory) { dirs++; continue; }
            const f = it.getAsFile(); if (f) out.push(f);
        }
    } else out.push(...dt.files);
    if (dirs) toast("Folders can't be dropped. Drop files, or create a folder first.", 'error');
    return out;
}

/* ------------------------------------------------------------------ preview */
let previewSeq = 0, previewUrls = [];
const trackUrl = u => { previewUrls.push(u); return u; };
const PREVIEW_LIMIT = 100 * 1024 * 1024;
const loadingEl = t => el('div', 'm-auto text-center text-xs text-slate-400', `<i class="fa-solid fa-circle-notch spin text-2xl text-cloudprimary mb-3 block"></i><span>${esc(t)}</span>`);
function msgEl(text, error) {
    const d = el('div', 'm-auto text-center max-w-sm p-6', `<i class="fa-solid ${error ? 'fa-circle-exclamation text-red-500' : 'fa-eye-slash text-slate-400'} text-3xl mb-3"></i><p class="text-xs text-slate-500 dark:text-slate-400"></p>`);
    d.querySelector('p').textContent = text;
    return d;
}
const setBody = node => $('preview-body').replaceChildren(node);
const loadScript = (() => {
    const cache = {};
    return url => cache[url] || (cache[url] = new Promise((res, rej) => {
        const s = document.createElement('script');
        s.src = url; s.onload = res; s.onerror = () => { delete cache[url]; rej(new Error('Could not load the viewer. Check your internet connection.')); };
        document.head.appendChild(s);
    }));
})();

function closePreview() {
    previewSeq++;
    previewUrls.forEach(u => URL.revokeObjectURL(u)); previewUrls = [];
    state.preview = null; slides = []; workbook = null;
    $('preview-body').replaceChildren();
    closeModal('preview-modal');
}

async function openPreview(id) {
    const item = state.items.find(i => i.id === id);
    const acc = getAcc(state.activeId);
    if (!item || !acc) return;
    const P = prov(acc);
    const my = ++previewSeq;
    state.preview = { item, acc };
    $('preview-title').textContent = item.name;
    $('preview-subtitle').textContent = `${P.name} • ${acc.name}`;
    $('preview-meta').textContent = `Size: ${formatBytes(item.size)} • Modified ${item.modified || '—'}`;
    $('preview-icon-container').className = `w-10 h-10 shrink-0 rounded-xl flex items-center justify-center ${kindStyle(item.type)}`;
    $('preview-icon-container').innerHTML = `<i class="fa-solid ${kindIcon(item.type)} text-lg"></i>`;
    const open = $('preview-open');
    open.classList.toggle('hidden', !item.webUrl); open.classList.toggle('inline-flex', !!item.webUrl);
    if (item.webUrl) { open.href = item.webUrl; $('preview-open-label').textContent = 'Open in ' + P.name; }
    setBody(loadingEl('Loading preview…'));
    openModal('preview-modal');
    try { await renderPreview(item, acc, P, my); }
    catch (e) { if (my === previewSeq) setBody(msgEl(e.message || 'Could not load the preview.', true)); }
}

async function blobFor(item, acc, P) {
    if (!item.native && item.size > PREVIEW_LIMIT) throw new Error(`This file is ${formatBytes(item.size)}. Previews are limited to ${formatBytes(PREVIEW_LIMIT)}. Use Download or Open.`);
    return P.blob(acc, item);
}

async function renderPreview(item, acc, P, my) {
    const stale = () => my !== previewSeq;
    const ext = extOf(item.name);

    if (item.type === 'image') {
        const url = (await P.directUrl(acc, item)) || trackUrl(URL.createObjectURL(await blobFor(item, acc, P)));
        if (stale()) return;
        const img = el('img', 'max-w-full max-h-full m-auto rounded-xl shadow'); img.alt = item.name; img.src = url;
        img.onerror = () => setBody(msgEl('This image could not be displayed. Try Download.', true));
        return setBody(img);
    }
    if (item.type === 'video' || item.type === 'audio') {
        const url = (await P.directUrl(acc, item)) || trackUrl(URL.createObjectURL(await blobFor(item, acc, P)));
        if (stale()) return;
        const m = el(item.type === 'video' ? 'video' : 'audio', 'max-w-full max-h-full m-auto rounded-xl w-full'); m.controls = true; m.src = url;
        return setBody(m);
    }
    if (item.type === 'pdf' || (item.native && item.type !== 'folder')) {
        const blob = await blobFor(item, acc, P);
        if (stale()) return;
        return renderPdf(blob, my);
    }
    if (item.type === 'text') {
        const blob = await blobFor(item, acc, P);
        if (stale()) return;
        const text = await blob.text();
        const pre = el('pre', 'w-full text-xs leading-relaxed whitespace-pre-wrap break-words bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 rounded-xl p-4 select-text');
        pre.textContent = text.length > 200000 ? text.slice(0, 200000) + '\n\n… (truncated, download for the full file)' : text;
        return setBody(pre);
    }
    if (item.type === 'doc' && ext === 'docx') {
        const blob = await blobFor(item, acc, P);
        await loadScript(LIB.mammoth);
        if (stale()) return;
        const result = await mammoth.convertToHtml({ arrayBuffer: await blob.arrayBuffer() });
        if (stale()) return;
        const f = el('iframe', 'w-full flex-1 min-h-[60vh] bg-white rounded-xl');
        f.setAttribute('sandbox', '');   // no scripts, fully isolated
        f.srcdoc = '<style>body{font-family:system-ui,sans-serif;padding:24px;max-width:800px;margin:auto;line-height:1.6;color:#111;background:#fff}img{max-width:100%}table{border-collapse:collapse}td,th{border:1px solid #ccc;padding:4px 8px}</style>' + result.value;
        return setBody(f);
    }
    if (item.type === 'sheet') {
        const blob = await blobFor(item, acc, P);
        await loadScript(LIB.xlsx);
        if (stale()) return;
        workbook = XLSX.read(await blob.arrayBuffer(), { type: 'array' });
        return renderSheet(0);
    }
    if (item.type === 'ppt' && ext === 'pptx') {
        const blob = await blobFor(item, acc, P);
        await loadScript(LIB.jszip);
        if (stale()) return;
        slides = await readPptx(await blob.arrayBuffer());
        if (stale()) return;
        if (!slides.length) return setBody(msgEl('No slides found in this presentation.'));
        pptIndex = 0;
        return renderSlide();
    }
    setBody(msgEl(item.type === 'folder' ? 'Folders cannot be previewed.' : `No in-app preview for .${ext || 'this'} files yet. Use Download, or Open in ${P.name}.`));
}

async function renderPdf(blob, my) {
    await loadScript(LIB.pdf);
    pdfjsLib.GlobalWorkerOptions.workerSrc = LIB.pdfWorker;
    const pdf = await pdfjsLib.getDocument({ data: await blob.arrayBuffer() }).promise;
    if (my !== previewSeq) return;
    const wrap = el('div', 'w-full flex flex-col items-center gap-3');
    setBody(wrap);
    const width = Math.min($('preview-body').clientWidth - 32 || 700, 900);
    const dpr = Math.min(window.devicePixelRatio || 1, 2);
    const max = Math.min(pdf.numPages, 40);
    for (let n = 1; n <= max; n++) {
        const page = await pdf.getPage(n);
        if (my !== previewSeq) return;
        const base = page.getViewport({ scale: 1 });
        const vp = page.getViewport({ scale: (width / base.width) * dpr });
        const c = el('canvas', 'bg-white rounded-lg shadow max-w-full');
        c.width = vp.width; c.height = vp.height; c.style.width = width + 'px'; c.style.height = (vp.height / dpr) + 'px';
        wrap.appendChild(c);
        await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
    }
    if (pdf.numPages > max) wrap.appendChild(el('p', 'text-xs text-slate-400 py-3', `Showing the first ${max} of ${pdf.numPages} pages. Download for the full file.`));
}

let workbook = null;
function renderSheet(index) {
    const name = workbook.SheetNames[index];
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, raw: false, defval: '' });
    const maxR = Math.min(rows.length, 300);
    const maxC = Math.min(26, rows.slice(0, maxR).reduce((m, r) => Math.max(m, r.length), 0)) || 1;
    const col = i => String.fromCharCode(65 + i);
    const head = Array.from({ length: maxC }, (_, i) => `<th class="p-2 text-center font-semibold text-slate-500 bg-slate-100 dark:bg-slate-800 border border-slate-200 dark:border-slate-700">${col(i)}</th>`).join('');
    const body = rows.slice(0, maxR).map((r, ri) => `<tr><td class="p-2 text-center text-slate-400 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700">${ri + 1}</td>` +
        Array.from({ length: maxC }, (_, ci) => `<td class="p-2 border border-slate-200 dark:border-slate-700 whitespace-nowrap">${esc(r[ci] ?? '')}</td>`).join('') + '</tr>').join('');
    const tabs = workbook.SheetNames.map((s, i) => `<button data-act="sheet-tab" data-i="${i}" class="px-3 py-1.5 rounded-lg text-xs font-semibold ${i === index ? 'bg-cloudprimary text-white' : 'bg-slate-200 dark:bg-slate-700 text-slate-600 dark:text-slate-200'}">${esc(s)}</button>`).join('');
    const wrap = el('div', 'w-full space-y-3', `
        <div class="flex flex-wrap gap-2">${tabs}</div>
        <div class="overflow-auto bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 max-h-[55vh] select-text">
            <table class="text-xs w-full border-collapse"><thead><tr><th class="p-2 bg-slate-100 dark:bg-slate-800 border border-slate-200 dark:border-slate-700"></th>${head}</tr></thead><tbody>${body}</tbody></table>
        </div>
        ${rows.length > maxR ? `<p class="text-xs text-slate-400">Showing the first ${maxR} of ${rows.length} rows. Download for everything.</p>` : ''}`);
    setBody(wrap);
}

let slides = [], pptIndex = 0;
async function readPptx(buf) {
    const zip = await JSZip.loadAsync(buf);
    const names = Object.keys(zip.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a, b) => parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10));
    const out = [];
    for (const n of names) {
        const xml = new DOMParser().parseFromString(await zip.file(n).async('string'), 'application/xml');
        const paras = [...xml.getElementsByTagName('a:p')].map(p => [...p.getElementsByTagName('a:t')].map(t => t.textContent).join('')).filter(Boolean);
        out.push({ title: paras[0] || `Slide ${out.length + 1}`, lines: paras.slice(1) });
    }
    return out;
}
function renderSlide() {
    const s = slides[pptIndex];
    const wrap = el('div', 'w-full max-w-2xl flex flex-col items-center m-auto', `
        <div class="w-full bg-gradient-to-br from-slate-900 to-slate-800 text-white rounded-2xl shadow-xl p-8 sm:p-10 min-h-[320px] flex flex-col justify-between border border-slate-700">
            <div class="flex justify-between items-center text-xs text-orange-400 font-semibold">
                <span><i class="fa-solid fa-file-powerpoint mr-1.5"></i>PowerPoint</span><span>Slide ${pptIndex + 1} / ${slides.length}</span>
            </div>
            <div class="my-auto space-y-3 text-center py-4"><h2 class="text-xl sm:text-2xl font-bold tracking-tight"></h2><div class="text-xs text-slate-300 max-w-lg mx-auto leading-relaxed space-y-1.5" data-lines></div></div>
            <div class="text-[10px] text-slate-400 border-t border-slate-700/60 pt-3">Text-only preview. Open the file for images and layout.</div>
        </div>
        <div class="flex items-center space-x-4 mt-4">
            <button data-act="ppt-prev" ${pptIndex === 0 ? 'disabled' : ''} class="px-4 py-2 rounded-xl text-xs font-semibold ${pptIndex === 0 ? 'bg-slate-200 dark:bg-slate-800 text-slate-400 cursor-not-allowed' : 'bg-slate-200 dark:bg-slate-800 hover:bg-slate-300 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200'}"><i class="fa-solid fa-chevron-left mr-1"></i> Previous</button>
            <span class="text-xs font-semibold text-slate-500">${pptIndex + 1} of ${slides.length}</span>
            <button data-act="ppt-next" ${pptIndex === slides.length - 1 ? 'disabled' : ''} class="px-4 py-2 rounded-xl text-xs font-semibold ${pptIndex === slides.length - 1 ? 'bg-slate-200 dark:bg-slate-800 text-slate-400 cursor-not-allowed' : 'bg-slate-200 dark:bg-slate-800 hover:bg-slate-300 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200'}">Next <i class="fa-solid fa-chevron-right ml-1"></i></button>
        </div>`);
    wrap.querySelector('h2').textContent = s.title;
    const box = wrap.querySelector('[data-lines]');
    s.lines.slice(0, 14).forEach(l => { const p = el('p'); p.textContent = l; box.appendChild(p); });
    setBody(wrap);
}

async function downloadCurrent() {
    const p = state.preview;
    if (!p) return;
    toast(`Preparing "${p.item.name}"…`);
    try { await prov(p.acc).download(p.acc, p.item); }
    catch (e) { toast(e.message || 'Download failed', 'error'); }
}

/* ----------------------------------------------------------------- settings */
function openSettings() {
    $('set-google').value = SAVED_CFG.googleClientId || CFG.googleClientId || '';
    $('set-ms').value = SAVED_CFG.msClientId || CFG.msClientId || '';
    $('reg-origin').textContent = APP_ORIGIN;
    $('reg-redirect').textContent = APP_URL;
    openModal('settings-modal');
}
function saveSettings() {
    writeStore(localStorage, 'cloudnest.config', { googleClientId: $('set-google').value.trim(), msClientId: $('set-ms').value.trim() });
    location.reload();
}
async function copyText(text) {
    try { await navigator.clipboard.writeText(text); toast('Copied'); }
    catch (e) {
        const t = el('textarea'); t.value = text; document.body.appendChild(t); t.select();
        try { document.execCommand('copy'); toast('Copied'); } catch (x) { toast('Could not copy', 'error'); }
        t.remove();
    }
}

/* -------------------------------------------------------------------- theme */
function toggleDark() {
    const on = document.documentElement.classList.toggle('dark');
    try { localStorage.setItem('cloudnest.dark', on ? '1' : '0'); } catch (e) { /* ignore */ }
}

/* ------------------------------------------------------------ event wiring */
const ACTIONS = {
    'toggle-dark': toggleDark,
    'connect-open': () => openModal('connect-modal'),
    'connect-google': () => connectProvider('google'),
    'connect-ms': () => connectProvider('microsoft'),
    'settings-open': openSettings,
    'settings-save': saveSettings,
    'copy': b => copyText($(b.dataset.src).textContent),
    'close-modal': b => closeModal(b.dataset.target),
    'select-account': b => selectAccount(b.dataset.id),
    'disconnect': b => disconnectAccount(b.dataset.id),
    'refresh-quota': () => { toast('Refreshing storage…'); refreshAllQuotas(); },
    'back': backToAccounts,
    'refresh': () => loadFolder({ interactive: true }),
    'reconnect': () => loadFolder({ interactive: true }),
    'load-more': () => loadFolder({ more: true }),
    'upload-open': openUpload,
    'upload-start': startUpload,
    'upload-cancel': b => cancelUpload(+b.dataset.job),
    'upload-retry': b => retryUpload(+b.dataset.job),
    'tray-toggle': () => { state.trayCollapsed = !state.trayCollapsed; renderTray(); },
    'tray-close': closeTray,
    'folder-open': openFolderModal,
    'folder-create': createFolder,
    'crumb': b => goCrumb(+b.dataset.i),
    'open-folder': b => openFolder(b.dataset.id, b.dataset.name),
    'preview': b => openPreview(b.dataset.id),
    'close-preview': closePreview,
    'download': downloadCurrent,
    'delete': b => deleteItem(b.dataset.id),
    'view-grid': () => { state.viewMode = 'grid'; writeStore(localStorage, 'cloudnest.viewMode', 'grid'); renderFiles(); },
    'view-list': () => { state.viewMode = 'list'; writeStore(localStorage, 'cloudnest.viewMode', 'list'); renderFiles(); },
    'sheet-tab': b => renderSheet(+b.dataset.i),
    'ppt-prev': () => { if (pptIndex > 0) { pptIndex--; renderSlide(); } },
    'ppt-next': () => { if (pptIndex < slides.length - 1) { pptIndex++; renderSlide(); } },
    'confirm-yes': () => settleConfirm(true),
    'confirm-no': () => settleConfirm(false)
};

function bindEvents() {
    document.addEventListener('click', e => {
        const b = e.target.closest('[data-act]');
        if (b) { const fn = ACTIONS[b.dataset.act]; if (fn) fn(b, e); return; }
        if (e.target.hasAttribute && e.target.hasAttribute('data-modal')) {   // click on the dark backdrop
            if (e.target.id === 'preview-modal') closePreview();
            else if (e.target.id === 'confirm-modal') settleConfirm(false);
            else closeModal(e.target.id);
        }
    });
    document.addEventListener('keydown', e => {
        if (e.key === 'Escape') {
            if (!$('confirm-modal').classList.contains('hidden')) return settleConfirm(false);
            if (!$('preview-modal').classList.contains('hidden')) return closePreview();
            document.querySelectorAll('[data-modal]:not(.hidden)').forEach(m => closeModal(m.id));
        } else if ((e.key === 'Enter' || e.key === ' ') && e.target.matches && e.target.matches('[role="button"][data-act]')) {
            e.preventDefault(); e.target.click();
        } else if (e.key === 'Enter' && e.target.id === 'new-folder-name') createFolder();
    });
    $('sort-by').addEventListener('change', renderFiles);
    $('upload-account-select').addEventListener('change', updateUploadDest);
    $('upload-file-input').addEventListener('change', updateUploadSummary);

    const zone = $('upload-drop');
    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('border-cloudprimary'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('border-cloudprimary'));
    zone.addEventListener('drop', e => {
        e.preventDefault(); e.stopPropagation(); zone.classList.remove('border-cloudprimary');
        const files = filesFromDrop(e.dataTransfer);
        if (files.length) { const dt = new DataTransfer(); files.forEach(f => dt.items.add(f)); $('upload-file-input').files = dt.files; updateUploadSummary(); }
    });

    window.addEventListener('dragenter', e => { if (state.view !== 'files' || !hasFiles(e) || anyModalOpen()) return; e.preventDefault(); dragDepth++; showDrop(true); });
    window.addEventListener('dragover', e => { if (hasFiles(e)) e.preventDefault(); });
    window.addEventListener('dragleave', () => { if (state.view !== 'files') return; dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) showDrop(false); });
    window.addEventListener('drop', e => {
        if (!hasFiles(e)) return;
        e.preventDefault(); dragDepth = 0; showDrop(false);
        if (state.view !== 'files' || anyModalOpen()) return;
        const acc = getAcc(state.activeId);
        const files = filesFromDrop(e.dataTransfer);
        if (acc && files.length) enqueueFiles(files, acc, currentFolder().id, currentFolder().name);
    });
}

/* --------------------------------------------------------------------- boot */
function boot() {
    bindEvents();
    if (location.protocol === 'file:') {
        const b = $('banner');
        b.classList.remove('hidden');
        b.innerHTML = '<i class="fa-solid fa-triangle-exclamation mr-2"></i>This page was opened as a local file. Google and Microsoft sign-in only work from <b>https://</b> or <b>http://localhost</b>. See README.md (deploy to GitHub Pages, or run <code>python3 -m http.server</code>).';
    }
    renderModeButtons();
    renderUI();
    initMsal();
    if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost')) navigator.serviceWorker.register('sw.js').catch(() => { });
    refreshAllQuotas();
}

if (IS_AUTH_POPUP) document.addEventListener('DOMContentLoaded', () => { document.body.style.display = 'none'; });
else boot();
