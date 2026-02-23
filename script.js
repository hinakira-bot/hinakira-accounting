document.addEventListener('DOMContentLoaded', () => {
    // ============================================================
    //  Section 1: Constants & State
    // ============================================================
    const CLIENT_ID = '353694435064-r6mlbk3mm2mflhl2mot2n94dpuactscc.apps.googleusercontent.com';
    const SCOPES = 'openid email profile https://www.googleapis.com/auth/drive';
    let tokenClient;
    let accessToken = localStorage.getItem('access_token');
    let tokenExpiration = localStorage.getItem('token_expiration');
    let accounts = [];            // Account master cache
    let scanResults = [];         // Scan tab working data
    let isLoggingOut = false;     // Prevent 401 cascade (multiple toasts)
    let refreshTimer = null;      // Token auto-refresh timer
    let heartbeatTimer = null;    // Periodic token validity check
    const thisYear = new Date().getFullYear();
    const thisMonth = new Date().getMonth() + 1;

    // ============================================================
    //  Section 2: DOM References
    // ============================================================
    const authBtn = document.getElementById('auth-btn');
    const settingsBtn = document.getElementById('settings-btn');
    const settingsModal = document.getElementById('settings-modal');
    const closeSettings = document.getElementById('close-settings');
    const saveSettingsBtn = document.getElementById('save-settings-btn');
    const loginOverlay = document.getElementById('login-overlay');
    const overlayLoginBtn = document.getElementById('overlay-login-btn');

    // Menu Grid navigation
    const menuGrid = document.getElementById('menu-grid');
    const backToMenuBtn = document.getElementById('back-to-menu');
    const logoTitle = document.getElementById('logo-title');

    // --- Global Fiscal Year Selector ---
    const globalFiscalYearSelect = document.getElementById('global-fiscal-year');
    (function initGlobalFiscalYear() {
        const saved = localStorage.getItem('hinakira_fiscal_year');
        const defaultYear = saved ? parseInt(saved) : thisYear;
        for (let y = thisYear + 1; y >= thisYear - 5; y--) {
            const opt = document.createElement('option');
            opt.value = y;
            opt.textContent = y + '年';
            if (y === defaultYear) opt.selected = true;
            globalFiscalYearSelect.appendChild(opt);
        }
    })();

    function getSelectedFiscalYear() {
        return parseInt(globalFiscalYearSelect.value) || thisYear;
    }

    // Callbacks registered later (after all views are initialized)
    const fiscalYearChangeCallbacks = [];
    globalFiscalYearSelect.addEventListener('change', () => {
        localStorage.setItem('hinakira_fiscal_year', globalFiscalYearSelect.value);
        fiscalYearChangeCallbacks.forEach(fn => fn());
    });

    // ============================================================
    //  Section 3: Google OAuth
    // ============================================================
    let _refreshRetryCount = 0;
    const MAX_REFRESH_RETRIES = 3;
    let _refreshResolve = null;  // For 401 retry: resolve when token refreshed
    let _popupRetried = false;   // Track if popup re-auth was already attempted
    let _isRefreshing = false;   // Prevent concurrent refresh attempts

    window.onload = function () {
        if (typeof google === 'undefined') {
            console.warn('Google Identity Services not loaded');
            return;
        }
        tokenClient = google.accounts.oauth2.initTokenClient({
            client_id: CLIENT_ID,
            scope: SCOPES,
            callback: (resp) => {
                if (resp.access_token) {
                    accessToken = resp.access_token;
                    _refreshRetryCount = 0;
                    _isRefreshing = false;
                    _popupRetried = false;
                    const expiresIn = resp.expires_in || 3600;
                    const exp = new Date().getTime() + (expiresIn * 1000);
                    localStorage.setItem('access_token', accessToken);
                    localStorage.setItem('token_expiration', exp);
                    scheduleTokenRefresh(expiresIn);
                    // Resolve any pending 401 retry
                    if (_refreshResolve) { _refreshResolve(true); _refreshResolve = null; }
                    onLoginSuccess();
                }
            },
            error_callback: (err) => {
                console.warn('Token refresh error:', err);
                _isRefreshing = false;
                _refreshRetryCount++;
                if (_refreshRetryCount < MAX_REFRESH_RETRIES) {
                    // Retry after 30 seconds
                    console.log(`Token refresh retry ${_refreshRetryCount}/${MAX_REFRESH_RETRIES} in 30s`);
                    setTimeout(() => {
                        if (tokenClient) tokenClient.requestAccessToken({ prompt: '' });
                    }, 30000);
                } else if (!_popupRetried) {
                    // Last resort: try popup re-authentication
                    console.log('Silent refresh failed, trying popup re-auth');
                    _popupRetried = true;
                    showToast('セッション更新中...再認証ウィンドウが開く場合があります');
                    setTimeout(() => {
                        if (tokenClient) tokenClient.requestAccessToken(); // with popup
                    }, 1000);
                } else {
                    // Popup also failed — show login overlay
                    console.warn('Token refresh failed after all attempts');
                    _popupRetried = false;
                    if (_refreshResolve) { _refreshResolve(false); _refreshResolve = null; }
                    if (!isLoggingOut) {
                        showToast('セッションの更新に失敗しました。再ログインしてください', true);
                        loginOverlay.classList.remove('hidden');
                        authBtn.textContent = 'Googleでログイン';
                        authBtn.onclick = handleLogin;
                    }
                }
            },
        });
        if (accessToken && tokenExpiration && new Date().getTime() < parseInt(tokenExpiration)) {
            // Schedule refresh for remaining time
            const remaining = Math.floor((parseInt(tokenExpiration) - new Date().getTime()) / 1000);
            if (remaining > 0) scheduleTokenRefresh(remaining);
            onLoginSuccess();
        } else {
            loginOverlay.classList.remove('hidden');
        }
    };

    // Background tab support: refresh token when tab becomes active
    document.addEventListener('visibilitychange', () => {
        if (document.visibilityState !== 'visible' || !accessToken) return;
        const exp = parseInt(localStorage.getItem('token_expiration') || '0');
        const now = new Date().getTime();
        const remainingSec = Math.floor((exp - now) / 1000);
        if (remainingSec <= 0) {
            // Token already expired — try immediate refresh
            console.log('Tab active: token expired, refreshing immediately');
            if (tokenClient && !_isRefreshing) {
                _isRefreshing = true;
                tokenClient.requestAccessToken({ prompt: '' });
                setTimeout(() => { _isRefreshing = false; }, 15000);
            }
        } else if (remainingSec < 300) {
            // Token expiring soon — refresh now
            console.log(`Tab active: token expires in ${remainingSec}s, refreshing`);
            if (tokenClient && !_isRefreshing) {
                _isRefreshing = true;
                tokenClient.requestAccessToken({ prompt: '' });
                setTimeout(() => { _isRefreshing = false; }, 15000);
            }
        } else {
            // Reschedule refresh (timer may have been delayed while tab was background)
            scheduleTokenRefresh(remainingSec);
        }
    });

    function handleLogin() { tokenClient && tokenClient.requestAccessToken(); }
    function handleLogout() {
        if (refreshTimer) { clearTimeout(refreshTimer); refreshTimer = null; }
        if (heartbeatTimer) { clearInterval(heartbeatTimer); heartbeatTimer = null; }
        const t = localStorage.getItem('access_token');
        // Clear session first
        accessToken = null;
        isLoggingOut = false;
        localStorage.removeItem('access_token');
        localStorage.removeItem('token_expiration');
        sessionStorage.removeItem('user_email');
        sessionStorage.removeItem('user_name');
        // Revoke token silently (don't wait for callback)
        try {
            if (t && typeof google !== 'undefined' && google.accounts && google.accounts.oauth2) {
                google.accounts.oauth2.revoke(t, () => {});
            }
        } catch (e) { /* ignore */ }
        // Show login overlay without page reload (avoids 502 if server is restarting)
        loginOverlay.classList.remove('hidden');
        authBtn.textContent = 'Googleでログイン';
        authBtn.onclick = handleLogin;
        settingsBtn.style.display = 'none';
        const userDisplay = document.getElementById('user-display');
        if (userDisplay) { userDisplay.textContent = ''; userDisplay.classList.add('hidden'); }
        // Clear displayed data
        const recentTbody = document.getElementById('recent-tbody');
        if (recentTbody) recentTbody.innerHTML = '';
        showToast('ログアウトしました');
    }
    function scheduleTokenRefresh(expiresInSec) {
        if (refreshTimer) clearTimeout(refreshTimer);
        // Refresh 5 minutes before expiry (or at 50% of lifetime if <10min)
        const refreshAt = expiresInSec > 600
            ? (expiresInSec - 300) * 1000
            : Math.floor(expiresInSec * 0.5) * 1000;
        refreshTimer = setTimeout(() => {
            console.log('Token refresh: requesting new token silently');
            if (tokenClient && !_isRefreshing) {
                _isRefreshing = true;
                tokenClient.requestAccessToken({ prompt: '' });
                // Safety valve: reset _isRefreshing if no callback within 15s
                setTimeout(() => { _isRefreshing = false; }, 15000);
            }
        }, refreshAt);
    }

    // Heartbeat: check token validity every 15 minutes
    function startHeartbeat() {
        if (heartbeatTimer) clearInterval(heartbeatTimer);
        heartbeatTimer = setInterval(async () => {
            if (!accessToken) return;
            try {
                const res = await fetch('/api/me', {
                    headers: { 'Authorization': 'Bearer ' + accessToken },
                    cache: 'no-store'
                });
                if (res.status === 401 && tokenClient && !_isRefreshing) {
                    console.log('Heartbeat: token expired, refreshing');
                    _isRefreshing = true;
                    tokenClient.requestAccessToken({ prompt: '' });
                    setTimeout(() => { _isRefreshing = false; }, 15000);
                } else if (res.ok) {
                    console.log('Heartbeat: token valid');
                }
            } catch (e) {
                console.warn('Heartbeat error:', e);
            }
        }, 15 * 60 * 1000); // 15 minutes
    }
    async function onLoginSuccess() {
        isLoggingOut = false;  // Reset 401 cascade guard
        loginOverlay.classList.add('hidden');
        authBtn.textContent = 'ログアウト';
        authBtn.onclick = handleLogout;
        settingsBtn.style.display = '';
        startHeartbeat(); // Start periodic token validity check

        // Verify token is valid by fetching user info first
        try {
            const me = await fetchAPI('/api/me');
            if (me && me.email) {
                const userDisplay = document.getElementById('user-display');
                if (userDisplay) {
                    userDisplay.textContent = me.name || me.email;
                    userDisplay.title = me.email;
                    userDisplay.classList.remove('hidden');
                }
            }
        } catch (e) {
            // If /api/me fails with 401, the fetchAPI handler already showed login overlay
            // Don't proceed with other API calls
            console.warn('Auth check failed, not loading data');
            return;
        }

        // License check
        try {
            const licStatus = await fetchAPI('/api/license/status');
            if (licStatus && !licStatus.has_license) {
                showLicenseOverlay();
                return;
            }
        } catch (e) {
            console.warn('License check failed:', e.message);
            return;
        }
        hideLicenseOverlay();

        // Load API key: prefer server-side setting, fall back to localStorage
        try {
            const settings = await fetchAPI('/api/settings');
            if (settings && settings.gemini_api_key) {
                localStorage.setItem('gemini_api_key', settings.gemini_api_key);
            }
        } catch (e) { /* ignore */ }

        const apiKey = localStorage.getItem('gemini_api_key');
        if (!apiKey) openSettings();

        loadAccounts();
        loadRecentEntries();
        loadCounterparties();
        // Initialize Drive folder structure (inbox/processed)
        if (accessToken) {
            fetchAPI('/api/drive/init', 'POST', { access_token: accessToken }).catch(() => {});
        }
    }

    authBtn.onclick = handleLogin;
    overlayLoginBtn.onclick = handleLogin;

    // ============================================================
    //  Section 4: Settings Modal
    // ============================================================
    settingsBtn.onclick = openSettings;
    closeSettings.onclick = () => settingsModal.classList.add('hidden');
    document.querySelector('.modal-backdrop')?.addEventListener('click', () => settingsModal.classList.add('hidden'));
    saveSettingsBtn.onclick = () => {
        const key = document.getElementById('api-key-input').value.trim();
        if (key) {
            localStorage.setItem('gemini_api_key', key);
            fetchAPI('/api/settings', 'POST', { gemini_api_key: key });
            settingsModal.classList.add('hidden');
            showToast('設定を保存しました');
        } else {
            showToast('APIキーを入力してください', true);
        }
    };
    function openSettings() {
        document.getElementById('api-key-input').value = localStorage.getItem('gemini_api_key') || '';
        settingsModal.classList.remove('hidden');
    }

    // ============================================================
    //  Section 5: Menu Grid Navigation (Hash Router)
    // ============================================================
    // View loaders: called when a view becomes active
    const VIEW_LOADERS = {
        'journal-input': () => { loadRecentEntries(); },
        'scan': () => {},
        'journal-book': () => initViewWithLatestDate('jb'),
        'ledger': () => initViewWithLatestDate('ledger'),
        'counterparty': () => loadCounterpartyList(),
        'accounts': () => loadAccountsList(),
        'opening-balance': () => loadOpeningBalances(),
        'backup': () => {},
        'output': () => { applyFiscalYearToOutput(); },
        'fixed-assets': () => loadFixedAssetsList(),
        'consumption-tax': () => { applyFiscalYearToTaxInputs(); loadConsumptionTax(); },
    };

    // --- Fetch latest entry date and set year/month selectors accordingly ---
    let _latestDateCache = {};

    async function initViewWithLatestDate(viewType) {
        const fy = getSelectedFiscalYear();
        const cacheKey = `${fy}_${viewType}`;

        // Determine target selectors
        const yearSelect = viewType === 'jb'
            ? document.getElementById('jb-year-select')
            : document.getElementById('ledger-year-select');
        const monthSelect = viewType === 'jb'
            ? document.getElementById('jb-month-select')
            : document.getElementById('ledger-month-select');

        // Set year to fiscal year
        yearSelect.value = String(fy);

        // Fetch latest date (with simple cache)
        if (!_latestDateCache[cacheKey]) {
            try {
                const data = await fetchAPI(`/api/journal/latest-date?fiscal_year=${fy}`);
                _latestDateCache[cacheKey] = data.latest_date || null;
            } catch (e) {
                _latestDateCache[cacheKey] = null;
            }
        }

        const latestDate = _latestDateCache[cacheKey];
        if (latestDate) {
            const latestMonth = parseInt(latestDate.substring(5, 7));
            monthSelect.value = String(latestMonth);
        } else {
            // No entries: default to January
            monthSelect.value = '1';
        }

        // Apply period and load
        if (viewType === 'jb') {
            applyJBPeriod();
            loadJournalBook();
        } else {
            applyPeriod();
            loadCurrentLedgerSubTab();
        }
    }

    const headerFiscalYear = document.getElementById('header-fiscal-year');

    function updateHeaderFiscalYear() {
        const fy = getSelectedFiscalYear();
        headerFiscalYear.textContent = `${fy}年度`;
    }

    function showMenu() {
        // Hide all content views
        document.querySelectorAll('.content-view').forEach(v => v.classList.remove('active'));
        // Show menu grid
        menuGrid.classList.add('active');
        // Hide back button, show logo, hide fiscal year badge
        backToMenuBtn.classList.add('hidden');
        headerFiscalYear.classList.add('hidden');
        logoTitle.classList.remove('hidden');
        location.hash = 'menu';
    }

    function showView(viewId) {
        // Hide menu grid
        menuGrid.classList.remove('active');
        // Hide all content views
        document.querySelectorAll('.content-view').forEach(v => v.classList.remove('active'));
        // Show target view
        const target = document.getElementById('view-' + viewId);
        if (target) {
            target.classList.add('active');
        }
        // Show back button, hide logo, show fiscal year badge
        backToMenuBtn.classList.remove('hidden');
        headerFiscalYear.classList.remove('hidden');
        updateHeaderFiscalYear();
        logoTitle.classList.add('hidden');
        // Update hash
        location.hash = viewId;
        // Call loader
        if (VIEW_LOADERS[viewId]) VIEW_LOADERS[viewId]();
    }

    // Menu tile click (both .menu-tile and .menu-tile-sm)
    menuGrid.addEventListener('click', (e) => {
        const tile = e.target.closest('.menu-tile') || e.target.closest('.menu-tile-sm');
        if (!tile) return;
        const viewId = tile.dataset.view;
        if (viewId) showView(viewId);
    });

    // Back button
    backToMenuBtn.addEventListener('click', showMenu);

    // Hash routing
    window.addEventListener('hashchange', () => {
        const hash = location.hash.replace('#', '');
        if (!hash || hash === 'menu') {
            showMenu();
        } else {
            showView(hash);
        }
    });

    // NOTE: Initial route from hash moved to end of script (after all declarations)

    // ============================================================
    //  Section 6: Shared Utilities
    // ============================================================
    async function fetchAPI(url, method = 'GET', body = null, _isRetry = false) {
        const opts = { method, headers: {}, cache: 'no-store' };
        // Add Authorization header for all API requests
        if (accessToken) {
            opts.headers['Authorization'] = 'Bearer ' + accessToken;
        }
        if (body && !(body instanceof FormData)) {
            opts.headers['Content-Type'] = 'application/json';
            opts.body = JSON.stringify(body);
        } else if (body instanceof FormData) {
            opts.body = body;
        }
        const res = await fetch(url, opts);
        if (res.status === 401) {
            // On first 401, try silent token refresh before logging out
            if (!_isRetry && tokenClient && !isLoggingOut) {
                console.log('401 received, attempting token refresh before logout');
                const refreshed = await new Promise((resolve) => {
                    _refreshResolve = resolve;
                    _refreshRetryCount = 0;
                    _isRefreshing = true;
                    tokenClient.requestAccessToken({ prompt: '' });
                    // Timeout: if no callback in 10s, consider failed
                    setTimeout(() => {
                        if (_refreshResolve) { _refreshResolve(false); _refreshResolve = null; }
                    }, 10000);
                });
                if (refreshed && accessToken) {
                    // Token refreshed — retry the original request
                    console.log('Token refreshed, retrying request');
                    // Update access_token in body if present (Drive API endpoints send token in body)
                    let retryBody = body;
                    if (body && !(body instanceof FormData) && body.access_token) {
                        retryBody = { ...body, access_token: accessToken };
                    } else if (body instanceof FormData && body.has('access_token')) {
                        retryBody = new FormData();
                        for (const [key, value] of body.entries()) {
                            retryBody.append(key, key === 'access_token' ? accessToken : value);
                        }
                    }
                    return fetchAPI(url, method, retryBody, true);
                }
            }
            // Refresh failed or already retried — show login overlay
            if (!isLoggingOut) {
                isLoggingOut = true;
                showToast('セッションが切れました。再ログインしてください', true);
                accessToken = null;
                localStorage.removeItem('access_token');
                localStorage.removeItem('token_expiration');
                loginOverlay.classList.remove('hidden');
                authBtn.textContent = 'Googleでログイン';
                authBtn.onclick = handleLogin;
                const userDisplay = document.getElementById('user-display');
                if (userDisplay) userDisplay.classList.add('hidden');
            }
            throw new Error('Unauthorized');
        }
        if (res.status === 403) {
            const data = await res.json();
            if (data.code === 'LICENSE_REQUIRED') {
                showLicenseOverlay();
                throw new Error('License required');
            }
        }
        if (!res.ok) {
            console.error(`API error: ${res.status} ${res.statusText} for ${url}`);
        }
        return res.json();
    }

    function loadAccounts() {
        fetchAPI('/api/accounts').then(data => {
            if (data.accounts) {
                accounts = data.accounts;
                populateAccountDatalist();
                populateJBAccountFilter();
            }
        }).catch(err => console.warn('loadAccounts failed:', err.message));
    }

    function populateAccountDatalist() {
        const dl = document.getElementById('account-list');
        if (!dl) return;
        dl.innerHTML = accounts.map(a => `<option value="${a.name}">`).join('');
    }

    function populateJBAccountFilter() {
        const sel = document.getElementById('jb-account');
        if (!sel) return;
        sel.innerHTML = '<option value="">全科目</option>' +
            accounts.map(a => `<option value="${a.id}">${a.name}</option>`).join('');
    }

    function fmt(n) {
        return Number(n || 0).toLocaleString();
    }

    function calcTax(amount, classification) {
        amount = parseInt(amount) || 0;
        if (classification === '10%') return Math.floor(amount * 10 / 110);
        if (classification === '8%') return Math.floor(amount * 8 / 108);
        return 0;
    }

    function showToast(msg, isError = false) {
        let toast = document.getElementById('toast');
        if (!toast) {
            toast = document.createElement('div');
            toast.id = 'toast';
            toast.className = 'toast';
            document.body.appendChild(toast);
        }
        toast.textContent = msg;
        toast.className = 'toast show' + (isError ? ' toast-error' : '');
        setTimeout(() => { toast.className = 'toast'; }, 3000);
    }

    function todayStr() {
        return new Date().toISOString().split('T')[0];
    }

    function fiscalYearDates() {
        const now = new Date();
        const y = now.getFullYear();
        return { start: `${y}-01-01`, end: `${y}-12-31` };
    }

    // ============================================================
    //  Section 6b: Help Button Tooltips (?)
    // ============================================================
    (function initHelpButtons() {
        let openTooltip = null;
        function closeTooltip() {
            if (openTooltip) { openTooltip.remove(); openTooltip = null; }
        }
        document.addEventListener('click', (e) => {
            const btn = e.target.closest('.help-btn');
            const prevBtn = openTooltip ? openTooltip._btn : null;
            closeTooltip();
            if (!btn || btn === prevBtn) return;   // toggle off
            e.preventDefault();
            e.stopPropagation();

            const tip = document.createElement('div');
            tip.className = 'help-tooltip';
            tip.textContent = btn.dataset.help;
            tip._btn = btn;
            document.body.appendChild(tip);
            openTooltip = tip;

            // Position with fixed: below the button
            const br = btn.getBoundingClientRect();
            const tw = 230;  // tooltip width
            let left = br.left + br.width / 2 - tw / 2;
            let arrowLeft = '50%';
            // Keep within viewport
            if (left < 8) {
                arrowLeft = Math.max(12, br.left + br.width / 2 - 8) + 'px';
                left = 8;
            } else if (left + tw > window.innerWidth - 8) {
                const oldLeft = left;
                left = window.innerWidth - 8 - tw;
                arrowLeft = Math.min(tw - 12, br.left + br.width / 2 - left) + 'px';
            }
            tip.style.top = (br.bottom + 8) + 'px';
            tip.style.left = left + 'px';
            tip.style.setProperty('--arrow-left', arrowLeft);
        });
        window.addEventListener('scroll', closeTooltip, true);
    })();

    // ============================================================
    //  Section 7: View 1 — 仕訳入力 (Journal Entry) [TKC FX2 Style]
    // ============================================================
    const journalForm = document.getElementById('journal-form');
    const jeDate = document.getElementById('je-date');
    const jeTaxCategory = document.getElementById('je-tax-category');
    const jeDebit = document.getElementById('je-debit');
    const jeCredit = document.getElementById('je-credit');
    const jeAmount = document.getElementById('je-amount');
    const jeCounterparty = document.getElementById('je-counterparty');
    const jeMemo = document.getElementById('je-memo');
    const jeTaxRate = document.getElementById('je-tax-rate');
    const jeAiBtn = document.getElementById('je-ai-btn');

    function applyFiscalYearConstraint() {
        const fy = getSelectedFiscalYear();
        jeDate.min = `${fy}-01-01`;
        jeDate.max = `${fy}-12-31`;
        if (jeDate.value < jeDate.min) jeDate.value = jeDate.min;
        if (jeDate.value > jeDate.max) jeDate.value = jeDate.max;
    }

    // Default date to today
    jeDate.value = todayStr();
    applyFiscalYearConstraint();

    // --- Tax category ↔ Tax rate linkage ---
    jeTaxCategory.addEventListener('change', () => {
        const cat = jeTaxCategory.value;
        if (cat === '非課税' || cat === '不課税') {
            jeTaxRate.value = '0%';
            updateNetAmount();
        }
        if ((cat === '課税仕入' || cat === '課税売上') && jeTaxRate.value === '0%') {
            jeTaxRate.value = '';
            updateNetAmount();
        }
    });

    jeTaxRate.addEventListener('change', () => {
        updateNetAmount();
    });

    // --- Counterparty autocomplete ---
    function loadCounterparties() {
        fetchAPI('/api/counterparties').then(data => {
            const dl = document.getElementById('counterparty-list');
            if (!dl) return;
            dl.innerHTML = (data.counterparties || []).map(c => `<option value="${c}">`).join('');
        });
    }

    // --- Resolve tax_classification for DB storage ---
    function resolveTaxClassification(taxCategory, taxRate) {
        if (taxCategory === '非課税') return '非課税';
        if (taxCategory === '不課税') return '不課税';
        if (taxRate === '10%') return '10%';
        if (taxRate === '8%') return '8%';
        return '10%';
    }

    // --- Reverse: DB tax_classification → display ---
    function parseTaxClassification(dbValue) {
        if (dbValue === '非課税') return { taxCategory: '非課税', taxRate: '0%' };
        if (dbValue === '不課税') return { taxCategory: '不課税', taxRate: '0%' };
        if (dbValue === '8%') return { taxCategory: '課税仕入', taxRate: '8%' };
        if (dbValue === '10%') return { taxCategory: '課税仕入', taxRate: '10%' };
        return { taxCategory: '', taxRate: '' };
    }

    // --- AI Auto-Detect Button ---
    jeAiBtn.addEventListener('click', async () => {
        const apiKey = localStorage.getItem('gemini_api_key');
        if (!apiKey) { showToast('設定画面でAPIキーを設定してください', true); openSettings(); return; }

        const amount = parseInt(jeAmount.value) || 0;
        const counterparty = jeCounterparty.value.trim();
        const memo = jeMemo.value.trim();
        if (!counterparty && !memo && !amount) {
            showToast('取引先・摘要・金額のいずれかを入力してください', true);
            return;
        }

        const predData = [{
            index: 0,
            counterparty: counterparty,
            memo: memo,
            amount: amount,
            debit: jeDebit.value.trim() || '',
            credit: jeCredit.value.trim() || '',
            tax_category: jeTaxCategory.value || '',
            tax_rate: jeTaxRate.value || '',
        }];

        jeAiBtn.disabled = true;
        jeAiBtn.textContent = 'AI判定中...';

        try {
            const predictions = await fetchAPI('/api/predict', 'POST', {
                data: predData,
                gemini_api_key: apiKey,
            });

            if (predictions.error) throw new Error(predictions.error);
            if (Array.isArray(predictions) && predictions.length > 0) {
                const p = predictions[0];
                if (!jeDebit.value.trim() && p.debit) jeDebit.value = p.debit;
                if (!jeCredit.value.trim() && p.credit) jeCredit.value = p.credit;
                if (!jeTaxCategory.value && p.tax_category) jeTaxCategory.value = p.tax_category;
                if (!jeTaxRate.value && p.tax_rate) jeTaxRate.value = p.tax_rate;
                if (jeTaxCategory.value === '非課税' || jeTaxCategory.value === '不課税') {
                    jeTaxRate.value = '0%';
                }
                updateNetAmount();
                showToast('AI判定が完了しました');
            } else {
                showToast('AI判定結果が空です', true);
            }
        } catch (err) {
            showToast('AI判定エラー: ' + err.message, true);
        } finally {
            jeAiBtn.disabled = false;
            jeAiBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2L2 7l10 5 10-5-10-5z"/><path d="M2 17l10 5 10-5"/><path d="M2 12l10 5 10-5"/></svg> AI自動判定`;
        }
    });

    // --- Form Submit ---
    journalForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        const amount = parseInt(jeAmount.value) || 0;
        const taxCategory = jeTaxCategory.value;
        const taxRate = jeTaxRate.value;
        const taxClassification = resolveTaxClassification(taxCategory, taxRate);

        const entry = {
            entry_date: jeDate.value,
            debit_account: jeDebit.value.trim(),
            credit_account: jeCredit.value.trim(),
            amount: amount,
            tax_classification: taxClassification,
            counterparty: jeCounterparty.value.trim(),
            memo: jeMemo.value.trim(),
            source: 'manual',
        };

        console.log('Journal submit:', JSON.stringify(entry));

        if (!entry.debit_account || !entry.credit_account || !entry.amount) {
            showToast('借方科目・貸方科目・金額は必須です', true);
            console.warn('Validation failed:', { debit: entry.debit_account, credit: entry.credit_account, amount: entry.amount });
            return;
        }

        // Fiscal year range check
        const fy = getSelectedFiscalYear();
        const entryYear = parseInt((entry.entry_date || '').substring(0, 4));
        if (entryYear !== fy) {
            showToast(`${fy}年度の範囲外の日付です（${fy}/01/01〜${fy}/12/31）`, true);
            return;
        }

        try {
            const res = await fetchAPI('/api/journal', 'POST', entry);
            if (res.status === 'success') {
                showToast('仕訳を登録しました');
                journalForm.reset();
                jeDate.value = todayStr();
                applyFiscalYearConstraint();
                loadRecentEntries();
                loadCounterparties();
            } else {
                showToast('登録に失敗しました: ' + (res.error || JSON.stringify(res)), true);
            }
        } catch (err) {
            console.error('Journal submit error:', err);
            if (err.message !== 'Unauthorized') {
                showToast('通信エラー: ' + err.message, true);
            }
        }
    });

    // --- Recent Entries (below form) ---
    let recentEntriesCache = [];

    function loadRecentEntries() {
        const fy = getSelectedFiscalYear();
        fetchAPI(`/api/journal?start_date=${fy}-01-01&end_date=${fy}-12-31&per_page=10&page=1`).then(data => {
            const entries = data.entries || [];
            recentEntriesCache = entries;
            const tbody = document.getElementById('recent-tbody');
            if (!tbody) return;
            tbody.innerHTML = entries.map(e => {
                const parsed = parseTaxClassification(e.tax_classification);
                return `
                <tr data-id="${e.id}">
                    <td>${e.entry_date || ''}</td>
                    <td>${parsed.taxCategory || ''}</td>
                    <td>${e.debit_account || ''}</td>
                    <td>${e.credit_account || ''}</td>
                    <td class="text-right">${fmt(e.amount)}</td>
                    <td>${e.counterparty || ''}</td>
                    <td>${e.memo || ''}</td>
                    <td>${parsed.taxRate || ''}</td>
                    <td>${e.evidence_url ? `<a href="${e.evidence_url}" target="_blank" class="evidence-link" title="証憑を表示">📎</a>` : ''}</td>
                    <td><button class="btn-row-delete" data-id="${e.id}" title="削除">✕</button></td>
                </tr>`;
            }).join('');

            // Attach row click → open edit modal
            tbody.querySelectorAll('tr').forEach(row => {
                row.addEventListener('click', (ev) => {
                    // Skip if delete button or evidence link was clicked
                    if (ev.target.closest('.btn-row-delete') || ev.target.closest('.evidence-link')) return;
                    const id = row.dataset.id;
                    const entry = recentEntriesCache.find(en => String(en.id) === String(id));
                    if (entry) openJEDetailModal(entry, loadRecentEntries);
                });
            });

            // Attach delete handlers
            tbody.querySelectorAll('.btn-row-delete').forEach(btn => {
                btn.addEventListener('click', async () => {
                    const id = btn.dataset.id;
                    if (!confirm('この仕訳を削除しますか？')) return;
                    try {
                        const res = await fetchAPI(`/api/journal/${id}`, 'DELETE');
                        if (res.status === 'success') {
                            showToast('仕訳を削除しました');
                            loadRecentEntries();
                        } else {
                            showToast('削除に失敗しました', true);
                        }
                    } catch (err) {
                        showToast('削除に失敗しました', true);
                    }
                });
            });
        });
    }

    // ============================================================
    //  Section 8: View 2 — 証憑読み取り (Document Scanning)
    // ============================================================
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const scanStatus = document.getElementById('scan-status');
    const scanStatusText = document.getElementById('scan-status-text');
    const scanResultsCard = document.getElementById('scan-results');
    const scanTbody = document.getElementById('scan-tbody');
    const scanSaveBtn = document.getElementById('scan-save');
    const scanAddFile = document.getElementById('scan-add-file');
    const scanAddInput = document.getElementById('scan-add-input');
    const scanDupAlert = document.getElementById('scan-duplicate-alert');
    const scanClearAll = document.getElementById('scan-clear-all');
    const scanAiBulk = document.getElementById('scan-ai-bulk');

    // Clear all scanned results
    scanClearAll.addEventListener('click', () => {
        if (!scanResults.length) return;
        if (!confirm('解析結果をすべて削除しますか？')) return;
        scanResults = [];
        scanTbody.innerHTML = '';
        scanResultsCard.classList.add('hidden');
        showToast('解析結果をすべて削除しました');
    });

    // AI bulk re-predict: re-judge accounts and tax for all scan results
    scanAiBulk.addEventListener('click', async () => {
        if (!scanResults.length) { showToast('判定するデータがありません', true); return; }
        const apiKey = localStorage.getItem('gemini_api_key');
        if (!apiKey) { showToast('設定画面でAPIキーを設定してください', true); openSettings(); return; }

        const predData = scanResults.map((r, i) => ({
            index: i,
            counterparty: r.counterparty || '',
            memo: r.memo || '',
            amount: r.amount || 0,
            debit: '',   // 空にして全て再判定
            credit: '',
        }));

        scanAiBulk.disabled = true;
        scanAiBulk.textContent = 'AI判定中...';

        try {
            const predictions = await fetchAPI('/api/predict', 'POST', {
                data: predData,
                gemini_api_key: apiKey,
            });

            if (predictions.error) throw new Error(predictions.error);
            if (Array.isArray(predictions)) {
                for (let i = 0; i < predictions.length && i < scanResults.length; i++) {
                    const pred = predictions[i];
                    if (pred.debit) scanResults[i].debit_account = pred.debit;
                    if (pred.credit) scanResults[i].credit_account = pred.credit;
                    // inferTaxClassification で消費税を連動
                    const newTax = inferTaxClassification(
                        scanResults[i].debit_account,
                        scanResults[i].credit_account
                    );
                    if (newTax) scanResults[i].tax_classification = newTax;
                }
                renderScanResults();
                showToast(`${predictions.length}件のAI判定が完了しました`);
            }
        } catch (err) {
            showToast('AI判定エラー: ' + err.message, true);
        } finally {
            scanAiBulk.disabled = false;
            scanAiBulk.textContent = 'AI一括判定';
        }
    });

    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        handleScanFiles(e.dataTransfer.files);
    });
    dropZone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', () => { handleScanFiles(fileInput.files); fileInput.value = ''; });
    scanAddFile.addEventListener('click', () => scanAddInput.click());
    scanAddInput.addEventListener('change', () => { handleScanFiles(scanAddInput.files, true); scanAddInput.value = ''; });

    // Statement import: track last parsed result for recording history
    let lastStatementParse = null;

    async function handleScanFiles(files, append = false) {
        if (!files.length) return;

        // Separate CSV files from image/PDF files
        const csvFiles = [];
        const otherFiles = [];
        for (const f of files) {
            if (f.name.toLowerCase().endsWith('.csv')) {
                csvFiles.push(f);
            } else {
                otherFiles.push(f);
            }
        }

        // Process CSV files through smart statement parser
        if (csvFiles.length > 0) {
            await handleCsvStatementFiles(csvFiles, append);
            // If there are also non-CSV files, process those with AI (appending)
            if (otherFiles.length > 0) {
                await handleAiScanFiles(otherFiles, true);
            }
        } else {
            // All non-CSV: use existing AI analysis
            await handleAiScanFiles(otherFiles, append);
        }
    }

    async function handleCsvStatementFiles(csvFiles, append = false) {
        scanStatus.classList.remove('hidden');
        scanStatusText.textContent = `${csvFiles.length}件のCSV明細を解析中...`;

        try {
            let allEntries = [];
            for (const csvFile of csvFiles) {
                const formData = new FormData();
                formData.append('file', csvFile);

                const result = await fetchAPI('/api/statement/parse', 'POST', formData);
                if (result.error) throw new Error(result.error);

                if (result.detected) {
                    // Smart parse succeeded — show source banner if first time
                    lastStatementParse = result;
                    showSourceBanner(result);
                    allEntries.push(...result.entries);

                    // Show duplicate file warning
                    if (result.duplicate_warning) {
                        showToast(`⚠️ ${result.duplicate_warning}`, true);
                    }

                    // AI prediction for entries missing debit/credit accounts
                    const needsPrediction = result.entries.filter(
                        e => !e.debit_account || !e.credit_account
                    );
                    if (needsPrediction.length > 0) {
                        const apiKey = localStorage.getItem('gemini_api_key');
                        if (apiKey) {
                            scanStatusText.textContent = `AI科目判定中... (${needsPrediction.length}件)`;
                            try {
                                const predictData = needsPrediction.map(e => ({
                                    description: e.counterparty || e.memo || '',
                                    amount: e.amount,
                                    date: e.date,
                                    current_debit: e.debit_account || '',
                                    current_credit: e.credit_account || '',
                                }));
                                const predictions = await fetchAPI('/api/predict', 'POST', {
                                    data: predictData,
                                    gemini_api_key: apiKey,
                                });
                                // Apply predictions to entries
                                if (Array.isArray(predictions)) {
                                    let pi = 0;
                                    for (const entry of result.entries) {
                                        if (!entry.debit_account || !entry.credit_account) {
                                            if (pi < predictions.length) {
                                                const pred = predictions[pi];
                                                if (!entry.debit_account && pred.debit_account) {
                                                    entry.debit_account = pred.debit_account;
                                                }
                                                if (!entry.credit_account && pred.credit_account) {
                                                    entry.credit_account = pred.credit_account;
                                                }
                                                if (pred.tax_classification) {
                                                    entry.tax_classification = pred.tax_classification;
                                                }
                                                pi++;
                                            }
                                        }
                                    }
                                }
                            } catch (predErr) {
                                console.warn('AI prediction failed, entries will need manual input:', predErr);
                            }
                        }
                    }

                    showToast(`${result.source_name}の明細を${result.total_rows}件読み取りました`);
                } else {
                    // Format not detected — fall back to AI analysis
                    const apiKey = localStorage.getItem('gemini_api_key');
                    if (!apiKey) { showToast('設定画面でAPIキーを設定してください', true); openSettings(); return; }
                    scanStatusText.textContent = `CSV形式を判定できません。AIで解析中...`;
                    const formData2 = new FormData();
                    formData2.append('files', csvFile);
                    formData2.append('gemini_api_key', apiKey);
                    if (accessToken) formData2.append('access_token', accessToken);
                    const data = await fetchAPI('/api/analyze', 'POST', formData2);
                    if (data.error) throw new Error(data.error);
                    allEntries.push(...data);
                }
            }

            if (allEntries.length > 0) {
                scanResults = append ? [...scanResults, ...allEntries] : allEntries;
                renderScanResults();
                scanResultsCard.classList.remove('hidden');
            }
            scanStatus.classList.add('hidden');
        } catch (err) {
            scanStatus.classList.add('hidden');
            showToast('CSV解析エラー: ' + err.message, true);
        }
    }

    async function handleAiScanFiles(files, append = false) {
        if (!files.length) return;
        const apiKey = localStorage.getItem('gemini_api_key');
        if (!apiKey) { showToast('設定画面でAPIキーを設定してください', true); openSettings(); return; }

        scanStatus.classList.remove('hidden');
        const BATCH_SIZE = 2;
        const fileArr = Array.from(files);
        const allResults = [];
        let errorCount = 0;

        for (let i = 0; i < fileArr.length; i += BATCH_SIZE) {
            const batch = fileArr.slice(i, i + BATCH_SIZE);
            const batchNum = Math.floor(i / BATCH_SIZE) + 1;
            scanStatusText.textContent = `AIで解析中... (${Math.min(i + BATCH_SIZE, fileArr.length)}/${fileArr.length}件)`;

            const formData = new FormData();
            for (let f of batch) formData.append('files', f);
            formData.append('gemini_api_key', apiKey);
            if (accessToken) formData.append('access_token', accessToken);

            try {
                const data = await fetchAPI('/api/analyze', 'POST', formData);
                if (data.error) throw new Error(data.error);
                allResults.push(...data);
            } catch (err) {
                console.error(`Batch ${batchNum} error:`, err);
                errorCount += batch.length;
            }
        }

        if (allResults.length > 0) {
            scanResults = append ? [...scanResults, ...allResults] : allResults;
            renderScanResults();
            scanResultsCard.classList.remove('hidden');
        }
        scanStatus.classList.add('hidden');

        if (errorCount > 0 && allResults.length > 0) {
            showToast(`${allResults.length}件解析成功、${errorCount}件失敗`, true);
        } else if (errorCount > 0) {
            showToast(`解析エラー: ${errorCount}件が失敗しました`, true);
        }
    }

    function showSourceBanner(result) {
        const banner = document.getElementById('import-source-banner');
        if (!banner) return;

        // Show banner only for detected formats without saved mapping
        if (!result.detected || (result.saved_debit || result.saved_credit)) {
            banner.classList.add('hidden');
            return;
        }

        const nameEl = document.getElementById('import-source-name');
        const debitEl = document.getElementById('import-source-debit');
        const creditEl = document.getElementById('import-source-credit');
        if (nameEl) nameEl.textContent = result.source_name;
        if (debitEl) debitEl.value = result.default_debit || '';
        if (creditEl) creditEl.value = result.default_credit || '';

        banner.classList.remove('hidden');

        // Apply button handler
        const applyBtn = document.getElementById('import-source-apply');
        if (applyBtn) {
            applyBtn.onclick = async () => {
                const debit = debitEl ? debitEl.value.trim() : '';
                const credit = creditEl ? creditEl.value.trim() : '';
                const saveCheck = document.getElementById('import-source-save');
                const shouldSave = saveCheck ? saveCheck.checked : true;

                // Update current entries
                for (const entry of scanResults) {
                    if (entry.source === 'csv_import') {
                        if (result.source_type === 'card' && credit) {
                            entry.credit_account = credit;
                        } else if (result.source_type === 'bank') {
                            // For bank: debit field is the bank account name
                            // Deposits: debit=bank account, Withdrawals: credit=bank account
                            // This is already handled by the parser
                        }
                    }
                }
                renderScanResults();

                // Save mapping if checkbox checked
                if (shouldSave && result.source_name) {
                    try {
                        await fetchAPI('/api/statement/sources', 'POST', {
                            source_name: result.source_name,
                            default_debit: debit,
                            default_credit: credit,
                        });
                        showToast(`${result.source_name}の設定を保存しました`);
                    } catch (e) {
                        console.warn('Failed to save source mapping:', e);
                    }
                }
                banner.classList.add('hidden');
            };
        }
    }

    const currentYear = new Date().getFullYear();

    function isPriorYear(dateStr) {
        if (!dateStr) return false;
        const y = parseInt(dateStr.substring(0, 4));
        const fy = getSelectedFiscalYear();
        return y < fy;
    }

    function isOutsideFiscalYear(dateStr) {
        if (!dateStr) return false;
        const y = parseInt(dateStr.substring(0, 4));
        const fy = getSelectedFiscalYear();
        return y !== fy;
    }

    // 不課税にすべき勘定科目（事業主勘定・資産間振替など）
    const NON_TAXABLE_ACCOUNTS = ['事業主貸', '事業主借', '元入金', '普通預金', '現金', '小口現金', '定期預金', '受取手形', '売掛金', '買掛金', '支払手形', '未払金', '前払金', '前受金', '仮払金', '仮受金', '貸付金', '借入金', '預り金', '立替金'];

    function inferTaxClassification(debitName, creditName) {
        // 両方が不課税対象科目 → 不課税（資産間振替・事業主勘定）
        const debitNonTax = NON_TAXABLE_ACCOUNTS.includes(debitName);
        const creditNonTax = NON_TAXABLE_ACCOUNTS.includes(creditName);
        if (debitNonTax && creditNonTax) return '不課税';

        // 事業主貸・事業主借が片方にある → 不課税
        if (['事業主貸', '事業主借', '元入金'].includes(debitName) ||
            ['事業主貸', '事業主借', '元入金'].includes(creditName)) {
            return '不課税';
        }

        // 勘定科目マスタの tax_default を参照
        // 費用科目（借方）の tax_default を優先、次に収益科目（貸方）
        const debitAcct = accounts.find(a => a.name === debitName);
        const creditAcct = accounts.find(a => a.name === creditName);

        // 借方が費用科目の場合はその tax_default
        if (debitAcct && debitAcct.account_type === '費用') return debitAcct.tax_default;
        // 貸方が収益科目の場合はその tax_default
        if (creditAcct && creditAcct.account_type === '収益') return creditAcct.tax_default;
        // それ以外：借方の tax_default → 貸方の tax_default
        if (debitAcct && debitAcct.tax_default) return debitAcct.tax_default;
        if (creditAcct && creditAcct.tax_default) return creditAcct.tax_default;

        return null; // 判定できない場合は変更しない
    }

    function renderScanResults() {
        let hasDup = false;
        let hasPriorYear = false;
        const fy = getSelectedFiscalYear();
        scanTbody.innerHTML = scanResults.map((item, i) => {
            if (item.is_duplicate) hasDup = true;
            const priorYear = isPriorYear(item.date);
            const outsideFY = isOutsideFiscalYear(item.date);
            if (priorYear) hasPriorYear = true;
            return `
            <tr class="${item.is_duplicate ? 'row-duplicate' : ''} ${outsideFY ? 'row-prior-year' : ''}">
                <td><input type="date" value="${item.date || ''}" min="${fy}-01-01" max="${fy}-12-31" data-i="${i}" data-k="date" class="scan-input ${outsideFY ? 'input-prior-year' : ''}"></td>
                <td><input type="text" value="${item.debit_account || ''}" list="account-list" data-i="${i}" data-k="debit_account" class="scan-input"></td>
                <td><input type="text" value="${item.credit_account || ''}" list="account-list" data-i="${i}" data-k="credit_account" class="scan-input"></td>
                <td><input type="number" value="${item.amount || 0}" data-i="${i}" data-k="amount" class="scan-input" style="width:100px;"></td>
                <td>
                    <select data-i="${i}" data-k="tax_classification" class="scan-input">
                        <option value="10%" ${(item.tax_classification === '10%') ? 'selected' : ''}>10%</option>
                        <option value="8%" ${(item.tax_classification === '8%') ? 'selected' : ''}>8%</option>
                        <option value="非課税" ${(item.tax_classification === '非課税') ? 'selected' : ''}>非課税</option>
                        <option value="不課税" ${(item.tax_classification === '不課税') ? 'selected' : ''}>不課税</option>
                    </select>
                </td>
                <td><input type="text" value="${item.counterparty || ''}" data-i="${i}" data-k="counterparty" class="scan-input"></td>
                <td><input type="text" value="${item.memo || ''}" data-i="${i}" data-k="memo" class="scan-input"></td>
                <td><button class="btn-icon scan-delete" data-i="${i}" title="削除">×</button></td>
            </tr>`;
        }).join('');

        scanDupAlert.classList.toggle('hidden', !hasDup);

        // Prior year alert
        let priorAlert = document.getElementById('scan-prior-year-alert');
        if (!priorAlert) {
            priorAlert = document.createElement('div');
            priorAlert.id = 'scan-prior-year-alert';
            priorAlert.className = 'alert alert-warning';
            priorAlert.innerHTML = '⚠️ 前年以前の日付が含まれています。登録時に当年1月1日の仕訳として登録され、摘要に実際の日付が記録されます。';
            scanDupAlert.parentNode.insertBefore(priorAlert, scanDupAlert.nextSibling);
        }
        priorAlert.classList.toggle('hidden', !hasPriorYear);

        scanTbody.querySelectorAll('.scan-input').forEach(el => {
            el.addEventListener('change', (e) => {
                const idx = parseInt(e.target.dataset.i);
                const key = e.target.dataset.k;
                scanResults[idx][key] = e.target.value;

                // Auto-update tax classification when account changes
                if (key === 'debit_account' || key === 'credit_account') {
                    const row = scanResults[idx];
                    const newTax = inferTaxClassification(row.debit_account, row.credit_account);
                    if (newTax && newTax !== row.tax_classification) {
                        row.tax_classification = newTax;
                        // Update the select element in the same row
                        const taxSelect = e.target.closest('tr').querySelector('select[data-k="tax_classification"]');
                        if (taxSelect) taxSelect.value = newTax;
                    }
                }
            });
        });
        scanTbody.querySelectorAll('.scan-delete').forEach(el => {
            el.addEventListener('click', (e) => {
                scanResults.splice(parseInt(e.target.dataset.i), 1);
                renderScanResults();
            });
        });

        // 読み取り件数・合計金額サマリ表示
        let scanCountEl = document.getElementById('scan-count-summary');
        if (!scanCountEl) {
            scanCountEl = document.createElement('div');
            scanCountEl.id = 'scan-count-summary';
            scanCountEl.style.cssText = 'padding:0.5rem 0;font-size:0.8125rem;color:var(--text-secondary);text-align:right;';
            const tableWrap = scanTbody.closest('.table-wrap');
            tableWrap.parentNode.insertBefore(scanCountEl, tableWrap.nextSibling);
        }
        const totalAmount = scanResults.reduce((sum, r) => sum + (parseInt(r.amount) || 0), 0);
        scanCountEl.textContent = `${scanResults.length}件の仕訳 ― 合計金額: ${totalAmount.toLocaleString()}円`;
    }

    scanSaveBtn.addEventListener('click', async () => {
        const valid = scanResults.filter(r => parseInt(r.amount) > 0);
        if (!valid.length) { showToast('保存するデータがありません', true); return; }

        // Check for entries outside fiscal year
        const fy = getSelectedFiscalYear();
        const outsideEntries = valid.filter(r => isOutsideFiscalYear(r.date));
        let entriesToSave = valid;
        if (outsideEntries.length > 0) {
            const insideCount = valid.length - outsideEntries.length;
            if (!confirm(`${outsideEntries.length}件の仕訳が${fy}年度の範囲外です。\n範囲外の仕訳は${fy}年1月1日に変更して登録します。よろしいですか？\n（キャンセルを押すと範囲外の${outsideEntries.length}件のみスキップし、残り${insideCount}件を登録します）`)) {
                // キャンセル → 年度外の仕訳だけ除外、年度内は登録続行
                entriesToSave = valid.filter(r => !isOutsideFiscalYear(r.date));
                if (!entriesToSave.length) {
                    showToast('登録対象の仕訳がありません', true);
                    return;
                }
                showToast(`年度外の${outsideEntries.length}件をスキップし、${entriesToSave.length}件を登録します`);
            } else {
                // OK → 年度外の仕訳の日付を変換＋摘要に元日付を記録
                outsideEntries.forEach(r => {
                    const originalDate = r.date;
                    r.memo = r.memo ? `[実際の日付: ${originalDate}] ${r.memo}` : `[実際の日付: ${originalDate}]`;
                    r.date = `${fy}-01-01`;
                });
            }
        }

        // Determine source: csv_import for statement-parsed entries, ai_receipt for others
        const hasCsvEntries = entriesToSave.some(r => r.source === 'csv_import');
        const entries = entriesToSave.map(r => ({
            entry_date: r.date,
            debit_account: r.debit_account,
            credit_account: r.credit_account,
            amount: parseInt(r.amount),
            tax_classification: r.tax_classification || '10%',
            counterparty: r.counterparty || '',
            memo: r.memo || '',
            evidence_url: r.evidence_url || '',
            source: r.source || 'ai_receipt',
        }));

        try {
            scanSaveBtn.disabled = true;
            scanSaveBtn.textContent = '保存中...';
            const res = await fetchAPI('/api/journal', 'POST', { entries });
            if (res.status === 'success') {
                showToast(`${res.created}件の仕訳を登録しました`);
                // Record import history for CSV statement imports
                if (hasCsvEntries && lastStatementParse) {
                    const dates = entriesToSave.filter(r => r.source === 'csv_import' && r.date).map(r => r.date).sort();
                    fetchAPI('/api/statement/history', 'POST', {
                        filename: lastStatementParse.filename || '',
                        file_hash: lastStatementParse.file_hash || '',
                        source_name: lastStatementParse.source_name || '',
                        row_count: lastStatementParse.total_rows || 0,
                        imported_count: entriesToSave.filter(r => r.source === 'csv_import').length,
                        date_range_start: dates[0] || '',
                        date_range_end: dates[dates.length - 1] || '',
                    }).catch(() => {});
                    lastStatementParse = null;
                }
                // Hide source banner
                const banner = document.getElementById('import-source-banner');
                if (banner) banner.classList.add('hidden');
                // Move inbox files to processed (if any)
                if (driveFileIdsInScan.length && accessToken) {
                    fetchAPI('/api/drive/inbox/move', 'POST', {
                        access_token: accessToken,
                        file_ids: driveFileIdsInScan,
                        categories: driveFileCategoryMap
                    }).then(() => {
                        driveFileIdsInScan = [];
                        driveFileCategoryMap = {};
                    }).catch(() => {});
                }
                scanResults = [];
                scanTbody.innerHTML = '';
                scanResultsCard.classList.add('hidden');
                loadRecentEntries();
            } else {
                showToast('登録に失敗しました', true);
            }
        } catch (err) {
            showToast('通信エラー', true);
        } finally {
            scanSaveBtn.disabled = false;
            scanSaveBtn.textContent = '仕訳を登録';
        }
    });

    // Google Drive inbox scan
    const drivePickBtn = document.getElementById('drive-pick-btn');
    let driveFileIdsInScan = []; // track inbox file IDs for move after save
    let driveFileCategoryMap = {}; // track file_id -> evidence_category for folder routing

    drivePickBtn.addEventListener('click', async () => {
        if (!accessToken) {
            showToast('Googleにログインしてください', true);
            return;
        }
        const apiKey = localStorage.getItem('gemini_api_key');
        if (!apiKey) { showToast('設定画面でAPIキーを設定してください', true); openSettings(); return; }

        // Step 1: List inbox files
        scanStatus.classList.remove('hidden');
        scanStatusText.textContent = 'ドライブのinboxフォルダを確認中...';

        try {
            const listData = await fetchAPI('/api/drive/inbox', 'POST', { access_token: accessToken });
            if (listData.error) throw new Error(listData.error);

            const files = listData.files || [];
            if (files.length === 0) {
                scanStatus.classList.add('hidden');
                showToast('inboxに未処理のファイルはありません');
                return;
            }

            // Step 2: Analyze all inbox files
            scanStatusText.textContent = `inboxの${files.length}件をAIで読み取り中...`;
            const fileIds = files.map(f => f.id);

            const results = await fetchAPI('/api/drive/inbox/analyze', 'POST', {
                access_token: accessToken,
                gemini_api_key: apiKey,
                file_ids: fileIds
            });
            if (results.error) throw new Error(results.error);

            // Track file IDs for moving after save + build category map
            driveFileIdsInScan = [...new Set(results.map(r => r.drive_file_id).filter(Boolean))];
            // Build file_id -> category mapping from AI analysis results
            for (const r of results) {
                if (r.drive_file_id && r.evidence_category) {
                    driveFileCategoryMap[r.drive_file_id] = r.evidence_category;
                }
            }

            scanResults = scanResults.length ? [...scanResults, ...results] : results;
            renderScanResults();
            scanStatus.classList.add('hidden');
            scanResultsCard.classList.remove('hidden');
            showToast(`${files.length}件の証憑を読み取りました。内容を確認して登録してください`);
        } catch (err) {
            scanStatus.classList.add('hidden');
            showToast('読み取りエラー: ' + err.message, true);
        }
    });

    // ============================================================
    //  Section 9: View 3 — 仕訳帳 (Journal Book)
    // ============================================================
    const jbStartInput = document.getElementById('jb-start');
    const jbEndInput = document.getElementById('jb-end');
    const jbTbody = document.getElementById('jb-tbody');
    const jbPagination = document.getElementById('jb-pagination');

    // Period nav for journal book (same pattern as ledger)
    const jbPeriodMode = document.getElementById('jb-period-mode');
    const jbYearSelect = document.getElementById('jb-year-select');
    const jbMonthSelect = document.getElementById('jb-month-select');
    const jbPeriodPrev = document.getElementById('jb-period-prev');
    const jbPeriodNext = document.getElementById('jb-period-next');
    const jbPeriodRange = document.getElementById('jb-period-range');

    // Advanced search modal
    const jbAdvBtn = document.getElementById('jb-advanced-btn');
    const jbSearchModal = document.getElementById('jb-search-modal');
    const jbSearchClose = document.getElementById('jb-search-close');
    const advStartInput = document.getElementById('adv-start');
    const advEndInput = document.getElementById('adv-end');
    const advAccountSelect = document.getElementById('adv-account');
    const advAmountMin = document.getElementById('adv-amount-min');
    const advAmountMax = document.getElementById('adv-amount-max');
    const advCounterparty = document.getElementById('adv-counterparty');
    const advMemo = document.getElementById('adv-memo');
    const advClearBtn = document.getElementById('adv-clear');
    const advSearchBtn = document.getElementById('adv-search');
    const jbActiveFilters = document.getElementById('jb-active-filters');

    let jbAdvancedFilters = {}; // stores current advanced filter state

    let jbPage = 1;
    const JB_PER_PAGE = 20;
    let jbCurrentPeriodMode = 'month';

    // Populate year/month selects (default to global fiscal year)
    {
        const fy = getSelectedFiscalYear();
        for (let y = thisYear + 1; y >= thisYear - 5; y--) {
            const opt = document.createElement('option');
            opt.value = y; opt.textContent = y + '年';
            if (y === fy) opt.selected = true;
            jbYearSelect.appendChild(opt);
        }
    }
    for (let m = 1; m <= 12; m++) {
        const opt = document.createElement('option');
        opt.value = m; opt.textContent = m + '月';
        if (m === thisMonth) opt.selected = true;
        jbMonthSelect.appendChild(opt);
    }

    function applyJBPeriod() {
        const y = parseInt(jbYearSelect.value) || getSelectedFiscalYear();
        const m = parseInt(jbMonthSelect.value) || thisMonth;
        const now = new Date();
        switch (jbCurrentPeriodMode) {
            case 'ytd':
                jbStartInput.value = `${y}-01-01`;
                jbEndInput.value = (y === now.getFullYear()) ? todayStr() : `${y}-12-31`;
                break;
            case 'year':
                jbStartInput.value = `${y}-01-01`;
                jbEndInput.value = `${y}-12-31`;
                break;
            case 'month': {
                const mStr = String(m).padStart(2, '0');
                const lastDay = new Date(y, m, 0).getDate();
                jbStartInput.value = `${y}-${mStr}-01`;
                jbEndInput.value = `${y}-${mStr}-${String(lastDay).padStart(2, '0')}`;
                break;
            }
        }
        jbPeriodRange.textContent = `${jbStartInput.value} 〜 ${jbEndInput.value}`;
        jbMonthSelect.style.display = (jbCurrentPeriodMode === 'month') ? '' : 'none';
        jbPeriodPrev.style.display = (jbCurrentPeriodMode === 'month') ? '' : 'none';
        jbPeriodNext.style.display = (jbCurrentPeriodMode === 'month') ? '' : 'none';
    }

    function moveJBPeriod(delta) {
        if (jbCurrentPeriodMode !== 'month') {
            jbYearSelect.value = (parseInt(jbYearSelect.value) || thisYear) + delta;
        } else {
            let m = (parseInt(jbMonthSelect.value) || thisMonth) + delta;
            let y = parseInt(jbYearSelect.value) || thisYear;
            if (m < 1) { m = 12; y--; }
            if (m > 12) { m = 1; y++; }
            jbYearSelect.value = y;
            jbMonthSelect.value = m;
        }
        applyJBPeriod();
        jbPage = 1;
        loadJournalBook();
    }

    applyJBPeriod();

    jbPeriodMode.addEventListener('change', () => {
        jbCurrentPeriodMode = jbPeriodMode.value;
        applyJBPeriod();
        jbPage = 1;
        loadJournalBook();
    });
    jbYearSelect.addEventListener('change', () => { applyJBPeriod(); jbPage = 1; loadJournalBook(); });
    jbMonthSelect.addEventListener('change', () => { applyJBPeriod(); jbPage = 1; loadJournalBook(); });
    jbPeriodPrev.addEventListener('click', () => moveJBPeriod(-1));
    jbPeriodNext.addEventListener('click', () => moveJBPeriod(1));

    // --- Advanced Search Modal ---
    jbAdvBtn.addEventListener('click', () => {
        // Pre-fill modal with current period
        advStartInput.value = jbStartInput.value;
        advEndInput.value = jbEndInput.value;
        // Populate account options from master
        advAccountSelect.innerHTML = '<option value="">すべて</option>';
        accounts.forEach(a => {
            advAccountSelect.innerHTML += `<option value="${a.id}">${a.code} ${a.name}</option>`;
        });
        // Restore last advanced filters
        if (jbAdvancedFilters.account_id) advAccountSelect.value = jbAdvancedFilters.account_id;
        if (jbAdvancedFilters.amount_min) advAmountMin.value = jbAdvancedFilters.amount_min;
        if (jbAdvancedFilters.amount_max) advAmountMax.value = jbAdvancedFilters.amount_max;
        if (jbAdvancedFilters.counterparty) advCounterparty.value = jbAdvancedFilters.counterparty;
        if (jbAdvancedFilters.memo) advMemo.value = jbAdvancedFilters.memo;
        jbSearchModal.classList.remove('hidden');
    });

    jbSearchClose.addEventListener('click', () => jbSearchModal.classList.add('hidden'));
    jbSearchModal.addEventListener('click', (e) => {
        if (e.target === jbSearchModal) jbSearchModal.classList.add('hidden');
    });

    advClearBtn.addEventListener('click', () => {
        advStartInput.value = jbStartInput.value;
        advEndInput.value = jbEndInput.value;
        advAccountSelect.value = '';
        advAmountMin.value = '';
        advAmountMax.value = '';
        advCounterparty.value = '';
        advMemo.value = '';
    });

    advSearchBtn.addEventListener('click', () => {
        // Override period with modal dates
        if (advStartInput.value) jbStartInput.value = advStartInput.value;
        if (advEndInput.value) jbEndInput.value = advEndInput.value;
        jbPeriodRange.textContent = `${jbStartInput.value} 〜 ${jbEndInput.value}`;

        jbAdvancedFilters = {};
        if (advAccountSelect.value) jbAdvancedFilters.account_id = advAccountSelect.value;
        if (advAmountMin.value) jbAdvancedFilters.amount_min = advAmountMin.value;
        if (advAmountMax.value) jbAdvancedFilters.amount_max = advAmountMax.value;
        if (advCounterparty.value.trim()) jbAdvancedFilters.counterparty = advCounterparty.value.trim();
        if (advMemo.value.trim()) jbAdvancedFilters.memo = advMemo.value.trim();

        updateActiveFiltersDisplay();
        jbSearchModal.classList.add('hidden');
        jbPage = 1;
        loadJournalBook();
    });

    function updateActiveFiltersDisplay() {
        const tags = [];
        if (jbAdvancedFilters.account_id) {
            const acc = accounts.find(a => String(a.id) === String(jbAdvancedFilters.account_id));
            tags.push(`科目: ${acc ? acc.name : jbAdvancedFilters.account_id}`);
        }
        if (jbAdvancedFilters.amount_min || jbAdvancedFilters.amount_max) {
            const min = jbAdvancedFilters.amount_min || '0';
            const max = jbAdvancedFilters.amount_max || '∞';
            tags.push(`金額: ${Number(min).toLocaleString()}〜${max === '∞' ? '∞' : Number(max).toLocaleString()}`);
        }
        if (jbAdvancedFilters.counterparty) tags.push(`取引先: ${jbAdvancedFilters.counterparty}`);
        if (jbAdvancedFilters.memo) tags.push(`摘要: ${jbAdvancedFilters.memo}`);

        if (tags.length) {
            jbActiveFilters.innerHTML = tags.map(t => `<span class="filter-tag">${t}</span>`).join('') +
                `<button class="btn-link filter-clear-all" id="jb-clear-all-filters">条件クリア</button>`;
            jbActiveFilters.classList.remove('hidden');
            jbAdvBtn.textContent = `詳細検索 (${tags.length})`;
            document.getElementById('jb-clear-all-filters').addEventListener('click', () => {
                jbAdvancedFilters = {};
                updateActiveFiltersDisplay();
                applyJBPeriod();
                jbPage = 1;
                loadJournalBook();
            });
        } else {
            jbActiveFilters.classList.add('hidden');
            jbActiveFilters.innerHTML = '';
            jbAdvBtn.textContent = '詳細検索';
        }
    }

    async function loadJournalBook() {
        const params = new URLSearchParams();
        if (jbStartInput.value) params.set('start_date', jbStartInput.value);
        if (jbEndInput.value) params.set('end_date', jbEndInput.value);
        if (jbAdvancedFilters.account_id) params.set('account_id', jbAdvancedFilters.account_id);
        if (jbAdvancedFilters.counterparty) params.set('counterparty', jbAdvancedFilters.counterparty);
        if (jbAdvancedFilters.memo) params.set('memo', jbAdvancedFilters.memo);
        if (jbAdvancedFilters.amount_min) params.set('amount_min', jbAdvancedFilters.amount_min);
        if (jbAdvancedFilters.amount_max) params.set('amount_max', jbAdvancedFilters.amount_max);
        params.set('page', jbPage);
        params.set('per_page', JB_PER_PAGE);

        try {
            const data = await fetchAPI('/api/journal?' + params.toString());
            renderJournalBook(data);
        } catch (err) {
            showToast('仕訳帳の読み込みに失敗しました', true);
        }
    }

    let jbEntriesCache = [];

    const jbSelectAll = document.getElementById('jb-select-all');
    const jbBulkDeleteBtn = document.getElementById('jb-bulk-delete-btn');

    function updateBulkDeleteBtn() {
        const checked = jbTbody.querySelectorAll('.jb-checkbox:checked');
        if (checked.length > 0) {
            jbBulkDeleteBtn.classList.remove('hidden');
            jbBulkDeleteBtn.textContent = `一括削除 (${checked.length}件)`;
        } else {
            jbBulkDeleteBtn.classList.add('hidden');
        }
    }

    jbSelectAll.addEventListener('change', () => {
        const checked = jbSelectAll.checked;
        jbTbody.querySelectorAll('.jb-checkbox').forEach(cb => { cb.checked = checked; });
        updateBulkDeleteBtn();
    });

    jbBulkDeleteBtn.addEventListener('click', async () => {
        const checkedBoxes = jbTbody.querySelectorAll('.jb-checkbox:checked');
        const ids = Array.from(checkedBoxes).map(cb => parseInt(cb.dataset.id));
        if (!ids.length) return;
        if (!confirm(`${ids.length}件の仕訳を削除しますか？この操作は元に戻せません。`)) return;
        try {
            jbBulkDeleteBtn.disabled = true;
            jbBulkDeleteBtn.textContent = '削除中...';
            const res = await fetchAPI('/api/journal/bulk-delete', 'POST', { ids });
            if (res.status === 'success') {
                showToast(`${res.deleted}件の仕訳を削除しました`);
                jbSelectAll.checked = false;
                loadJournalBook();
                loadRecentEntries();
            } else {
                showToast('削除に失敗しました', true);
            }
        } catch (err) {
            showToast('通信エラー', true);
        } finally {
            jbBulkDeleteBtn.disabled = false;
            updateBulkDeleteBtn();
        }
    });

    function renderJournalBook(data) {
        const entries = data.entries || [];
        jbEntriesCache = entries;
        jbSelectAll.checked = false;
        jbBulkDeleteBtn.classList.add('hidden');
        jbTbody.innerHTML = entries.map(e => `
            <tr data-id="${e.id}" style="cursor:pointer;">
                <td><input type="checkbox" class="jb-checkbox" data-id="${e.id}"></td>
                <td>${e.entry_date || ''}</td>
                <td>${e.debit_account || ''}</td>
                <td>${e.credit_account || ''}</td>
                <td class="text-right">${fmt(e.amount)}</td>
                <td>${e.tax_classification || ''}</td>
                <td>${e.counterparty || ''}</td>
                <td>${e.memo || ''}</td>
                <td>${e.evidence_url ? `<a href="${e.evidence_url}" target="_blank" class="evidence-link" title="証憑を表示">📎</a>` : ''}</td>
                <td>
                    <button class="btn-icon jb-edit" data-id="${e.id}" title="編集">✎</button>
                    <button class="btn-icon jb-delete" data-id="${e.id}" title="削除">×</button>
                </td>
            </tr>
        `).join('');

        if (!entries.length) {
            jbTbody.innerHTML = '<tr><td colspan="10" style="text-align:center;color:var(--text-dim);padding:2rem;">仕訳データがありません</td></tr>';
        }

        const total = data.total || 0;
        const totalPages = Math.ceil(total / JB_PER_PAGE);
        let pgHtml = '';
        if (totalPages > 1) {
            if (jbPage > 1) pgHtml += `<button class="btn btn-sm btn-ghost jb-page" data-p="${jbPage - 1}">← 前</button>`;
            pgHtml += `<span style="padding:0.5rem;">${jbPage} / ${totalPages} (${total}件)</span>`;
            if (jbPage < totalPages) pgHtml += `<button class="btn btn-sm btn-ghost jb-page" data-p="${jbPage + 1}">次 →</button>`;
        }
        jbPagination.innerHTML = pgHtml;

        // Checkbox change handlers
        jbTbody.querySelectorAll('.jb-checkbox').forEach(cb => {
            cb.addEventListener('change', () => {
                updateBulkDeleteBtn();
                // Update select-all state
                const allBoxes = jbTbody.querySelectorAll('.jb-checkbox');
                const allChecked = jbTbody.querySelectorAll('.jb-checkbox:checked');
                jbSelectAll.checked = allBoxes.length > 0 && allBoxes.length === allChecked.length;
                jbSelectAll.indeterminate = allChecked.length > 0 && allChecked.length < allBoxes.length;
            });
        });

        jbTbody.querySelectorAll('.jb-delete').forEach(btn => {
            btn.addEventListener('click', async (e) => {
                const id = e.target.dataset.id;
                if (!confirm('この仕訳を削除しますか？')) return;
                try {
                    const res = await fetchAPI(`/api/journal/${id}`, 'DELETE');
                    if (res.status === 'success') {
                        showToast('削除しました');
                        loadJournalBook();
                        loadRecentEntries();
                    } else {
                        showToast('削除に失敗しました', true);
                    }
                } catch (err) {
                    showToast('通信エラー', true);
                }
            });
        });

        jbTbody.querySelectorAll('.jb-edit').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = e.target.dataset.id;
                const entry = jbEntriesCache.find(en => String(en.id) === String(id));
                if (entry) openJEDetailModal(entry, () => { loadJournalBook(); loadRecentEntries(); });
            });
        });

        // Double-click to open detail modal
        jbTbody.querySelectorAll('tr[data-id]').forEach(row => {
            row.addEventListener('dblclick', (e) => {
                if (e.target.closest('a') || e.target.closest('button') || e.target.closest('input[type=checkbox]')) return;
                const id = row.dataset.id;
                const entry = jbEntriesCache.find(en => String(en.id) === String(id));
                if (entry) openJEDetailModal(entry, () => { loadJournalBook(); loadRecentEntries(); });
            });
        });

        jbPagination.querySelectorAll('.jb-page').forEach(btn => {
            btn.addEventListener('click', (e) => {
                jbPage = parseInt(e.target.dataset.p);
                loadJournalBook();
            });
        });
    }

    function startInlineEdit(row, entryId) {
        const cells = row.querySelectorAll('td');
        const original = {
            entry_date: cells[0].textContent.trim(),
            debit_account: cells[1].textContent.trim(),
            credit_account: cells[2].textContent.trim(),
            amount: cells[3].textContent.trim().replace(/,/g, ''),
            tax_classification: cells[4].textContent.trim(),
            counterparty: cells[5].textContent.trim(),
            memo: cells[6].textContent.trim(),
        };

        row.innerHTML = `
            <td><input type="date" value="${original.entry_date}" class="edit-input" data-k="entry_date"></td>
            <td><input type="text" value="${original.debit_account}" list="account-list" class="edit-input" data-k="debit_account"></td>
            <td><input type="text" value="${original.credit_account}" list="account-list" class="edit-input" data-k="credit_account"></td>
            <td><input type="number" value="${original.amount}" class="edit-input" data-k="amount" style="width:100px;"></td>
            <td>
                <select class="edit-input" data-k="tax_classification">
                    <option value="10%" ${original.tax_classification === '10%' ? 'selected' : ''}>10%</option>
                    <option value="8%" ${original.tax_classification === '8%' ? 'selected' : ''}>8%</option>
                    <option value="非課税" ${original.tax_classification === '非課税' ? 'selected' : ''}>非課税</option>
                    <option value="不課税" ${original.tax_classification === '不課税' ? 'selected' : ''}>不課税</option>
                </select>
            </td>
            <td><input type="text" value="${original.counterparty}" class="edit-input" data-k="counterparty"></td>
            <td><input type="text" value="${original.memo}" class="edit-input" data-k="memo"></td>
            <td>
                <button class="btn-icon edit-save" title="保存">✓</button>
                <button class="btn-icon edit-cancel" title="キャンセル">✕</button>
            </td>
        `;

        row.querySelector('.edit-save').addEventListener('click', async () => {
            const updated = {};
            row.querySelectorAll('.edit-input').forEach(inp => {
                updated[inp.dataset.k] = inp.value;
            });
            updated.amount = parseInt(updated.amount) || 0;
            try {
                const res = await fetchAPI(`/api/journal/${entryId}`, 'PUT', updated);
                if (res.status === 'success') {
                    showToast('更新しました');
                    loadJournalBook();
                    loadRecentEntries();
                } else {
                    showToast('更新に失敗しました', true);
                }
            } catch (err) {
                showToast('通信エラー', true);
            }
        });

        row.querySelector('.edit-cancel').addEventListener('click', () => loadJournalBook());
    }

    // ============================================================
    //  Section 10: View 4 — 総勘定元帳 (General Ledger with Sub-tabs)
    //  Tabs: B/S(資産) / B/S(負債・純資産) / 損益計算書
    // ============================================================
    const ledgerSubNav = document.getElementById('ledger-sub-nav');
    const ledgerStartInput = document.getElementById('ledger-start');
    const ledgerEndInput = document.getElementById('ledger-end');
    const ledgerDetail = document.getElementById('ledger-detail');
    const ledgerDetailBack = document.getElementById('ledger-detail-back');
    const ledgerDetailTitle = document.getElementById('ledger-detail-title');
    const ledgerDetailContent = document.getElementById('ledger-detail-content');
    const ledgerYearSelect = document.getElementById('ledger-year-select');
    const ledgerMonthSelect = document.getElementById('ledger-month-select');
    const ledgerPeriodMode = document.getElementById('ledger-period-mode');
    const periodPrevBtn = document.getElementById('period-prev');
    const periodNextBtn = document.getElementById('period-next');
    const periodRangeDisplay = document.getElementById('period-range-display');

    const bsAssetsContent = document.getElementById('bs-assets-content');
    const bsLiabilitiesContent = document.getElementById('bs-liabilities-content');
    const plContent = document.getElementById('pl-content');

    let currentLedgerSubTab = 'bs-assets';
    let currentPeriodMode = 'month';  // 'ytd' | 'year' | 'month'

    // --- Year & Month selector population (default to global fiscal year) ---
    {
        const fy = getSelectedFiscalYear();
        for (let y = thisYear + 1; y >= thisYear - 5; y--) {
            const opt = document.createElement('option');
            opt.value = y;
            opt.textContent = y + '年';
            if (y === fy) opt.selected = true;
            ledgerYearSelect.appendChild(opt);
        }
    }
    for (let m = 1; m <= 12; m++) {
        const opt = document.createElement('option');
        opt.value = m;
        opt.textContent = m + '月';
        if (m === thisMonth) opt.selected = true;
        ledgerMonthSelect.appendChild(opt);
    }

    function getSelectedYear() { return parseInt(ledgerYearSelect.value) || getSelectedFiscalYear(); }
    function getSelectedMonth() { return parseInt(ledgerMonthSelect.value) || thisMonth; }

    function isMonthlyMode() { return currentPeriodMode === 'month'; }

    function updateMonthSelectVisibility() {
        ledgerMonthSelect.style.display = (currentPeriodMode === 'month') ? '' : 'none';
        periodPrevBtn.style.display = (currentPeriodMode === 'month') ? '' : 'none';
        periodNextBtn.style.display = (currentPeriodMode === 'month') ? '' : 'none';
    }

    function applyPeriod() {
        const y = getSelectedYear();
        const m = getSelectedMonth();
        const now = new Date();
        switch (currentPeriodMode) {
            case 'ytd':
                ledgerStartInput.value = `${y}-01-01`;
                ledgerEndInput.value = (y === now.getFullYear()) ? todayStr() : `${y}-12-31`;
                break;
            case 'year':
                ledgerStartInput.value = `${y}-01-01`;
                ledgerEndInput.value = `${y}-12-31`;
                break;
            case 'month': {
                const mStr = String(m).padStart(2, '0');
                const lastDay = new Date(y, m, 0).getDate();
                ledgerStartInput.value = `${y}-${mStr}-01`;
                ledgerEndInput.value = `${y}-${mStr}-${String(lastDay).padStart(2, '0')}`;
                break;
            }
        }
        // Update display
        periodRangeDisplay.textContent = `${ledgerStartInput.value} 〜 ${ledgerEndInput.value}`;
        updateMonthSelectVisibility();
    }

    function movePeriod(delta) {
        if (currentPeriodMode !== 'month') {
            // Year-based: move year
            ledgerYearSelect.value = getSelectedYear() + delta;
        } else {
            // Month-based: move month
            let m = getSelectedMonth() + delta;
            let y = getSelectedYear();
            if (m < 1) { m = 12; y--; }
            if (m > 12) { m = 1; y++; }
            ledgerYearSelect.value = y;
            ledgerMonthSelect.value = m;
        }
        applyPeriod();
        loadCurrentLedgerSubTab();
    }

    // Initial setup
    applyPeriod();

    // Event listeners
    ledgerPeriodMode.addEventListener('change', () => {
        currentPeriodMode = ledgerPeriodMode.value;
        applyPeriod();
        loadCurrentLedgerSubTab();
    });
    ledgerYearSelect.addEventListener('change', () => { applyPeriod(); loadCurrentLedgerSubTab(); });
    ledgerMonthSelect.addEventListener('change', () => { applyPeriod(); loadCurrentLedgerSubTab(); });
    periodPrevBtn.addEventListener('click', () => movePeriod(-1));
    periodNextBtn.addEventListener('click', () => movePeriod(1));

    // Sub-tab switching
    ledgerSubNav.addEventListener('click', (e) => {
        const btn = e.target.closest('.sub-tab-btn');
        if (!btn) return;
        const subtab = btn.dataset.subtab;
        currentLedgerSubTab = subtab;

        ledgerSubNav.querySelectorAll('.sub-tab-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');

        document.querySelectorAll('.ledger-sub-panel').forEach(p => p.classList.remove('active'));
        const panel = document.getElementById('ledger-tab-' + subtab);
        if (panel) panel.classList.add('active');

        hideLedgerDetail();
        loadCurrentLedgerSubTab();
    });

    function loadCurrentLedgerSubTab() {
        // Always reload the sub-tab list
        if (currentLedgerSubTab === 'bs-assets') loadBSAssets();
        else if (currentLedgerSubTab === 'bs-liabilities') loadBSLiabilities();
        else if (currentLedgerSubTab === 'profit-loss') loadProfitLoss();

        // If drill-down detail is open, also refresh it with new period
        if (currentDrillAccountId && !ledgerDetail.classList.contains('hidden')) {
            showAccountDetail(currentDrillAccountId);
        }
    }

    function hideLedgerDetail() {
        ledgerDetail.classList.add('hidden');
        updateHeaderFiscalYear();  // ヘッダーを「2025年度」に戻す
        // Show sub-panels again
        document.querySelectorAll('.ledger-sub-panel').forEach(p => {
            if (p.id === 'ledger-tab-' + currentLedgerSubTab) p.classList.add('active');
        });
    }

    ledgerDetailBack.addEventListener('click', hideLedgerDetail);

    // --- Account Drill-down ---
    let currentDrillAccountId = null;

    async function showAccountDetail(accountId) {
        currentDrillAccountId = accountId;
        const params = new URLSearchParams();
        if (ledgerStartInput.value) params.set('start_date', ledgerStartInput.value);
        if (ledgerEndInput.value) params.set('end_date', ledgerEndInput.value);

        try {
            const data = await fetchAPI(`/api/ledger/${accountId}?${params.toString()}`);
            const acc = data.account || {};
            ledgerDetailTitle.textContent = `${acc.code || ''} ${acc.name || ''}`;
            headerFiscalYear.textContent = `${getSelectedFiscalYear()}年度 ― ${acc.name || ''}`;

            // Hide sub-panels, show detail
            document.querySelectorAll('.ledger-sub-panel').forEach(p => p.classList.remove('active'));
            ledgerDetail.classList.remove('hidden');

            const entries = data.entries || [];
            const openBal = data.opening_balance || 0;

            let html = `<p style="margin-bottom:0.75rem;font-size:0.8125rem;color:var(--text-secondary);">期首残高: <strong>${fmt(openBal)}</strong></p>`;
            html += `<div class="table-wrap"><table class="tb-table ledger-detail-table">
                <thead><tr>
                    <th>日付</th><th>相手科目</th><th>摘要</th><th>取引先</th>
                    <th class="text-right">借方</th><th class="text-right">貸方</th>
                    <th class="text-right">差引残高</th><th style="width:70px;">操作</th>
                </tr></thead><tbody>`;

            entries.forEach(e => {
                html += `<tr data-entry-id="${e.id}" style="cursor:pointer;">
                    <td>${e.entry_date || ''}</td>
                    <td>${e.counter_account || ''}</td>
                    <td>${e.memo || ''}</td>
                    <td>${e.counterparty || ''}</td>
                    <td class="text-right">${e.debit_amount ? fmt(e.debit_amount) : ''}</td>
                    <td class="text-right">${e.credit_amount ? fmt(e.credit_amount) : ''}</td>
                    <td class="text-right" style="font-weight:600;">${fmt(e.balance)}</td>
                    <td class="action-cell">
                        <button class="btn-icon ledger-edit" data-id="${e.id}" title="編集">✎</button>
                        <button class="btn-icon ledger-delete" data-id="${e.id}" title="削除">×</button>
                    </td>
                </tr>`;
            });

            if (!entries.length) {
                html += '<tr><td colspan="8" style="text-align:center;padding:2rem;color:var(--text-dim);">該当する仕訳がありません</td></tr>';
            }

            html += '</tbody></table></div>';
            ledgerDetailContent.innerHTML = html;

            // Attach delete handlers
            ledgerDetailContent.querySelectorAll('.ledger-delete').forEach(btn => {
                btn.addEventListener('click', async () => {
                    const id = btn.dataset.id;
                    if (!confirm('この仕訳を削除しますか？')) return;
                    try {
                        const res = await fetchAPI(`/api/journal/${id}`, 'DELETE');
                        if (res.status === 'success') {
                            showToast('削除しました');
                            showAccountDetail(accountId);
                        } else {
                            showToast('削除に失敗しました', true);
                        }
                    } catch (err) {
                        showToast('通信エラー', true);
                    }
                });
            });

            // Attach edit handlers (open modal)
            ledgerDetailContent.querySelectorAll('.ledger-edit').forEach(btn => {
                btn.addEventListener('click', () => {
                    const id = btn.dataset.id;
                    const entry = entries.find(e => String(e.id) === String(id));
                    if (entry) openJEDetailModal(entry, () => showAccountDetail(accountId));
                });
            });

            // Double-click to open detail modal
            ledgerDetailContent.querySelectorAll('tr[data-entry-id]').forEach(row => {
                row.addEventListener('dblclick', (e) => {
                    if (e.target.closest('a') || e.target.closest('button')) return;
                    const id = row.dataset.entryId;
                    const entry = entries.find(en => String(en.id) === String(id));
                    if (entry) openJEDetailModal(entry, () => showAccountDetail(accountId));
                });
            });
        } catch (err) {
            showToast('元帳明細の読み込みに失敗しました', true);
        }
    }

    function startLedgerInlineEdit(row, entry, accountId) {
        const taxClass = entry.tax_classification || '10%';
        row.innerHTML = `
            <td><input type="date" value="${entry.entry_date || ''}" class="edit-input" data-k="entry_date"></td>
            <td colspan="2">
                <div style="display:flex;gap:4px;">
                    <input type="text" value="${entry.debit_account || ''}" list="account-list" class="edit-input" data-k="debit_account" placeholder="借方" style="flex:1;">
                    <input type="text" value="${entry.credit_account || ''}" list="account-list" class="edit-input" data-k="credit_account" placeholder="貸方" style="flex:1;">
                </div>
            </td>
            <td><input type="text" value="${entry.counterparty || ''}" class="edit-input" data-k="counterparty"></td>
            <td><input type="number" value="${entry.amount || 0}" class="edit-input" data-k="amount" style="width:100px;"></td>
            <td>
                <select class="edit-input" data-k="tax_classification">
                    <option value="10%" ${taxClass === '10%' ? 'selected' : ''}>10%</option>
                    <option value="8%" ${taxClass === '8%' ? 'selected' : ''}>8%</option>
                    <option value="非課税" ${taxClass === '非課税' ? 'selected' : ''}>非課税</option>
                    <option value="不課税" ${taxClass === '不課税' ? 'selected' : ''}>不課税</option>
                </select>
            </td>
            <td><input type="text" value="${entry.memo || ''}" class="edit-input" data-k="memo"></td>
            <td class="action-cell">
                <button class="btn-icon edit-save" title="保存">✓</button>
                <button class="btn-icon edit-cancel" title="キャンセル">✕</button>
            </td>
        `;

        row.querySelector('.edit-save').addEventListener('click', async () => {
            const updated = {};
            row.querySelectorAll('.edit-input').forEach(inp => {
                updated[inp.dataset.k] = inp.value;
            });
            updated.amount = parseInt(updated.amount) || 0;
            try {
                const res = await fetchAPI(`/api/journal/${entry.id}`, 'PUT', updated);
                if (res.status === 'success') {
                    showToast('更新しました');
                    showAccountDetail(accountId);
                } else {
                    showToast('更新に失敗しました', true);
                }
            } catch (err) {
                showToast('通信エラー', true);
            }
        });

        row.querySelector('.edit-cancel').addEventListener('click', () => {
            showAccountDetail(accountId);
        });
    }

    // Helper: build a clickable account table with drill-down
    function buildAccountTable(items, showSectionTitle = '') {
        const monthly = isMonthlyMode();
        let html = '';
        if (showSectionTitle) {
            html += `<div class="pl-section-title" style="margin:0.75rem 0 0.5rem;">${showSectionTitle}</div>`;
        }
        if (monthly) {
            html += `<table class="tb-table">
                <thead><tr>
                    <th>コード</th>
                    <th>勘定科目</th>
                    <th class="text-right">前月繰越</th>
                    <th class="text-right">当月借方</th>
                    <th class="text-right">当月貸方</th>
                    <th class="text-right">当月残高</th>
                </tr></thead><tbody>`;
            items.forEach(b => {
                html += `<tr class="tb-row clickable" data-account-id="${b.account_id}" style="cursor:pointer;">
                    <td>${b.code}</td>
                    <td>${b.name}</td>
                    <td class="text-right">${fmt(b.carry_forward)}</td>
                    <td class="text-right">${fmt(b.debit_total)}</td>
                    <td class="text-right">${fmt(b.credit_total)}</td>
                    <td class="text-right" style="font-weight:600;">${fmt(b.closing_balance)}</td>
                </tr>`;
            });
        } else {
            html += `<table class="tb-table">
                <thead><tr>
                    <th>コード</th>
                    <th>勘定科目</th>
                    <th class="text-right">期首残高</th>
                    <th class="text-right">借方合計</th>
                    <th class="text-right">貸方合計</th>
                    <th class="text-right">残高</th>
                </tr></thead><tbody>`;
            items.forEach(b => {
                html += `<tr class="tb-row clickable" data-account-id="${b.account_id}" style="cursor:pointer;">
                    <td>${b.code}</td>
                    <td>${b.name}</td>
                    <td class="text-right">${fmt(b.opening_balance)}</td>
                    <td class="text-right">${fmt(b.debit_total)}</td>
                    <td class="text-right">${fmt(b.credit_total)}</td>
                    <td class="text-right" style="font-weight:600;">${fmt(b.closing_balance)}</td>
                </tr>`;
            });
        }
        html += `</tbody></table>`;
        return html;
    }

    function attachDrilldown(container) {
        container.querySelectorAll('.clickable').forEach(row => {
            row.addEventListener('click', () => {
                showAccountDetail(row.dataset.accountId);
            });
        });
    }

    // --- B/S（資産） ---
    async function loadBSAssets() {
        const params = new URLSearchParams();
        if (ledgerStartInput.value) params.set('start_date', ledgerStartInput.value);
        if (ledgerEndInput.value) params.set('end_date', ledgerEndInput.value);

        try {
            const data = await fetchAPI('/api/trial-balance?' + params.toString());
            renderBSAssets(data.balances || []);
        } catch (err) {
            showToast('貸借対照表（資産）の読み込みに失敗しました', true);
        }
    }

    function renderBSAssets(balances) {
        const assets = balances.filter(b => b.account_type === '資産');
        const total = assets.reduce((s, b) => s + b.closing_balance, 0);

        let html = buildAccountTable(assets);

        html += `<div class="tb-grand-total">
            <span>資産合計: <strong>${fmt(total)}</strong></span>
        </div>`;

        bsAssetsContent.innerHTML = html;
        attachDrilldown(bsAssetsContent);
    }

    // --- B/S（負債・純資産） ---
    async function loadBSLiabilities() {
        const params = new URLSearchParams();
        if (ledgerStartInput.value) params.set('start_date', ledgerStartInput.value);
        if (ledgerEndInput.value) params.set('end_date', ledgerEndInput.value);

        try {
            const data = await fetchAPI('/api/trial-balance?' + params.toString());
            renderBSLiabilities(data.balances || []);
        } catch (err) {
            showToast('貸借対照表（負債・純資産）の読み込みに失敗しました', true);
        }
    }

    function renderBSLiabilities(balances) {
        const liabilities = balances.filter(b => b.account_type === '負債');
        const equity = balances.filter(b => b.account_type === '純資産');
        const liabilityTotal = liabilities.reduce((s, b) => s + b.closing_balance, 0);
        const equityTotal = equity.reduce((s, b) => s + b.closing_balance, 0);
        const grandTotal = liabilityTotal + equityTotal;

        let html = '';

        // 負債セクション
        html += buildAccountTable(liabilities, '負債の部');
        html += `<div class="tb-grand-total" style="margin-bottom:1rem;">
            <span>負債合計: <strong>${fmt(liabilityTotal)}</strong></span>
        </div>`;

        // 純資産セクション
        html += buildAccountTable(equity, '純資産の部');
        html += `<div class="tb-grand-total" style="margin-bottom:1rem;">
            <span>純資産合計: <strong>${fmt(equityTotal)}</strong></span>
        </div>`;

        // 合計
        html += `<div class="tb-grand-total" style="font-size:1rem;">
            <span>負債・純資産合計: <strong>${fmt(grandTotal)}</strong></span>
        </div>`;

        bsLiabilitiesContent.innerHTML = html;
        attachDrilldown(bsLiabilitiesContent);
    }

    // --- 損益計算書 ---
    async function loadProfitLoss() {
        const params = new URLSearchParams();
        if (ledgerStartInput.value) params.set('start_date', ledgerStartInput.value);
        if (ledgerEndInput.value) params.set('end_date', ledgerEndInput.value);

        try {
            const data = await fetchAPI('/api/trial-balance?' + params.toString());
            renderProfitLoss(data.balances || []);
        } catch (err) {
            showToast('損益計算書の読み込みに失敗しました', true);
        }
    }

    function renderProfitLoss(balances) {
        const revenues = balances.filter(b => b.account_type === '収益');
        const expenses = balances.filter(b => b.account_type === '費用');
        const revenueTotal = revenues.reduce((s, b) => s + b.closing_balance, 0);
        const expenseTotal = expenses.reduce((s, b) => s + b.closing_balance, 0);
        const netIncome = revenueTotal - expenseTotal;

        let html = '';

        // 収益の部
        html += buildAccountTable(revenues, '収益の部');
        html += `<div class="tb-grand-total" style="margin-bottom:1rem;">
            <span>収益合計: <strong>${fmt(revenueTotal)}</strong></span>
        </div>`;

        // 費用の部
        html += buildAccountTable(expenses, '費用の部');
        html += `<div class="tb-grand-total" style="margin-bottom:1rem;">
            <span>費用合計: <strong>${fmt(expenseTotal)}</strong></span>
        </div>`;

        // 当期純利益 / 当期純損失
        const isProfit = netIncome >= 0;
        html += `<div class="pl-net-income ${isProfit ? 'profit' : 'loss'}">
            <span>${isProfit ? '当期純利益' : '当期純損失'}</span>
            <span>${fmt(Math.abs(netIncome))}</span>
        </div>`;

        plContent.innerHTML = html;
        attachDrilldown(plContent);
    }

    // ============================================================
    //  Section 11: View 5 — 取引先 (Counterparty Management)
    // ============================================================
    const cpAddBtn = document.getElementById('cp-add-btn');
    const cpForm = document.getElementById('cp-form');
    const cpEditId = document.getElementById('cp-edit-id');
    const cpNameInput = document.getElementById('cp-name');
    const cpCodeInput = document.getElementById('cp-code');
    const cpContactInput = document.getElementById('cp-contact');
    const cpNotesInput = document.getElementById('cp-notes');
    const cpSaveBtn = document.getElementById('cp-save-btn');
    const cpCancelBtn = document.getElementById('cp-cancel-btn');
    const cpTbody = document.getElementById('cp-tbody');

    cpAddBtn.addEventListener('click', () => {
        cpEditId.value = '';
        cpNameInput.value = '';
        cpCodeInput.value = '';
        cpContactInput.value = '';
        cpNotesInput.value = '';
        cpForm.classList.remove('hidden');
        cpNameInput.focus();
    });

    cpCancelBtn.addEventListener('click', () => {
        cpForm.classList.add('hidden');
    });

    cpSaveBtn.addEventListener('click', async () => {
        const name = cpNameInput.value.trim();
        if (!name) { showToast('取引先名は必須です', true); return; }

        const payload = {
            name: name,
            code: cpCodeInput.value.trim(),
            contact_info: cpContactInput.value.trim(),
            notes: cpNotesInput.value.trim(),
        };

        const editId = cpEditId.value;
        try {
            let res;
            if (editId) {
                res = await fetchAPI(`/api/counterparties/${editId}`, 'PUT', payload);
            } else {
                res = await fetchAPI('/api/counterparties', 'POST', payload);
            }
            if (res.status === 'success') {
                showToast(editId ? '取引先を更新しました' : '取引先を登録しました');
                cpForm.classList.add('hidden');
                loadCounterpartyList();
                loadCounterparties(); // refresh datalist
            } else {
                showToast('保存に失敗しました', true);
            }
        } catch (err) {
            showToast('通信エラー', true);
        }
    });

    async function loadCounterpartyList() {
        try {
            const data = await fetchAPI('/api/counterparties/list');
            const items = data.counterparties || [];
            cpTbody.innerHTML = items.map(cp => `
                <tr>
                    <td>${cp.name || ''}</td>
                    <td>${cp.code || ''}</td>
                    <td>${cp.contact_info || ''}</td>
                    <td>${cp.notes || ''}</td>
                    <td>
                        <button class="btn-icon cp-edit" data-id="${cp.id}" data-name="${escAttr(cp.name)}" data-code="${escAttr(cp.code)}" data-contact="${escAttr(cp.contact_info)}" data-notes="${escAttr(cp.notes)}" title="編集">✎</button>
                        <button class="btn-icon jb-delete cp-delete" data-id="${cp.id}" title="削除">×</button>
                    </td>
                </tr>
            `).join('');

            if (!items.length) {
                cpTbody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:2rem;color:var(--text-dim);">取引先が登録されていません</td></tr>';
            }

            // Bind edit
            cpTbody.querySelectorAll('.cp-edit').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    const b = e.target.closest('.cp-edit');
                    cpEditId.value = b.dataset.id;
                    cpNameInput.value = b.dataset.name || '';
                    cpCodeInput.value = b.dataset.code || '';
                    cpContactInput.value = b.dataset.contact || '';
                    cpNotesInput.value = b.dataset.notes || '';
                    cpForm.classList.remove('hidden');
                    cpNameInput.focus();
                });
            });

            // Bind delete
            cpTbody.querySelectorAll('.cp-delete').forEach(btn => {
                btn.addEventListener('click', async (e) => {
                    const id = e.target.closest('.cp-delete').dataset.id;
                    if (!confirm('この取引先を削除しますか？')) return;
                    try {
                        const res = await fetchAPI(`/api/counterparties/${id}`, 'DELETE');
                        if (res.status === 'success') {
                            showToast('取引先を削除しました');
                            loadCounterpartyList();
                            loadCounterparties();
                        } else {
                            showToast('削除に失敗しました', true);
                        }
                    } catch (err) {
                        showToast('通信エラー', true);
                    }
                });
            });
        } catch (err) {
            showToast('取引先一覧の読み込みに失敗しました', true);
        }
    }

    function escAttr(s) {
        return (s || '').replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    // ============================================================
    //  Section 11b: View — 勘定科目管理 (Accounts Management)
    // ============================================================
    const accTbody = document.getElementById('acc-tbody');
    const accAddBtn = document.getElementById('acc-add-btn');
    const accCodeInput = document.getElementById('acc-code');
    const accNameInput = document.getElementById('acc-name');
    const accTypeSelect = document.getElementById('acc-type');
    const accTaxSelect = document.getElementById('acc-tax');

    async function loadAccountsList() {
        const data = await fetchAPI('/api/accounts');
        const accounts = data.accounts || [];
        accTbody.innerHTML = accounts.map(a => `
            <tr>
                <td>${a.code}</td>
                <td>${a.name}</td>
                <td>${a.account_type}</td>
                <td>${a.tax_default}</td>
                <td><button class="btn-icon acc-del" data-id="${a.id}" title="削除">🗑</button></td>
            </tr>
        `).join('');

        accTbody.querySelectorAll('.acc-del').forEach(btn => {
            btn.addEventListener('click', async () => {
                const id = btn.dataset.id;
                if (!confirm('この勘定科目を削除しますか？')) return;
                try {
                    const res = await fetchAPI(`/api/accounts/${id}`, 'DELETE');
                    if (res.status === 'success') {
                        showToast('勘定科目を削除しました');
                        loadAccountsList();
                        refreshAccountDatalist();
                    } else {
                        showToast(res.error || '削除に失敗しました', true);
                    }
                } catch (err) {
                    showToast('通信エラー', true);
                }
            });
        });
    }

    accAddBtn.addEventListener('click', async () => {
        const code = accCodeInput.value.trim();
        const name = accNameInput.value.trim();
        const account_type = accTypeSelect.value;
        const tax_default = accTaxSelect.value;
        if (!code || !name) {
            showToast('コードと科目名は必須です', true);
            return;
        }
        try {
            const res = await fetchAPI('/api/accounts', 'POST', { code, name, account_type, tax_default });
            if (res.status === 'success') {
                showToast(`勘定科目「${name}」を追加しました`);
                accCodeInput.value = '';
                accNameInput.value = '';
                loadAccountsList();
                refreshAccountDatalist();
            } else {
                showToast(res.error || '追加に失敗しました', true);
            }
        } catch (err) {
            showToast('通信エラー', true);
        }
    });

    // Refresh account datalist used in journal entry form
    function refreshAccountDatalist() {
        fetchAPI('/api/accounts').then(data => {
            const dl = document.getElementById('account-list');
            if (!dl) return;
            dl.innerHTML = (data.accounts || []).map(a => `<option value="${a.name}">`).join('');
        });
    }

    // ============================================================
    //  Section 12: View 6 — 期首残高設定 (Opening Balances)
    // ============================================================
    const obFiscalYear = document.getElementById('ob-fiscal-year');
    const obLoadBtn = document.getElementById('ob-load');
    const obSaveBtn = document.getElementById('ob-save');
    const obTbody = document.getElementById('ob-tbody');

    // Populate fiscal year options (currentYear already defined in Section 7)
    for (let y = currentYear - 2; y <= currentYear + 1; y++) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.textContent = y + '年度';
        if (y === currentYear) opt.selected = true;
        obFiscalYear.appendChild(opt);
    }

    obLoadBtn.addEventListener('click', loadOpeningBalances);
    obSaveBtn.addEventListener('click', saveOpeningBalances);

    async function loadOpeningBalances() {
        const year = obFiscalYear.value;
        try {
            const data = await fetchAPI(`/api/opening-balances?fiscal_year=${year}`);
            const balances = data.balances || [];

            // Build map of existing balances
            const balMap = {};
            balances.forEach(b => { balMap[b.account_id] = b; });

            // Render all accounts (BS accounts only: 資産, 負債, 純資産)
            const bsAccounts = accounts.filter(a => ['資産', '負債', '純資産'].includes(a.account_type));

            obTbody.innerHTML = bsAccounts.map(a => {
                const existing = balMap[a.id];
                const amount = existing ? existing.amount : 0;
                const note = existing ? (existing.note || '') : '';
                return `
                <tr>
                    <td>${a.code}</td>
                    <td>${a.name}</td>
                    <td>${a.account_type}</td>
                    <td><input type="number" class="ob-amount" data-account-id="${a.id}" value="${amount}" style="width:120px;text-align:right;"></td>
                    <td><input type="text" class="ob-note" data-account-id="${a.id}" value="${escAttr(note)}" placeholder="備考"></td>
                </tr>`;
            }).join('');

            if (!bsAccounts.length) {
                obTbody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:2rem;color:var(--text-dim);">勘定科目がありません</td></tr>';
            }
        } catch (err) {
            showToast('期首残高の読み込みに失敗しました', true);
        }
    }

    async function saveOpeningBalances() {
        const year = obFiscalYear.value;
        const balances = [];

        obTbody.querySelectorAll('.ob-amount').forEach(input => {
            const accountId = parseInt(input.dataset.accountId);
            const amount = parseInt(input.value) || 0;
            const noteInput = obTbody.querySelector(`.ob-note[data-account-id="${accountId}"]`);
            const note = noteInput ? noteInput.value.trim() : '';
            if (amount !== 0 || note) {
                balances.push({ account_id: accountId, amount: amount, note: note });
            }
        });

        try {
            obSaveBtn.disabled = true;
            obSaveBtn.textContent = '保存中...';
            const res = await fetchAPI('/api/opening-balances', 'POST', {
                fiscal_year: year,
                balances: balances,
            });
            if (res.status === 'success') {
                showToast('期首残高を保存しました');
            } else {
                showToast('保存に失敗しました', true);
            }
        } catch (err) {
            showToast('通信エラー', true);
        } finally {
            obSaveBtn.disabled = false;
            obSaveBtn.textContent = '期首残高を保存';
        }
    }

    // ============================================================
    //  Section 13: View 7 — データのバックアップ (Backup & Restore)
    // ============================================================
    const backupJsonBtn = document.getElementById('backup-json');
    const backupDriveBtn = document.getElementById('backup-drive');

    // --- JSON Download ---
    backupJsonBtn.addEventListener('click', async () => {
        try {
            backupJsonBtn.disabled = true;
            backupJsonBtn.textContent = 'ダウンロード中...';
            const res = await fetch('/api/backup/download?format=json', {
                headers: accessToken ? { 'Authorization': 'Bearer ' + accessToken } : {},
                cache: 'no-store'
            });
            if (!res.ok) {
                showToast('バックアップのダウンロードに失敗しました', true);
                return;
            }
            const blob = await res.blob();
            downloadBlob(blob, `hinakira_backup_${todayStr()}.json`, 'application/json');
            showToast('JSONバックアップをダウンロードしました');
        } catch (err) {
            showToast('バックアップに失敗しました', true);
        } finally {
            backupJsonBtn.disabled = false;
            backupJsonBtn.textContent = 'JSONバックアップ';
        }
    });

    // --- Google Drive Backup ---
    backupDriveBtn.addEventListener('click', async () => {
        if (!accessToken) {
            showToast('Googleにログインしてください', true);
            return;
        }
        try {
            backupDriveBtn.disabled = true;
            backupDriveBtn.textContent = 'アップロード中...';
            const result = await fetchAPI('/api/backup/drive', 'POST', { access_token: accessToken });
            if (result.status === 'success') {
                showToast(`Driveに保存しました: ${result.filename}`);
            } else {
                showToast(result.error || 'Drive保存に失敗しました', true);
            }
        } catch (err) {
            showToast('Drive保存に失敗しました', true);
        } finally {
            backupDriveBtn.disabled = false;
            backupDriveBtn.textContent = 'Driveにバックアップ';
        }
    });

    function downloadBlob(blob, filename, mimeType) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    // --- Restore Tab Switching ---
    const restoreTabs = document.querySelectorAll('.restore-tab');
    const restoreLocalPanel = document.getElementById('restore-local-panel');
    const restoreDrivePanel = document.getElementById('restore-drive-panel');

    restoreTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            restoreTabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            const mode = tab.dataset.restore;
            restoreLocalPanel.style.display = mode === 'local' ? '' : 'none';
            restoreDrivePanel.style.display = mode === 'drive' ? '' : 'none';
        });
    });

    // --- Local File Restore ---
    const restoreFileInput = document.getElementById('restore-file');
    const restoreSelectBtn = document.getElementById('restore-select-btn');
    const restoreFilename = document.getElementById('restore-filename');
    const restorePreview = document.getElementById('restore-preview');
    const restoreSummary = document.getElementById('restore-summary');
    const restoreBtn = document.getElementById('restore-btn');
    let restoreData = null;

    const TABLE_LABELS = {
        accounts_master: '勘定科目',
        journal_entries: '仕訳',
        opening_balances: '期首残高',
        counterparties: '取引先',
        settings: '設定'
    };

    restoreSelectBtn.addEventListener('click', () => restoreFileInput.click());

    restoreFileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        restoreFilename.textContent = file.name;
        const reader = new FileReader();
        reader.onload = (ev) => {
            try {
                restoreData = JSON.parse(ev.target.result);
                let html = '<table>';
                let hasData = false;
                for (const [key, label] of Object.entries(TABLE_LABELS)) {
                    const arr = restoreData[key];
                    if (Array.isArray(arr)) {
                        html += `<tr><td>${label}</td><td>${arr.length}件</td></tr>`;
                        hasData = true;
                    }
                }
                html += '</table>';
                if (!hasData) {
                    restoreSummary.innerHTML = '<p style="color:#dc2626;">有効なバックアップデータが見つかりません。</p>';
                    restoreBtn.style.display = 'none';
                } else {
                    restoreSummary.innerHTML = html;
                    restoreBtn.style.display = '';
                }
                restorePreview.style.display = '';
            } catch (err) {
                restoreData = null;
                restoreSummary.innerHTML = '<p style="color:#dc2626;">JSONファイルの解析に失敗しました。</p>';
                restorePreview.style.display = '';
                restoreBtn.style.display = 'none';
            }
        };
        reader.readAsText(file);
    });

    restoreBtn.addEventListener('click', async () => {
        if (!restoreData) return;
        const ok = confirm('⚠️ 現在のデータはすべて上書きされます。\n本当に復元しますか？');
        if (!ok) return;
        try {
            restoreBtn.disabled = true;
            restoreBtn.textContent = '復元中...';
            const formData = new FormData();
            const blob = new Blob([JSON.stringify(restoreData)], { type: 'application/json' });
            formData.append('file', blob, 'restore.json');
            const result = await fetchAPI('/api/backup/restore', 'POST', formData);
            if (result.status === 'success') {
                showRestoreSuccess(result.summary);
                restoreData = null;
                restoreFileInput.value = '';
                restoreFilename.textContent = 'ファイル未選択';
                restorePreview.style.display = 'none';
                restoreBtn.style.display = 'none';
            } else {
                showToast(result.error || '復元に失敗しました', true);
            }
        } catch (err) {
            showToast('復元に失敗しました', true);
        } finally {
            restoreBtn.disabled = false;
            restoreBtn.textContent = 'データを復元する';
        }
    });

    // --- Drive Restore ---
    const driveListBtn = document.getElementById('drive-list-btn');
    const driveFileList = document.getElementById('drive-file-list');

    driveListBtn.addEventListener('click', async () => {
        if (!accessToken) {
            showToast('Googleにログインしてください', true);
            return;
        }
        try {
            driveListBtn.disabled = true;
            driveListBtn.textContent = '取得中...';
            const result = await fetchAPI('/api/backup/drive/list', 'POST', { access_token: accessToken });
            if (result.error) {
                showToast(result.error, true);
                driveFileList.innerHTML = '';
                return;
            }
            const files = result.files || [];
            if (files.length === 0) {
                driveFileList.innerHTML = '<p style="font-size:0.8125rem; color:#64748b;">Driveにバックアップファイルがありません。</p>';
                return;
            }
            let html = '';
            for (const f of files) {
                const d = new Date(f.createdTime);
                const dateStr = `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
                html += `<div class="drive-file-item">
                    <div class="drive-file-info">
                        <div class="name">${f.name}</div>
                        <div class="date">${dateStr}</div>
                    </div>
                    <button class="btn btn-danger btn-sm" onclick="window.__driveRestore('${f.id}','${f.name}')">復元</button>
                </div>`;
            }
            driveFileList.innerHTML = html;
        } catch (err) {
            showToast('一覧取得に失敗しました', true);
        } finally {
            driveListBtn.disabled = false;
            driveListBtn.textContent = 'Driveのバックアップ一覧を取得';
        }
    });

    // Expose drive restore to onclick handlers
    window.__driveRestore = async (fileId, fileName) => {
        const ok = confirm(`⚠️ 「${fileName}」から復元します。\n現在のデータはすべて上書きされます。\n本当に復元しますか？`);
        if (!ok) return;
        try {
            showToast('Driveから復元中...');
            const result = await fetchAPI('/api/backup/drive/restore', 'POST', { access_token: accessToken, file_id: fileId });
            if (result.status === 'success') {
                showRestoreSuccess(result.summary);
            } else {
                showToast(result.error || '復元に失敗しました', true);
            }
        } catch (err) {
            showToast('復元に失敗しました', true);
        }
    };

    function showRestoreSuccess(summary) {
        const s = summary || {};
        const parts = [];
        for (const [key, label] of Object.entries(TABLE_LABELS)) {
            if (s[key] !== undefined) parts.push(`${label}: ${s[key]}件`);
        }
        showToast(`復元完了: ${parts.join(', ')}`);
    }

    // ============================================================
    //  Section 14: View 8 — アウトプット (Output / Export)
    // ============================================================
    const outStartInput = document.getElementById('out-start');
    const outEndInput = document.getElementById('out-end');

    function applyFiscalYearToOutput() {
        const fy = getSelectedFiscalYear();
        outStartInput.value = `${fy}-01-01`;
        outEndInput.value = `${fy}-12-31`;
    }
    applyFiscalYearToOutput();

    function outParams() {
        const p = new URLSearchParams();
        if (outStartInput.value) p.set('start_date', outStartInput.value);
        if (outEndInput.value) p.set('end_date', outEndInput.value);
        return p;
    }
    function periodLabel() {
        return `${outStartInput.value || '?'} ～ ${outEndInput.value || '?'}`;
    }

    // --- 1. 仕訳帳 ---
    document.getElementById('out-journal-csv').addEventListener('click', async () => {
        const p = outParams(); p.set('format', 'csv');
        try {
            const res = await fetch('/api/export/journal?' + p.toString());
            const blob = await res.blob();
            downloadBlob(blob, `仕訳帳_${todayStr()}.csv`, 'text/csv');
            showToast('仕訳帳CSVをダウンロードしました');
        } catch (err) { showToast('エクスポートに失敗しました', true); }
    });
    document.getElementById('out-journal-pdf').addEventListener('click', async () => {
        const p = outParams(); p.set('format', 'json');
        try {
            const data = await fetchAPI('/api/export/journal?' + p.toString());
            openPrintView('仕訳帳', periodLabel(), buildJournalPrintTable(data.entries || []));
        } catch (err) { showToast('エクスポートに失敗しました', true); }
    });

    // --- 2. 総勘定元帳 ---
    document.getElementById('out-ledger-csv').addEventListener('click', async () => {
        try {
            const data = await fetchAPI('/api/export/ledger?' + outParams().toString());
            const csv = buildLedgerCsv(data.accounts || []);
            const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
            downloadBlob(blob, `総勘定元帳_${todayStr()}.csv`, 'text/csv');
            showToast('総勘定元帳CSVをダウンロードしました');
        } catch (err) { showToast('エクスポートに失敗しました', true); }
    });
    document.getElementById('out-ledger-pdf').addEventListener('click', async () => {
        try {
            const data = await fetchAPI('/api/export/ledger?' + outParams().toString());
            openPrintView('総勘定元帳', periodLabel(), buildLedgerPrintTable(data.accounts || []));
        } catch (err) { showToast('エクスポートに失敗しました', true); }
    });

    // --- 3. 貸借対照表 (B/S) ---
    document.getElementById('out-bs-csv').addEventListener('click', async () => {
        try {
            const data = await fetchAPI('/api/export/trial-balance?' + outParams().toString());
            const balances = (data.balances || []).filter(b => b.account_type === '資産' || b.account_type === '負債' || b.account_type === '純資産');
            const csv = buildBSCsv(balances);
            const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
            downloadBlob(blob, `貸借対照表_${todayStr()}.csv`, 'text/csv');
            showToast('貸借対照表CSVをダウンロードしました');
        } catch (err) { showToast('エクスポートに失敗しました', true); }
    });
    document.getElementById('out-bs-pdf').addEventListener('click', async () => {
        try {
            const data = await fetchAPI('/api/export/trial-balance?' + outParams().toString());
            const balances = (data.balances || []).filter(b => b.account_type === '資産' || b.account_type === '負債' || b.account_type === '純資産');
            openPrintView('貸借対照表', periodLabel(), buildBSPrintTable(balances));
        } catch (err) { showToast('エクスポートに失敗しました', true); }
    });

    // --- 4. 損益計算書 (P/L) ---
    document.getElementById('out-pl-csv').addEventListener('click', async () => {
        try {
            const data = await fetchAPI('/api/export/trial-balance?' + outParams().toString());
            const balances = (data.balances || []).filter(b => b.account_type === '収益' || b.account_type === '費用');
            const csv = buildPLCsv(balances);
            const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
            downloadBlob(blob, `損益計算書_${todayStr()}.csv`, 'text/csv');
            showToast('損益計算書CSVをダウンロードしました');
        } catch (err) { showToast('エクスポートに失敗しました', true); }
    });
    document.getElementById('out-pl-pdf').addEventListener('click', async () => {
        try {
            const data = await fetchAPI('/api/export/trial-balance?' + outParams().toString());
            const balances = (data.balances || []).filter(b => b.account_type === '収益' || b.account_type === '費用');
            openPrintView('損益計算書', periodLabel(), buildPLPrintTable(balances));
        } catch (err) { showToast('エクスポートに失敗しました', true); }
    });

    // --- 5. 青色申告決算書 ---
    document.getElementById('out-blue-return').addEventListener('click', async () => {
        try {
            showToast('青色申告決算書を生成中...');
            const data = await fetchAPI('/api/export/blue-return?' + outParams().toString());
            const balances = data.balances || [];
            const monthly = data.monthly || {};
            const fiscalYear = data.fiscal_year || thisYear.toString();
            openBlueReturnPrintView(fiscalYear, balances, monthly);
        } catch (err) {
            console.error('Blue return export error:', err);
            showToast('青色申告決算書の生成に失敗しました', true);
        }
    });

    // ============================================================
    //  Print / CSV Builders
    // ============================================================
    const TBL_STYLE = 'border-collapse:collapse;width:100%;font-size:11px;';
    const TH_STYLE = 'background:#eef2ff;color:#1e3a8a;';
    const R = 'text-align:right;';
    const B = 'font-weight:bold;';

    // -- 仕訳帳 --
    function buildJournalPrintTable(entries) {
        let html = `<table border="1" cellpadding="4" cellspacing="0" style="${TBL_STYLE}">
            <thead><tr style="${TH_STYLE}">
                <th>日付</th><th>借方科目</th><th>貸方科目</th><th style="${R}">金額</th><th>税区分</th><th>取引先</th><th>摘要</th>
            </tr></thead><tbody>`;
        entries.forEach(e => {
            html += `<tr>
                <td>${e.entry_date||''}</td><td>${e.debit_account||''}</td><td>${e.credit_account||''}</td>
                <td style="${R}">${fmt(e.amount)}</td><td>${e.tax_classification||''}</td>
                <td>${e.counterparty||''}</td><td>${e.memo||''}</td></tr>`;
        });
        html += '</tbody></table>';
        return html;
    }

    // -- 総勘定元帳 --
    function buildLedgerCsv(accts) {
        let csv = '勘定科目,日付,相手科目,摘要,借方,貸方,残高\n';
        accts.forEach(a => {
            a.entries.forEach(e => {
                csv += `"${a.account_name}",${e.entry_date},"${e.counter_account||''}","${(e.memo||'').replace(/"/g,'""')}",${e.debit_amount||0},${e.credit_amount||0},${e.balance||0}\n`;
            });
        });
        return csv;
    }
    function buildLedgerPrintTable(accts) {
        let html = '';
        accts.forEach(a => {
            html += `<h3 style="margin:1.5em 0 0.3em;font-size:13px;border-bottom:1px solid #ccc;padding-bottom:4px;">${a.account_code} ${a.account_name}</h3>`;
            html += `<table border="1" cellpadding="3" cellspacing="0" style="${TBL_STYLE}">
                <thead><tr style="${TH_STYLE}">
                    <th>日付</th><th>相手科目</th><th>摘要</th><th style="${R}">借方</th><th style="${R}">貸方</th><th style="${R}">残高</th>
                </tr></thead><tbody>`;
            html += `<tr style="background:#f8fafc;"><td colspan="3" style="${B}">前期繰越</td><td></td><td></td><td style="${R}${B}">${fmt(a.opening_balance)}</td></tr>`;
            a.entries.forEach(e => {
                html += `<tr>
                    <td>${e.entry_date||''}</td><td>${e.counter_account||''}</td><td>${e.memo||''}</td>
                    <td style="${R}">${e.debit_amount ? fmt(e.debit_amount) : ''}</td>
                    <td style="${R}">${e.credit_amount ? fmt(e.credit_amount) : ''}</td>
                    <td style="${R}${B}">${fmt(e.balance)}</td></tr>`;
            });
            html += '</tbody></table>';
        });
        return html;
    }

    // -- 貸借対照表 (B/S) -- freee風 左右対照フォーマット --
    function buildBSCsv(balances) {
        let csv = '区分,勘定科目,残高\n';
        const assets = balances.filter(b => b.account_type === '資産');
        const liab = balances.filter(b => b.account_type === '負債');
        const equity = balances.filter(b => b.account_type === '純資産');
        csv += '【資産の部】,,\n';
        let aTotal = 0;
        assets.forEach(b => { aTotal += b.closing_balance; csv += `資産,"${b.name}",${b.closing_balance}\n`; });
        csv += `,資産合計,${aTotal}\n`;
        csv += '【負債の部】,,\n';
        let lTotal = 0;
        liab.forEach(b => { lTotal += b.closing_balance; csv += `負債,"${b.name}",${b.closing_balance}\n`; });
        csv += `,負債合計,${lTotal}\n`;
        csv += '【純資産の部】,,\n';
        let eTotal = 0;
        equity.forEach(b => { eTotal += b.closing_balance; csv += `純資産,"${b.name}",${b.closing_balance}\n`; });
        csv += `,純資産合計,${eTotal}\n`;
        csv += `,負債・純資産合計,${lTotal + eTotal}\n`;
        return csv;
    }
    function buildBSPrintTable(balances) {
        const assets = balances.filter(b => b.account_type === '資産');
        const liab = balances.filter(b => b.account_type === '負債');
        const equity = balances.filter(b => b.account_type === '純資産');
        let aTotal = 0, lTotal = 0, eTotal = 0;
        assets.forEach(b => aTotal += b.closing_balance);
        liab.forEach(b => lTotal += b.closing_balance);
        equity.forEach(b => eTotal += b.closing_balance);

        const S = 'border:none;padding:6px 12px;font-size:12px;';
        const SH = `${S}font-weight:700;font-size:13px;color:#1e3a8a;padding-top:14px;`;
        const ST = `${S}font-weight:700;background:#eef2ff;`;
        const SG = `${S}font-weight:700;background:#1e3a8a;color:#fff;font-size:13px;`;

        function buildSide(sections) {
            let h = '';
            sections.forEach(sec => {
                h += `<tr><td style="${SH}" colspan="2">${sec.title}</td></tr>`;
                sec.items.forEach(b => {
                    h += `<tr><td style="${S}padding-left:24px;">${b.name}</td><td style="${S}${R}">${fmt(b.closing_balance)}</td></tr>`;
                });
                h += `<tr><td style="${ST}">${sec.subtotalLabel}</td><td style="${ST}${R}">${fmt(sec.subtotal)}</td></tr>`;
            });
            return h;
        }

        const leftRows = buildSide([
            { title: '資産の部', items: assets, subtotalLabel: '資産合計', subtotal: aTotal }
        ]);
        const rightRows = buildSide([
            { title: '負債の部', items: liab, subtotalLabel: '負債合計', subtotal: lTotal },
            { title: '純資産の部', items: equity, subtotalLabel: '純資産合計', subtotal: eTotal }
        ]);

        let html = `<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead><tr>
            <th colspan="2" style="background:#1e3a8a;color:#fff;padding:10px;text-align:center;width:50%;border-right:2px solid #fff;">資産の部</th>
            <th colspan="2" style="background:#1e3a8a;color:#fff;padding:10px;text-align:center;width:50%;">負債・純資産の部</th>
        </tr></thead>
        <tbody><tr>
            <td colspan="2" style="vertical-align:top;border-right:1px solid #cbd5e1;"><table style="width:100%;border-collapse:collapse;">${leftRows}</table></td>
            <td colspan="2" style="vertical-align:top;"><table style="width:100%;border-collapse:collapse;">${rightRows}</table></td>
        </tr></tbody></table>`;

        // 合計一致バー
        html += `<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;margin-top:4px;">
            <tr>
                <td style="${SG}width:50%;border-right:2px solid #334155;">資産合計　${fmt(aTotal)}</td>
                <td style="${SG}width:50%;">負債・純資産合計　${fmt(lTotal + eTotal)}</td>
            </tr></table>`;
        return html;
    }

    // -- 損益計算書 (P/L) -- freee風 階層フォーマット --
    function buildPLCsv(balances) {
        const revenues = balances.filter(b => b.account_type === '収益');
        const expenses = balances.filter(b => b.account_type === '費用');
        const sales = revenues.filter(b => b.code === '400');
        const costOfSales = expenses.filter(b => b.code === '500');
        const sgaExpenses = expenses.filter(b => b.code !== '500');
        const otherRevenues = revenues.filter(b => b.code !== '400');

        const salesTotal = sales.reduce((s, b) => s + Math.abs(b.closing_balance), 0);
        const costTotal = costOfSales.reduce((s, b) => s + Math.abs(b.closing_balance), 0);
        const grossProfit = salesTotal - costTotal;
        const sgaTotal = sgaExpenses.reduce((s, b) => s + Math.abs(b.closing_balance), 0);
        const operatingIncome = grossProfit - sgaTotal;
        const otherRevTotal = otherRevenues.reduce((s, b) => s + Math.abs(b.closing_balance), 0);
        const ordinaryIncome = operatingIncome + otherRevTotal;

        let csv = '項目,金額\n';
        csv += '【売上高】,\n';
        sales.forEach(b => csv += `"  ${b.name}",${Math.abs(b.closing_balance)}\n`);
        csv += `売上高合計,${salesTotal}\n`;
        csv += '【売上原価】,\n';
        costOfSales.forEach(b => csv += `"  ${b.name}",${Math.abs(b.closing_balance)}\n`);
        csv += `売上原価合計,${costTotal}\n`;
        csv += `売上総利益,${grossProfit}\n`;
        csv += '【販売費及び一般管理費】,\n';
        sgaExpenses.forEach(b => csv += `"  ${b.name}",${Math.abs(b.closing_balance)}\n`);
        csv += `販売費及び一般管理費合計,${sgaTotal}\n`;
        csv += `営業利益,${operatingIncome}\n`;
        if (otherRevenues.length > 0) {
            csv += '【営業外収益】,\n';
            otherRevenues.forEach(b => csv += `"  ${b.name}",${Math.abs(b.closing_balance)}\n`);
        }
        csv += `経常利益,${ordinaryIncome}\n`;
        csv += `税引前当期純利益,${ordinaryIncome}\n`;
        csv += `当期純利益,${ordinaryIncome}\n`;
        return csv;
    }
    function buildPLPrintTable(balances) {
        const revenues = balances.filter(b => b.account_type === '収益');
        const expenses = balances.filter(b => b.account_type === '費用');

        // 科目分類
        const sales = revenues.filter(b => b.code === '400');
        const costOfSales = expenses.filter(b => b.code === '500');
        const sgaExpenses = expenses.filter(b => b.code !== '500');
        const otherRevenues = revenues.filter(b => b.code !== '400');

        const salesTotal = sales.reduce((s, b) => s + Math.abs(b.closing_balance), 0);
        const costTotal = costOfSales.reduce((s, b) => s + Math.abs(b.closing_balance), 0);
        const grossProfit = salesTotal - costTotal;
        const sgaTotal = sgaExpenses.reduce((s, b) => s + Math.abs(b.closing_balance), 0);
        const operatingIncome = grossProfit - sgaTotal;
        const otherRevTotal = otherRevenues.reduce((s, b) => s + Math.abs(b.closing_balance), 0);
        const ordinaryIncome = operatingIncome + otherRevTotal;

        const S = 'border:none;padding:5px 12px;font-size:12px;';
        const SH = `${S}font-weight:700;font-size:13px;color:#1e3a8a;background:#f8fafc;padding-top:12px;`;
        const SST = `${S}font-weight:700;background:#eef2ff;border-top:1px solid #cbd5e1;border-bottom:1px solid #cbd5e1;`;
        const SBG = `${S}font-weight:700;background:#1e3a8a;color:#fff;font-size:13px;`;

        function row(label, val, style) { return `<tr><td style="${style || S}">${label}</td><td style="${(style || S)}${R}">${fmt(val)}</td></tr>`; }
        function itemRow(name, val) { return `<tr><td style="${S}padding-left:24px;">${name}</td><td style="${S}${R}">${fmt(val)}</td></tr>`; }
        function secHeader(title) { return `<tr><td colspan="2" style="${SH}">${title}</td></tr>`; }
        function subtotalRow(label, val) { return row(label, val, SST); }
        function totalRow(label, val) { return row(label, val, SBG); }

        let html = `<table cellpadding="0" cellspacing="0" style="width:100%;max-width:600px;border-collapse:collapse;border:1px solid #cbd5e1;">`;

        // 売上高
        html += secHeader('売上高');
        sales.forEach(b => html += itemRow(b.name, Math.abs(b.closing_balance)));
        html += subtotalRow('売上高合計', salesTotal);

        // 売上原価
        html += secHeader('売上原価');
        costOfSales.forEach(b => html += itemRow(b.name, Math.abs(b.closing_balance)));
        html += subtotalRow('売上原価合計', costTotal);

        // 売上総利益
        html += totalRow('売上総利益', grossProfit);

        // 販売費及び一般管理費
        html += secHeader('販売費及び一般管理費');
        sgaExpenses.forEach(b => html += itemRow(b.name, Math.abs(b.closing_balance)));
        html += subtotalRow('販売費及び一般管理費合計', sgaTotal);

        // 営業利益
        html += totalRow('営業利益', operatingIncome);

        // 営業外収益
        if (otherRevenues.length > 0) {
            html += secHeader('営業外収益');
            otherRevenues.forEach(b => html += itemRow(b.name, Math.abs(b.closing_balance)));
            html += subtotalRow('営業外収益合計', otherRevTotal);
        }

        // 経常利益
        html += totalRow('経常利益', ordinaryIncome);

        // 税引前当期純利益 = 経常利益（特別損益なし）
        html += totalRow('税引前当期純利益', ordinaryIncome);

        // 当期純利益
        const netStyle = `${S}font-weight:700;background:#0f172a;color:#fff;font-size:14px;`;
        html += row('当期純利益', ordinaryIncome, netStyle);

        html += '</table>';
        return html;
    }

    function openPrintView(title, period, tableHtml) {
        const win = window.open('', '_blank');
        win.document.write(`<!DOCTYPE html><html><head><title>${title}</title>
            <style>body{font-family:'Inter',sans-serif;padding:20px;color:#1e293b;}
            h1{font-size:18px;margin-bottom:4px;}
            .meta{font-size:12px;color:#64748b;margin-bottom:16px;}
            table{page-break-inside:auto;} tr{page-break-inside:avoid;}
            @media print{.no-print{display:none;}}</style>
        </head><body>
            <h1>${title}</h1>
            <p class="meta">期間: ${period}　|　出力日: ${todayStr()}</p>
            ${tableHtml}
            <br><button class="no-print" onclick="window.print()" style="padding:8px 20px;font-size:14px;cursor:pointer;border:1px solid #cbd5e1;border-radius:6px;background:#fff;">印刷 / PDF保存</button>
        </body></html>`);
        win.document.close();
    }

    // ============================================================
    //  Blue Return Tax Form (青色申告決算書) — 4-page print view
    // ============================================================
    function openBlueReturnPrintView(fiscalYear, balances, monthly) {
        // --- Data preparation ---
        const revenues = balances.filter(b => b.account_type === '収益');
        const expenses = balances.filter(b => b.account_type === '費用');
        const assets = balances.filter(b => b.account_type === '資産');
        const liabilities = balances.filter(b => b.account_type === '負債');
        const equity = balances.filter(b => b.account_type === '純資産');

        const sales = revenues.filter(b => b.code === '400');
        const costOfSales = expenses.filter(b => b.code === '500');
        const otherRevenues = revenues.filter(b => b.code !== '400');
        const sgaExpenses = expenses.filter(b => b.code !== '500');

        const salesTotal = sales.reduce((s, b) => s + Math.abs(b.closing_balance), 0);
        const costTotal = costOfSales.reduce((s, b) => s + Math.abs(b.closing_balance), 0);
        const grossProfit = salesTotal - costTotal;
        const sgaTotal = sgaExpenses.reduce((s, b) => s + Math.abs(b.closing_balance), 0);
        const operatingIncome = grossProfit - sgaTotal;
        const otherRevTotal = otherRevenues.reduce((s, b) => s + Math.abs(b.closing_balance), 0);
        const ordinaryIncome = operatingIncome + otherRevTotal;

        const assetTotal = assets.reduce((s, b) => s + b.closing_balance, 0);
        const liabTotal = liabilities.reduce((s, b) => s + b.closing_balance, 0);
        const equityTotal = equity.reduce((s, b) => s + b.closing_balance, 0);

        const monthlyData = monthly.monthly || [];
        const expenseAccounts = monthly.expense_accounts || [];

        // Account mapping for 青色申告決算書 line items
        const BR_EXPENSE_MAP = [
            {code: '510', label: '給料賃金'},
            {code: '520', label: '外注工賃'},
            {code: '580', label: '修繕費'},
            {code: '530', label: '旅費交通費'},
            {code: '531', label: '通信費'},
            {code: '540', label: '広告宣伝費'},
            {code: '541', label: '接待交際費'},
            {code: '550', label: '消耗品費'},
            {code: '560', label: '水道光熱費'},
            {code: '570', label: '地代家賃'},
            {code: '590', label: '支払手数料'},
            {code: '600', label: '租税公課'},
            {code: '620', label: '保険料'},
            {code: '630', label: '減価償却費'},
            {code: '610', label: '新聞図書費'},
            {code: '551', label: '会議費'},
            {code: '511', label: '給料手当'},
            {code: '900', label: '雑費'},
        ];

        function findExpense(code) {
            const b = expenses.find(e => e.code === code);
            return b ? Math.abs(b.closing_balance) : 0;
        }

        const f = (n) => (n || 0).toLocaleString();

        // --- CSS for A4 print layout ---
        const css = `
            * { margin:0; padding:0; box-sizing:border-box; }
            body { font-family: 'Noto Sans JP', 'Inter', sans-serif; font-size: 10px; color: #000; }
            .page { width: 210mm; min-height: 297mm; padding: 12mm 15mm; margin: 0 auto; page-break-after: always; position: relative; }
            .page:last-child { page-break-after: auto; }
            .page-title { text-align: center; font-size: 16px; font-weight: 700; margin-bottom: 4px; letter-spacing: 2px; }
            .page-subtitle { text-align: center; font-size: 11px; color: #333; margin-bottom: 8px; }
            .form-header { display: flex; justify-content: space-between; margin-bottom: 8px; font-size: 10px; }
            .form-header span { border: 1px solid #999; padding: 2px 8px; }
            table { border-collapse: collapse; width: 100%; font-size: 10px; }
            th, td { border: 1px solid #666; padding: 3px 6px; vertical-align: middle; }
            th { background: #f0f0f0; font-weight: 600; text-align: center; }
            td.num { text-align: right; font-variant-numeric: tabular-nums; }
            td.label { background: #f8f8f8; font-weight: 500; }
            .section-header { background: #e0e0e0; font-weight: 700; text-align: center; font-size: 11px; }
            .total-row td { font-weight: 700; background: #eef2ff; }
            .grand-total td { font-weight: 700; background: #1e3a8a; color: #fff; font-size: 11px; }
            .note { font-size: 9px; color: #666; margin-top: 4px; }
            .two-col { display: flex; gap: 8px; }
            .two-col > div { flex: 1; }
            .print-btn { padding: 10px 24px; font-size: 14px; cursor: pointer; border: 1px solid #cbd5e1; border-radius: 6px; background: #fff; margin: 16px auto; display: block; }
            @media print {
                .no-print { display: none !important; }
                .page { padding: 8mm 10mm; margin: 0; }
                body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
            }
            @media screen {
                body { background: #e2e8f0; }
                .page { background: #fff; box-shadow: 0 2px 8px rgba(0,0,0,0.15); margin: 20px auto; }
            }
        `;

        // --- Page 1: 損益計算書 ---
        function buildPage1() {
            let html = `<div class="page">`;
            html += `<div class="page-title">青色申告決算書（一般用）</div>`;
            html += `<div class="page-subtitle">所得税青色申告決算書（一般用）　令和${parseInt(fiscalYear) - 2018}年分</div>`;
            html += `<div class="form-header"><span>自 ${fiscalYear}年1月1日</span><span>至 ${fiscalYear}年12月31日</span></div>`;

            html += `<table>`;
            html += `<tr class="section-header"><td colspan="3">損益計算書</td></tr>`;
            html += `<tr><th style="width:50%;">科目</th><th style="width:15%;">行</th><th style="width:35%;">金額</th></tr>`;

            // Revenue
            html += `<tr><td class="label">売上（収入）金額 ①</td><td class="num">①</td><td class="num">${f(salesTotal)}</td></tr>`;

            // Cost of sales
            html += `<tr class="section-header"><td colspan="3">売上原価</td></tr>`;
            html += `<tr><td class="label">　仕入金額 ③</td><td class="num">③</td><td class="num">${f(costTotal)}</td></tr>`;
            html += `<tr class="total-row"><td>差引原価 ⑤</td><td class="num">⑤</td><td class="num">${f(costTotal)}</td></tr>`;
            html += `<tr class="total-row"><td>差引金額（売上総利益）⑥</td><td class="num">⑥</td><td class="num">${f(grossProfit)}</td></tr>`;

            // Expenses
            html += `<tr class="section-header"><td colspan="3">経費</td></tr>`;
            let lineNum = 7;
            let expensesShown = 0;
            BR_EXPENSE_MAP.forEach(item => {
                const val = findExpense(item.code);
                if (val > 0) {
                    html += `<tr><td>　${item.label}</td><td class="num">${lineNum}</td><td class="num">${f(val)}</td></tr>`;
                    expensesShown++;
                }
                lineNum++;
            });

            // Any remaining expenses not in the map
            expenses.forEach(e => {
                if (e.code !== '500' && !BR_EXPENSE_MAP.find(m => m.code === e.code) && Math.abs(e.closing_balance) > 0) {
                    html += `<tr><td>　${e.name}</td><td class="num">${lineNum}</td><td class="num">${f(Math.abs(e.closing_balance))}</td></tr>`;
                    lineNum++;
                }
            });

            html += `<tr class="total-row"><td>経費計 ㉕</td><td class="num">㉕</td><td class="num">${f(sgaTotal)}</td></tr>`;
            html += `<tr class="total-row"><td>差引金額 ㉖ (⑥−㉕)</td><td class="num">㉖</td><td class="num">${f(operatingIncome)}</td></tr>`;

            // Other revenue
            if (otherRevTotal > 0) {
                html += `<tr><td>　その他の収入 ㉗</td><td class="num">㉗</td><td class="num">${f(otherRevTotal)}</td></tr>`;
            }

            html += `<tr class="grand-total"><td>所得金額 ㉙</td><td class="num">㉙</td><td class="num">${f(ordinaryIncome)}</td></tr>`;

            html += `</table>`;
            html += `<p class="note">※ Hinakira会計から自動生成。最終確認は税務署の様式と照合してください。</p>`;
            html += `</div>`;
            return html;
        }

        // --- Page 2: 月別売上（収入）金額及び仕入金額 ---
        function buildPage2() {
            let html = `<div class="page">`;
            html += `<div class="page-title">月別売上（収入）金額及び仕入金額</div>`;
            html += `<div class="page-subtitle">令和${parseInt(fiscalYear) - 2018}年分</div>`;

            html += `<table>`;
            html += `<tr><th>月</th><th>売上（収入）金額</th><th>仕入金額</th></tr>`;
            let revTotal = 0, purTotal = 0;
            monthlyData.forEach(m => {
                revTotal += m.revenue;
                purTotal += m.purchases;
                html += `<tr>
                    <td style="text-align:center;">${m.month}月</td>
                    <td class="num">${m.revenue > 0 ? f(m.revenue) : ''}</td>
                    <td class="num">${m.purchases > 0 ? f(m.purchases) : ''}</td>
                </tr>`;
            });
            html += `<tr class="total-row">
                <td style="text-align:center;">合計</td>
                <td class="num">${f(revTotal)}</td>
                <td class="num">${f(purTotal)}</td>
            </tr>`;
            html += `</table>`;

            // Expense breakdown by account
            if (expenseAccounts.length > 0) {
                html += `<div style="margin-top:12px;">`;
                html += `<table>`;
                html += `<tr class="section-header"><td colspan="14">経費の内訳（月別）</td></tr>`;
                html += `<tr><th>科目</th>`;
                for (let m = 1; m <= 12; m++) html += `<th>${m}月</th>`;
                html += `<th>合計</th></tr>`;

                expenseAccounts.forEach(ea => {
                    html += `<tr><td class="label" style="font-size:9px;white-space:nowrap;">${ea.name}</td>`;
                    for (let m = 1; m <= 12; m++) {
                        const val = ea.months[m] || 0;
                        html += `<td class="num" style="font-size:9px;">${val > 0 ? f(val) : ''}</td>`;
                    }
                    html += `<td class="num" style="font-size:9px;font-weight:600;">${f(ea.total)}</td></tr>`;
                });
                html += `</table></div>`;
            }

            html += `</div>`;
            return html;
        }

        // --- Page 3: 減価償却費の計算・地代家賃の内訳 ---
        function buildPage3() {
            const depreciationVal = findExpense('630');
            const rentVal = findExpense('570');
            const insuranceVal = findExpense('620');
            const taxVal = findExpense('600');

            let html = `<div class="page">`;
            html += `<div class="page-title">減価償却費の計算・地代家賃の内訳等</div>`;
            html += `<div class="page-subtitle">令和${parseInt(fiscalYear) - 2018}年分</div>`;

            // Depreciation
            html += `<table style="margin-bottom:16px;">`;
            html += `<tr class="section-header"><td colspan="4">減価償却費の計算</td></tr>`;
            html += `<tr><th>資産名称</th><th>取得年月</th><th>取得価額</th><th>本年分の償却費</th></tr>`;
            if (depreciationVal > 0) {
                html += `<tr><td colspan="3" style="color:#666;">（個別資産の明細は別途管理してください）</td><td class="num">${f(depreciationVal)}</td></tr>`;
            } else {
                html += `<tr><td colspan="4" style="color:#999;text-align:center;">減価償却資産なし</td></tr>`;
            }
            html += `<tr class="total-row"><td colspan="3">減価償却費合計</td><td class="num">${f(depreciationVal)}</td></tr>`;
            html += `</table>`;

            // Rent
            html += `<table style="margin-bottom:16px;">`;
            html += `<tr class="section-header"><td colspan="4">地代家賃の内訳</td></tr>`;
            html += `<tr><th>支払先</th><th>物件名</th><th>賃貸料（月額）</th><th>本年中の賃借料</th></tr>`;
            if (rentVal > 0) {
                html += `<tr><td colspan="2" style="color:#666;">（明細は摘要を参照）</td><td class="num">—</td><td class="num">${f(rentVal)}</td></tr>`;
            } else {
                html += `<tr><td colspan="4" style="color:#999;text-align:center;">地代家賃なし</td></tr>`;
            }
            html += `<tr class="total-row"><td colspan="3">地代家賃合計</td><td class="num">${f(rentVal)}</td></tr>`;
            html += `</table>`;

            // Tax and Insurance summary
            html += `<table style="margin-bottom:16px;">`;
            html += `<tr class="section-header"><td colspan="2">租税公課・保険料の内訳</td></tr>`;
            if (taxVal > 0) html += `<tr><td class="label">租税公課</td><td class="num">${f(taxVal)}</td></tr>`;
            if (insuranceVal > 0) html += `<tr><td class="label">保険料</td><td class="num">${f(insuranceVal)}</td></tr>`;
            if (taxVal === 0 && insuranceVal === 0) {
                html += `<tr><td colspan="2" style="color:#999;text-align:center;">該当なし</td></tr>`;
            }
            html += `</table>`;

            html += `<p class="note">※ 減価償却費の詳細（定額法・定率法、耐用年数等）は別途管理が必要です。</p>`;
            html += `</div>`;
            return html;
        }

        // --- Page 4: 貸借対照表 ---
        function buildPage4() {
            let html = `<div class="page">`;
            html += `<div class="page-title">貸借対照表</div>`;
            html += `<div class="page-subtitle">令和${parseInt(fiscalYear) - 2018}年分　${fiscalYear}年12月31日現在</div>`;

            html += `<div class="two-col">`;

            // Left: Assets
            html += `<div><table>`;
            html += `<tr class="section-header"><td colspan="2">資産の部</td></tr>`;
            html += `<tr><th>科目</th><th>金額</th></tr>`;
            assets.forEach(a => {
                if (a.closing_balance !== 0) {
                    html += `<tr><td>${a.name}</td><td class="num">${f(a.closing_balance)}</td></tr>`;
                }
            });
            // Add net income to assets as 元入金 adjustment or show separately
            html += `<tr class="total-row"><td>資産合計</td><td class="num">${f(assetTotal)}</td></tr>`;
            html += `</table></div>`;

            // Right: Liabilities + Equity
            html += `<div><table>`;
            html += `<tr class="section-header"><td colspan="2">負債・資本の部</td></tr>`;
            html += `<tr><th>科目</th><th>金額</th></tr>`;
            liabilities.forEach(l => {
                if (l.closing_balance !== 0) {
                    html += `<tr><td>${l.name}</td><td class="num">${f(l.closing_balance)}</td></tr>`;
                }
            });
            // Equity section
            html += `<tr><td colspan="2" style="background:#f0f0f0;font-weight:600;text-align:center;">【元入金等】</td></tr>`;
            equity.forEach(e => {
                if (e.closing_balance !== 0) {
                    html += `<tr><td>${e.name}</td><td class="num">${f(e.closing_balance)}</td></tr>`;
                }
            });
            // Show current year income
            html += `<tr><td>青色申告特別控除前の所得金額</td><td class="num">${f(ordinaryIncome)}</td></tr>`;
            html += `<tr class="total-row"><td>負債・資本合計</td><td class="num">${f(liabTotal + equityTotal + ordinaryIncome)}</td></tr>`;
            html += `</table></div>`;

            html += `</div>`; // end two-col

            // Balance check
            const diff = assetTotal - (liabTotal + equityTotal + ordinaryIncome);
            if (diff !== 0) {
                html += `<p style="color:red;font-weight:bold;margin-top:8px;">⚠ 貸借差額: ${f(diff)}円（資産合計と負債・資本合計が一致していません）</p>`;
            }

            html += `<p class="note" style="margin-top:12px;">※ 元入金の期末残高 = 期首元入金 + 事業主借 − 事業主貸 + 所得金額</p>`;
            html += `</div>`;
            return html;
        }

        // --- Open print window ---
        const win = window.open('', '_blank');
        win.document.write(`<!DOCTYPE html><html lang="ja"><head>
            <meta charset="UTF-8">
            <title>青色申告決算書 ${fiscalYear}年分</title>
            <link rel="preconnect" href="https://fonts.googleapis.com">
            <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&display=swap" rel="stylesheet">
            <style>${css}</style>
        </head><body>
            <div class="no-print" style="text-align:center;padding:16px;">
                <button class="print-btn" onclick="window.print()">印刷 / PDF保存</button>
                <span style="font-size:12px;color:#666;">4ページ構成の青色申告決算書が印刷されます</span>
            </div>
            ${buildPage1()}
            ${buildPage2()}
            ${buildPage3()}
            ${buildPage4()}
            <div class="no-print" style="text-align:center;padding:16px;">
                <button class="print-btn" onclick="window.print()">印刷 / PDF保存</button>
            </div>
        </body></html>`);
        win.document.close();
    }

    // ============================================================
    //  Section 15a: Journal Entry Detail Modal (shared)
    // ============================================================
    const jeDetailModal = document.getElementById('je-detail-modal');
    const jeDetailClose = document.getElementById('je-detail-close');
    const jedDate = document.getElementById('jed-date');
    const jedDebit = document.getElementById('jed-debit');
    const jedCredit = document.getElementById('jed-credit');
    const jedAmount = document.getElementById('jed-amount');
    const jedTax = document.getElementById('jed-tax');
    const jedCounterparty = document.getElementById('jed-counterparty');
    const jedMemo = document.getElementById('jed-memo');
    const jedEvidenceRow = document.getElementById('jed-evidence-row');
    const jedEvidenceLink = document.getElementById('jed-evidence-link');
    const jedDelete = document.getElementById('jed-delete');
    const jedCancel = document.getElementById('jed-cancel');
    const jedSave = document.getElementById('jed-save');

    let jedCurrentId = null;
    let jedOnSaved = null; // callback after save/delete

    function openJEDetailModal(entry, onSaved) {
        jedCurrentId = entry.id;
        jedOnSaved = onSaved;

        jedDate.value = entry.entry_date || '';
        jedDebit.value = entry.debit_account || '';
        jedCredit.value = entry.credit_account || '';
        jedAmount.value = entry.amount || 0;
        jedTax.value = entry.tax_classification || '10%';
        jedCounterparty.value = entry.counterparty || '';
        jedMemo.value = entry.memo || '';

        if (entry.evidence_url) {
            jedEvidenceRow.classList.remove('hidden');
            jedEvidenceLink.href = entry.evidence_url;
        } else {
            jedEvidenceRow.classList.add('hidden');
        }

        jeDetailModal.classList.remove('hidden');
    }

    function closeJEDetailModal() {
        jeDetailModal.classList.add('hidden');
        jedCurrentId = null;
    }

    jeDetailClose.addEventListener('click', closeJEDetailModal);
    jedCancel.addEventListener('click', closeJEDetailModal);
    jeDetailModal.addEventListener('click', (e) => {
        if (e.target === jeDetailModal) closeJEDetailModal();
    });

    jedSave.addEventListener('click', async () => {
        if (!jedCurrentId) return;
        const updated = {
            entry_date: jedDate.value,
            debit_account: jedDebit.value,
            credit_account: jedCredit.value,
            amount: parseInt(jedAmount.value) || 0,
            tax_classification: jedTax.value,
            counterparty: jedCounterparty.value,
            memo: jedMemo.value,
        };
        try {
            const res = await fetchAPI(`/api/journal/${jedCurrentId}`, 'PUT', updated);
            if (res.status === 'success') {
                showToast('更新しました');
                closeJEDetailModal();
                if (jedOnSaved) jedOnSaved();
            } else {
                showToast('更新に失敗しました', true);
            }
        } catch (err) {
            showToast('通信エラー', true);
        }
    });

    jedDelete.addEventListener('click', async () => {
        if (!jedCurrentId) return;
        if (!confirm('この仕訳を削除しますか？')) return;
        try {
            const res = await fetchAPI(`/api/journal/${jedCurrentId}`, 'DELETE');
            if (res.status === 'success') {
                showToast('削除しました');
                closeJEDetailModal();
                if (jedOnSaved) jedOnSaved();
            } else {
                showToast('削除に失敗しました', true);
            }
        } catch (err) {
            showToast('通信エラー', true);
        }
    });

    // ============================================================
    //  Section 15b: AI Chat Bot
    // ============================================================
    const chatFab = document.getElementById('ai-chat-fab');
    const chatPanel = document.getElementById('ai-chat-panel');
    const chatClose = document.getElementById('ai-chat-close');
    const chatMessages = document.getElementById('ai-chat-messages');
    const chatInput = document.getElementById('ai-chat-input');
    const chatSendBtn = document.getElementById('ai-chat-send');
    const chatImageBtn = document.getElementById('ai-chat-image-btn');
    const chatImageInput = document.getElementById('ai-chat-image-input');
    const chatImagePreview = document.getElementById('ai-chat-image-preview');
    let chatHistory = [];
    let chatPendingImage = null; // {data: base64, mimeType: string, name: string}

    chatFab.addEventListener('click', () => {
        chatPanel.classList.toggle('hidden');
        if (!chatPanel.classList.contains('hidden')) {
            chatInput.focus();
        }
    });
    chatClose.addEventListener('click', () => chatPanel.classList.add('hidden'));

    // Image attachment handling
    if (chatImageBtn) {
        chatImageBtn.addEventListener('click', () => chatImageInput.click());
    }
    if (chatImageInput) {
        chatImageInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            if (!file.type.startsWith('image/')) {
                addChatMsg('bot', '画像ファイルのみアップロードできます。');
                return;
            }
            if (file.size > 10 * 1024 * 1024) {
                addChatMsg('bot', 'ファイルサイズは10MB以下にしてください。');
                return;
            }
            const reader = new FileReader();
            reader.onload = () => {
                const base64 = reader.result.split(',')[1];
                chatPendingImage = { data: base64, mimeType: file.type, name: file.name };
                showChatImagePreview(reader.result, file.name);
            };
            reader.readAsDataURL(file);
            chatImageInput.value = '';
        });
    }

    function showChatImagePreview(dataUrl, name) {
        if (!chatImagePreview) return;
        chatImagePreview.innerHTML = `
            <div class="chat-preview-wrap">
                <img src="${dataUrl}" alt="${name}" class="chat-preview-img">
                <button type="button" class="chat-preview-remove" title="削除">&times;</button>
            </div>`;
        chatImagePreview.classList.remove('hidden');
        chatImagePreview.querySelector('.chat-preview-remove').addEventListener('click', clearChatImage);
    }

    function clearChatImage() {
        chatPendingImage = null;
        if (chatImagePreview) {
            chatImagePreview.innerHTML = '';
            chatImagePreview.classList.add('hidden');
        }
    }

    async function sendChatMessage() {
        const msg = chatInput.value.trim();
        if (!msg && !chatPendingImage) return;

        const apiKey = localStorage.getItem('gemini_api_key');
        if (!apiKey) {
            addChatMsg('bot', 'Gemini APIキーが設定されていません。右上の設定からAPIキーを登録してください。');
            return;
        }

        // Show user message with optional image
        if (chatPendingImage) {
            addChatMsg('user', msg || '(画像を送信)', `data:${chatPendingImage.mimeType};base64,${chatPendingImage.data}`);
        } else {
            addChatMsg('user', msg);
        }
        chatInput.value = '';
        chatHistory.push({ role: 'user', text: msg || '画像について質問' });

        // Build request body
        const reqBody = {
            message: msg,
            history: chatHistory,
            gemini_api_key: apiKey,
        };
        if (chatPendingImage) {
            reqBody.image = { data: chatPendingImage.data, mimeType: chatPendingImage.mimeType };
        }
        clearChatImage();

        // Show loading
        const loadingEl = addChatMsg('loading', '考え中...');
        chatSendBtn.disabled = true;

        try {
            const data = await fetchAPI('/api/chat', 'POST', reqBody);
            loadingEl.remove();

            if (data.reply) {
                addChatMsg('bot', data.reply);
                chatHistory.push({ role: 'model', text: data.reply });
            } else {
                addChatMsg('bot', data.error || 'エラーが発生しました。');
            }
        } catch (err) {
            loadingEl.remove();
            addChatMsg('bot', '通信エラーが発生しました。');
        } finally {
            chatSendBtn.disabled = false;
            chatInput.focus();
        }
    }

    /** Escape HTML to prevent XSS */
    function escapeHtml(str) {
        const d = document.createElement('div');
        d.textContent = str;
        return d.innerHTML;
    }

    /** Convert simple markdown to HTML (fallback if Gemini still uses markdown) */
    function formatChatText(text) {
        let html = escapeHtml(text);
        // **bold** → <strong>
        html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
        // *italic* → <em> (but not inside strong)
        html = html.replace(/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/g, '<em>$1</em>');
        // newlines → <br>
        html = html.replace(/\n/g, '<br>');
        return html;
    }

    function addChatMsg(type, text, imageUrl) {
        const div = document.createElement('div');
        div.className = `ai-msg ai-msg-${type}`;
        if (type === 'loading') {
            div.textContent = text;
        } else {
            let html = '';
            if (imageUrl) {
                html += `<img src="${imageUrl}" class="chat-msg-img" alt="添付画像">`;
            }
            html += formatChatText(text);
            div.innerHTML = html;
        }
        chatMessages.appendChild(div);
        chatMessages.scrollTop = chatMessages.scrollHeight;
        return div;
    }

    chatSendBtn.addEventListener('click', sendChatMessage);
    chatInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            sendChatMessage();
        }
    });

    // ============================================================
    //  Section: Fixed Assets (固定資産台帳)
    // ============================================================
    let faAssetsCache = [];
    let faDepreciationCache = [];

    // Sub-tab switching
    document.querySelectorAll('[data-fa-tab]').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('[data-fa-tab]').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.fa-tab-content').forEach(c => c.classList.remove('active'));
            btn.classList.add('active');
            const target = document.getElementById(btn.dataset.faTab);
            if (target) target.classList.add('active');
            // Load data when switching tabs
            if (btn.dataset.faTab === 'fa-list') loadFixedAssetsList();
            if (btn.dataset.faTab === 'fa-disposal') loadDisposalList();
            if (btn.dataset.faTab === 'fa-depreciation') loadDepreciationSchedule();
        });
    });

    // --- Registration form ---
    const faForm = document.getElementById('fa-form');
    const faEditId = document.getElementById('fa-edit-id');
    const faName = document.getElementById('fa-name');
    const faDate = document.getElementById('fa-date');
    const faLife = document.getElementById('fa-life');
    const faCost = document.getElementById('fa-cost');
    const faMethod = document.getElementById('fa-method');
    const faCategory = document.getElementById('fa-category');
    const faNotes = document.getElementById('fa-notes');
    const faCancelBtn = document.getElementById('fa-cancel-btn');
    const faAiBtn = document.getElementById('fa-ai-btn');
    const faAiHint = document.getElementById('fa-ai-hint');

    if (faForm) {
        faForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const data = {
                asset_name: faName.value.trim(),
                acquisition_date: faDate.value,
                useful_life: parseInt(faLife.value) || 0,
                acquisition_cost: parseInt(faCost.value) || 0,
                depreciation_method: faMethod.value,
                asset_category: faCategory ? faCategory.value : '',
                notes: faNotes.value.trim(),
            };
            if (!data.asset_name || !data.acquisition_date || !data.useful_life || !data.acquisition_cost) {
                showToast('必須項目を入力してください', true);
                return;
            }
            try {
                const editId = faEditId.value;
                if (editId) {
                    await fetchAPI(`/api/fixed-assets/${editId}`, 'PUT', data);
                    showToast('更新しました');
                } else {
                    await fetchAPI('/api/fixed-assets', 'POST', data);
                    showToast('登録しました');
                }
                resetFaForm();
                loadFixedAssetsList();
            } catch (err) {
                showToast('保存に失敗しました', true);
            }
        });
    }

    function resetFaForm() {
        faEditId.value = '';
        faName.value = '';
        faDate.value = '';
        faLife.value = '';
        faCost.value = '';
        faMethod.value = '定額法';
        if (faCategory) faCategory.value = '';
        faNotes.value = '';
        if (faAiHint) faAiHint.textContent = '💡 資産名称を入力して「AI判定」を押すと、耐用年数を自動で判定します。';
        faCancelBtn.style.display = 'none';
    }

    if (faCancelBtn) {
        faCancelBtn.addEventListener('click', resetFaForm);
    }

    // --- AI useful life estimation ---
    if (faAiBtn) {
        faAiBtn.addEventListener('click', async () => {
            const name = faName.value.trim();
            if (!name) {
                showToast('資産名称を入力してからAI判定を押してください', true);
                return;
            }
            faAiBtn.disabled = true;
            faAiBtn.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2L2 7l10 5 10-5-10-5z"/><path d="M2 17l10 5 10-5"/><path d="M2 12l10 5 10-5"/></svg> 判定中...';
            if (faAiHint) faAiHint.textContent = '🔄 AIが耐用年数を判定中...';
            try {
                const apiKey = localStorage.getItem('gemini_api_key');
                const res = await fetchAPI('/api/fixed-assets/ai-useful-life', 'POST', { asset_name: name, gemini_api_key: apiKey });
                if (res.useful_life) {
                    faLife.value = res.useful_life;
                    // Auto-fill asset_category from AI result
                    if (faCategory && res.asset_category) {
                        const catMap = {
                            '器具備品': '器具備品', '車両運搬具': '車両運搬具',
                            '建物': '建物', '建物附属設備': '建物附属設備',
                            '機械装置': '機械装置', '工具': '工具器具', '工具器具': '工具器具',
                            '無形固定資産': 'ソフトウェア', 'ソフトウェア': 'ソフトウェア',
                        };
                        const mapped = catMap[res.asset_category] || '';
                        if (mapped) {
                            faCategory.value = mapped;
                        } else {
                            // Try partial match
                            for (const [key, val] of Object.entries(catMap)) {
                                if (res.asset_category.includes(key)) {
                                    faCategory.value = val;
                                    break;
                                }
                            }
                        }
                    }
                    const detail = res.detail_category ? ` / ${res.detail_category}` : '';
                    if (faAiHint) {
                        faAiHint.textContent = `✅ AI判定: ${res.asset_category || ''}${detail} → ${res.useful_life}年（${res.reasoning || ''})`;
                    }
                } else {
                    showToast('AI判定できませんでした', true);
                    if (faAiHint) faAiHint.textContent = '❌ AI判定できませんでした。手動で入力してください。';
                }
            } catch (err) {
                showToast('AI判定エラー', true);
                if (faAiHint) faAiHint.textContent = '❌ AI判定エラー。手動で入力してください。';
            } finally {
                faAiBtn.disabled = false;
                faAiBtn.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2L2 7l10 5 10-5-10-5z"/><path d="M2 17l10 5 10-5"/><path d="M2 12l10 5 10-5"/></svg> AI判定';
            }
        });
    }

    // --- Fixed assets list ---
    async function loadFixedAssetsList() {
        try {
            const data = await fetchAPI('/api/fixed-assets');
            faAssetsCache = data.assets || [];
            const tbody = document.getElementById('fa-list-tbody');
            const empty = document.getElementById('fa-list-empty');
            if (!tbody) return;

            if (faAssetsCache.length === 0) {
                tbody.innerHTML = '';
                if (empty) empty.style.display = '';
                return;
            }
            if (empty) empty.style.display = 'none';

            tbody.innerHTML = faAssetsCache.map(a => {
                const dispBadge = a.disposal_type ? `<span class="badge-disp badge-${a.disposal_type === '売却' ? 'sale' : 'retire'}">${a.disposal_type}</span>` : '';
                return `
                <tr data-id="${a.id}" ${a.disposal_type ? 'class="row-disposed"' : ''}>
                    <td>${a.asset_name || ''} ${dispBadge}</td>
                    <td>${a.acquisition_date || ''}</td>
                    <td class="text-right">${fmt(a.acquisition_cost)}</td>
                    <td>${a.useful_life}年</td>
                    <td>${a.depreciation_method || '定額法'}</td>
                    <td>${a.asset_category || ''}</td>
                    <td>${a.notes || ''}</td>
                    <td><button class="btn-row-delete" data-id="${a.id}" title="削除">✕</button></td>
                </tr>`;
            }).join('');

            // Row click → edit
            tbody.querySelectorAll('tr').forEach(row => {
                row.addEventListener('click', (ev) => {
                    if (ev.target.closest('.btn-row-delete')) return;
                    const id = row.dataset.id;
                    const asset = faAssetsCache.find(a => String(a.id) === String(id));
                    if (asset) editFixedAsset(asset);
                });
            });

            // Delete button
            tbody.querySelectorAll('.btn-row-delete').forEach(btn => {
                btn.addEventListener('click', async (ev) => {
                    ev.stopPropagation();
                    if (!confirm('この固定資産を削除しますか？')) return;
                    try {
                        await fetchAPI(`/api/fixed-assets/${btn.dataset.id}`, 'DELETE');
                        showToast('削除しました');
                        loadFixedAssetsList();
                    } catch (err) {
                        showToast('削除に失敗しました', true);
                    }
                });
            });
        } catch (err) {
            console.error('Fixed assets load error:', err);
        }
    }

    function editFixedAsset(asset) {
        // Switch to register tab
        document.querySelectorAll('[data-fa-tab]').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.fa-tab-content').forEach(c => c.classList.remove('active'));
        const regBtn = document.querySelector('[data-fa-tab="fa-register"]');
        const regTab = document.getElementById('fa-register');
        if (regBtn) regBtn.classList.add('active');
        if (regTab) regTab.classList.add('active');

        // Fill form
        faEditId.value = asset.id;
        faName.value = asset.asset_name || '';
        faDate.value = asset.acquisition_date || '';
        faLife.value = asset.useful_life || '';
        faCost.value = asset.acquisition_cost || '';
        faMethod.value = asset.depreciation_method || '定額法';
        if (faCategory) faCategory.value = asset.asset_category || '';
        faNotes.value = asset.notes || '';
        faCancelBtn.style.display = '';
        faAiHint.textContent = '';

        // Scroll to form
        faForm.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    // --- Disposal (売却・除却) ---
    const faDispForm = document.getElementById('fa-disposal-form');
    const faDispAssetId = document.getElementById('fa-disp-asset-id');
    const faDispAssetName = document.getElementById('fa-disp-asset-name');
    const faDispType = document.getElementById('fa-disp-type');
    const faDispDate = document.getElementById('fa-disp-date');
    const faDispPrice = document.getElementById('fa-disp-price');

    // Toggle sale price field visibility based on disposal type
    if (faDispType) {
        faDispType.addEventListener('change', () => {
            const priceCell = faDispPrice?.closest('td');
            if (faDispType.value === '除却') {
                if (faDispPrice) faDispPrice.value = '0';
                if (priceCell) priceCell.style.opacity = '0.4';
            } else {
                if (priceCell) priceCell.style.opacity = '1';
            }
        });
    }

    async function loadDisposalList() {
        try {
            const data = await fetchAPI('/api/fixed-assets');
            faAssetsCache = data.assets || [];
            const tbody = document.getElementById('fa-disp-list-tbody');
            const empty = document.getElementById('fa-disp-list-empty');
            if (!tbody) return;

            if (faAssetsCache.length === 0) {
                tbody.innerHTML = '';
                if (empty) empty.style.display = '';
                return;
            }
            if (empty) empty.style.display = 'none';

            tbody.innerHTML = faAssetsCache.map(a => {
                const isDisposed = a.disposal_type && a.disposal_type !== '';
                const statusText = isDisposed
                    ? `<span class="badge-disp badge-${a.disposal_type === '売却' ? 'sale' : 'retire'}">${a.disposal_type} (${a.disposal_date || ''})</span>`
                    : '<span style="color:#16a34a;">使用中</span>';
                const actionBtn = isDisposed
                    ? `<button class="btn btn-ghost btn-sm fa-cancel-disp-btn" data-id="${a.id}" title="取消">取消</button>`
                    : `<button class="btn btn-primary btn-sm fa-select-disp-btn" data-id="${a.id}" title="選択">選択</button>`;
                return `
                <tr data-id="${a.id}" ${isDisposed ? 'class="row-disposed"' : ''}>
                    <td>${a.asset_name || ''}</td>
                    <td>${a.acquisition_date || ''}</td>
                    <td class="text-right">${fmt(a.acquisition_cost)}</td>
                    <td>${a.useful_life}年</td>
                    <td>${statusText}</td>
                    <td>${actionBtn}</td>
                </tr>`;
            }).join('');

            // Select button → fill form
            tbody.querySelectorAll('.fa-select-disp-btn').forEach(btn => {
                btn.addEventListener('click', (ev) => {
                    ev.stopPropagation();
                    const id = btn.dataset.id;
                    const asset = faAssetsCache.find(a => String(a.id) === String(id));
                    if (asset) {
                        faDispAssetId.value = asset.id;
                        faDispAssetName.value = asset.asset_name;
                        faDispDate.focus();
                        // Scroll form into view
                        faDispForm.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    }
                });
            });

            // Cancel disposal button
            tbody.querySelectorAll('.fa-cancel-disp-btn').forEach(btn => {
                btn.addEventListener('click', async (ev) => {
                    ev.stopPropagation();
                    if (!confirm('売却/除却を取り消しますか？')) return;
                    try {
                        await fetchAPI(`/api/fixed-assets/${btn.dataset.id}/cancel-disposal`, 'POST');
                        showToast('取り消しました');
                        loadDisposalList();
                    } catch (err) {
                        showToast('取消に失敗しました', true);
                    }
                });
            });
        } catch (err) {
            console.error('Disposal list load error:', err);
        }
    }

    if (faDispForm) {
        faDispForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const assetId = faDispAssetId.value;
            if (!assetId) {
                showToast('対象資産を選択してください', true);
                return;
            }
            if (!faDispDate.value) {
                showToast('売却/除却日を入力してください', true);
                return;
            }
            const payload = {
                disposal_type: faDispType.value,
                disposal_date: faDispDate.value,
                disposal_price: parseInt(faDispPrice.value) || 0,
            };
            if (payload.disposal_type === '売却' && !payload.disposal_price) {
                if (!confirm('売却額が0円ですが、よろしいですか？')) return;
            }
            try {
                await fetchAPI(`/api/fixed-assets/${assetId}/dispose`, 'POST', payload);
                showToast(`${payload.disposal_type}を登録しました`);
                // Reset form
                faDispAssetId.value = '';
                faDispAssetName.value = '';
                faDispDate.value = '';
                faDispPrice.value = '0';
                faDispType.value = '除却';
                loadDisposalList();
            } catch (err) {
                showToast('処理に失敗しました', true);
            }
        });
    }

    // --- Depreciation schedule ---
    async function loadDepreciationSchedule() {
        const yearEl = document.getElementById('fa-depr-year');
        const year = yearEl ? yearEl.value : '2025';
        try {
            const data = await fetchAPI(`/api/fixed-assets/depreciation?fiscal_year=${year}`);
            faDepreciationCache = data.schedule || [];
            renderDepreciationTable(faDepreciationCache);
        } catch (err) {
            console.error('Depreciation load error:', err);
        }
    }

    function renderDepreciationTable(schedule) {
        const tbody = document.getElementById('fa-depr-tbody');
        const tfoot = document.getElementById('fa-depr-tfoot');
        const empty = document.getElementById('fa-depr-empty');
        const notice = document.getElementById('fa-depr-notice');
        if (!tbody) return;

        if (schedule.length === 0) {
            tbody.innerHTML = '';
            tfoot.innerHTML = '';
            if (empty) empty.style.display = '';
            if (notice) notice.style.display = 'none';
            return;
        }
        if (empty) empty.style.display = 'none';

        let totalDepreciation = 0;
        let hasSale = false;
        tbody.innerHTML = schedule.map(s => {
            totalDepreciation += s.depreciation_amount;
            if (s.disposal_type === '売却') hasSale = true;
            const remarkClass = s.disposal_remark ? (s.disposal_type === '売却' ? 'remark-sale' : 'remark-retire') : '';
            return `
                <tr ${s.disposal_type ? 'class="row-disposed"' : ''}>
                    <td>${s.asset_name}</td>
                    <td>${s.acquisition_date}</td>
                    <td class="text-right">${fmt(s.acquisition_cost)}</td>
                    <td>${s.useful_life}年</td>
                    <td>${s.depreciation_method}</td>
                    <td>${s.annual_rate ? (s.annual_rate * 100).toFixed(1) + '%' : '-'}</td>
                    <td class="text-right">${fmt(s.opening_book_value)}</td>
                    <td class="text-right">${fmt(s.depreciation_amount)}</td>
                    <td class="text-right">${fmt(s.closing_book_value)}</td>
                    <td class="${remarkClass}">${s.disposal_remark || ''}</td>
                </tr>`;
        }).join('');

        tfoot.innerHTML = `
            <tr style="font-weight:600;background:#f8fafc;">
                <td colspan="7" class="text-right">合計</td>
                <td class="text-right">${fmt(totalDepreciation)}</td>
                <td></td>
                <td></td>
            </tr>`;

        // Show/hide 譲渡所得 notice
        if (notice) notice.style.display = hasSale ? '' : 'none';
    }

    const faDeprCalcBtn = document.getElementById('fa-depr-calc-btn');
    if (faDeprCalcBtn) {
        faDeprCalcBtn.addEventListener('click', loadDepreciationSchedule);
    }

    // --- Generate depreciation journals ---
    const faDeprGenBtn = document.getElementById('fa-depr-generate-btn');
    if (faDeprGenBtn) {
        faDeprGenBtn.addEventListener('click', async () => {
            const yearEl = document.getElementById('fa-depr-year');
            const year = yearEl ? yearEl.value : '2025';
            if (!faDepreciationCache || faDepreciationCache.length === 0) {
                showToast('先に「計算」を実行してください', true);
                return;
            }
            // Check if any asset is missing asset_category
            const missing = faDepreciationCache.filter(s => !s.asset_category && s.depreciation_amount > 0);
            if (missing.length > 0) {
                const names = missing.map(s => s.asset_name).join('、');
                if (!confirm(`以下の資産に「資産区分」が設定されていません。\n${names}\n\nデフォルト（器具備品）として処理しますか？`)) return;
            }
            if (!confirm(`${year}年度の減価償却費・売却/除却の仕訳を仕訳帳に一括登録します。\nよろしいですか？`)) return;

            faDeprGenBtn.disabled = true;
            faDeprGenBtn.textContent = '生成中...';
            try {
                const res = await fetchAPI('/api/fixed-assets/generate-journals', 'POST', { fiscal_year: year });
                if (res.error) {
                    showToast(res.error, true);
                } else {
                    showToast(res.message || `${res.created}件の仕訳を生成しました`);
                }
            } catch (err) {
                showToast('仕訳生成に失敗しました', true);
            } finally {
                faDeprGenBtn.disabled = false;
                faDeprGenBtn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg> 仕訳を一括生成';
            }
        });
    }

    // --- Fixed assets output ---
    const faOutListCsv = document.getElementById('fa-out-list-csv');
    const faOutDeprCsv = document.getElementById('fa-out-depr-csv');
    const faOutDeprPdf = document.getElementById('fa-out-depr-pdf');

    if (faOutListCsv) {
        faOutListCsv.addEventListener('click', async () => {
            try {
                const data = await fetchAPI('/api/fixed-assets');
                const assets = data.assets || [];
                if (assets.length === 0) { showToast('データがありません', true); return; }
                let csv = '\uFEFF資産名称,取得日,取得原価,耐用年数,償却方法,資産区分,備考,売却/除却,処分日,売却額\n';
                assets.forEach(a => {
                    csv += `"${a.asset_name}",${a.acquisition_date},${a.acquisition_cost},${a.useful_life},"${a.depreciation_method}","${a.asset_category || ''}","${a.notes || ''}","${a.disposal_type || ''}","${a.disposal_date || ''}",${a.disposal_price || 0}\n`;
                });
                downloadBlob(csv, `固定資産一覧_${todayStr()}.csv`, 'text/csv;charset=utf-8');
                showToast('CSVをダウンロードしました');
            } catch (err) { showToast('エラー', true); }
        });
    }

    if (faOutDeprCsv) {
        faOutDeprCsv.addEventListener('click', async () => {
            const year = document.getElementById('fa-out-year')?.value || '2025';
            try {
                const data = await fetchAPI(`/api/fixed-assets/depreciation?fiscal_year=${year}`);
                const schedule = data.schedule || [];
                if (schedule.length === 0) { showToast('データがありません', true); return; }
                let csv = '\uFEFF資産名称,取得日,取得原価,耐用年数,償却方法,償却率,期首帳簿価額,本年分償却費,期末帳簿価額,備考\n';
                schedule.forEach(s => {
                    csv += `"${s.asset_name}",${s.acquisition_date},${s.acquisition_cost},${s.useful_life},"${s.depreciation_method}",${s.annual_rate},${s.opening_book_value},${s.depreciation_amount},${s.closing_book_value},"${s.disposal_remark || ''}"\n`;
                });
                downloadBlob(csv, `減価償却明細書_${year}年_${todayStr()}.csv`, 'text/csv;charset=utf-8');
                showToast('CSVをダウンロードしました');
            } catch (err) { showToast('エラー', true); }
        });
    }

    if (faOutDeprPdf) {
        faOutDeprPdf.addEventListener('click', async () => {
            const year = document.getElementById('fa-out-year')?.value || '2025';
            try {
                const data = await fetchAPI(`/api/fixed-assets/depreciation?fiscal_year=${year}`);
                const schedule = data.schedule || [];
                if (schedule.length === 0) { showToast('データがありません', true); return; }
                printDepreciationSchedule(schedule, year);
            } catch (err) { showToast('エラー', true); }
        });
    }

    function printDepreciationSchedule(schedule, year) {
        let totalDep = 0;
        let hasSale = false;
        const rows = schedule.map(s => {
            totalDep += s.depreciation_amount;
            if (s.disposal_type === '売却') hasSale = true;
            return `<tr>
                <td>${s.asset_name}</td>
                <td>${s.acquisition_date}</td>
                <td class="r">${fmt(s.acquisition_cost)}</td>
                <td>${s.useful_life}年</td>
                <td>${s.depreciation_method}</td>
                <td>${s.annual_rate ? (s.annual_rate * 100).toFixed(1) + '%' : '-'}</td>
                <td class="r">${fmt(s.opening_book_value)}</td>
                <td class="r">${fmt(s.depreciation_amount)}</td>
                <td class="r">${fmt(s.closing_book_value)}</td>
                <td style="font-size:10px;${s.disposal_type === '売却' ? 'color:#b91c1c;' : s.disposal_type === '除却' ? 'color:#92400e;' : ''}">${s.disposal_remark || ''}</td>
            </tr>`;
        }).join('');

        const saleNotice = hasSale ? `<p style="margin-top:12px;padding:8px 12px;background:#fffbeb;border:1px solid #fbbf24;font-size:11px;color:#92400e;">⚠️ 注意：売却損益は事業所得ではなく譲渡所得となります。確定申告書Bの「譲渡所得」欄に記載してください。</p>` : '';

        const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
        <title>減価償却費の計算明細書 ${year}年</title>
        <style>
            body{font-family:'Noto Sans JP',sans-serif;margin:20px;font-size:12px;}
            h1{font-size:16px;text-align:center;margin-bottom:4px;}
            h2{font-size:12px;text-align:center;color:#666;margin-bottom:16px;}
            table{width:100%;border-collapse:collapse;margin-top:12px;}
            th,td{border:1px solid #333;padding:6px 8px;font-size:11px;}
            th{background:#f0f0f0;font-weight:600;}
            .r{text-align:right;}
            tfoot td{font-weight:600;background:#f8f8f8;}
            @media print{body{margin:10mm;}}
        </style></head><body>
        <h1>減価償却費の計算明細書</h1>
        <h2>${year}年分</h2>
        <table>
            <thead><tr>
                <th>資産名称</th><th>取得日</th><th>取得原価</th><th>耐用年数</th>
                <th>償却方法</th><th>償却率</th><th>期首帳簿価額</th><th>本年分償却費</th><th>期末帳簿価額</th><th>備考</th>
            </tr></thead>
            <tbody>${rows}</tbody>
            <tfoot><tr>
                <td colspan="7" class="r">合計</td>
                <td class="r">${fmt(totalDep)}</td>
                <td></td>
                <td></td>
            </tr></tfoot>
        </table>
        ${saleNotice}
        <script>window.onload=()=>window.print();<\/script>
        </body></html>`;

        const w = window.open('', '_blank');
        w.document.write(html);
        w.document.close();
    }

    // ============================================================
    //  Section 15c: View — 消費税 (Consumption Tax)
    // ============================================================
    // --- Tax sub-tab navigation ---
    const taxSubNav = document.getElementById('tax-sub-nav');
    if (taxSubNav) {
        taxSubNav.addEventListener('click', (e) => {
            const btn = e.target.closest('.sub-tab-btn');
            if (!btn) return;
            const tabId = btn.dataset.taxTab;
            if (!tabId) return;
            taxSubNav.querySelectorAll('.sub-tab-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            document.querySelectorAll('.tax-tab-content').forEach(p => p.classList.remove('active'));
            const target = document.getElementById(tabId);
            if (target) target.classList.add('active');
            // Load data when switching to data tabs
            if (tabId === 'tax-summary') loadTaxSummary();
            if (tabId === 'tax-calc') loadTaxCalculation();
        });
    }

    // --- Tax settings state ---
    let taxSettings = {
        tax_business_type: 'exempt',
        tax_calculation_method: 'simplified',
        tax_simplified_category: '5',
    };

    // --- Toggle simplified section visibility ---
    function updateSimplifiedVisibility() {
        const section = document.getElementById('tax-simplified-section');
        const method = document.querySelector('input[name="tax-calc-method"]:checked');
        if (section && method) {
            section.style.display = method.value === 'simplified' ? '' : 'none';
        }
    }

    // --- Radio change handlers ---
    document.querySelectorAll('input[name="tax-calc-method"]').forEach(r => {
        r.addEventListener('change', updateSimplifiedVisibility);
    });

    // --- Load consumption tax view ---
    function loadConsumptionTax() {
        loadTaxSettings();
    }

    async function loadTaxSettings() {
        try {
            const data = await fetchAPI('/api/settings');
            taxSettings.tax_business_type = data.tax_business_type || 'exempt';
            taxSettings.tax_calculation_method = data.tax_calculation_method || 'simplified';
            taxSettings.tax_simplified_category = data.tax_simplified_category || '5';

            // Set radio buttons
            const btRadio = document.querySelector(`input[name="tax-business-type"][value="${taxSettings.tax_business_type}"]`);
            if (btRadio) btRadio.checked = true;
            const cmRadio = document.querySelector(`input[name="tax-calc-method"][value="${taxSettings.tax_calculation_method}"]`);
            if (cmRadio) cmRadio.checked = true;
            const catSelect = document.getElementById('tax-simplified-category');
            if (catSelect) catSelect.value = taxSettings.tax_simplified_category;

            updateSimplifiedVisibility();
            updateExemptNotices();
        } catch (err) {
            console.warn('loadTaxSettings failed:', err.message);
        }
    }

    function updateExemptNotices() {
        const isExempt = taxSettings.tax_business_type === 'exempt';
        const n1 = document.getElementById('tax-exempt-notice');
        const n2 = document.getElementById('tax-exempt-notice2');
        if (n1) n1.classList.toggle('hidden', !isExempt);
        if (n2) n2.classList.toggle('hidden', !isExempt);
    }

    // --- Save tax settings ---
    const taxSettingsSaveBtn = document.getElementById('tax-settings-save');
    if (taxSettingsSaveBtn) {
        taxSettingsSaveBtn.addEventListener('click', async () => {
            const bt = document.querySelector('input[name="tax-business-type"]:checked');
            const cm = document.querySelector('input[name="tax-calc-method"]:checked');
            const cat = document.getElementById('tax-simplified-category');
            const settings = {
                tax_business_type: bt ? bt.value : 'exempt',
                tax_calculation_method: cm ? cm.value : 'simplified',
                tax_simplified_category: cat ? cat.value : '5',
            };
            try {
                await fetchAPI('/api/settings', 'POST', settings);
                taxSettings = settings;
                updateExemptNotices();
                showToast('消費税設定を保存しました');
            } catch (err) {
                showToast('設定の保存に失敗しました', true);
            }
        });
    }

    // --- Set default dates for tax period inputs ---
    const taxDateInputs = [
        ['tax-start', 'tax-end'],
        ['tax-calc-start', 'tax-calc-end'],
    ];
    function applyFiscalYearToTaxInputs() {
        const fy = getSelectedFiscalYear();
        taxDateInputs.forEach(([startId, endId]) => {
            const s = document.getElementById(startId);
            const e = document.getElementById(endId);
            if (s) s.value = `${fy}-01-01`;
            if (e) e.value = `${fy}-12-31`;
        });
    }
    applyFiscalYearToTaxInputs();

    // --- Load tax summary (科目別消費税一覧) ---
    const taxSummaryLoadBtn = document.getElementById('tax-summary-load');
    if (taxSummaryLoadBtn) {
        taxSummaryLoadBtn.addEventListener('click', loadTaxSummary);
    }

    async function loadTaxSummary() {
        const startDate = document.getElementById('tax-start').value;
        const endDate = document.getElementById('tax-end').value;
        if (!startDate || !endDate) { showToast('期間を指定してください', true); return; }

        try {
            const p = new URLSearchParams({ start_date: startDate, end_date: endDate });
            const data = await fetchAPI('/api/tax/summary?' + p.toString());

            renderTaxSalesTable(data.sales || [], data.sales_agg || {});
            renderTaxPurchaseTable(data.purchases || [], data.purchase_agg || {});
            renderTaxNontaxTable(data.sales || [], data.purchases || []);
        } catch (err) {
            showToast('消費税集計に失敗しました', true);
        }
    }

    function renderTaxSalesTable(rows, agg) {
        const tbody = document.getElementById('tax-sales-tbody');
        const tfoot = document.getElementById('tax-sales-tfoot');
        if (!tbody) return;
        const taxableRows = rows.filter(r => r.tax_classification === '10%' || r.tax_classification === '8%');
        if (taxableRows.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" class="text-muted">該当データなし</td></tr>';
            tfoot.innerHTML = '';
            return;
        }
        tbody.innerHTML = taxableRows.map(r => {
            const net = r.total_amount - r.total_tax;
            return `<tr class="tax-drill-row" data-account-id="${r.account_id || ''}" data-tax-class="${r.tax_classification}" data-side="credit" data-name="${r.name}" style="cursor:pointer;"><td>${r.name}</td><td class="text-right">${fmt(r.total_amount)}</td><td>${r.tax_classification}</td><td class="text-right">${fmt(net)}</td><td class="text-right">${fmt(r.total_tax)}</td></tr>`;
        }).join('');
        tfoot.innerHTML = `<tr class="tax-tfoot-row"><td><strong>合計</strong></td><td class="text-right"><strong>${fmt(agg.taxable_total)}</strong></td><td></td><td class="text-right"><strong>${fmt(agg.net_total)}</strong></td><td class="text-right"><strong>${fmt(agg.tax_total)}</strong></td></tr>`;
    }

    function renderTaxPurchaseTable(rows, agg) {
        const tbody = document.getElementById('tax-purchase-tbody');
        const tfoot = document.getElementById('tax-purchase-tfoot');
        if (!tbody) return;
        const taxableRows = rows.filter(r => r.tax_classification === '10%' || r.tax_classification === '8%');
        if (taxableRows.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" class="text-muted">該当データなし</td></tr>';
            tfoot.innerHTML = '';
            return;
        }
        tbody.innerHTML = taxableRows.map(r => {
            const net = r.total_amount - r.total_tax;
            return `<tr class="tax-drill-row" data-account-id="${r.account_id || ''}" data-tax-class="${r.tax_classification}" data-side="debit" data-name="${r.name}" style="cursor:pointer;"><td>${r.name}</td><td class="text-right">${fmt(r.total_amount)}</td><td>${r.tax_classification}</td><td class="text-right">${fmt(net)}</td><td class="text-right">${fmt(r.total_tax)}</td></tr>`;
        }).join('');
        tfoot.innerHTML = `<tr class="tax-tfoot-row"><td><strong>合計</strong></td><td class="text-right"><strong>${fmt(agg.taxable_total)}</strong></td><td></td><td class="text-right"><strong>${fmt(agg.net_total)}</strong></td><td class="text-right"><strong>${fmt(agg.tax_total)}</strong></td></tr>`;
    }

    function renderTaxNontaxTable(sales, purchases) {
        const tbody = document.getElementById('tax-nontax-tbody');
        if (!tbody) return;
        const nontaxSales = sales.filter(r => r.tax_classification === '非課税' || r.tax_classification === '不課税');
        const nontaxPurchases = purchases.filter(r => r.tax_classification === '非課税' || r.tax_classification === '不課税');
        const allNontax = [
            ...nontaxSales.map(r => ({ ...r, side: r.account_type === '収益' ? 'credit' : 'debit' })),
            ...nontaxPurchases.map(r => ({ ...r, side: 'debit' })),
        ];
        if (allNontax.length === 0) {
            tbody.innerHTML = '<tr><td colspan="3" class="text-muted">該当データなし</td></tr>';
            return;
        }
        tbody.innerHTML = allNontax.map(r => {
            const sideLabel = r.side === 'credit' ? '売上' : '仕入';
            return `<tr class="tax-drill-row" data-account-id="${r.account_id || ''}" data-tax-class="${r.tax_classification}" data-side="${r.side}" data-name="${r.name}" style="cursor:pointer;"><td>${r.name}（${sideLabel}）</td><td>${r.tax_classification}</td><td class="text-right">${fmt(r.total_amount)}</td></tr>`;
        }).join('');
    }

    // --- Tax drill-down: show journal entries for a clicked account ---
    const taxDrillPanel = document.getElementById('tax-drill-panel');
    const taxDrillBack = document.getElementById('tax-drill-back');
    const taxDrillTitle = document.getElementById('tax-drill-title');
    const taxDrillTbody = document.getElementById('tax-drill-tbody');
    const taxDrillPager = document.getElementById('tax-drill-pager');
    let taxDrillCurrentParams = {};

    // Delegate click on tax summary table rows
    document.getElementById('view-consumption-tax').addEventListener('click', (e) => {
        const row = e.target.closest('.tax-drill-row');
        if (!row) return;
        const accountId = row.dataset.accountId;
        const taxClass = row.dataset.taxClass;
        const side = row.dataset.side;
        const name = row.dataset.name;
        if (!accountId) return;

        taxDrillCurrentParams = { accountId, taxClass, side, name, page: 1 };
        loadTaxDrillDown();
    });

    if (taxDrillBack) {
        taxDrillBack.addEventListener('click', () => {
            taxDrillPanel.classList.add('hidden');
        });
    }

    async function loadTaxDrillDown(page = 1) {
        const { accountId, taxClass, side, name } = taxDrillCurrentParams;
        const startDate = document.getElementById('tax-start').value;
        const endDate = document.getElementById('tax-end').value;

        const p = new URLSearchParams({
            account_id: accountId,
            account_side: side,
            tax_classification: taxClass,
            per_page: '50',
            page: String(page),
        });
        if (startDate) p.set('start_date', startDate);
        if (endDate) p.set('end_date', endDate);

        try {
            const data = await fetchAPI('/api/journal?' + p.toString());
            const entries = data.entries || [];
            taxDrillTitle.textContent = `${name}（${taxClass}）の仕訳明細`;
            taxDrillPanel.classList.remove('hidden');

            if (!entries.length) {
                taxDrillTbody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:2rem;color:var(--text-dim);">該当する仕訳がありません</td></tr>';
                taxDrillPager.innerHTML = '';
                return;
            }

            taxDrillTbody.innerHTML = entries.map(e => {
                return `<tr data-entry-id="${e.id}" style="cursor:pointer;">
                    <td>${e.entry_date || ''}</td>
                    <td>${e.debit_account || ''}</td>
                    <td>${e.credit_account || ''}</td>
                    <td class="text-right">${fmt(e.amount)}</td>
                    <td>${e.tax_classification || ''}</td>
                    <td>${e.counterparty || ''}</td>
                    <td>${e.memo || ''}</td>
                </tr>`;
            }).join('');

            // Click row to open edit modal
            taxDrillTbody.querySelectorAll('tr[data-entry-id]').forEach(row => {
                row.addEventListener('click', () => {
                    const id = row.dataset.entryId;
                    const entry = entries.find(en => String(en.id) === String(id));
                    if (entry) openJEDetailModal(entry, () => {
                        loadTaxDrillDown(page);
                        loadTaxSummary();
                    });
                });
            });

            // Pager
            const total = data.total || 0;
            const perPage = data.per_page || 50;
            const totalPages = Math.ceil(total / perPage);
            if (totalPages > 1) {
                let pagerHtml = '';
                for (let i = 1; i <= totalPages; i++) {
                    pagerHtml += `<button class="btn btn-sm ${i === page ? 'btn-primary' : 'btn-ghost'}" data-page="${i}">${i}</button> `;
                }
                taxDrillPager.innerHTML = pagerHtml;
                taxDrillPager.querySelectorAll('button').forEach(btn => {
                    btn.addEventListener('click', () => {
                        loadTaxDrillDown(parseInt(btn.dataset.page));
                    });
                });
            } else {
                taxDrillPager.innerHTML = '';
            }
        } catch (err) {
            showToast('仕訳明細の読み込みに失敗しました', true);
        }
    }

    // --- Load tax calculation (消費税計算) ---
    const taxCalcLoadBtn = document.getElementById('tax-calc-load');
    if (taxCalcLoadBtn) {
        taxCalcLoadBtn.addEventListener('click', loadTaxCalculation);
    }

    async function loadTaxCalculation() {
        const startDate = document.getElementById('tax-calc-start').value;
        const endDate = document.getElementById('tax-calc-end').value;
        if (!startDate || !endDate) { showToast('期間を指定してください', true); return; }

        try {
            const p = new URLSearchParams({
                start_date: startDate,
                end_date: endDate,
                method: taxSettings.tax_calculation_method,
                simplified_category: taxSettings.tax_simplified_category,
            });
            const data = await fetchAPI('/api/tax/calculation?' + p.toString());
            renderTaxCalculation(data);
        } catch (err) {
            showToast('消費税計算に失敗しました', true);
        }
    }

    function renderTaxCalculation(data) {
        const s = data.standard || {};
        const si = data.simplified || {};

        // Standard method
        document.getElementById('tc-sales-10').textContent = fmt(s.sales_net_10);
        document.getElementById('tc-sales-8').textContent = fmt(s.sales_net_8);
        document.getElementById('tc-sales-total').textContent = fmt(s.sales_net_total);
        document.getElementById('tc-tax-sales-10').textContent = fmt(s.sales_tax_10);
        document.getElementById('tc-tax-sales-8').textContent = fmt(s.sales_tax_8);
        document.getElementById('tc-tax-sales-total').textContent = fmt(s.sales_tax_total);
        document.getElementById('tc-purchase-10').textContent = fmt(s.purchase_net_10);
        document.getElementById('tc-purchase-8').textContent = fmt(s.purchase_net_8);
        document.getElementById('tc-purchase-total').textContent = fmt(s.purchase_net_total);
        document.getElementById('tc-tax-purchase-10').textContent = fmt(s.purchase_tax_10);
        document.getElementById('tc-tax-purchase-8').textContent = fmt(s.purchase_tax_8);
        document.getElementById('tc-tax-purchase-total').textContent = fmt(s.purchase_tax_total);
        document.getElementById('tc-national-tax').textContent = fmt(s.national_tax);
        document.getElementById('tc-local-tax').textContent = fmt(s.local_tax);
        document.getElementById('tc-total-due').textContent = fmt(s.total_due);

        // Simplified method
        document.getElementById('ts-sales').textContent = fmt(si.sales_net);
        document.getElementById('ts-tax-sales').textContent = fmt(si.tax_on_sales);
        document.getElementById('ts-category').textContent = si.category_label || '';
        document.getElementById('ts-deemed-rate').textContent = (si.deemed_rate || 0) + '%';
        document.getElementById('ts-deemed-credit').textContent = fmt(si.deemed_credit);
        document.getElementById('ts-national-tax').textContent = fmt(si.national_tax);
        document.getElementById('ts-local-tax').textContent = fmt(si.local_tax);
        document.getElementById('ts-total-due').textContent = fmt(si.total_due);

        // Comparison
        const compSection = document.getElementById('tax-calc-comparison');
        if (compSection) {
            compSection.style.display = '';
            document.getElementById('tc-comp-standard').textContent = fmt(s.total_due) + ' 円';
            document.getElementById('tc-comp-simplified').textContent = fmt(si.total_due) + ' 円';
            const diff = (s.total_due || 0) - (si.total_due || 0);
            const diffEl = document.getElementById('tc-comp-diff');
            diffEl.textContent = (diff >= 0 ? '+' : '') + fmt(diff) + ' 円';
            diffEl.style.color = diff > 0 ? '#dc2626' : '#16a34a';
        }

        // Show/hide sections based on user's method setting
        const stdSection = document.getElementById('tax-calc-standard-section');
        const simpSection = document.getElementById('tax-calc-simplified-section');
        if (taxSettings.tax_calculation_method === 'standard') {
            if (stdSection) stdSection.style.display = '';
            if (simpSection) simpSection.style.display = '';
        } else {
            if (stdSection) stdSection.style.display = '';
            if (simpSection) simpSection.style.display = '';
        }
    }

    // --- Tax output (アウトプット) ---
    // CSV
    const taxOutSummaryCsvBtn = document.getElementById('tax-out-summary-csv');
    if (taxOutSummaryCsvBtn) {
        taxOutSummaryCsvBtn.addEventListener('click', async () => {
            const startDate = outStartInput.value;
            const endDate = outEndInput.value;
            if (!startDate || !endDate) { showToast('期間を指定してください', true); return; }
            try {
                const p = new URLSearchParams({ start_date: startDate, end_date: endDate });
                const data = await fetchAPI('/api/tax/summary?' + p.toString());
                const sa = data.sales_agg || {};
                const pa = data.purchase_agg || {};

                let csv = '\ufeff勘定科目,税込金額,税率,税抜金額,消費税額,区分\n';
                (data.sales || []).forEach(r => {
                    csv += `${r.name},${r.total_amount},${r.tax_classification},${r.total_amount - r.total_tax},${r.total_tax},課税売上\n`;
                });
                csv += `課税売上合計,${sa.taxable_total},,${sa.net_total},${sa.tax_total},\n`;
                csv += '\n';
                (data.purchases || []).forEach(r => {
                    csv += `${r.name},${r.total_amount},${r.tax_classification},${r.total_amount - r.total_tax},${r.total_tax},課税仕入\n`;
                });
                csv += `課税仕入合計,${pa.taxable_total},,${pa.net_total},${pa.tax_total},\n`;

                const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
                downloadBlob(blob, `科目別消費税一覧_${todayStr()}.csv`, 'text/csv');
                showToast('CSVをダウンロードしました');
            } catch (err) { showToast('エクスポートに失敗しました', true); }
        });
    }

    // Summary PDF
    const taxOutSummaryPdfBtn = document.getElementById('tax-out-summary-pdf');
    if (taxOutSummaryPdfBtn) {
        taxOutSummaryPdfBtn.addEventListener('click', async () => {
            const startDate = outStartInput.value;
            const endDate = outEndInput.value;
            if (!startDate || !endDate) { showToast('期間を指定してください', true); return; }
            try {
                const p = new URLSearchParams({ start_date: startDate, end_date: endDate });
                const data = await fetchAPI('/api/tax/summary?' + p.toString());
                openTaxSummaryPrintView(startDate, endDate, data);
            } catch (err) { showToast('エクスポートに失敗しました', true); }
        });
    }

    function openTaxSummaryPrintView(startDate, endDate, data) {
        const sa = data.sales_agg || {};
        const pa = data.purchase_agg || {};

        function buildTable(rows, agg, title) {
            const taxableRows = rows.filter(r => r.tax_classification === '10%' || r.tax_classification === '8%');
            let html = `<h3 style="margin:16px 0 6px;">${title}</h3>`;
            html += '<table><thead><tr><th>勘定科目</th><th class="r">税込金額</th><th>税率</th><th class="r">税抜金額</th><th class="r">消費税額</th></tr></thead><tbody>';
            taxableRows.forEach(r => {
                const net = r.total_amount - r.total_tax;
                html += `<tr><td>${r.name}</td><td class="r">${fmt(r.total_amount)}</td><td>${r.tax_classification}</td><td class="r">${fmt(net)}</td><td class="r">${fmt(r.total_tax)}</td></tr>`;
            });
            html += `</tbody><tfoot><tr style="font-weight:bold;background:#f1f5f9;"><td>合計</td><td class="r">${fmt(agg.taxable_total)}</td><td></td><td class="r">${fmt(agg.net_total)}</td><td class="r">${fmt(agg.tax_total)}</td></tr></tfoot></table>`;
            return html;
        }

        let tableHtml = buildTable(data.sales || [], sa, '課税売上');
        tableHtml += buildTable(data.purchases || [], pa, '課税仕入');

        openPrintView('科目別消費税一覧', `${startDate} ～ ${endDate}`, tableHtml);
    }

    // Calculation PDF
    const taxOutCalcPdfBtn = document.getElementById('tax-out-calc-pdf');
    if (taxOutCalcPdfBtn) {
        taxOutCalcPdfBtn.addEventListener('click', async () => {
            const startDate = outStartInput.value;
            const endDate = outEndInput.value;
            if (!startDate || !endDate) { showToast('期間を指定してください', true); return; }
            try {
                const p = new URLSearchParams({
                    start_date: startDate,
                    end_date: endDate,
                    method: taxSettings.tax_calculation_method,
                    simplified_category: taxSettings.tax_simplified_category,
                });
                const data = await fetchAPI('/api/tax/calculation?' + p.toString());
                openTaxCalcPrintView(startDate, endDate, data);
            } catch (err) { showToast('エクスポートに失敗しました', true); }
        });
    }

    function openTaxCalcPrintView(startDate, endDate, data) {
        const s = data.standard || {};
        const si = data.simplified || {};

        function row(label, val) {
            return `<tr><td style="padding:6px 12px;">${label}</td><td class="r" style="padding:6px 12px;">${fmt(val)}</td></tr>`;
        }
        function headerRow(label) {
            return `<tr><th colspan="2" style="background:#e2e8f0;padding:6px 12px;text-align:left;font-size:12px;">${label}</th></tr>`;
        }
        function totalRow(label, val) {
            return `<tr style="font-weight:bold;background:#f1f5f9;font-size:13px;"><td style="padding:8px 12px;">${label}</td><td class="r" style="padding:8px 12px;">${fmt(val)} 円</td></tr>`;
        }

        let html = '<h3 style="margin:12px 0 6px;">本則課税による計算</h3>';
        html += '<table style="margin-bottom:20px;"><tbody>';
        html += headerRow('課税売上');
        html += row('課税売上高（税抜・10%）', s.sales_net_10);
        html += row('課税売上高（税抜・8%）', s.sales_net_8);
        html += row('売上に係る消費税額（10%）', s.sales_tax_10);
        html += row('売上に係る消費税額（8%）', s.sales_tax_8);
        html += headerRow('課税仕入');
        html += row('課税仕入高（税抜・10%）', s.purchase_net_10);
        html += row('課税仕入高（税抜・8%）', s.purchase_net_8);
        html += row('仕入に係る消費税額（10%）', s.purchase_tax_10);
        html += row('仕入に係る消費税額（8%）', s.purchase_tax_8);
        html += headerRow('納付税額');
        html += row('消費税額（国税）', s.national_tax);
        html += row('地方消費税額', s.local_tax);
        html += totalRow('納付すべき消費税額 合計', s.total_due);
        html += '</tbody></table>';

        html += '<h3 style="margin:20px 0 6px;">簡易課税による計算</h3>';
        html += '<table><tbody>';
        html += row('課税売上高（税抜）', si.sales_net);
        html += row('売上に係る消費税額', si.tax_on_sales);
        html += row('事業区分', si.category_label);
        html += row('みなし仕入率', (si.deemed_rate || 0) + '%');
        html += row('控除対象仕入税額（みなし）', si.deemed_credit);
        html += row('消費税額（国税）', si.national_tax);
        html += row('地方消費税額', si.local_tax);
        html += totalRow('納付すべき消費税額 合計', si.total_due);
        html += '</tbody></table>';

        openPrintView('消費税計算書', `${startDate} ～ ${endDate}`, html);
    }

    // ============================================================
    //  Section 16: Keyboard Shortcuts
    // ============================================================
    document.addEventListener('keydown', (e) => {
        // Ctrl+Enter to submit journal form when in journal-input view
        if (e.ctrlKey && e.key === 'Enter') {
            const activeView = document.querySelector('.content-view.active');
            if (activeView && activeView.id === 'view-journal-input') {
                journalForm.dispatchEvent(new Event('submit'));
            }
        }
        // Escape to close chat or go back to menu
        if (e.key === 'Escape') {
            if (!chatPanel.classList.contains('hidden')) {
                chatPanel.classList.add('hidden');
            } else if (!settingsModal.classList.contains('hidden')) {
                settingsModal.classList.add('hidden');
            } else if (!menuGrid.classList.contains('active')) {
                showMenu();
            }
        }
    });

    // ============================================================
    //  Section: License Activation
    // ============================================================
    const licenseOverlay = document.getElementById('license-overlay');
    const licenseKeyInput = document.getElementById('license-key-input');
    const licenseActivateBtn = document.getElementById('license-activate-btn');
    const licenseLogoutBtn = document.getElementById('license-logout-btn');

    function showLicenseOverlay() {
        if (licenseOverlay) licenseOverlay.classList.remove('hidden');
    }
    function hideLicenseOverlay() {
        if (licenseOverlay) licenseOverlay.classList.add('hidden');
    }

    if (licenseActivateBtn) {
        licenseActivateBtn.addEventListener('click', async () => {
            const key = (licenseKeyInput.value || '').trim().toUpperCase();
            if (!key) {
                showToast('ライセンスキーを入力してください', true);
                return;
            }
            licenseActivateBtn.disabled = true;
            licenseActivateBtn.textContent = '確認中...';
            try {
                const res = await fetchAPI('/api/license/activate', 'POST', { license_key: key });
                if (res.status === 'success') {
                    showToast('ライセンスが有効化されました！');
                    hideLicenseOverlay();
                    licenseKeyInput.value = '';
                    // Reload app data
                    onLoginSuccess();
                } else {
                    showToast(res.error || 'ライセンスの有効化に失敗しました', true);
                }
            } catch (e) {
                if (e.message !== 'License required') {
                    showToast('エラーが発生しました', true);
                }
            } finally {
                licenseActivateBtn.disabled = false;
                licenseActivateBtn.textContent = 'ライセンスを有効化';
            }
        });
    }

    if (licenseLogoutBtn) {
        licenseLogoutBtn.addEventListener('click', () => {
            hideLicenseOverlay();
            handleLogout();
        });
    }

    // Allow Enter key to activate license
    if (licenseKeyInput) {
        licenseKeyInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && licenseActivateBtn) licenseActivateBtn.click();
        });
    }

    // ============================================================
    //  Register global fiscal year change callbacks
    // ============================================================
    fiscalYearChangeCallbacks.push(() => {
        _latestDateCache = {}; // Clear cache when year changes
        applyFiscalYearConstraint(); // Update journal input date range
        applyFiscalYearToOutput();   // Update output date range
        applyFiscalYearToTaxInputs(); // Update tax date range
        updateHeaderFiscalYear();    // Update header fiscal year badge
    });

    // ============================================================
    //  Initial route from hash (must be at end after all declarations)
    // ============================================================
    const initHash = location.hash.replace('#', '');
    if (initHash && initHash !== 'menu') {
        showView(initHash);
    }
});
