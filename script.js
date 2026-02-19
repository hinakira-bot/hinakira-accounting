document.addEventListener('DOMContentLoaded', () => {
    // ============================================================
    //  Section 1: Constants & State
    // ============================================================
    const CLIENT_ID = '353694435064-r6mlbk3mm2mflhl2mot2n94dpuactscc.apps.googleusercontent.com';
    const SCOPES = 'openid email profile https://www.googleapis.com/auth/drive';
    let tokenClient;
    let accessToken = sessionStorage.getItem('access_token');
    let tokenExpiration = sessionStorage.getItem('token_expiration');
    let accounts = [];            // Account master cache
    let scanResults = [];         // Scan tab working data
    let isLoggingOut = false;     // Prevent 401 cascade (multiple toasts)
    let refreshTimer = null;      // Token auto-refresh timer
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

    // ============================================================
    //  Section 3: Google OAuth
    // ============================================================
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
                    const expiresIn = resp.expires_in || 3600;
                    const exp = new Date().getTime() + (expiresIn * 1000);
                    sessionStorage.setItem('access_token', accessToken);
                    sessionStorage.setItem('token_expiration', exp);
                    scheduleTokenRefresh(expiresIn);
                    onLoginSuccess();
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

    function handleLogin() { tokenClient && tokenClient.requestAccessToken(); }
    function handleLogout() {
        if (refreshTimer) { clearTimeout(refreshTimer); refreshTimer = null; }
        const t = sessionStorage.getItem('access_token');
        // Clear session first
        accessToken = null;
        isLoggingOut = false;
        sessionStorage.removeItem('access_token');
        sessionStorage.removeItem('token_expiration');
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
            if (tokenClient) {
                // prompt: '' means silent refresh (no popup if user already consented)
                tokenClient.requestAccessToken({ prompt: '' });
            }
        }, refreshAt);
    }
    async function onLoginSuccess() {
        isLoggingOut = false;  // Reset 401 cascade guard
        loginOverlay.classList.add('hidden');
        authBtn.textContent = 'ログアウト';
        authBtn.onclick = handleLogout;
        settingsBtn.style.display = '';

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
        'journal-book': () => loadJournalBook(),
        'ledger': () => loadCurrentLedgerSubTab(),
        'counterparty': () => loadCounterpartyList(),
        'accounts': () => loadAccountsList(),
        'opening-balance': () => loadOpeningBalances(),
        'backup': () => {},
        'output': () => {},
    };

    function showMenu() {
        // Hide all content views
        document.querySelectorAll('.content-view').forEach(v => v.classList.remove('active'));
        // Show menu grid
        menuGrid.classList.add('active');
        // Hide back button, show logo
        backToMenuBtn.classList.add('hidden');
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
        // Show back button, hide logo
        backToMenuBtn.classList.remove('hidden');
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
    async function fetchAPI(url, method = 'GET', body = null) {
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
            // Token expired or invalid — show login overlay (only once)
            if (!isLoggingOut) {
                isLoggingOut = true;
                showToast('セッションが切れました。再ログインしてください', true);
                accessToken = null;
                sessionStorage.removeItem('access_token');
                sessionStorage.removeItem('token_expiration');
                loginOverlay.classList.remove('hidden');
                authBtn.textContent = 'Googleでログイン';
                authBtn.onclick = handleLogin;
                const userDisplay = document.getElementById('user-display');
                if (userDisplay) userDisplay.classList.add('hidden');
            }
            throw new Error('Unauthorized');
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

    // Default date to today
    jeDate.value = todayStr();

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

        // Prior year date check for manual input
        if (isPriorYear(entry.entry_date)) {
            const origDate = entry.entry_date;
            if (!confirm(`⚠️ 前年以前の日付（${origDate}）が入力されています。\n\n当年1月1日（${currentYear}-01-01）の仕訳として登録され、摘要に実際の日付が記録されます。\n\n登録しますか？`)) {
                return;
            }
        }

        try {
            const res = await fetchAPI('/api/journal', 'POST', entry);
            if (res.status === 'success') {
                showToast('仕訳を登録しました');
                journalForm.reset();
                jeDate.value = todayStr();
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
    function loadRecentEntries() {
        fetchAPI('/api/journal/recent?limit=10').then(data => {
            const tbody = document.getElementById('recent-tbody');
            if (!tbody) return;
            tbody.innerHTML = (data.entries || []).map(e => {
                const parsed = parseTaxClassification(e.tax_classification);
                return `
                <tr>
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

    // Clear all scanned results
    scanClearAll.addEventListener('click', () => {
        if (!scanResults.length) return;
        if (!confirm('解析結果をすべて削除しますか？')) return;
        scanResults = [];
        scanTbody.innerHTML = '';
        scanResultsCard.classList.add('hidden');
        showToast('解析結果をすべて削除しました');
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
        scanStatusText.textContent = `${files.length}件のファイルをAIで解析中...`;

        const formData = new FormData();
        for (let f of files) formData.append('files', f);
        formData.append('gemini_api_key', apiKey);
        if (accessToken) formData.append('access_token', accessToken);

        try {
            const data = await fetchAPI('/api/analyze', 'POST', formData);
            if (data.error) throw new Error(data.error);
            scanResults = append ? [...scanResults, ...data] : data;
            renderScanResults();
            scanStatus.classList.add('hidden');
            scanResultsCard.classList.remove('hidden');
        } catch (err) {
            scanStatus.classList.add('hidden');
            showToast('解析エラー: ' + err.message, true);
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
        return y < currentYear;
    }

    function renderScanResults() {
        let hasDup = false;
        let hasPriorYear = false;
        scanTbody.innerHTML = scanResults.map((item, i) => {
            if (item.is_duplicate) hasDup = true;
            const priorYear = isPriorYear(item.date);
            if (priorYear) hasPriorYear = true;
            return `
            <tr class="${item.is_duplicate ? 'row-duplicate' : ''} ${priorYear ? 'row-prior-year' : ''}">
                <td><input type="date" value="${item.date || ''}" data-i="${i}" data-k="date" class="scan-input ${priorYear ? 'input-prior-year' : ''}"></td>
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
                scanResults[idx][e.target.dataset.k] = e.target.value;
            });
        });
        scanTbody.querySelectorAll('.scan-delete').forEach(el => {
            el.addEventListener('click', (e) => {
                scanResults.splice(parseInt(e.target.dataset.i), 1);
                renderScanResults();
            });
        });
    }

    scanSaveBtn.addEventListener('click', async () => {
        const valid = scanResults.filter(r => parseInt(r.amount) > 0);
        if (!valid.length) { showToast('保存するデータがありません', true); return; }

        // Determine source: csv_import for statement-parsed entries, ai_receipt for others
        const hasCsvEntries = valid.some(r => r.source === 'csv_import');
        const entries = valid.map(r => ({
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
                    const dates = valid.filter(r => r.source === 'csv_import' && r.date).map(r => r.date).sort();
                    fetchAPI('/api/statement/history', 'POST', {
                        filename: lastStatementParse.filename || '',
                        file_hash: lastStatementParse.file_hash || '',
                        source_name: lastStatementParse.source_name || '',
                        row_count: lastStatementParse.total_rows || 0,
                        imported_count: valid.filter(r => r.source === 'csv_import').length,
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
                        file_ids: driveFileIdsInScan
                    }).then(() => {
                        driveFileIdsInScan = [];
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

            // Track file IDs for moving after save
            driveFileIdsInScan = [...new Set(results.map(r => r.drive_file_id).filter(Boolean))];

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

    // Populate year/month selects
    for (let y = thisYear - 3; y <= thisYear + 1; y++) {
        const opt = document.createElement('option');
        opt.value = y; opt.textContent = y + '年';
        if (y === thisYear) opt.selected = true;
        jbYearSelect.appendChild(opt);
    }
    for (let m = 1; m <= 12; m++) {
        const opt = document.createElement('option');
        opt.value = m; opt.textContent = m + '月';
        if (m === thisMonth) opt.selected = true;
        jbMonthSelect.appendChild(opt);
    }

    function applyJBPeriod() {
        const y = parseInt(jbYearSelect.value) || thisYear;
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

    function renderJournalBook(data) {
        const entries = data.entries || [];
        jbEntriesCache = entries;
        jbTbody.innerHTML = entries.map(e => `
            <tr data-id="${e.id}" style="cursor:pointer;">
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
            jbTbody.innerHTML = '<tr><td colspan="9" style="text-align:center;color:var(--text-dim);padding:2rem;">仕訳データがありません</td></tr>';
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
                if (e.target.closest('a') || e.target.closest('button')) return;
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

    // --- Year & Month selector population ---
    for (let y = thisYear - 3; y <= thisYear + 1; y++) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.textContent = y + '年';
        if (y === thisYear) opt.selected = true;
        ledgerYearSelect.appendChild(opt);
    }
    for (let m = 1; m <= 12; m++) {
        const opt = document.createElement('option');
        opt.value = m;
        opt.textContent = m + '月';
        if (m === thisMonth) opt.selected = true;
        ledgerMonthSelect.appendChild(opt);
    }

    function getSelectedYear() { return parseInt(ledgerYearSelect.value) || thisYear; }
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
            const res = await fetch('/api/backup/download?format=json');
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
            const res = await fetch('/api/backup/drive', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ access_token: accessToken })
            });
            const result = await res.json();
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
            const res = await fetch('/api/backup/restore', { method: 'POST', body: formData });
            const result = await res.json();
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
            const res = await fetch('/api/backup/drive/list', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ access_token: accessToken })
            });
            const result = await res.json();
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
            const res = await fetch('/api/backup/drive/restore', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ access_token: accessToken, file_id: fileId })
            });
            const result = await res.json();
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

    outStartInput.value = `${thisYear}-01-01`;
    outEndInput.value = `${thisYear}-12-31`;

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
    let chatHistory = [];

    chatFab.addEventListener('click', () => {
        chatPanel.classList.toggle('hidden');
        if (!chatPanel.classList.contains('hidden')) {
            chatInput.focus();
        }
    });
    chatClose.addEventListener('click', () => chatPanel.classList.add('hidden'));

    async function sendChatMessage() {
        const msg = chatInput.value.trim();
        if (!msg) return;

        const apiKey = localStorage.getItem('gemini_api_key');
        if (!apiKey) {
            addChatMsg('bot', 'Gemini APIキーが設定されていません。右上の設定からAPIキーを登録してください。');
            return;
        }

        // Add user message
        addChatMsg('user', msg);
        chatInput.value = '';
        chatHistory.push({ role: 'user', text: msg });

        // Show loading
        const loadingEl = addChatMsg('loading', '考え中...');
        chatSendBtn.disabled = true;

        try {
            const data = await fetchAPI('/api/chat', 'POST', {
                message: msg,
                history: chatHistory,
                gemini_api_key: apiKey,
            });

            // Remove loading
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

    function addChatMsg(type, text) {
        const div = document.createElement('div');
        div.className = `ai-msg ai-msg-${type}`;
        div.textContent = text;
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
    //  Initial route from hash (must be at end after all declarations)
    // ============================================================
    const initHash = location.hash.replace('#', '');
    if (initHash && initHash !== 'menu') {
        showView(initHash);
    }
});
