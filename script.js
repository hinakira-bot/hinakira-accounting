document.addEventListener('DOMContentLoaded', () => {
    // ============================================================
    //  Section 1: Constants & State
    // ============================================================
    const CLIENT_ID = '353694435064-r6mlbk3mm2mflhl2mot2n94dpuactscc.apps.googleusercontent.com';
    const SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file';
    let tokenClient;
    let accessToken = sessionStorage.getItem('access_token');
    let tokenExpiration = sessionStorage.getItem('token_expiration');
    let accounts = [];            // Account master cache
    let scanResults = [];         // Scan tab working data

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
                    const exp = new Date().getTime() + (resp.expires_in * 1000);
                    sessionStorage.setItem('access_token', accessToken);
                    sessionStorage.setItem('token_expiration', exp);
                    onLoginSuccess();
                }
            },
        });
        if (accessToken && tokenExpiration && new Date().getTime() < parseInt(tokenExpiration)) {
            onLoginSuccess();
        } else {
            loginOverlay.classList.remove('hidden');
        }
    };

    function handleLogin() { tokenClient && tokenClient.requestAccessToken(); }
    function handleLogout() {
        const t = sessionStorage.getItem('access_token');
        if (t) google.accounts.oauth2.revoke(t, () => {});
        sessionStorage.removeItem('access_token');
        sessionStorage.removeItem('token_expiration');
        location.reload();
    }
    function onLoginSuccess() {
        loginOverlay.classList.add('hidden');
        authBtn.textContent = 'ログアウト';
        authBtn.onclick = handleLogout;
        settingsBtn.style.display = '';
        const apiKey = localStorage.getItem('gemini_api_key');
        if (!apiKey) openSettings();
        loadAccounts();
        loadRecentEntries();
        loadCounterparties();
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
        const sid = document.getElementById('spreadsheet-id-input').value.trim();
        if (key) {
            localStorage.setItem('gemini_api_key', key);
            if (sid) localStorage.setItem('spreadsheet_id', sid);
            fetchAPI('/api/settings', 'POST', { gemini_api_key: key, spreadsheet_id: sid });
            settingsModal.classList.add('hidden');
            showToast('設定を保存しました');
        } else {
            showToast('APIキーを入力してください', true);
        }
    };
    function openSettings() {
        document.getElementById('api-key-input').value = localStorage.getItem('gemini_api_key') || '';
        document.getElementById('spreadsheet-id-input').value = localStorage.getItem('spreadsheet_id') || '';
        settingsModal.classList.remove('hidden');
    }

    // ============================================================
    //  Section 5: Menu Grid Navigation (Hash Router)
    // ============================================================
    // View loaders: called when a view becomes active
    const VIEW_LOADERS = {
        'journal-input': () => { /* already loaded on login */ },
        'scan': () => {},
        'journal-book': () => loadJournalBook(),
        'ledger': () => loadCurrentLedgerSubTab(),
        'counterparty': () => loadCounterpartyList(),
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

    // Menu tile click
    menuGrid.addEventListener('click', (e) => {
        const tile = e.target.closest('.menu-tile');
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

    // Initial route from hash
    const initHash = location.hash.replace('#', '');
    if (initHash && initHash !== 'menu') {
        showView(initHash);
    }
    // If no hash, menu is already visible (active class in HTML)

    // ============================================================
    //  Section 6: Shared Utilities
    // ============================================================
    async function fetchAPI(url, method = 'GET', body = null) {
        const opts = { method, headers: {} };
        if (body && !(body instanceof FormData)) {
            opts.headers['Content-Type'] = 'application/json';
            opts.body = JSON.stringify(body);
        } else if (body instanceof FormData) {
            opts.body = body;
        }
        const res = await fetch(url, opts);
        return res.json();
    }

    function loadAccounts() {
        fetchAPI('/api/accounts').then(data => {
            if (data.accounts) {
                accounts = data.accounts;
                populateAccountDatalist();
                populateJBAccountFilter();
            }
        });
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
    //  Section 7: View 1 — 仕訳入力 (Journal Entry) [TKC FX2 Style]
    // ============================================================
    const journalForm = document.getElementById('journal-form');
    const jeDate = document.getElementById('je-date');
    const jeTaxCategory = document.getElementById('je-tax-category');
    const jeDebit = document.getElementById('je-debit');
    const jeCredit = document.getElementById('je-credit');
    const jeAmount = document.getElementById('je-amount');
    const jeNetAmount = document.getElementById('je-net-amount');
    const jeCounterparty = document.getElementById('je-counterparty');
    const jeMemo = document.getElementById('je-memo');
    const jeTaxRate = document.getElementById('je-tax-rate');
    const jeAiBtn = document.getElementById('je-ai-btn');

    // Default date to today
    jeDate.value = todayStr();

    // --- Tax-exclusive amount auto-calculation ---
    function updateNetAmount() {
        const amt = parseInt(jeAmount.value) || 0;
        const rate = jeTaxRate.value;
        if (!amt) { jeNetAmount.value = ''; return; }
        let tax = 0;
        if (rate === '10%') tax = Math.floor(amt * 10 / 110);
        else if (rate === '8%') tax = Math.floor(amt * 8 / 108);
        jeNetAmount.value = fmt(amt - tax);
    }
    jeAmount.addEventListener('input', updateNetAmount);
    jeTaxRate.addEventListener('change', updateNetAmount);

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

        if (!entry.debit_account || !entry.credit_account || !entry.amount) {
            showToast('借方科目・貸方科目・金額は必須です', true);
            return;
        }

        try {
            const res = await fetchAPI('/api/journal', 'POST', entry);
            if (res.status === 'success') {
                showToast('仕訳を登録しました');
                journalForm.reset();
                jeDate.value = todayStr();
                jeNetAmount.value = '';
                loadRecentEntries();
                loadCounterparties();
            } else {
                showToast('登録に失敗しました: ' + (res.error || ''), true);
            }
        } catch (err) {
            showToast('通信エラー', true);
        }
    });

    // --- Recent Entries (below form) ---
    function loadRecentEntries() {
        fetchAPI('/api/journal/recent?limit=5').then(data => {
            const tbody = document.getElementById('recent-tbody');
            if (!tbody) return;
            tbody.innerHTML = (data.entries || []).map(e => {
                const parsed = parseTaxClassification(e.tax_classification);
                const amt = parseInt(e.amount) || 0;
                let tax = 0;
                if (e.tax_classification === '10%') tax = Math.floor(amt * 10 / 110);
                else if (e.tax_classification === '8%') tax = Math.floor(amt * 8 / 108);
                const net = amt - tax;
                return `
                <tr>
                    <td>${e.entry_date || ''}</td>
                    <td>${parsed.taxCategory || ''}</td>
                    <td>${e.debit_account || ''}</td>
                    <td>${e.credit_account || ''}</td>
                    <td class="text-right">${fmt(e.amount)}</td>
                    <td class="text-right">${fmt(net)}</td>
                    <td>${e.counterparty || ''}</td>
                    <td>${e.memo || ''}</td>
                    <td>${parsed.taxRate || ''}</td>
                </tr>`;
            }).join('');
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

    async function handleScanFiles(files, append = false) {
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

    function renderScanResults() {
        let hasDup = false;
        scanTbody.innerHTML = scanResults.map((item, i) => {
            if (item.is_duplicate) hasDup = true;
            return `
            <tr class="${item.is_duplicate ? 'row-duplicate' : ''}">
                <td><input type="date" value="${item.date || ''}" data-i="${i}" data-k="date" class="scan-input"></td>
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

        const entries = valid.map(r => ({
            entry_date: r.date,
            debit_account: r.debit_account,
            credit_account: r.credit_account,
            amount: parseInt(r.amount),
            tax_classification: r.tax_classification || '10%',
            counterparty: r.counterparty || '',
            memo: r.memo || '',
            evidence_url: r.evidence_url || '',
            source: 'ai_receipt',
        }));

        try {
            scanSaveBtn.disabled = true;
            scanSaveBtn.textContent = '保存中...';
            const res = await fetchAPI('/api/journal', 'POST', { entries });
            if (res.status === 'success') {
                showToast(`${res.created}件の仕訳を登録しました`);
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

    // ============================================================
    //  Section 9: View 3 — 仕訳帳 (Journal Book)
    // ============================================================
    const jbStartInput = document.getElementById('jb-start');
    const jbEndInput = document.getElementById('jb-end');
    const jbAccountSelect = document.getElementById('jb-account');
    const jbSearchInput = document.getElementById('jb-search');
    const jbApplyBtn = document.getElementById('jb-apply');
    const jbTbody = document.getElementById('jb-tbody');
    const jbPagination = document.getElementById('jb-pagination');

    let jbPage = 1;
    const JB_PER_PAGE = 20;

    const fy = fiscalYearDates();
    jbStartInput.value = fy.start;
    jbEndInput.value = fy.end;

    jbApplyBtn.addEventListener('click', () => { jbPage = 1; loadJournalBook(); });

    async function loadJournalBook() {
        const params = new URLSearchParams();
        if (jbStartInput.value) params.set('start_date', jbStartInput.value);
        if (jbEndInput.value) params.set('end_date', jbEndInput.value);
        if (jbAccountSelect.value) params.set('account_id', jbAccountSelect.value);
        const search = jbSearchInput.value.trim();
        if (search) {
            params.set('counterparty', search);
            params.set('memo', search);
        }
        params.set('page', jbPage);
        params.set('per_page', JB_PER_PAGE);

        try {
            const data = await fetchAPI('/api/journal?' + params.toString());
            renderJournalBook(data);
        } catch (err) {
            showToast('仕訳帳の読み込みに失敗しました', true);
        }
    }

    function renderJournalBook(data) {
        const entries = data.entries || [];
        jbTbody.innerHTML = entries.map(e => `
            <tr data-id="${e.id}">
                <td>${e.entry_date || ''}</td>
                <td>${e.debit_account || ''}</td>
                <td>${e.credit_account || ''}</td>
                <td class="text-right">${fmt(e.amount)}</td>
                <td>${e.tax_classification || ''}</td>
                <td>${e.counterparty || ''}</td>
                <td>${e.memo || ''}</td>
                <td>
                    <button class="btn-icon jb-edit" data-id="${e.id}" title="編集">✎</button>
                    <button class="btn-icon jb-delete" data-id="${e.id}" title="削除">×</button>
                </td>
            </tr>
        `).join('');

        if (!entries.length) {
            jbTbody.innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--text-dim);padding:2rem;">仕訳データがありません</td></tr>';
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
                const row = e.target.closest('tr');
                startInlineEdit(row, id);
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
    // ============================================================
    const ledgerSubNav = document.getElementById('ledger-sub-nav');
    const ledgerStartInput = document.getElementById('ledger-start');
    const ledgerEndInput = document.getElementById('ledger-end');
    const ledgerApplyBtn = document.getElementById('ledger-apply');
    const ledgerDetail = document.getElementById('ledger-detail');
    const ledgerDetailBack = document.getElementById('ledger-detail-back');
    const ledgerDetailTitle = document.getElementById('ledger-detail-title');
    const ledgerDetailContent = document.getElementById('ledger-detail-content');

    const tbContent = document.getElementById('tb-content');
    const bsContent = document.getElementById('bs-content');
    const plContent = document.getElementById('pl-content');

    // Set default dates
    ledgerStartInput.value = fy.start;
    ledgerEndInput.value = fy.end;

    let currentLedgerSubTab = 'trial-balance';

    // Sub-tab switching
    ledgerSubNav.addEventListener('click', (e) => {
        const btn = e.target.closest('.sub-tab-btn');
        if (!btn) return;
        const subtab = btn.dataset.subtab;
        currentLedgerSubTab = subtab;

        // Update button states
        ledgerSubNav.querySelectorAll('.sub-tab-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');

        // Show/hide panels
        document.querySelectorAll('.ledger-sub-panel').forEach(p => p.classList.remove('active'));
        const panel = document.getElementById('ledger-tab-' + subtab);
        if (panel) panel.classList.add('active');

        // Hide drill-down detail when switching tabs
        hideLedgerDetail();

        // Load data
        loadCurrentLedgerSubTab();
    });

    ledgerApplyBtn.addEventListener('click', loadCurrentLedgerSubTab);

    function loadCurrentLedgerSubTab() {
        if (currentLedgerSubTab === 'trial-balance') loadTrialBalance();
        else if (currentLedgerSubTab === 'balance-sheet') loadBalanceSheet();
        else if (currentLedgerSubTab === 'profit-loss') loadProfitLoss();
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
    async function showAccountDetail(accountId) {
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
            html += `<div class="table-wrap"><table class="tb-table">
                <thead><tr>
                    <th>日付</th><th>相手科目</th><th>摘要</th><th>取引先</th>
                    <th class="text-right">借方</th><th class="text-right">貸方</th>
                    <th class="text-right">差引残高</th>
                </tr></thead><tbody>`;

            entries.forEach(e => {
                html += `<tr>
                    <td>${e.entry_date || ''}</td>
                    <td>${e.counter_account || ''}</td>
                    <td>${e.memo || ''}</td>
                    <td>${e.counterparty || ''}</td>
                    <td class="text-right">${e.debit_amount ? fmt(e.debit_amount) : ''}</td>
                    <td class="text-right">${e.credit_amount ? fmt(e.credit_amount) : ''}</td>
                    <td class="text-right" style="font-weight:600;">${fmt(e.balance)}</td>
                </tr>`;
            });

            if (!entries.length) {
                html += '<tr><td colspan="7" style="text-align:center;padding:2rem;color:var(--text-dim);">該当する仕訳がありません</td></tr>';
            }

            html += '</tbody></table></div>';
            ledgerDetailContent.innerHTML = html;
        } catch (err) {
            showToast('元帳明細の読み込みに失敗しました', true);
        }
    }

    // --- Trial Balance ---
    async function loadTrialBalance() {
        const params = new URLSearchParams();
        if (ledgerStartInput.value) params.set('start_date', ledgerStartInput.value);
        if (ledgerEndInput.value) params.set('end_date', ledgerEndInput.value);

        try {
            const data = await fetchAPI('/api/trial-balance?' + params.toString());
            renderTrialBalance(data.balances || []);
        } catch (err) {
            showToast('残高一覧表の読み込みに失敗しました', true);
        }
    }

    function renderTrialBalance(balances) {
        if (!balances.length) {
            tbContent.innerHTML = '<p style="text-align:center;color:var(--text-dim);padding:2rem;">データがありません</p>';
            return;
        }

        let grandDebit = 0, grandCredit = 0;

        let html = `<table class="tb-table">
            <thead>
                <tr>
                    <th>コード</th>
                    <th>勘定科目</th>
                    <th>科目区分</th>
                    <th class="text-right">期首残高</th>
                    <th class="text-right">借方合計</th>
                    <th class="text-right">貸方合計</th>
                    <th class="text-right">残高</th>
                </tr>
            </thead>
            <tbody>`;

        balances.forEach(b => {
            grandDebit += b.debit_total;
            grandCredit += b.credit_total;
            html += `
                <tr class="tb-row" data-account-id="${b.account_id}" style="cursor:pointer;">
                    <td>${b.code}</td>
                    <td>${b.name}</td>
                    <td>${b.account_type}</td>
                    <td class="text-right">${fmt(b.opening_balance)}</td>
                    <td class="text-right">${fmt(b.debit_total)}</td>
                    <td class="text-right">${fmt(b.credit_total)}</td>
                    <td class="text-right" style="font-weight:600;">${fmt(b.closing_balance)}</td>
                </tr>`;
        });

        html += `</tbody></table>`;

        html += `<div class="tb-grand-total">
            <span>借方合計: <strong>${fmt(grandDebit)}</strong></span>
            <span>貸方合計: <strong>${fmt(grandCredit)}</strong></span>
            <span>差額: <strong>${fmt(grandDebit - grandCredit)}</strong></span>
        </div>`;

        tbContent.innerHTML = html;

        // Drill-down: click account row → show detail within ledger view
        tbContent.querySelectorAll('.tb-row').forEach(row => {
            row.addEventListener('click', () => {
                showAccountDetail(row.dataset.accountId);
            });
        });
    }

    // --- Balance Sheet ---
    async function loadBalanceSheet() {
        const params = new URLSearchParams();
        if (ledgerStartInput.value) params.set('start_date', ledgerStartInput.value);
        if (ledgerEndInput.value) params.set('end_date', ledgerEndInput.value);

        try {
            const data = await fetchAPI('/api/trial-balance?' + params.toString());
            renderBalanceSheet(data.balances || []);
        } catch (err) {
            showToast('貸借対照表の読み込みに失敗しました', true);
        }
    }

    function renderBalanceSheet(balances) {
        const assets = balances.filter(b => b.account_type === '資産');
        const liabilities = balances.filter(b => b.account_type === '負債');
        const equity = balances.filter(b => b.account_type === '純資産');

        if (!assets.length && !liabilities.length && !equity.length) {
            bsContent.innerHTML = '<p style="text-align:center;color:var(--text-dim);padding:2rem;">データがありません</p>';
            return;
        }

        const assetTotal = assets.reduce((s, b) => s + b.closing_balance, 0);
        const liabilityTotal = liabilities.reduce((s, b) => s + b.closing_balance, 0);
        const equityTotal = equity.reduce((s, b) => s + b.closing_balance, 0);
        const rightTotal = liabilityTotal + equityTotal;

        function buildRows(items) {
            return items.map(b => `
                <tr class="clickable" data-account-id="${b.account_id}">
                    <td>${b.code}</td>
                    <td>${b.name}</td>
                    <td class="text-right">${fmt(b.closing_balance)}</td>
                </tr>`).join('');
        }

        let html = '<div class="bs-container">';

        html += `<div class="bs-side">
            <div class="bs-side-title">資産の部</div>
            <table class="bs-table">
                <thead><tr><th>コード</th><th>勘定科目</th><th class="text-right">残高</th></tr></thead>
                <tbody>
                    ${buildRows(assets)}
                    <tr class="bs-grand-total">
                        <td colspan="2">資産合計</td>
                        <td class="text-right">${fmt(assetTotal)}</td>
                    </tr>
                </tbody>
            </table>
        </div>`;

        html += `<div class="bs-side">
            <div class="bs-side-title">負債・純資産の部</div>
            <table class="bs-table">
                <thead><tr><th>コード</th><th>勘定科目</th><th class="text-right">残高</th></tr></thead>
                <tbody>
                    ${buildRows(liabilities)}
                    <tr class="bs-subtotal">
                        <td colspan="2">負債合計</td>
                        <td class="text-right">${fmt(liabilityTotal)}</td>
                    </tr>
                    ${buildRows(equity)}
                    <tr class="bs-subtotal">
                        <td colspan="2">純資産合計</td>
                        <td class="text-right">${fmt(equityTotal)}</td>
                    </tr>
                    <tr class="bs-grand-total">
                        <td colspan="2">負債・純資産合計</td>
                        <td class="text-right">${fmt(rightTotal)}</td>
                    </tr>
                </tbody>
            </table>
        </div>`;

        html += '</div>';

        const isBalanced = assetTotal === rightTotal;
        html += `<div class="bs-balance-check ${isBalanced ? 'balanced' : 'unbalanced'}">
            ${isBalanced
                ? '✓ 貸借一致 (資産 = 負債 + 純資産)'
                : `✗ 貸借不一致: 資産 ${fmt(assetTotal)} ≠ 負債+純資産 ${fmt(rightTotal)} (差額: ${fmt(assetTotal - rightTotal)})`
            }
        </div>`;

        bsContent.innerHTML = html;

        // Drill-down
        bsContent.querySelectorAll('.clickable').forEach(row => {
            row.addEventListener('click', () => {
                showAccountDetail(row.dataset.accountId);
            });
        });
    }

    // --- Profit & Loss ---
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

        if (!revenues.length && !expenses.length) {
            plContent.innerHTML = '<p style="text-align:center;color:var(--text-dim);padding:2rem;">データがありません</p>';
            return;
        }

        const revenueTotal = revenues.reduce((s, b) => s + b.closing_balance, 0);
        const expenseTotal = expenses.reduce((s, b) => s + b.closing_balance, 0);
        const netIncome = revenueTotal - expenseTotal;

        function buildSection(title, items, subtotalLabel, subtotal) {
            let html = `<div class="pl-section">
                <div class="pl-section-title">${title}</div>
                <table class="pl-table">
                    <thead><tr>
                        <th>コード</th>
                        <th>勘定科目</th>
                        <th class="text-right">借方合計</th>
                        <th class="text-right">貸方合計</th>
                        <th class="text-right">金額</th>
                    </tr></thead>
                    <tbody>`;

            items.forEach(b => {
                html += `<tr class="clickable" data-account-id="${b.account_id}">
                    <td>${b.code}</td>
                    <td>${b.name}</td>
                    <td class="text-right">${fmt(b.debit_total)}</td>
                    <td class="text-right">${fmt(b.credit_total)}</td>
                    <td class="text-right" style="font-weight:600;">${fmt(b.closing_balance)}</td>
                </tr>`;
            });

            html += `<tr class="pl-subtotal">
                    <td colspan="4">${subtotalLabel}</td>
                    <td class="text-right">${fmt(subtotal)}</td>
                </tr>
                </tbody></table></div>`;
            return html;
        }

        let html = '';
        html += buildSection('収益の部', revenues, '収益合計', revenueTotal);
        html += buildSection('費用の部', expenses, '費用合計', expenseTotal);

        const isProfit = netIncome >= 0;
        html += `<div class="pl-net-income ${isProfit ? 'profit' : 'loss'}">
            <span>${isProfit ? '当期純利益' : '当期純損失'}</span>
            <span>${fmt(Math.abs(netIncome))}</span>
        </div>`;

        plContent.innerHTML = html;

        // Drill-down
        plContent.querySelectorAll('.clickable').forEach(row => {
            row.addEventListener('click', () => {
                showAccountDetail(row.dataset.accountId);
            });
        });
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
    //  Section 12: View 6 — 期首残高設定 (Opening Balances)
    // ============================================================
    const obFiscalYear = document.getElementById('ob-fiscal-year');
    const obLoadBtn = document.getElementById('ob-load');
    const obSaveBtn = document.getElementById('ob-save');
    const obTbody = document.getElementById('ob-tbody');

    // Populate fiscal year options
    const currentYear = new Date().getFullYear();
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
    //  Section 13: View 7 — データのバックアップ (Backup)
    // ============================================================
    const backupJsonBtn = document.getElementById('backup-json');
    const backupSqliteBtn = document.getElementById('backup-sqlite');

    backupJsonBtn.addEventListener('click', async () => {
        try {
            backupJsonBtn.disabled = true;
            backupJsonBtn.textContent = 'ダウンロード中...';
            const res = await fetch('/api/backup/download?format=json');
            const blob = await res.blob();
            downloadBlob(blob, `accounting_backup_${todayStr()}.json`, 'application/json');
            showToast('JSONバックアップをダウンロードしました');
        } catch (err) {
            showToast('バックアップに失敗しました', true);
        } finally {
            backupJsonBtn.disabled = false;
            backupJsonBtn.textContent = 'JSONバックアップ';
        }
    });

    backupSqliteBtn.addEventListener('click', async () => {
        try {
            backupSqliteBtn.disabled = true;
            backupSqliteBtn.textContent = 'ダウンロード中...';
            const res = await fetch('/api/backup/download?format=sqlite');
            const blob = await res.blob();
            downloadBlob(blob, `accounting_backup_${todayStr()}.db`, 'application/x-sqlite3');
            showToast('SQLiteバックアップをダウンロードしました');
        } catch (err) {
            showToast('バックアップに失敗しました', true);
        } finally {
            backupSqliteBtn.disabled = false;
            backupSqliteBtn.textContent = 'SQLiteダウンロード';
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

    // ============================================================
    //  Section 14: View 8 — アウトプット (Output / Export)
    // ============================================================
    const outStartInput = document.getElementById('out-start');
    const outEndInput = document.getElementById('out-end');
    const outJournalCsvBtn = document.getElementById('out-journal-csv');
    const outJournalPdfBtn = document.getElementById('out-journal-pdf');
    const outTbCsvBtn = document.getElementById('out-tb-csv');
    const outTbPdfBtn = document.getElementById('out-tb-pdf');

    outStartInput.value = fy.start;
    outEndInput.value = fy.end;

    // Journal CSV download
    outJournalCsvBtn.addEventListener('click', async () => {
        const params = new URLSearchParams();
        if (outStartInput.value) params.set('start_date', outStartInput.value);
        if (outEndInput.value) params.set('end_date', outEndInput.value);
        params.set('format', 'csv');

        try {
            const res = await fetch('/api/export/journal?' + params.toString());
            const blob = await res.blob();
            downloadBlob(blob, `journal_export_${todayStr()}.csv`, 'text/csv');
            showToast('仕訳帳CSVをダウンロードしました');
        } catch (err) {
            showToast('エクスポートに失敗しました', true);
        }
    });

    // Journal PDF (print view)
    outJournalPdfBtn.addEventListener('click', async () => {
        const params = new URLSearchParams();
        if (outStartInput.value) params.set('start_date', outStartInput.value);
        if (outEndInput.value) params.set('end_date', outEndInput.value);
        params.set('format', 'json');

        try {
            const data = await fetchAPI('/api/export/journal?' + params.toString());
            const entries = data.entries || [];
            openPrintView('仕訳帳', buildJournalPrintTable(entries));
        } catch (err) {
            showToast('エクスポートに失敗しました', true);
        }
    });

    // Trial Balance CSV
    outTbCsvBtn.addEventListener('click', async () => {
        const params = new URLSearchParams();
        if (outStartInput.value) params.set('start_date', outStartInput.value);
        if (outEndInput.value) params.set('end_date', outEndInput.value);

        try {
            const data = await fetchAPI('/api/export/trial-balance?' + params.toString());
            const balances = data.balances || [];
            const csv = buildTrialBalanceCsv(balances);
            const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
            downloadBlob(blob, `trial_balance_${todayStr()}.csv`, 'text/csv');
            showToast('残高試算表CSVをダウンロードしました');
        } catch (err) {
            showToast('エクスポートに失敗しました', true);
        }
    });

    // Trial Balance PDF (print view)
    outTbPdfBtn.addEventListener('click', async () => {
        const params = new URLSearchParams();
        if (outStartInput.value) params.set('start_date', outStartInput.value);
        if (outEndInput.value) params.set('end_date', outEndInput.value);

        try {
            const data = await fetchAPI('/api/export/trial-balance?' + params.toString());
            const balances = data.balances || [];
            openPrintView('残高試算表', buildTrialBalancePrintTable(balances));
        } catch (err) {
            showToast('エクスポートに失敗しました', true);
        }
    });

    function buildJournalPrintTable(entries) {
        let html = `<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;width:100%;font-size:11px;">
            <thead><tr style="background:#f0f0f0;">
                <th>日付</th><th>借方科目</th><th>貸方科目</th><th style="text-align:right;">金額</th><th>税区分</th><th>取引先</th><th>摘要</th>
            </tr></thead><tbody>`;
        entries.forEach(e => {
            html += `<tr>
                <td>${e.entry_date || ''}</td>
                <td>${e.debit_account || ''}</td>
                <td>${e.credit_account || ''}</td>
                <td style="text-align:right;">${fmt(e.amount)}</td>
                <td>${e.tax_classification || ''}</td>
                <td>${e.counterparty || ''}</td>
                <td>${e.memo || ''}</td>
            </tr>`;
        });
        html += '</tbody></table>';
        return html;
    }

    function buildTrialBalanceCsv(balances) {
        let csv = 'コード,勘定科目,科目区分,借方合計,貸方合計,残高\n';
        balances.forEach(b => {
            csv += `${b.code},"${b.name}",${b.account_type},${b.debit_total},${b.credit_total},${b.closing_balance}\n`;
        });
        return csv;
    }

    function buildTrialBalancePrintTable(balances) {
        let html = `<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;width:100%;font-size:11px;">
            <thead><tr style="background:#f0f0f0;">
                <th>コード</th><th>勘定科目</th><th>科目区分</th><th style="text-align:right;">借方合計</th><th style="text-align:right;">貸方合計</th><th style="text-align:right;">残高</th>
            </tr></thead><tbody>`;
        balances.forEach(b => {
            html += `<tr>
                <td>${b.code}</td>
                <td>${b.name}</td>
                <td>${b.account_type}</td>
                <td style="text-align:right;">${fmt(b.debit_total)}</td>
                <td style="text-align:right;">${fmt(b.credit_total)}</td>
                <td style="text-align:right;font-weight:bold;">${fmt(b.closing_balance)}</td>
            </tr>`;
        });
        html += '</tbody></table>';
        return html;
    }

    function openPrintView(title, tableHtml) {
        const win = window.open('', '_blank');
        win.document.write(`<!DOCTYPE html><html><head><title>${title}</title>
            <style>body{font-family:sans-serif;padding:20px;}h1{font-size:18px;margin-bottom:10px;}
            @media print{button{display:none;}}</style>
        </head><body>
            <h1>${title}</h1>
            <p style="font-size:12px;color:#666;">出力日: ${todayStr()}</p>
            ${tableHtml}
            <br><button onclick="window.print()" style="padding:8px 16px;font-size:14px;cursor:pointer;">印刷 / PDF保存</button>
        </body></html>`);
        win.document.close();
    }

    // ============================================================
    //  Section 15: Keyboard Shortcuts
    // ============================================================
    document.addEventListener('keydown', (e) => {
        // Ctrl+Enter to submit journal form when in journal-input view
        if (e.ctrlKey && e.key === 'Enter') {
            const activeView = document.querySelector('.content-view.active');
            if (activeView && activeView.id === 'view-journal-input') {
                journalForm.dispatchEvent(new Event('submit'));
            }
        }
        // Escape to go back to menu
        if (e.key === 'Escape' && !settingsModal.classList.contains('hidden')) {
            settingsModal.classList.add('hidden');
        } else if (e.key === 'Escape' && !menuGrid.classList.contains('active')) {
            showMenu();
        }
    });
});
