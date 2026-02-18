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
    const tabNav = document.getElementById('tab-nav');

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
            // Also save to server
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
    //  Section 5: Tab Navigation (Hash Router)
    // ============================================================
    function switchTab(tabId) {
        document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        const panel = document.getElementById('tab-' + tabId);
        const btn = document.querySelector(`.tab-btn[data-tab="${tabId}"]`);
        if (panel) panel.classList.add('active');
        if (btn) btn.classList.add('active');
        // Trigger data loading for specific tabs
        if (tabId === 'journal-book') loadJournalBook();
        if (tabId === 'trial-balance') loadTrialBalance();
    }

    tabNav.addEventListener('click', (e) => {
        const btn = e.target.closest('.tab-btn');
        if (!btn) return;
        const tab = btn.dataset.tab;
        location.hash = tab;
        switchTab(tab);
    });

    window.addEventListener('hashchange', () => {
        const hash = location.hash.replace('#', '') || 'journal-input';
        switchTab(hash);
    });

    // Initial tab from hash
    const initTab = location.hash.replace('#', '') || 'journal-input';
    switchTab(initTab);

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
    //  Section 7: Tab 1 — 仕訳入力 (Journal Entry) [TKC FX2 Style]
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
        const rate = jeTaxRate.value;   // "10%", "8%", "0%", or ""
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
        // If switching to 課税 and rate is 0%, reset to blank for AI
        if ((cat === '課税仕入' || cat === '課税売上') && jeTaxRate.value === '0%') {
            jeTaxRate.value = '';
            updateNetAmount();
        }
    });

    jeTaxRate.addEventListener('change', () => {
        const rate = jeTaxRate.value;
        const cat = jeTaxCategory.value;
        // If rate is 0% and category is taxable, auto-set to 非課税
        if (rate === '0%' && (cat === '課税仕入' || cat === '課税売上' || cat === '')) {
            // Don't auto-change, just update net amount
        }
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
    // Maps (tax_category, tax_rate) → DB tax_classification value
    function resolveTaxClassification(taxCategory, taxRate) {
        if (taxCategory === '非課税') return '非課税';
        if (taxCategory === '不課税') return '不課税';
        // For 課税仕入/課税売上 or blank, use the tax rate
        if (taxRate === '10%') return '10%';
        if (taxRate === '8%') return '8%';
        // Default
        return '10%';
    }

    // --- Reverse: DB tax_classification → display (tax_category, tax_rate) ---
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

        // Build prediction request: send only user-filled fields as fixed
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
                // Only fill empty fields
                if (!jeDebit.value.trim() && p.debit) jeDebit.value = p.debit;
                if (!jeCredit.value.trim() && p.credit) jeCredit.value = p.credit;
                if (!jeTaxCategory.value && p.tax_category) jeTaxCategory.value = p.tax_category;
                if (!jeTaxRate.value && p.tax_rate) jeTaxRate.value = p.tax_rate;
                // Sync linkage: if category is 非課税/不課税, force rate to 0%
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
                loadCounterparties(); // refresh counterparty suggestions
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
    //  Section 8: Tab 2 — 証憑読み取り (Document Scanning)
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

        // Bind events
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

    // Save scan results to DB
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
    //  Section 9: Tab 3 — 仕訳帳 (Journal Book)
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

    // Set default date range to current year
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
            jbTbody.innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--text-muted);padding:2rem;">仕訳データがありません</td></tr>';
        }

        // Pagination
        const total = data.total || 0;
        const totalPages = Math.ceil(total / JB_PER_PAGE);
        let pgHtml = '';
        if (totalPages > 1) {
            if (jbPage > 1) pgHtml += `<button class="btn btn-sm btn-ghost jb-page" data-p="${jbPage - 1}">← 前</button>`;
            pgHtml += `<span style="padding:0.5rem;">${jbPage} / ${totalPages} (${total}件)</span>`;
            if (jbPage < totalPages) pgHtml += `<button class="btn btn-sm btn-ghost jb-page" data-p="${jbPage + 1}">次 →</button>`;
        }
        jbPagination.innerHTML = pgHtml;

        // Bind events
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
    //  Section 10: Tab 4 — 残高一覧表 (Trial Balance)
    // ============================================================
    const tbStartInput = document.getElementById('tb-start');
    const tbEndInput = document.getElementById('tb-end');
    const tbApplyBtn = document.getElementById('tb-apply');
    const tbContent = document.getElementById('tb-content');

    tbStartInput.value = fy.start;
    tbEndInput.value = fy.end;

    tbApplyBtn.addEventListener('click', loadTrialBalance);

    async function loadTrialBalance() {
        const params = new URLSearchParams();
        if (tbStartInput.value) params.set('start_date', tbStartInput.value);
        if (tbEndInput.value) params.set('end_date', tbEndInput.value);

        try {
            const data = await fetchAPI('/api/trial-balance?' + params.toString());
            renderTrialBalance(data.balances || []);
        } catch (err) {
            showToast('残高一覧表の読み込みに失敗しました', true);
        }
    }

    function renderTrialBalance(balances) {
        if (!balances.length) {
            tbContent.innerHTML = '<p style="text-align:center;color:var(--text-muted);padding:2rem;">データがありません</p>';
            return;
        }

        // Group by account_type
        const groups = {};
        const typeOrder = ['資産', '負債', '純資産', '収益', '費用'];
        typeOrder.forEach(t => { groups[t] = []; });
        balances.forEach(b => {
            const t = b.account_type || '費用';
            if (!groups[t]) groups[t] = [];
            groups[t].push(b);
        });

        let html = '';
        let grandDebit = 0, grandCredit = 0;

        typeOrder.forEach(type => {
            const items = groups[type];
            if (!items || !items.length) return;

            let typeDebit = 0, typeCredit = 0;

            html += `<div class="tb-section">
                <h3 class="tb-section-title">${type}</h3>
                <table class="tb-table">
                    <thead>
                        <tr>
                            <th>コード</th>
                            <th>勘定科目</th>
                            <th class="text-right">期首残高</th>
                            <th class="text-right">借方合計</th>
                            <th class="text-right">貸方合計</th>
                            <th class="text-right">残高</th>
                        </tr>
                    </thead>
                    <tbody>`;

            items.forEach(b => {
                typeDebit += b.debit_total;
                typeCredit += b.credit_total;
                html += `
                    <tr class="tb-row" data-account-id="${b.account_id}" style="cursor:pointer;">
                        <td>${b.code}</td>
                        <td>${b.name}</td>
                        <td class="text-right">${fmt(b.opening_balance)}</td>
                        <td class="text-right">${fmt(b.debit_total)}</td>
                        <td class="text-right">${fmt(b.credit_total)}</td>
                        <td class="text-right" style="font-weight:600;">${fmt(b.closing_balance)}</td>
                    </tr>`;
            });

            html += `<tr class="tb-subtotal">
                        <td colspan="3">${type} 合計</td>
                        <td class="text-right">${fmt(typeDebit)}</td>
                        <td class="text-right">${fmt(typeCredit)}</td>
                        <td></td>
                    </tr>
                    </tbody></table></div>`;

            grandDebit += typeDebit;
            grandCredit += typeCredit;
        });

        // Grand total
        html += `<div class="tb-grand-total">
            <span>借方合計: <strong>${fmt(grandDebit)}</strong></span>
            <span>貸方合計: <strong>${fmt(grandCredit)}</strong></span>
            <span>差額: <strong>${fmt(grandDebit - grandCredit)}</strong></span>
        </div>`;

        tbContent.innerHTML = html;

        // Drill-down: click account row → switch to journal book with that account filter
        tbContent.querySelectorAll('.tb-row').forEach(row => {
            row.addEventListener('click', () => {
                const accId = row.dataset.accountId;
                jbAccountSelect.value = accId;
                jbStartInput.value = tbStartInput.value;
                jbEndInput.value = tbEndInput.value;
                jbPage = 1;
                location.hash = 'journal-book';
                switchTab('journal-book');
            });
        });
    }

    // ============================================================
    //  Section 11: Keyboard Shortcuts
    // ============================================================
    document.addEventListener('keydown', (e) => {
        // Ctrl+Enter to submit journal form when in Tab 1
        if (e.ctrlKey && e.key === 'Enter') {
            const activePanel = document.querySelector('.tab-panel.active');
            if (activePanel && activePanel.id === 'tab-journal-input') {
                journalForm.dispatchEvent(new Event('submit'));
            }
        }
    });
});
