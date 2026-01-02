document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const statusArea = document.getElementById('status-area');
    const resultsArea = document.getElementById('results-area');
    const resultsTable = document.querySelector('#results-table tbody');
    const sendBtn = document.getElementById('send-btn');
    const resetBtn = document.getElementById('reset-btn');

    const addMoreBtn = document.getElementById('add-more-btn');
    const addMoreInput = document.getElementById('add-more-input');

    let extractedData = [];

    // Drag & Drop event listeners
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('active');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('active');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('active');
        const files = e.dataTransfer.files;
        handleFiles(files);
    });

    dropZone.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', () => {
        handleFiles(fileInput.files);
    });

    addMoreBtn.addEventListener('click', () => {
        addMoreInput.click();
    });

    addMoreInput.addEventListener('change', () => {
        handleFiles(addMoreInput.files, true);
    });

    function handleFiles(files, append = false) {
        if (files.length === 0) return;

        if (!append) {
            dropZone.classList.add('hidden');
        }
        statusArea.classList.remove('hidden');
        document.getElementById('status-text').textContent = `${files.length}枚のファイルを解析中...`;

        const formData = new FormData();
        for (let i = 0; i < files.length; i++) {
            formData.append('files', files[i]);
        }

        // Call backend API
        fetch('/api/analyze', {
            method: 'POST',
            body: formData
        })
            .then(response => response.json())
            .then(data => {
                if (append) {
                    extractedData = [...extractedData, ...data];
                } else {
                    extractedData = data;
                }
                renderResults(extractedData);
                statusArea.classList.add('hidden');
                resultsArea.classList.remove('hidden');
            })
            .catch(err => {
                console.error(err);
                alert('解析中にエラーが発生しました。');
                statusArea.classList.add('hidden');
                if (!append && extractedData.length === 0) {
                    dropZone.classList.remove('hidden');
                }
            });
    }

    function renderResults(data) {
        resultsTable.innerHTML = '';
        let hasDuplicate = false;

        data.forEach((item, index) => {
            if (item.is_duplicate) hasDuplicate = true;

            const row = document.createElement('tr');
            if (item.is_duplicate) row.classList.add('duplicate');

            row.innerHTML = `
                <td><input type="text" value="${item.date || ''}" data-index="${index}" data-key="date"></td>
                <td><input type="text" value="${item.debit_account || ''}" data-index="${index}" data-key="debit_account"></td>
                <td><input type="text" value="${item.credit_account || ''}" data-index="${index}" data-key="credit_account"></td>
                <td><input type="number" value="${item.amount || 0}" data-index="${index}" data-key="amount"></td>
                <td><input type="text" value="${item.counterparty || ''}" data-index="${index}" data-key="counterparty"></td>
                <td><input type="text" value="${item.memo || ''}" data-index="${index}" data-key="memo"></td>
            `;
            resultsTable.appendChild(row);
        });

        // 重複警告の表示制御
        const duplicateAlert = document.getElementById('duplicate-alert');
        if (hasDuplicate) {
            duplicateAlert.classList.remove('hidden');
            duplicateAlert.style.display = 'flex';
        } else {
            duplicateAlert.classList.add('hidden');
            duplicateAlert.style.display = 'none';
        }

        // Add event listeners to inputs to update extractedData
        resultsTable.querySelectorAll('input').forEach(input => {
            input.addEventListener('change', (e) => {
                const idx = parseInt(e.target.dataset.index);
                const key = e.target.dataset.key;
                extractedData[idx][key] = e.target.value;
            });
        });
    }

    sendBtn.addEventListener('click', () => {
        sendBtn.disabled = true;
        sendBtn.textContent = '送信中...';

        fetch('/api/save', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(extractedData)
        })
            .then(response => response.json())
            .then(data => {
                alert('スプレッドシートに保存しました！');
                resetUI();
            })
            .catch(err => {
                console.error(err);
                alert('保存中にエラーが発生しました。');
            })
            .finally(() => {
                sendBtn.disabled = false;
                sendBtn.textContent = 'スプレッドシートに書き込む';
            });
    });

    resetBtn.addEventListener('click', resetUI);

    function resetUI() {
        resultsArea.classList.add('hidden');
        dropZone.classList.remove('hidden');
        fileInput.value = '';
        extractedData = [];
    }
});
