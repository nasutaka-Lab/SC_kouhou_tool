const DEFAULT_ISSUERS = ["ここに発行者名を入力"];
let customIssuers = [];
let savedGradients = [
    { primary: "#D32F2F", accent: "#FF5252" },
    { primary: "#1976D2", accent: "#42A5F5" },
    { primary: "#388E3C", accent: "#66BB6A" },
    { primary: "#7B1FA2", accent: "#AB47BC" },
    { primary: "#FBC02D", accent: "#FFF176" }
];

document.addEventListener('DOMContentLoaded', () => {
    // --- Elements ---
    const inputs = {
        title: document.getElementById('input-title'),
        date: document.getElementById('input-date'),
        issuer: document.getElementById('input-issuer'),
        vol: document.getElementById('input-vol'),
        color: document.getElementById('input-color-primary'),
        colorAccent: document.getElementById('input-color-accent'),
        font: document.getElementById('input-font'),
        footer: document.getElementById('input-footer'),
        imageFile: document.getElementById('input-image'),
        imageSize: document.getElementById('input-image-size'),
        titleSize: document.getElementById('input-size-title'),
        bodySize: document.getElementById('input-size-body'),
        issuerSelect: document.getElementById('input-issuer-select'),
        showPageNumber: document.getElementById('input-show-page-number'),
    };

    const sectionsContainer = document.getElementById('sections-container');
    const btnAddSection = document.getElementById('btn-add-section');
    const btnAddPageBreak = document.getElementById('btn-add-page-break');
    const btnResetDoc = document.getElementById('btn-reset-doc');
    const layoutBtns = document.querySelectorAll('.layout-btn');
    const colorCode = document.getElementById('color-code');
    const colorCodeAccent = document.getElementById('color-code-accent');
    const exportBtn = document.getElementById('btn-export');
    const exportPngBtn = document.getElementById('btn-export-png');
    const exportJpgBtn = document.getElementById('btn-export-jpg');
    const dataExportBtn = document.getElementById('btn-data-export');
    const dataImportBtn = document.getElementById('btn-data-import');
    const dataCopyBtn = document.getElementById('btn-data-copy');
    const inputImportJson = document.getElementById('input-import-json');
    const removeImgBtn = document.getElementById('btn-remove-image');
    const btnDeleteIssuer = document.getElementById('btn-delete-issuer');
    const historyContainer = document.getElementById('history-container');
    const btnSaveHistory = document.getElementById('btn-save-history');
    const btnSaveGradient = document.getElementById('btn-save-gradient');
    const gradientsContainer = document.getElementById('saved-gradients-container');
    const progressPopup = document.getElementById('export-progress');
    const progressBar = document.getElementById('progress-bar');
    const progressText = document.getElementById('progress-text');
    const progressTime = document.getElementById('progress-time');
    const paperContainer = document.querySelector('.paper-container');

    const valDisplays = {
        titleSize: document.getElementById('val-size-title'),
        bodySize: document.getElementById('val-size-body'),
    };

    let zoomLevel = 0.8;
    const zoomInBtn = document.getElementById('zoom-in');
    const zoomOutBtn = document.getElementById('zoom-out');
    const zoomDisplay = document.getElementById('zoom-level');

    // --- Debounce Control ---
    let updatePending = false;
    function requestUpdate() {
        if (!updatePending) {
            updatePending = true;
            requestAnimationFrame(() => {
                updatePreview();
                updatePending = false;
            });
        }
    }

    // --- Initialization ---
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    inputs.date.value = `${yyyy}-${mm}-${dd}`;

    loadGradients();
    loadState();
    updateIssuerSelect();
    renderHistory();
    renderGradients();
    requestUpdate();

    // --- Event Listeners ---

    btnResetDoc.addEventListener('click', () => {
        if (confirm('全ての入力をリセットしますか？')) {
            localStorage.removeItem('announcement_tool_state');
            location.reload();
        }
    });

    inputs.titleSize.addEventListener('input', (e) => {
        valDisplays.titleSize.textContent = e.target.value;
        requestUpdate();
        saveState();
    });
    inputs.bodySize.addEventListener('input', (e) => {
        valDisplays.bodySize.textContent = e.target.value;
        requestUpdate();
        saveState();
    });

    Object.keys(inputs).forEach(key => {
        const el = inputs[key];
        if (el && el.type !== 'file' && el !== inputs.titleSize && el !== inputs.bodySize && el !== inputs.issuerSelect) {
            const eventType = (el.tagName === 'SELECT' || el.type === 'checkbox') ? 'change' : 'input';
            el.addEventListener(eventType, () => {
                requestUpdate();
                saveState();
            });
        }
    });

    inputs.issuerSelect.addEventListener('change', (e) => {
        if (e.target.value === 'ADD_NEW') {
            inputs.issuerSelect.parentElement.style.display = 'none';
            inputs.issuer.style.display = 'block';
            inputs.issuer.value = '';
            inputs.issuer.focus();
        } else {
            inputs.issuer.value = e.target.value;
            toggleDeleteBtn();
            requestUpdate();
            saveState();
        }
    });

    btnDeleteIssuer.addEventListener('click', () => {
        const val = inputs.issuerSelect.value;
        if (customIssuers.includes(val)) {
            customIssuers = customIssuers.filter(i => i !== val);
            inputs.issuer.value = DEFAULT_ISSUERS[0];
            updateIssuerSelect();
            requestUpdate();
            saveState();
        }
    });

    function toggleDeleteBtn() {
        if (customIssuers.includes(inputs.issuerSelect.value)) {
            btnDeleteIssuer.style.display = 'flex';
            inputs.issuerSelect.classList.add('is-custom');
        } else {
            btnDeleteIssuer.style.display = 'none';
            inputs.issuerSelect.classList.remove('is-custom');
        }
    }

    inputs.issuer.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') addIssuerIfNew(inputs.issuer.value);
    });

    inputs.issuer.addEventListener('blur', () => addIssuerIfNew(inputs.issuer.value));

    btnSaveGradient.addEventListener('click', () => {
        const p = inputs.color.value;
        const a = inputs.colorAccent.value;
        if (!savedGradients.some(g => g.primary === p && g.accent === a)) {
            savedGradients.push({ primary: p, accent: a });
            saveGradients();
            renderGradients();
        }
    });

    function saveGradients() {
        localStorage.setItem('announcement_tool_gradients', JSON.stringify(savedGradients));
    }

    function loadGradients() {
        const saved = localStorage.getItem('announcement_tool_gradients');
        if (saved) savedGradients = JSON.parse(saved);
    }

    function renderGradients() {
        if (!gradientsContainer) return;
        gradientsContainer.innerHTML = '';
        savedGradients.forEach(grad => {
            const swatch = document.createElement('div');
            swatch.className = 'color-swatch';
            if (inputs.color.value.toLowerCase() === grad.primary.toLowerCase() &&
                inputs.colorAccent.value.toLowerCase() === grad.accent.toLowerCase()) {
                swatch.classList.add('active');
            }
            swatch.style.background = `linear-gradient(135deg, ${grad.primary} 0%, ${grad.accent} 100%)`;

            const removeBtn = document.createElement('button');
            removeBtn.className = 'btn-remove-color';
            removeBtn.innerHTML = '<i class="fa-solid fa-xmark"></i>';
            removeBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                savedGradients = savedGradients.filter(g => !(g.primary === grad.primary && g.accent === grad.accent));
                saveGradients();
                renderGradients();
            });
            swatch.appendChild(removeBtn);
            swatch.addEventListener('click', () => {
                inputs.color.value = grad.primary;
                inputs.colorAccent.value = grad.accent;
                colorCode.textContent = grad.primary;
                if (colorCodeAccent) colorCodeAccent.textContent = grad.accent;
                renderGradients();
                requestUpdate();
                saveState();
            });
            gradientsContainer.appendChild(swatch);
        });
    }

    inputs.color.addEventListener('input', (e) => {
        colorCode.textContent = e.target.value;
        requestUpdate();
        saveState();
    });

    inputs.colorAccent.addEventListener('input', (e) => {
        if (colorCodeAccent) colorCodeAccent.textContent = e.target.value;
        requestUpdate();
        saveState();
    });

    layoutBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            layoutBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            requestUpdate();
            saveState();
        });
    });

    btnAddSection.addEventListener('click', () => {
        createSectionUI("", "", "text");
        requestUpdate();
        saveState();
    });

    btnAddPageBreak.addEventListener('click', () => {
        createSectionUI("改ページ", "", "page-break");
        requestUpdate();
        saveState();
    });

    dataExportBtn.addEventListener('click', () => {
        const state = getCurrentState();
        const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `document_data_${new Date().getTime()}.json`;
        a.click();
        URL.revokeObjectURL(url);
    });

    dataCopyBtn.addEventListener('click', () => {
        const state = getCurrentState();
        const str = btoa(unescape(encodeURIComponent(JSON.stringify(state))));
        navigator.clipboard.writeText(str).then(() => {
            alert('データを「識別キー」としてコピーしました。他の端末で「読み込み」からペーストして復元できます。');
        });
    });

    dataImportBtn.addEventListener('click', () => {
        const key = prompt('識別キー（文字列）を貼り付けるか、JSONファイルを選択してください（キャンセルでファイル選択）');
        if (key && key.trim().length > 10) {
            try {
                const state = JSON.parse(decodeURIComponent(escape(atob(key))));
                applyState(state);
                requestUpdate();
                saveState();
                alert('キーからデータを復元しました。');
            } catch (err) {
                alert('無効なキーです。');
            }
        } else {
            inputImportJson.click();
        }
    });

    inputImportJson.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const state = JSON.parse(e.target.result);
                    applyState(state);
                    requestUpdate();
                    saveState();
                    alert('データを読み込みました。');
                } catch (err) {
                    alert('ファイルの読み込みに失敗しました。');
                }
            };
            reader.readAsText(file);
        }
    });

    btnSaveHistory.addEventListener('click', () => saveToHistory());

    function saveToHistory() {
        const history = JSON.parse(localStorage.getItem('announcement_tool_history') || '[]');
        const currentState = getCurrentState();
        const entry = { id: Date.now(), savedAt: new Date().toLocaleString(), state: currentState };
        history.unshift(entry);
        if (history.length > 20) history.pop();
        localStorage.setItem('announcement_tool_history', JSON.stringify(history));
        renderHistory();
        alert('現在の状態を履歴に保存しました。');
    }

    function renderHistory() {
        if (!historyContainer) return;
        const history = JSON.parse(localStorage.getItem('announcement_tool_history') || '[]');
        if (history.length === 0) {
            historyContainer.innerHTML = '<p class="empty-msg">履歴はありません</p>';
            return;
        }
        historyContainer.innerHTML = '';
        history.forEach(item => {
            const historyItem = document.createElement('div');
            historyItem.className = 'history-item';
            historyItem.innerHTML = `
                <div class="history-item-info"><span class="history-item-title">${item.state.title}</span><span class="history-item-date">${item.savedAt}</span></div>
                <div class="history-item-actions"><button class="btn-delete-history" title="削除"><i class="fa-solid fa-trash"></i></button></div>
            `;
            historyItem.addEventListener('click', (e) => {
                if (e.target.closest('.btn-delete-history')) return;
                if (confirm('読み込みますか？')) {
                    applyState(item.state);
                    requestUpdate();
                    saveState();
                }
            });
            historyItem.querySelector('.btn-delete-history').addEventListener('click', (e) => {
                e.stopPropagation();
                if (confirm('削除しますか？')) {
                    const updatedHistory = history.filter(h => h.id !== item.id);
                    localStorage.setItem('announcement_tool_history', JSON.stringify(updatedHistory));
                    renderHistory();
                }
            });
            historyContainer.appendChild(historyItem);
        });
    }

    function createSectionUI(headingValue, bodyValue, type = "text") {
        const sectionDiv = document.createElement('div');
        sectionDiv.className = `sidebar-section ${type === 'page-break' ? 'is-page-break' : ''}`;
        sectionDiv.dataset.type = type;

        if (type === 'page-break') {
            sectionDiv.innerHTML = `
                <div class="section-controls">
                    <span class="section-label"><i class="fa-solid fa-scissors"></i> 改ページ</span>
                    <button class="btn-remove-section"><i class="fa-solid fa-trash"></i></button>
                </div>
                <p>このページ以後のヘッダーは非表示になります</p>
                <input type="hidden" class="section-heading" value="[PAGE_BREAK]">
                <textarea style="display:none" class="section-body"></textarea>
            `;
        } else {
            sectionDiv.innerHTML = `
                <div class="section-controls">
                    <span class="section-label">セクション</span>
                    <button class="btn-remove-section"><i class="fa-solid fa-trash"></i></button>
                </div>
                <div class="input-wrap"><input type="text" class="section-heading" value="${headingValue}" placeholder="項目名を入力"></div>
                <div class="section-toolbar">
                    <button class="tool-btn btn-bold" title="太字にする (**テキスト**)"><i class="fa-solid fa-bold"></i> 太字</button>
                </div>
                <div class="input-wrap"><textarea class="section-body" rows="4" placeholder="本文を入力">${bodyValue}</textarea></div>
            `;

            const textarea = sectionDiv.querySelector('.section-body');
            sectionDiv.querySelector('.btn-bold').addEventListener('click', () => {
                const start = textarea.selectionStart;
                const end = textarea.selectionEnd;
                const text = textarea.value;
                const selectedText = text.substring(start, end);
                const beforeText = text.substring(0, start);
                const afterText = text.substring(end);

                textarea.value = `${beforeText}**${selectedText}**${afterText}`;
                textarea.focus();
                textarea.setSelectionRange(start + 2, end + 2);
                requestUpdate();
                saveState();
            });
        }

        sectionDiv.querySelector('.btn-remove-section').addEventListener('click', () => {
            sectionDiv.remove();
            requestUpdate();
            saveState();
        });

        sectionDiv.querySelectorAll('input, textarea').forEach(el => {
            el.addEventListener('input', () => {
                requestUpdate();
                saveState();
            });
        });

        sectionsContainer.appendChild(sectionDiv);
        return sectionDiv;
    }

    inputs.imageFile.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                document.getElementById('temp-image-storage').dataset.img = e.target.result;
                removeImgBtn.style.display = 'inline-block';
                requestUpdate();
                saveState();
            };
            reader.readAsDataURL(file);
        }
    });

    removeImgBtn.addEventListener('click', () => {
        inputs.imageFile.value = '';
        document.getElementById('temp-image-storage').dataset.img = '';
        removeImgBtn.style.display = 'none';
        requestUpdate();
        saveState();
    });

    function setZoom(level) {
        zoomLevel = Math.max(0.1, Math.min(2.0, level));
        document.querySelectorAll('.paper').forEach(p => p.style.transform = `scale(${zoomLevel})`);
        zoomDisplay.textContent = `${Math.round(zoomLevel * 100)}%`;
    }
    zoomInBtn.addEventListener('click', () => setZoom(zoomLevel + 0.1));
    zoomOutBtn.addEventListener('click', () => setZoom(zoomLevel - 0.1));
    setZoom(0.75);

    function updateIssuerSelect() {
        if (!inputs.issuerSelect) return;
        const currentVal = inputs.issuer.value;
        inputs.issuerSelect.innerHTML = `<option value="${DEFAULT_ISSUERS[0]}">${DEFAULT_ISSUERS[0]}</option>`;
        customIssuers.forEach(issuer => {
            const opt = document.createElement('option');
            opt.value = opt.textContent = issuer;
            inputs.issuerSelect.appendChild(opt);
        });
        const addNewOpt = document.createElement('option');
        addNewOpt.value = 'ADD_NEW';
        addNewOpt.textContent = '-- 新規追加 --';
        inputs.issuerSelect.appendChild(addNewOpt);
        inputs.issuerSelect.value = currentVal;
        toggleDeleteBtn();
    }

    function addIssuerIfNew(value) {
        if (!value) {
            inputs.issuerSelect.value = inputs.issuer.value = DEFAULT_ISSUERS[0];
        } else if (!DEFAULT_ISSUERS.includes(value) && !customIssuers.includes(value)) {
            customIssuers.push(value);
        }
        updateIssuerSelect();
        inputs.issuer.style.display = 'none';
        inputs.issuerSelect.parentElement.style.display = 'flex';
        inputs.issuerSelect.value = value || DEFAULT_ISSUERS[0];
        requestUpdate();
        saveState();
    }

    function parseText(text) {
        if (!text) return "";
        return text.replace(/\*\*(.*?)\*\*/g, '<b>$1</b>').replace(/\n/g, '<br>');
    }

    function createPageTemplate(pageNumber, totalPages, headerVisible) {
        const layout = document.querySelector('.layout-btn.active').dataset.layout;
        const page = document.createElement('div');
        page.className = `paper a4 layout-${layout}`;
        page.style.setProperty('--paper-primary', inputs.color.value);
        page.style.setProperty('--paper-accent', inputs.colorAccent.value);
        page.style.fontFamily = inputs.font.value;
        page.style.transform = `scale(${zoomLevel})`;

        const footerNumVisible = inputs.showPageNumber.checked;

        page.innerHTML = `
            <header class="paper-header" style="display: ${headerVisible ? 'block' : 'none'}">
                <div class="header-decoration"></div>
                <div class="header-content">
                    <div class="issue-info">
                        <span class="issue-date">${formatDate(inputs.date.value)}</span>
                        <span class="issue-vol">${inputs.vol.value}</span>
                    </div>
                    <h1 class="paper-title" style="font-size: ${inputs.titleSize.value}pt">${inputs.title.value}</h1>
                    <div class="issuer">${inputs.issuer.value}</div>
                </div>
            </header>
            <div class="paper-body">
                <div class="page-sections"></div>
            </div>
            <footer class="paper-footer">
                <div class="footer-line"></div>
                <p class="footer-text">${inputs.footer.value}</p>
                <div class="page-number-preview" style="display: ${footerNumVisible ? 'block' : 'none'}">
                    ${pageNumber} / <span class="total-pages">${totalPages}</span>
                </div>
            </footer>
        `;
        return page;
    }

    function formatDate(dateStr) {
        if (!dateStr) return "";
        const parts = dateStr.split('-');
        return parts.length === 3 ? `${parts[0]}年${parseInt(parts[1], 10)}月${parseInt(parts[2], 10)}日` : dateStr;
    }

    function updatePreview() {
        const sectionsData = [];
        sectionsContainer.querySelectorAll('.sidebar-section').forEach(sec => {
            sectionsData.push({
                type: sec.dataset.type || "text",
                heading: sec.querySelector('.section-heading').value,
                body: sec.querySelector('.section-body').value
            });
        });

        const imageData = document.getElementById('temp-image-storage').dataset.img;
        const fragment = document.createDocumentFragment();

        let currentPageNum = 1;
        let isHeaderHiddenForSubsequent = false;
        let currentPage = createPageTemplate(currentPageNum, 1, true);
        fragment.appendChild(currentPage);

        let sectionsArea = currentPage.querySelector('.page-sections');
        const maxContentHeight = 850;

        sectionsData.forEach((sec, index) => {
            if (sec.type === 'page-break') {
                isHeaderHiddenForSubsequent = true;
                const breakMarker = document.createElement('div');
                breakMarker.className = 'manual-page-break-preview';
                sectionsArea.appendChild(breakMarker);

                currentPageNum++;
                currentPage = createPageTemplate(currentPageNum, 1, false);
                fragment.appendChild(currentPage);
                sectionsArea = currentPage.querySelector('.page-sections');
                return;
            }

            const sectionDiv = document.createElement('div');
            sectionDiv.className = 'preview-section';
            sectionDiv.innerHTML = `<div class="main-heading-container"><h2 class="main-heading">${sec.heading}</h2></div><div class="text-content" style="font-size: ${inputs.bodySize.value}pt; line-height: 1.8;"><p>${parseText(sec.body)}</p></div>`;

            sectionsArea.appendChild(sectionDiv);

            if (sectionsArea.offsetHeight > maxContentHeight && index < sectionsData.length - 1) {
                sectionDiv.remove();
                currentPageNum++;
                const shouldShowHeader = (currentPageNum === 1 && !isHeaderHiddenForSubsequent);
                currentPage = createPageTemplate(currentPageNum, 1, shouldShowHeader);
                fragment.appendChild(currentPage);
                sectionsArea = currentPage.querySelector('.page-sections');
                sectionsArea.appendChild(sectionDiv);
            }
        });

        if (imageData) {
            const imgWrapper = document.createElement('div');
            imgWrapper.className = 'content-wrapper';
            imgWrapper.innerHTML = `<div class="image-container"><img src="${imageData}" style="width: ${inputs.imageSize.value}%; max-width: ${inputs.imageSize.value}%"></div>`;
            sectionsArea.appendChild(imgWrapper);
            if (sectionsArea.offsetHeight > maxContentHeight) {
                imgWrapper.remove();
                currentPageNum++;
                const shouldShowHeader = (currentPageNum === 1 && !isHeaderHiddenForSubsequent);
                currentPage = createPageTemplate(currentPageNum, 1, shouldShowHeader);
                fragment.appendChild(currentPage);
                sectionsArea = currentPage.querySelector('.page-sections');
                sectionsArea.appendChild(imgWrapper);
            }
        }

        const totalPages = currentPageNum;
        fragment.querySelectorAll('.total-pages').forEach(el => el.textContent = totalPages);

        paperContainer.innerHTML = '';
        paperContainer.appendChild(fragment);
    }

    function getCurrentState() {
        const sectionsData = [];
        sectionsContainer.querySelectorAll('.sidebar-section').forEach(sec => {
            sectionsData.push({
                type: sec.dataset.type || "text",
                heading: sec.querySelector('.section-heading').value,
                body: sec.querySelector('.section-body').value
            });
        });
        return {
            title: inputs.title.value, date: inputs.date.value, issuer: inputs.issuer.value, vol: inputs.vol.value,
            color: inputs.color.value, colorAccent: inputs.colorAccent.value, font: inputs.font.value, footer: inputs.footer.value,
            titleSize: inputs.titleSize.value, bodySize: inputs.bodySize.value, imageSize: inputs.imageSize.value,
            sections: sectionsData, customIssuers: customIssuers,
            layout: document.querySelector('.layout-btn.active').dataset.layout,
            imageData: document.getElementById('temp-image-storage').dataset.img,
            showPageNumber: inputs.showPageNumber.checked
        };
    }

    function saveState() {
        localStorage.setItem('announcement_tool_state', JSON.stringify(getCurrentState()));
    }

    function applyState(state) {
        try {
            inputs.title.value = state.title || "";
            inputs.date.value = state.date || "";
            inputs.issuer.value = state.issuer || DEFAULT_ISSUERS[0];
            inputs.vol.value = state.vol || "ID-001";
            inputs.color.value = state.color || "#D32F2F";
            inputs.colorAccent.value = state.colorAccent || "#FF5252";
            inputs.font.value = state.font || "'Noto Sans JP', sans-serif";
            inputs.footer.value = state.footer || "";
            inputs.titleSize.value = state.titleSize || 32;
            inputs.bodySize.value = state.bodySize || 11;
            inputs.imageSize.value = state.imageSize || 100;
            inputs.showPageNumber.checked = state.showPageNumber !== undefined ? state.showPageNumber : true;
            customIssuers = state.customIssuers || [];
            updateIssuerSelect();
            if (inputs.issuerSelect) inputs.issuerSelect.value = state.issuer || inputs.issuer.value;
            colorCode.textContent = inputs.color.value;
            if (colorCodeAccent) colorCodeAccent.textContent = inputs.colorAccent.value;
            valDisplays.titleSize.textContent = inputs.titleSize.value;
            valDisplays.bodySize.textContent = inputs.bodySize.value;
            layoutBtns.forEach(btn => {
                if (btn.dataset.layout === state.layout) {
                    layoutBtns.forEach(b => b.classList.remove('active'));
                    btn.classList.add('active');
                }
            });
            document.getElementById('temp-image-storage').dataset.img = state.imageData || '';
            removeImgBtn.style.display = state.imageData ? 'inline-block' : 'none';
            sectionsContainer.innerHTML = '';
            if (state.sections && state.sections.length > 0) {
                state.sections.forEach(s => createSectionUI(s.heading, s.body, s.type));
            } else {
                createSectionUI("ここに項目を入力", "本文を入力してください。", "text");
            }
        } catch (e) {
            console.error(e);
        }
    }

    function loadState() {
        const saved = localStorage.getItem('announcement_tool_state');
        if (!saved) {
            createSectionUI("ここに項目を入力", "本文を入力してください。", "text");
        } else {
            applyState(JSON.parse(saved));
        }
    }

    function showProgress(format, estimatedTotalSeconds = 5) {
        progressPopup.style.display = 'block';
        progressText.textContent = `${format.toUpperCase()}生成中...`;
        progressTime.textContent = `約 ${estimatedTotalSeconds} 秒`;
        progressBar.style.width = '0%';
        let elapsed = 0;
        const interval = setInterval(() => {
            elapsed += 0.2;
            progressBar.style.width = `${Math.min((elapsed / estimatedTotalSeconds) * 90, 90)}%`;
            progressTime.textContent = `約 ${Math.max(0, Math.ceil(estimatedTotalSeconds - elapsed))} 秒`;
            if (elapsed >= estimatedTotalSeconds * 0.9) clearInterval(interval);
        }, 200);
        return () => {
            clearInterval(interval);
            progressTime.textContent = "完了！";
            progressBar.style.width = '100%';
            setTimeout(() => progressPopup.style.display = 'none', 800);
        };
    }

    exportBtn.addEventListener('click', async () => {
        const papers = document.querySelectorAll('.paper');
        const finishProgress = showProgress('PDF', papers.length * 4);
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('p', 'mm', 'a4');
        exportBtn.disabled = true;
        try {
            for (let i = 0; i < papers.length; i++) {
                const paper = papers[i];
                const originalTransform = paper.style.transform;
                paper.style.transform = 'none';
                paper.style.boxShadow = 'none';
                const canvas = await html2canvas(paper, { scale: 2, useCORS: true, logging: false, width: 793.7, height: 1122 });
                if (i > 0) pdf.addPage();
                pdf.addImage(canvas.toDataURL('image/jpeg', 0.95), 'JPEG', 0, 0, 210, 297);
                paper.style.transform = originalTransform;
                paper.style.boxShadow = '';
            }
            pdf.save(`${inputs.title.value || 'document'}.pdf`);
        } catch (err) {
            console.error(err);
        } finally {
            exportBtn.disabled = false;
            finishProgress();
        }
    });

    async function exportImage(format) {
        const papers = document.querySelectorAll('.paper');
        if (papers.length > 1) alert('1ページ目のみ対応しています。');
        const finishProgress = showProgress(format, 5);
        try {
            const paper = papers[0];
            const originalTransform = paper.style.transform;
            paper.style.transform = 'none';
            const canvas = await html2canvas(paper, { scale: 2, useCORS: true, logging: false, width: 793.7, height: 1122 });
            const link = document.createElement('a');
            link.download = `${inputs.title.value}.${format}`;
            link.href = canvas.toDataURL(`image/${format === 'png' ? 'png' : 'jpeg'}`, 0.9);
            link.click();
            paper.style.transform = originalTransform;
        } catch (err) {
            console.error(err);
        } finally {
            finishProgress();
        }
    }

    if (exportPngBtn) exportPngBtn.addEventListener('click', () => exportImage('png'));
    if (exportJpgBtn) exportJpgBtn.addEventListener('click', () => exportImage('jpg'));
});
