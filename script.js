const DOC_TEMPLATES = {
    announcement: {
        title: "生徒会からのお知らせ",
        issuer: "生徒会執行部",
        sections: [
            { heading: "文化祭の開催について", body: "ここに本文を入力してください。\n改行も反映されます。\n\n・箇条書きなども\n・使いやすいです" }
        ]
    },
    proposal: {
        title: "事業提案書",
        issuer: "生徒会執行部",
        sections: [
            { heading: "1. 目的", body: "本プロジェクトの目的と背景を記載してください。" },
            { heading: "2. 提案内容", body: "具体的な施策や導入メリットを詳しく説明します。" }
        ]
    },
    planning: {
        title: "プロジェクト企画書",
        issuer: "生徒会執行部",
        sections: [
            { heading: "企画概要", body: "企画の核心となるアイデアを簡潔にまとめます。" },
            { heading: "ターゲット層", body: "主要なターゲットと市場ニーズを分析します。" }
        ]
    },

    contract: {
        title: "業務委託契約書",
        issuer: "生徒会執行部",
        sections: [
            { heading: "第1条（目的）", body: "本契約は、甲が乙に対し、業務を委託することを目的とする。" },
            { heading: "第2条（報酬）", body: "報酬の支払時期および方法について規定します。" }
        ]
    },
    purchase_order: {
        title: "発注書",
        issuer: "生徒会執行部",
        sections: [
            { heading: "発注内容", body: "品名：商品A\n数量：10個\n単価：1,000円" },
            { heading: "納期・場所", body: "2024年X月Y日まで。弊社倉庫へ納品。" }
        ]
    }
};

const DEFAULT_ISSUERS = ["生徒会執行部"];
let customIssuers = [];
let savedGradients = [
    { primary: "#D32F2F", accent: "#FF5252" },
    { primary: "#1976D2", accent: "#42A5F5" },
    { primary: "#388E3C", accent: "#66BB6A" },
    { primary: "#7B1FA2", accent: "#AB47BC" },
    { primary: "#FBC02D", accent: "#FFF176" }
]; // Default gradient themes

document.addEventListener('DOMContentLoaded', () => {
    // --- Elements ---
    const inputs = {
        docType: document.getElementById('input-doc-type'),
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
    };

    const issuerList = document.getElementById('issuer-list');

    const preview = {
        title: document.getElementById('preview-title'),
        date: document.getElementById('preview-date'),
        issuer: document.getElementById('preview-issuer'),
        vol: document.getElementById('preview-vol'),
        sections: document.getElementById('preview-sections'),
        footer: document.getElementById('preview-footer'),
        image: document.getElementById('preview-image'),
        imageContainer: document.getElementById('preview-image-container'),
        paper: document.getElementById('paper'),
    };

    const sectionsContainer = document.getElementById('sections-container');
    const btnAddSection = document.getElementById('btn-add-section');
    const layoutBtns = document.querySelectorAll('.layout-btn');
    const colorCode = document.getElementById('color-code');
    const exportBtn = document.getElementById('btn-export');
    const removeImgBtn = document.getElementById('btn-remove-image');
    const btnDeleteIssuer = document.getElementById('btn-delete-issuer');
    const historyContainer = document.getElementById('history-container');
    const btnSaveHistory = document.getElementById('btn-save-history');
    const btnSaveGradient = document.getElementById('btn-save-gradient');
    const gradientsContainer = document.getElementById('saved-gradients-container');
    const colorCodeAccent = document.getElementById('color-code-accent');

    const valDisplays = {
        titleSize: document.getElementById('val-size-title'),
        bodySize: document.getElementById('val-size-body'),
    };

    let zoomLevel = 0.8;
    const zoomInBtn = document.getElementById('zoom-in');
    const zoomOutBtn = document.getElementById('zoom-out');
    const zoomDisplay = document.getElementById('zoom-level');

    // --- Initialization ---
    // Set default date to today
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    inputs.date.value = `${yyyy}-${mm}-${dd}`;

    loadState(); // Initial load
    loadGradients();
    updateIssuerSelect();
    renderHistory();
    renderGradients();
    updatePreview();

    // --- Event Listeners ---

    // Doc Type Change
    inputs.docType.addEventListener('change', (e) => {
        const templateKey = e.target.value.replace('-', '_');
        const template = DOC_TEMPLATES[templateKey];
        if (template) {
            inputs.title.value = template.title;
            inputs.issuer.value = template.issuer;
            sectionsContainer.innerHTML = '';
            template.sections.forEach(s => createSectionUI(s.heading, s.body));
            updatePreview();
            saveState();
        }
    });

    // Font size display update
    inputs.titleSize.addEventListener('input', (e) => {
        valDisplays.titleSize.textContent = e.target.value;
        updatePreview();
        saveState();
    });
    inputs.bodySize.addEventListener('input', (e) => {
        valDisplays.bodySize.textContent = e.target.value;
        updatePreview();
        saveState();
    });

    // Generic listener for interactive inputs
    Object.keys(inputs).forEach(key => {
        const el = inputs[key];
        if (el && el.type !== 'file' && el !== inputs.docType && el !== inputs.titleSize && el !== inputs.bodySize && el !== inputs.issuerSelect) {
            const eventType = el.tagName === 'SELECT' ? 'change' : 'input';
            el.addEventListener(eventType, () => {
                updatePreview();
                saveState();
            });
        }
    });

    // Issuer Select Handling
    inputs.issuerSelect.addEventListener('change', (e) => {
        if (e.target.value === 'ADD_NEW') {
            inputs.issuerSelect.parentElement.style.display = 'none';
            inputs.issuer.style.display = 'block';
            inputs.issuer.value = '';
            inputs.issuer.focus();
        } else {
            inputs.issuer.value = e.target.value;
            toggleDeleteBtn();
            updatePreview();
            saveState();
        }
    });

    btnDeleteIssuer.addEventListener('click', () => {
        const val = inputs.issuerSelect.value;
        if (customIssuers.includes(val)) {
            customIssuers = customIssuers.filter(i => i !== val);
            inputs.issuer.value = "生徒会執行部";
            updateIssuerSelect();
            updatePreview();
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

    // Special listener for issuer text input to "finish" adding
    inputs.issuer.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            addIssuerIfNew(inputs.issuer.value);
        }
    });

    inputs.issuer.addEventListener('blur', () => {
        addIssuerIfNew(inputs.issuer.value);
    });

    // Color Picker
    inputs.color.addEventListener('input', (e) => {
        const color = e.target.value;
        colorCode.textContent = color;
        document.documentElement.style.setProperty('--primary-color', color);
        updatePreview();
        saveState();
    });

    btnSaveGradient.addEventListener('click', () => {
        const p = inputs.color.value;
        const a = inputs.colorAccent.value;
        // Check if already exists
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
        if (saved) {
            savedGradients = JSON.parse(saved);
        }
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
            swatch.title = `${grad.primary} / ${grad.accent}`;

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
                preview.paper.style.setProperty('--paper-primary', grad.primary);
                preview.paper.style.setProperty('--paper-accent', grad.accent);
                renderGradients();
                updatePreview();
                saveState();
            });

            gradientsContainer.appendChild(swatch);
        });
    }

    // Color Pickers
    inputs.color.addEventListener('input', (e) => {
        const color = e.target.value;
        colorCode.textContent = color;
        preview.paper.style.setProperty('--paper-primary', color);
        updatePreview();
        saveState();
    });

    inputs.colorAccent.addEventListener('input', (e) => {
        const color = e.target.value;
        if (colorCodeAccent) colorCodeAccent.textContent = color;
        preview.paper.style.setProperty('--paper-accent', color);
        updatePreview();
        saveState();
    });
    // Layout Switching
    layoutBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            layoutBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            const layout = btn.dataset.layout;
            preview.paper.className = `paper a4 layout-${layout}`;
            updatePreview();
            saveState();
        });
    });

    // Section Management
    btnAddSection.addEventListener('click', () => {
        createSectionUI("", "");
        updatePreview();
        saveState();
    });

    // History Actions
    btnSaveHistory.addEventListener('click', () => {
        saveToHistory();
    });

    function saveToHistory() {
        const history = JSON.parse(localStorage.getItem('announcement_tool_history') || '[]');
        const currentState = getCurrentState();

        // Add timestamp for the history entry
        const entry = {
            id: Date.now(),
            savedAt: new Date().toLocaleString(),
            state: currentState
        };

        history.unshift(entry); // Add to beginning
        // Keep only last 20 entries to avoid overflow
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
                <div class="history-item-info">
                    <span class="history-item-title">${item.state.title}</span>
                    <span class="history-item-date">${item.savedAt}</span>
                </div>
                <div class="history-item-actions">
                    <button class="btn-delete-history" title="削除"><i class="fa-solid fa-trash"></i></button>
                </div>
            `;

            // Click to load
            historyItem.addEventListener('click', (e) => {
                if (e.target.closest('.btn-delete-history')) return;
                if (confirm('この履歴を読み込みますか？現在の編集内容は上書きされます。')) {
                    applyState(item.state);
                    updatePreview();
                    saveState();
                }
            });

            // Delete specific history
            historyItem.querySelector('.btn-delete-history').addEventListener('click', (e) => {
                e.stopPropagation();
                if (confirm('この履歴を削除しますか？')) {
                    const updatedHistory = history.filter(h => h.id !== item.id);
                    localStorage.setItem('announcement_tool_history', JSON.stringify(updatedHistory));
                    renderHistory();
                }
            });

            historyContainer.appendChild(historyItem);
        });
    }

    function createSectionUI(headingValue, bodyValue) {
        const sectionDiv = document.createElement('div');
        sectionDiv.className = 'sidebar-section';
        sectionDiv.innerHTML = `
            <div class="section-controls">
                <span class="section-label">セクション</span>
                <button class="btn-remove-section"><i class="fa-solid fa-trash"></i></button>
            </div>
            <div class="input-wrap">
                <input type="text" class="section-heading" value="${headingValue}" placeholder="見出し">
            </div>
            <div class="input-wrap">
                <textarea class="section-body" rows="4" placeholder="本文">${bodyValue}</textarea>
            </div>
        `;

        sectionDiv.querySelector('.btn-remove-section').addEventListener('click', () => {
            sectionDiv.remove();
            updatePreview();
            saveState();
        });

        sectionDiv.querySelectorAll('input, textarea').forEach(el => {
            el.addEventListener('input', () => {
                updatePreview();
                saveState();
            });
        });

        sectionsContainer.appendChild(sectionDiv);
    }

    // Image Upload
    inputs.imageFile.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                preview.image.src = e.target.result;
                preview.imageContainer.style.display = 'block';
                removeImgBtn.style.display = 'inline-block';
                updatePreview();
                saveState();
            };
            reader.readAsDataURL(file);
        }
    });

    removeImgBtn.addEventListener('click', () => {
        inputs.imageFile.value = '';
        preview.image.src = '';
        preview.imageContainer.style.display = 'none';
        removeImgBtn.style.display = 'none';
        updatePreview();
        saveState();
    });

    // Zoom
    function setZoom(level) {
        zoomLevel = level;
        preview.paper.style.transform = `scale(${zoomLevel})`;
        zoomDisplay.textContent = `${Math.round(zoomLevel * 100)}%`;
    }
    zoomInBtn.addEventListener('click', () => zoomLevel < 1.5 && setZoom(zoomLevel + 0.1));
    zoomOutBtn.addEventListener('click', () => zoomLevel > 0.3 && setZoom(zoomLevel - 0.1));
    setZoom(0.75);

    function updateIssuerSelect() {
        if (!inputs.issuerSelect) return;
        // Keep only standard options + ADD_NEW
        const currentVal = inputs.issuer.value;
        inputs.issuerSelect.innerHTML = `
            <option value="生徒会執行部">生徒会執行部</option>
        `;

        customIssuers.forEach(issuer => {
            const opt = document.createElement('option');
            opt.value = issuer;
            opt.textContent = issuer;
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
            // If empty, just switch back
            inputs.issuer.style.display = 'none';
            inputs.issuerSelect.parentElement.style.display = 'flex';
            inputs.issuerSelect.value = "生徒会執行部";
            inputs.issuer.value = "生徒会執行部";
            updatePreview();
            return;
        }
        if (!DEFAULT_ISSUERS.includes(value) && !customIssuers.includes(value) && value !== "生徒会執行部") {
            customIssuers.push(value);
        }
        updateIssuerSelect();
        inputs.issuer.style.display = 'none';
        inputs.issuerSelect.parentElement.style.display = 'flex';
        inputs.issuerSelect.value = value;
        updatePreview();
        saveState();
    }

    // --- Core Functions ---

    function updatePreview() {
        // Date
        const dateVal = inputs.date.value;
        if (dateVal) {
            const parts = dateVal.split('-');
            if (parts.length === 3) {
                preview.date.textContent = `${parts[0]}年${parseInt(parts[1], 10)}月${parseInt(parts[2], 10)}日`;
            }
        }

        preview.title.textContent = inputs.title.value;
        preview.issuer.textContent = inputs.issuer.value;
        preview.vol.textContent = inputs.vol.value;

        if (inputs.font) preview.paper.style.fontFamily = inputs.font.value;
        if (inputs.titleSize) preview.title.style.fontSize = `${inputs.titleSize.value}pt`;

        // Render sections in preview
        preview.sections.innerHTML = '';
        const sidebarSections = sectionsContainer.querySelectorAll('.sidebar-section');
        sidebarSections.forEach(sec => {
            const h = sec.querySelector('.section-heading').value;
            const b = sec.querySelector('.section-body').value;

            const sectionDiv = document.createElement('div');
            sectionDiv.className = 'preview-section';
            sectionDiv.innerHTML = `
                <div class="main-heading-container">
                    <h2 class="main-heading">${h}</h2>
                </div>
                <div class="text-content" style="font-size: ${inputs.bodySize.value}pt">
                    <p style="white-space: pre-wrap;">${b}</p>
                </div>
            `;
            preview.sections.appendChild(sectionDiv);
        });

        if (inputs.imageSize && preview.image) {
            preview.image.style.width = `${inputs.imageSize.value}%`;
            preview.image.style.maxWidth = `${inputs.imageSize.value}%`;
        }

        preview.footer.textContent = inputs.footer.value;
    }

    function getCurrentState() {
        const sectionsData = [];
        sectionsContainer.querySelectorAll('.sidebar-section').forEach(sec => {
            sectionsData.push({
                heading: sec.querySelector('.section-heading').value,
                body: sec.querySelector('.section-body').value
            });
        });

        return {
            docType: inputs.docType.value,
            title: inputs.title.value,
            date: inputs.date.value,
            issuer: inputs.issuer.value,
            vol: inputs.vol.value,
            color: inputs.color.value,
            colorAccent: inputs.colorAccent.value,
            font: inputs.font.value,
            footer: inputs.footer.value,
            titleSize: inputs.titleSize.value,
            bodySize: inputs.bodySize.value,
            imageSize: inputs.imageSize.value,
            sections: sectionsData,
            customIssuers: customIssuers,
            layout: document.querySelector('.layout-btn.active').dataset.layout,
            imageData: preview.image.src
        };
    }

    function saveState() {
        const state = getCurrentState();
        localStorage.setItem('announcement_tool_state', JSON.stringify(state));
    }

    function applyState(state) {
        try {
            inputs.docType.value = state.docType || 'announcement';
            inputs.title.value = state.title;
            inputs.date.value = state.date;
            inputs.issuer.value = state.issuer;
            inputs.vol.value = state.vol;
            inputs.color.value = state.color;
            inputs.colorAccent.value = state.colorAccent || '#FF5252';
            inputs.font.value = state.font;
            inputs.footer.value = state.footer;
            inputs.titleSize.value = state.titleSize;
            inputs.bodySize.value = state.bodySize;
            inputs.imageSize.value = state.imageSize;

            customIssuers = state.customIssuers || [];
            updateIssuerSelect();
            if (inputs.issuerSelect) inputs.issuerSelect.value = state.issuer || inputs.issuer.value;

            colorCode.textContent = state.color;
            if (colorCodeAccent) colorCodeAccent.textContent = state.colorAccent || '#FF5252';
            preview.paper.style.setProperty('--paper-primary', state.color);
            preview.paper.style.setProperty('--paper-accent', state.colorAccent || '#FF5252');
            valDisplays.titleSize.textContent = state.titleSize;
            valDisplays.bodySize.textContent = state.bodySize;

            layoutBtns.forEach(btn => {
                if (btn.dataset.layout === state.layout) {
                    layoutBtns.forEach(b => b.classList.remove('active'));
                    btn.classList.add('active');
                    preview.paper.className = `paper a4 layout-${state.layout}`;
                }
            });

            if (state.imageData && state.imageData !== "" && state.imageData.startsWith('data:')) {
                preview.image.src = state.imageData;
                preview.imageContainer.style.display = 'block';
                removeImgBtn.style.display = 'inline-block';
            } else {
                preview.image.src = '';
                preview.imageContainer.style.display = 'none';
                removeImgBtn.style.display = 'none';
            }

            sectionsContainer.innerHTML = '';
            state.sections.forEach(s => createSectionUI(s.heading, s.body));
        } catch (e) {
            console.error("Failed to apply state", e);
        }
    }

    function loadState() {
        const saved = localStorage.getItem('announcement_tool_state');
        if (!saved) {
            createSectionUI("文化祭の開催について", "ここに本文を入力してください。");
            return;
        }
        applyState(JSON.parse(saved));
    }

    // --- PDF Export ---
    exportBtn.addEventListener('click', () => {
        const element = document.getElementById('paper');

        // Sanitize filename
        const safeTitle = (inputs.title.value || 'document').replace(/[\\/:*?"<>|]/g, '_');
        const safeDate = (inputs.date.value || '').replace(/[\\/:*?"<>|]/g, '_');
        const filename = `${safeTitle}_${safeDate}.pdf`;

        const opt = {
            margin: 0,
            filename: filename,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: {
                scale: 2,
                useCORS: true,
                letterRendering: true,
                scrollX: 0,
                scrollY: 0,
                // Critical for correct layout scaling
                windowWidth: 794 // Approx A4 width in px at 96dpi
            },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
            pagebreak: { mode: 'avoid-all' }
        };

        const originalTransform = element.style.transform;
        const originalHeight = element.style.height;
        const originalWidth = element.style.width;
        const originalOverflow = element.style.overflow;
        const originalShadow = element.style.boxShadow;

        // Force single page dimensions (793.7px wide by 1122.5px high is A4 at 96dpi)
        element.style.transform = 'none';
        element.style.width = '793.7px';
        element.style.height = '1122px'; // Just slightly shorter to be safe
        element.style.overflow = 'hidden';
        element.style.boxShadow = 'none';

        exportBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> 生成中...';
        exportBtn.disabled = true;

        html2pdf().set(opt).from(element).save().then(() => {
            element.style.transform = originalTransform;
            element.style.width = originalWidth;
            element.style.height = originalHeight;
            element.style.overflow = originalOverflow;
            element.style.boxShadow = originalShadow;
            exportBtn.innerHTML = '<i class="fa-solid fa-file-pdf"></i> PDFとして保存';
            exportBtn.disabled = false;
        }).catch(err => {
            console.error("PDF Export Error:", err);
            alert('PDFの書き出し中にエラーが発生しました。コンソールで詳細を確認してください。');
            element.style.transform = originalTransform;
            element.style.width = originalWidth;
            element.style.height = originalHeight;
            element.style.overflow = originalOverflow;
            element.style.boxShadow = originalShadow;
            exportBtn.innerHTML = '<i class="fa-solid fa-file-pdf"></i> PDFとして保存';
            exportBtn.disabled = false;
        });
    });
});
