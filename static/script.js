document.addEventListener("DOMContentLoaded", () => {
    const form = document.getElementById("mainForm");
    const submitBtn = document.getElementById("submitBtn");
    const btnText = document.getElementById("btnLabel");
    const btnLoading = submitBtn.querySelector(".btn-loading");
    const loadingLabel = document.getElementById("loadingLabel");
    const statusMessage = document.getElementById("statusMessage");
    const analysisResult = document.getElementById("analysisResult");
    const analysisText = document.getElementById("analysisText");
    const copyBtn = document.getElementById("copyBtn");
    const summaryToggle = document.getElementById("summaryToggle");
    const includeSummary = document.getElementById("includeSummary");
    const pptxOptions = document.getElementById("pptxOptions");
    const topicInput = document.getElementById("topic");
    const unifiedInput = document.getElementById("unifiedInput");
    const unifiedPlaceholder = document.getElementById("unifiedPlaceholder");
    const fileInput = document.getElementById("fileInput");
    const fileListEl = document.getElementById("fileList");
    const attachBtn = document.getElementById("attachBtn");
    const modelSelect = document.getElementById("modelSelect");
    const modelRefresh = document.getElementById("modelRefresh");
    const progressLog = document.getElementById("progressLog");
    const progressLogContent = document.getElementById("progressLogContent");
    const stopBtn = document.getElementById("stopBtn");
    let activeAbortController = null;
    let currentRequestStartMs = 0;

    // ===== ヘルプモーダル =====
    const helpBtn = document.getElementById("helpBtn");
    const helpOverlay = document.getElementById("helpOverlay");
    const helpClose = document.getElementById("helpClose");

    helpBtn.addEventListener("click", () => helpOverlay.classList.add("open"));
    helpClose.addEventListener("click", () => helpOverlay.classList.remove("open"));
    helpOverlay.addEventListener("click", (e) => {
        if (e.target === helpOverlay) helpOverlay.classList.remove("open");
    });
    document.addEventListener("keydown", (e) => {
        if (e.key === "Escape") helpOverlay.classList.remove("open");
    });

    // ===== モデル読み込み =====
    async function loadModels(refresh) {
        try {
            const url = refresh ? "/models?refresh=true" : "/models";
            const resp = await fetch(url);
            if (!resp.ok) return;
            const data = await resp.json();

            while (modelSelect.options.length > 1) {
                modelSelect.remove(1);
            }

            for (const m of data.models) {
                const opt = document.createElement("option");
                opt.value = m.id;
                if (m.available) {
                    opt.textContent = m.label;
                } else {
                    opt.textContent = m.label + " (未設定)";
                    opt.disabled = true;
                }
                if (m.description) opt.title = m.description;
                modelSelect.appendChild(opt);
            }
        } catch (e) {
            console.error("モデル一覧の取得に失敗:", e);
        }
    }

    loadModels(false);

    // ===== PPTXテンプレート読み込み =====
    async function loadPptxTemplates() {
        const templateSelect = document.getElementById("templateSelect");
        if (!templateSelect) return;
        try {
            const resp = await fetch("/pptx-templates");
            if (!resp.ok) return;
            const data = await resp.json();
            while (templateSelect.options.length > 0) {
                templateSelect.remove(0);
            }
            const templates = data.templates || [];
            if (templates.length === 0) {
                const opt = document.createElement("option");
                opt.value = "";
                opt.textContent = "テンプレートが見つかりません";
                templateSelect.appendChild(opt);
                templateSelect.disabled = true;
                return;
            }
            for (const t of templates) {
                const opt = document.createElement("option");
                opt.value = t.id;
                opt.textContent = t.label || t.id;
                templateSelect.appendChild(opt);
            }
            templateSelect.disabled = false;
            templateSelect.selectedIndex = 0;
        } catch (e) {
            console.error("PPTXテンプレート一覧の取得に失敗:", e);
        }
    }

    loadPptxTemplates();

    modelRefresh.addEventListener("click", async () => {
        modelRefresh.classList.add("spinning");
        await loadModels(true);
        modelRefresh.classList.remove("spinning");
    });

    let currentMode = "analyze";
    const selectedFiles = [];

    const ACCEPTED_EXTENSIONS = new Set([
        ".pdf", ".csv", ".xlsx", ".xls", ".msg",
        ".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif",
        ".webp", ".gif", ".ico", ".heic", ".heif", ".svg",
    ]);

    const FILE_ICONS = {
        ".pdf": "📄", ".csv": "📊", ".xlsx": "📗", ".xls": "📗",
        ".msg": "📧",
        ".jpg": "🖼️", ".jpeg": "🖼️", ".png": "🖼️", ".bmp": "🖼️",
        ".tiff": "🖼️", ".tif": "🖼️", ".webp": "🖼️", ".gif": "🖼️",
        ".ico": "🖼️", ".heic": "🖼️", ".heif": "🖼️", ".svg": "🎨",
    };

    // ===== モード切替 =====
    document.querySelectorAll(".mode-tab").forEach(tab => {
        tab.addEventListener("click", () => {
            document.querySelectorAll(".mode-tab").forEach(t => t.classList.remove("active"));
            tab.classList.add("active");
            currentMode = tab.dataset.mode;
            updateModeUI();
        });
    });

    const excelOptions = document.getElementById("excelOptions");

    function updateModeUI() {
        pptxOptions.style.display = "none";
        excelOptions.style.display = "none";
        submitBtn.classList.remove("analyze-mode", "excel-mode");

        if (currentMode === "pptx") {
            pptxOptions.style.display = "block";
            btnText.textContent = "PPTX を生成する";
            loadingLabel.textContent = "生成中...（20〜40秒かかります）";
        } else if (currentMode === "excel") {
            excelOptions.style.display = "block";
            btnText.textContent = "Excel を生成する";
            submitBtn.classList.add("excel-mode");
            loadingLabel.textContent = "Excel 生成中...（15〜30秒）";
        } else {
            btnText.textContent = "分析を実行する";
            submitBtn.classList.add("analyze-mode");
            loadingLabel.textContent = "AI が分析中...（10〜30秒）";
        }
        updateSummaryVisibility();
        analysisResult.style.display = "none";
    }

    // ===== プレースホルダー制御 =====
    function updatePlaceholder() {
        const hasText = topicInput.value.trim().length > 0;
        const hasFiles = selectedFiles.length > 0;
        unifiedPlaceholder.classList.toggle("hidden", hasText || hasFiles);
    }

    topicInput.addEventListener("input", () => {
        updatePlaceholder();
        autoResize();
    });

    function autoResize() {
        topicInput.style.height = "auto";
        topicInput.style.height = Math.min(Math.max(130, topicInput.scrollHeight), 300) + "px";
    }

    // ===== ファイル管理 =====
    function hasFiles() {
        return selectedFiles.length > 0;
    }

    function updateSummaryVisibility() {
        summaryToggle.style.display =
            (currentMode === "pptx" && hasFiles()) ? "block" : "none";
    }

    function setRunningState(isRunning) {
        btnText.style.display = isRunning ? "none" : "inline";
        btnLoading.style.display = isRunning ? "flex" : "none";
        submitBtn.disabled = isRunning;
        if (stopBtn) stopBtn.style.display = isRunning ? "inline-block" : "none";
    }

    function getExt(filename) {
        const dot = filename.lastIndexOf(".");
        return dot >= 0 ? filename.substring(dot).toLowerCase() : "";
    }

    function formatSize(bytes) {
        if (bytes < 1024) return bytes + " B";
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
        return (bytes / (1024 * 1024)).toFixed(1) + " MB";
    }

    function addFiles(fileListObj) {
        for (const file of fileListObj) {
            const ext = getExt(file.name);
            if (!ACCEPTED_EXTENSIONS.has(ext)) {
                showStatus(`未対応のファイル形式です: ${file.name}`, "error");
                continue;
            }
            if (selectedFiles.some(f => f.name === file.name && f.size === file.size)) {
                continue;
            }
            selectedFiles.push(file);
        }
        renderFileList();
        updatePlaceholder();
        updateSummaryVisibility();
    }

    function removeFile(index) {
        selectedFiles.splice(index, 1);
        renderFileList();
        updatePlaceholder();
        updateSummaryVisibility();
    }

    function clearAllFiles() {
        selectedFiles.length = 0;
        renderFileList();
        updatePlaceholder();
        updateSummaryVisibility();
    }

    function renderFileList() {
        if (selectedFiles.length === 0) {
            fileListEl.innerHTML = "";
            return;
        }

        const chips = selectedFiles.map((file, i) => {
            const ext = getExt(file.name);
            const icon = FILE_ICONS[ext] || "📎";
            return `<span class="file-chip">
                <span class="file-chip-icon">${icon}</span>
                <span class="file-chip-name" title="${file.name}">${file.name}</span>
                <span class="file-chip-size">${formatSize(file.size)}</span>
                <button type="button" class="file-chip-remove" data-index="${i}" title="削除">✕</button>
            </span>`;
        }).join("");

        fileListEl.innerHTML = chips;
    }

    fileListEl.addEventListener("click", (e) => {
        const btn = e.target.closest(".file-chip-remove");
        if (btn) {
            removeFile(parseInt(btn.dataset.index, 10));
        }
    });

    // ===== ドラッグ＆ドロップ =====
    attachBtn.addEventListener("click", (e) => {
        e.preventDefault();
        fileInput.click();
    });

    unifiedInput.addEventListener("dragover", (e) => {
        e.preventDefault();
        unifiedInput.classList.add("drag-over");
    });

    unifiedInput.addEventListener("dragleave", (e) => {
        if (!unifiedInput.contains(e.relatedTarget)) {
            unifiedInput.classList.remove("drag-over");
        }
    });

    unifiedInput.addEventListener("drop", (e) => {
        e.preventDefault();
        unifiedInput.classList.remove("drag-over");
        if (e.dataTransfer.files.length) {
            addFiles(e.dataTransfer.files);
        }
    });

    fileInput.addEventListener("change", () => {
        if (fileInput.files.length) {
            addFiles(fileInput.files);
        }
        fileInput.value = "";
    });

    // ===== クリップボードペースト =====
    document.addEventListener("paste", (e) => {
        const items = e.clipboardData?.items;
        if (!items) return;

        for (const item of items) {
            if (item.kind === "file" && item.type.startsWith("image/")) {
                e.preventDefault();
                const file = item.getAsFile();
                if (!file) continue;

                const extMap = {
                    "image/png": "png", "image/jpeg": "jpg", "image/gif": "gif",
                    "image/webp": "webp", "image/bmp": "bmp", "image/tiff": "tiff",
                };
                const ext = extMap[file.type] || "png";
                const ts = new Date().toISOString().slice(11, 19).replace(/:/g, "");
                const namedFile = new File([file], `paste_${ts}.${ext}`, { type: file.type });
                addFiles([namedFile]);
            }
        }
    });

    // ===== コピーボタン =====
    copyBtn.addEventListener("click", () => {
        navigator.clipboard.writeText(analysisText.textContent).then(() => {
            const orig = copyBtn.textContent;
            copyBtn.textContent = "コピーしました!";
            setTimeout(() => { copyBtn.textContent = orig; }, 2000);
        });
    });

    // ===== フォーム送信 =====
    form.addEventListener("submit", async (e) => {
        e.preventDefault();

        const text = topicInput.value.trim();

        if ((currentMode === "pptx" || currentMode === "excel") && !text) {
            showStatus("トピックを入力してください。", "error");
            return;
        }

        if (currentMode === "analyze" && !text && !hasFiles()) {
            showStatus("テキストを入力するか、ファイルを添付してください。", "error");
            return;
        }

        setRunningState(true);
        statusMessage.style.display = "none";
        analysisResult.style.display = "none";

        try {
            if (currentMode === "pptx") {
                await handlePptxGeneration();
            } else if (currentMode === "excel") {
                await handleExcelGeneration();
            } else {
                await handleAnalysis();
            }
        } finally {
            setRunningState(false);
        }
    });

    if (stopBtn) {
        stopBtn.addEventListener("click", () => {
            if (activeAbortController) {
                activeAbortController.abort();
                const elapsed = currentRequestStartMs
                    ? Math.floor((Date.now() - currentRequestStartMs) / 1000)
                    : 0;
                appendProgressLog("処理停止を要求しました。", elapsed);
            }
        });
    }

    function hasImageFiles() {
        const imgExts = [".jpg",".jpeg",".png",".bmp",".tiff",".tif",".webp",".gif",".ico",".heic",".heif",".svg"];
        return selectedFiles.some(f => imgExts.includes(getExt(f.name)));
    }

    async function handlePptxGeneration() {
        const templateSelect = document.getElementById("templateSelect");
        if (!templateSelect || !templateSelect.value) {
            showStatus("PPTXテンプレートを選択してください。", "error");
            return;
        }
        const instructions = document.getElementById("additionalInstructions").value.trim();

        const formData = new FormData();
        formData.append("topic", topicInput.value.trim());
        formData.append("num_slides", document.getElementById("numSlides").value);
        formData.append("style", document.getElementById("style").value);
        formData.append("language", document.getElementById("language").value);
        formData.append("additional_instructions", instructions);
        formData.append("model", modelSelect.value);
        formData.append("template", templateSelect.value);

        for (const file of selectedFiles) {
            formData.append("files", file);
        }

        if (hasFiles() && includeSummary.checked) {
            formData.append("include_summary", "true");
        }

        const phaseMessages = hasFiles()
            ? hasImageFiles()
                ? ["リクエスト送信中...", "ファイル解析中...", "画像をAIで分析中...", "AIでスライド構成を生成中...", "PPTX生成中..."]
                : ["リクエスト送信中...", "ファイル解析中...", "AIでスライド構成を生成中...", "PPTX生成中..."]
            : ["リクエスト送信中...", "AIでコンテンツ生成中...", "チャート・画像を準備中...", "PPTX生成中..."];

        try {
            clearProgressLog();
            appendProgressLog(phaseMessages[0], 0);

            const response = await runWithProgressLog(
                phaseMessages,
                async (signal) => fetch("/generate", { method: "POST", body: formData, signal }),
                { timeoutMs: 180000, intervalMs: 3000 }
            );
            if (!response.ok) {
                let errMsg = `HTTP ${response.status}`;
                try {
                    const err = await response.json();
                    if (err.detail) errMsg = typeof err.detail === "string" ? err.detail : JSON.stringify(err.detail);
                } catch (_) {
                    const text = await response.text();
                    if (text) errMsg = text.slice(0, 200);
                }
                throw new Error(errMsg);
            }
            appendUsedModelLog(response);

            const warnings = response.headers.get("X-Processing-Warnings");

            const blob = await response.blob();
            const cd = response.headers.get("content-disposition");
            let filename = "presentation.pptx";
            if (cd) {
                const match = cd.match(/filename\*?=(?:UTF-8'')?([^;]+)/i);
                if (match) filename = decodeURIComponent(match[1].replace(/"/g, ""));
            }

            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);

            if (warnings) {
                showStatus(`生成完了（警告あり）: ${warnings}`, "warning");
            } else {
                showStatus("生成が完了しました！ ダウンロードが開始されます。", "success");
            }
        } catch (err) {
            const msg = err.name === "AbortError"
                ? "処理がタイムアウトしました（3分）。AIの応答が遅い場合があります。スライド枚数を減らすか、しばらく待って再試行してください。"
                : `エラーが発生しました: ${err.message}`;
            showStatus(msg, "error");
        }
    }

    async function handleExcelGeneration() {
        const formData = new FormData();
        formData.append("topic", topicInput.value.trim());
        formData.append("style", document.getElementById("excelStyle").value);
        formData.append("additional_instructions",
            (document.getElementById("excelInstructions").value || "").trim());
        formData.append("model", modelSelect.value);

        for (const file of selectedFiles) {
            formData.append("files", file);
        }

        try {
            const phaseMessages = hasFiles()
                ? hasImageFiles()
                    ? ["リクエスト送信中...", "ファイル解析中...", "画像をAIで分析中...", "Excelレポートを生成中..."]
                    : ["リクエスト送信中...", "ファイル解析中...", "Excelレポートを生成中..."]
                : ["リクエスト送信中...", "AIでレポート構成を生成中...", "Excelファイルを生成中..."];

            clearProgressLog();
            appendProgressLog(phaseMessages[0], 0);

            const response = await runWithProgressLog(
                phaseMessages,
                async (signal) => fetch("/generate-excel", { method: "POST", body: formData, signal }),
                { timeoutMs: 180000, intervalMs: 3000 }
            );
            if (!response.ok) {
                let errMsg = `HTTP ${response.status}`;
                try {
                    const err = await response.json();
                    if (err.detail) errMsg = typeof err.detail === "string" ? err.detail : JSON.stringify(err.detail);
                } catch (_) {
                    const text = await response.text();
                    if (text) errMsg = text.slice(0, 200);
                }
                throw new Error(errMsg);
            }
            appendUsedModelLog(response);

            const warnings = response.headers.get("X-Processing-Warnings");

            const blob = await response.blob();
            const cd = response.headers.get("content-disposition");
            let filename = "report.xlsx";
            if (cd) {
                const match = cd.match(/filename\*?=(?:UTF-8'')?([^;]+)/i);
                if (match) filename = decodeURIComponent(match[1].replace(/"/g, ""));
            }

            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);

            if (warnings) {
                showStatus(`生成完了（警告あり）: ${warnings}`, "warning");
            } else {
                showStatus("Excel の生成が完了しました！ ダウンロードが開始されます。", "success");
            }
        } catch (err) {
            const msg = err.name === "AbortError"
                ? "処理を停止しました。"
                : `エラーが発生しました: ${err.message}`;
            showStatus(msg, "error");
        }
    }

    async function handleAnalysis() {
        const formData = new FormData();
        formData.append("topic", topicInput.value.trim());
        formData.append("model", modelSelect.value);

        for (const file of selectedFiles) {
            formData.append("files", file);
        }

        try {
            const phaseMessages = hasFiles()
                ? hasImageFiles()
                    ? ["リクエスト送信中...", "ファイル解析中...", "画像をAIで分析中...", "分析レポートを生成中..."]
                    : ["リクエスト送信中...", "ファイルを読み込み中...", "分析レポートを生成中..."]
                : ["リクエスト送信中...", "AIが回答を生成中..."];

            clearProgressLog();
            appendProgressLog(phaseMessages[0], 0);

            const response = await runWithProgressLog(
                phaseMessages,
                async (signal) => fetch("/analyze", { method: "POST", body: formData, signal }),
                { timeoutMs: 180000, intervalMs: 3000 }
            );
            if (!response.ok) {
                let errMsg = `HTTP ${response.status}`;
                try {
                    const err = await response.json();
                    if (err.detail) errMsg = typeof err.detail === "string" ? err.detail : JSON.stringify(err.detail);
                } catch (_) {
                    const text = await response.text();
                    if (text) errMsg = text.slice(0, 200);
                }
                throw new Error(errMsg);
            }
            appendUsedModelLog(response);

            const data = await response.json();
            analysisText.textContent = data.analysis;
            analysisResult.style.display = "block";

            if (data.warnings && data.warnings.length > 0) {
                showStatus(`分析完了（警告: ${data.warnings.join(", ")}）`, "warning");
            } else {
                showStatus("分析が完了しました！", "success");
            }

            analysisResult.scrollIntoView({ behavior: "smooth", block: "start" });
        } catch (err) {
            const msg = err.name === "AbortError"
                ? "処理を停止しました。"
                : `エラーが発生しました: ${err.message}`;
            showStatus(msg, "error");
        }
    }

    function showStatus(message, type) {
        if (progressLogContent.children.length === 0) {
            progressLog.style.display = "none";
        } else {
            progressLog.style.display = "block";
        }
        statusMessage.textContent = message;
        statusMessage.className = `status-message ${type}`;
        statusMessage.style.display = "block";
    }

    function clearProgressLog() {
        progressLogContent.innerHTML = "";
        progressLog.style.display = "none";
        statusMessage.style.display = "none";
    }

    function appendProgressLog(message, elapsedSec) {
        progressLog.style.display = "block";
        statusMessage.style.display = "none";
        const line = document.createElement("div");
        line.className = "progress-log-line";
        const timeStr = elapsedSec >= 60
            ? `${Math.floor(elapsedSec / 60)}分${elapsedSec % 60}秒`
            : `${elapsedSec}秒`;
        line.innerHTML = `<span class="progress-log-time">[${timeStr}]</span><span>${message}</span>`;
        progressLogContent.appendChild(line);
        progressLogContent.scrollTop = progressLogContent.scrollHeight;
    }

    function appendUsedModelLog(response, elapsedSec = null) {
        const usedModel = response?.headers?.get("X-AI-Model-Used");
        if (!usedModel) return;
        const sec = elapsedSec ?? (currentRequestStartMs
            ? Math.floor((Date.now() - currentRequestStartMs) / 1000)
            : 0);
        appendProgressLog(`使用モデル: ${usedModel}`, sec);
    }

    function runWithProgressLog(phaseMessages, fetchFn, options = {}) {
        const timeoutMs = options.timeoutMs || 120000;
        const intervalMs = options.intervalMs || 3000;
        const startTime = Date.now();
        currentRequestStartMs = startTime;
        let phaseIdx = 1; // 0番は呼び出し元で既に表示済み

        const timer = setInterval(() => {
            const elapsed = Math.floor((Date.now() - startTime) / 1000);
            if (phaseIdx < phaseMessages.length) {
                appendProgressLog(phaseMessages[phaseIdx], elapsed);
                phaseIdx++;
            } else {
                appendProgressLog(`${phaseMessages[phaseMessages.length - 1]} (${elapsed}秒経過)`, elapsed);
            }
        }, intervalMs);

        activeAbortController = new AbortController();
        const timeoutId = setTimeout(() => activeAbortController.abort(), timeoutMs);

        return fetchFn(activeAbortController.signal)
            .finally(() => {
                clearInterval(timer);
                clearTimeout(timeoutId);
                activeAbortController = null;
                currentRequestStartMs = 0;
            });
    }
});
