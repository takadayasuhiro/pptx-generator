document.addEventListener("DOMContentLoaded", () => {
    const form = document.getElementById("generateForm");
    const submitBtn = document.getElementById("submitBtn");
    const btnText = submitBtn.querySelector(".btn-text");
    const btnLoading = submitBtn.querySelector(".btn-loading");
    const statusMessage = document.getElementById("statusMessage");

    const pdfFile = document.getElementById("pdfFile");
    const csvFile = document.getElementById("csvFile");
    const pdfFileName = document.getElementById("pdfFileName");
    const csvFileName = document.getElementById("csvFileName");
    const pdfClear = document.getElementById("pdfClear");
    const csvClear = document.getElementById("csvClear");
    const summaryWrap = document.getElementById("summaryOptionWrap");
    const includeSummary = document.getElementById("includeSummary");

    function updateSummaryVisibility() {
        const hasFiles = pdfFile.files.length > 0 || csvFile.files.length > 0;
        summaryWrap.style.display = hasFiles ? "block" : "none";
    }

    pdfFile.addEventListener("change", () => {
        if (pdfFile.files.length) {
            pdfFileName.textContent = pdfFile.files[0].name;
            pdfClear.style.display = "inline-block";
        }
        updateSummaryVisibility();
    });

    csvFile.addEventListener("change", () => {
        if (csvFile.files.length) {
            csvFileName.textContent = csvFile.files[0].name;
            csvClear.style.display = "inline-block";
        }
        updateSummaryVisibility();
    });

    pdfClear.addEventListener("click", () => {
        pdfFile.value = "";
        pdfFileName.textContent = "未選択";
        pdfClear.style.display = "none";
        updateSummaryVisibility();
    });

    csvClear.addEventListener("click", () => {
        csvFile.value = "";
        csvFileName.textContent = "未選択";
        csvClear.style.display = "none";
        updateSummaryVisibility();
    });

    form.addEventListener("submit", async (e) => {
        e.preventDefault();

        btnText.style.display = "none";
        btnLoading.style.display = "flex";
        submitBtn.disabled = true;
        statusMessage.style.display = "none";

        const selectedTheme = document.querySelector('input[name="theme"]:checked')?.value || "auto";

        let instructions = document.getElementById("additionalInstructions").value.trim();
        if (selectedTheme !== "auto") {
            instructions += `\n\nカラーテーマは必ず「${selectedTheme}」を使用してください。`;
        }

        const formData = new FormData();
        formData.append("topic", document.getElementById("topic").value.trim());
        formData.append("num_slides", document.getElementById("numSlides").value);
        formData.append("style", document.getElementById("style").value);
        formData.append("language", document.getElementById("language").value);
        formData.append("additional_instructions", instructions);

        if (pdfFile.files.length) {
            formData.append("pdf_file", pdfFile.files[0]);
        }
        if (csvFile.files.length) {
            formData.append("csv_file", csvFile.files[0]);
        }

        const hasFiles = pdfFile.files.length > 0 || csvFile.files.length > 0;

        if (hasFiles && includeSummary.checked) {
            formData.append("include_summary", "true");
        }

        try {
            showStatus(
                hasFiles
                    ? "ファイルを解析中... → AI でスライド構成を生成中..."
                    : "AI でコンテンツ生成中... チャート・画像も準備します",
                "info"
            );

            const response = await fetch("/generate", {
                method: "POST",
                body: formData,
            });

            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.detail || `HTTP ${response.status}`);
            }

            const blob = await response.blob();
            const contentDisposition = response.headers.get("content-disposition");
            let filename = "presentation.pptx";
            if (contentDisposition) {
                const match = contentDisposition.match(/filename\*?=(?:UTF-8'')?([^;]+)/i);
                if (match) {
                    filename = decodeURIComponent(match[1].replace(/"/g, ""));
                }
            }

            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);

            showStatus("生成が完了しました！ダウンロードが開始されます。", "success");
        } catch (err) {
            showStatus(`エラーが発生しました: ${err.message}`, "error");
        } finally {
            btnText.style.display = "inline";
            btnLoading.style.display = "none";
            submitBtn.disabled = false;
        }
    });

    function showStatus(message, type) {
        statusMessage.textContent = message;
        statusMessage.className = `status-message ${type}`;
        statusMessage.style.display = "block";
    }
});
