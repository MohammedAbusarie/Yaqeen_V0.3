import { processMultipleImages, generateTextFile } from "../ocr.js";
import { downloadBlob } from "../dom.js";

/**
 * @param {{ els: any, state: import('../state.js').AppState }} refs
 */
export function createOcrHandlers(refs) {
  const { els, state } = refs;

  // ========== OCR Handlers ==========

  /**
   * Handle OCR image upload
   */
  async function handleOcrImageUpload() {
    const input = els.ocrImageUpload;
    if (!input || !input.files || input.files.length === 0) {
      setOcrStatus("Please select at least one image file.", "error");
      return;
    }

    const files = Array.from(input.files);
    state.ocr.uploadedImages = files;
    state.ocr.currentStep = 1;

    // Show image previews
    const previewContainer = els.ocrImagePreview;
    if (previewContainer) {
      previewContainer.innerHTML = "";
      files.forEach((file, index) => {
        const img = document.createElement("img");
        img.src = URL.createObjectURL(file);
        img.style.maxWidth = "200px";
        img.style.maxHeight = "150px";
        img.style.margin = "8px";
        img.style.border = "1px solid var(--border)";
        img.style.borderRadius = "8px";
        previewContainer.appendChild(img);
      });
    }

    if (els.btnOcrProcess) {
      els.btnOcrProcess.disabled = false;
    }

    setOcrStatus(`Loaded ${files.length} image(s). Click "Process Images with OCR" to begin.`, "ok");
  }

  /**
   * Handle OCR processing
   */
  async function handleOcrProcess() {
    if (!state.ocr.uploadedImages || state.ocr.uploadedImages.length === 0) {
      setOcrStatus("Please upload images first.", "error");
      return;
    }

    if (!window.Tesseract) {
      setOcrStatus("Tesseract.js is not loaded. Please refresh the page.", "error");
      return;
    }

    state.ocr.currentStep = 2;
    updateOcrUI();

    try {
      const results = await processMultipleImages(
        state.ocr.uploadedImages,
        (progress, message) => {
          updateOcrProgress(progress, message);
        }
      );

      state.ocr.processingResults = results;
      state.ocr.currentStep = 3;
      updateOcrUI();
      renderOcrResults(results);

      setOcrStatus(
        `Processing complete! Found ${results.confident.length} confident IDs and ${results.uncertain.length} uncertain IDs.`,
        "ok"
      );
    } catch (error) {
      console.error("OCR processing error:", error);
      setOcrStatus(`Error processing images: ${error.message}`, "error");
      state.ocr.currentStep = 1;
      updateOcrUI();
    }
  }

  /**
   * Handle OCR approve and generate text file
   */
  function handleOcrApprove() {
    if (!state.ocr.processingResults) {
      setOcrStatus("No results to approve. Please process images first.", "error");
      return;
    }

    // Collect all approved IDs (confident + manually approved uncertain)
    const approvedIds = [...state.ocr.processingResults.confident];
    
    // Get manually approved uncertain IDs from UI
    const uncertainList = els.ocrUncertainList;
    if (uncertainList) {
      const approvedCheckboxes = uncertainList.querySelectorAll('input[type="checkbox"]:checked');
      approvedCheckboxes.forEach(checkbox => {
        const index = parseInt(checkbox.dataset.index);
        if (state.ocr.processingResults.uncertain[index]) {
          approvedIds.push(state.ocr.processingResults.uncertain[index]);
        }
      });
    }

    if (approvedIds.length === 0) {
      setOcrStatus("No IDs selected for approval.", "error");
      return;
    }

    // Generate text file
    const textContent = generateTextFile(approvedIds);
    const blob = new Blob([textContent], { type: "text/plain" });
    const filename = `attendance_ids_${new Date().toISOString().split('T')[0]}.txt`;
    
    downloadBlob(blob, filename);

    setOcrStatus(`Generated text file with ${approvedIds.length} IDs. File downloaded!`, "ok");
  }

  /**
   * Handle OCR reset
   */
  function handleOcrReset() {
    state.ocr.uploadedImages = [];
    state.ocr.processingResults = null;
    state.ocr.approvedIds = [];
    state.ocr.currentStep = 1;

    if (els.ocrImageUpload) els.ocrImageUpload.value = "";
    if (els.ocrImagePreview) els.ocrImagePreview.innerHTML = "";
    if (els.btnOcrProcess) els.btnOcrProcess.disabled = true;

    updateOcrUI();
    setOcrStatus("Reset complete. You can upload new images.", "info");
  }

  /**
   * Handle uncertain ID edit
   */
  function handleOcrUncertainEdit(index, newId) {
    if (!state.ocr.processingResults || !state.ocr.processingResults.uncertain[index]) {
      return;
    }

    // Update the ID
    state.ocr.processingResults.uncertain[index].id = newId;
    state.ocr.processingResults.uncertain[index].confidence = 100; // Mark as manually edited

    // Re-render results
    renderOcrResults(state.ocr.processingResults);
  }

  /**
   * Update OCR UI based on current step
   */
  function updateOcrUI() {
    // Show/hide steps
    for (let i = 1; i <= 3; i++) {
      const stepEl = document.querySelector(`[data-ocr-step="${i}"]`);
      if (stepEl) {
        stepEl.style.display = state.ocr.currentStep === i ? "block" : "none";
      }
    }
  }

  /**
   * Update OCR progress
   */
  function updateOcrProgress(progress, message) {
    const progressFill = els.ocrProgressFill;
    const progressText = els.ocrProgressText;
    const processingLog = els.ocrProcessingLog;

    if (progressFill) {
      progressFill.style.width = `${Math.min(100, Math.max(0, progress))}%`;
    }
    if (progressText) {
      progressText.textContent = message || `Processing... ${Math.round(progress)}%`;
    }
    if (processingLog) {
      const logEntry = document.createElement("div");
      logEntry.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;
      logEntry.style.fontSize = "12px";
      logEntry.style.color = "var(--text-secondary)";
      logEntry.style.marginTop = "4px";
      processingLog.appendChild(logEntry);
      processingLog.scrollTop = processingLog.scrollHeight;
    }
  }

  /**
   * Render OCR results
   */
  function renderOcrResults(results) {
    // Update summary
    if (els.ocrTotalIds) {
      els.ocrTotalIds.textContent = results.confident.length + results.uncertain.length;
    }
    if (els.ocrConfidentIds) {
      els.ocrConfidentIds.textContent = results.confident.length;
    }
    if (els.ocrUncertainIds) {
      els.ocrUncertainIds.textContent = results.uncertain.length;
    }

    // Render confident IDs
    const confidentList = els.ocrConfidentList;
    if (confidentList) {
      confidentList.innerHTML = "";
      if (results.confident.length === 0) {
        confidentList.innerHTML = '<p style="color: var(--text-muted);">No confident matches found.</p>';
      } else {
        results.confident.forEach((result, index) => {
          const item = document.createElement("div");
          item.className = "ocrIdItem";
          item.innerHTML = `
            <span class="ocrIdItem__id">${result.id}</span>
            <span class="ocrIdItem__confidence">${Math.round(result.confidence)}%</span>
            <span class="ocrIdItem__source">${result.imageName}</span>
          `;
          confidentList.appendChild(item);
        });
      }
    }

    // Render uncertain IDs with edit capability
    const uncertainList = els.ocrUncertainList;
    if (uncertainList) {
      uncertainList.innerHTML = "";
      if (results.uncertain.length === 0) {
        uncertainList.innerHTML = '<p style="color: var(--text-muted);">No uncertain matches found.</p>';
      } else {
        results.uncertain.forEach((result, index) => {
          const item = document.createElement("div");
          item.className = "ocrIdItem ocrIdItem--uncertain";
          item.innerHTML = `
            <label class="ocrIdItem__checkbox">
              <input type="checkbox" data-index="${index}" />
              <span>Approve</span>
            </label>
            <input 
              type="text" 
              class="ocrIdItem__input" 
              value="${result.id}" 
              data-index="${index}"
              placeholder="Edit ID..."
            />
            <span class="ocrIdItem__confidence">${Math.round(result.confidence)}%</span>
            <span class="ocrIdItem__source">${result.imageName}</span>
            <button class="ocrIdItem__edit" data-index="${index}" type="button">Save</button>
          `;
          uncertainList.appendChild(item);

          // Wire up edit button
          const editBtn = item.querySelector('.ocrIdItem__edit');
          const input = item.querySelector('.ocrIdItem__input');
          if (editBtn && input) {
            editBtn.addEventListener('click', () => {
              handleOcrUncertainEdit(index, input.value);
            });
            input.addEventListener('keypress', (e) => {
              if (e.key === 'Enter') {
                handleOcrUncertainEdit(index, input.value);
              }
            });
          }
        });
      }
    }
  }

  /**
   * Set OCR status message
   */
  function setOcrStatus(msg, kind = "info") {
    const statusEl = els.ocrStatus;
    if (!statusEl) return;
    
    statusEl.textContent = msg || "";
    statusEl.classList.remove("is-error", "is-ok");
    if (kind === "error") statusEl.classList.add("is-error");
    if (kind === "ok") statusEl.classList.add("is-ok");
  }

  return {
    handleOcrImageUpload,
    handleOcrProcess,
    handleOcrApprove,
    handleOcrReset,
    updateOcrUI,
  };
}
