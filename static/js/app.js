/**
 * app.js — State machine and UI logic for Traffic Notice Generator.
 * Plain vanilla JS; no external frameworks or libraries.
 */

(function () {
  "use strict";

  // Application State
  const state = {
    selectedFile: null,
    jobId: null,
    reportData: null,
    statusPollTimer: null,
    currentState: 1, // 1 to 9
  };

  // DOM Elements
  const el = {
    dropZone: document.getElementById("drop-zone"),
    fileInput: document.getElementById("file-input"),
    fileStrip: document.getElementById("file-strip"),
    fileName: document.getElementById("file-name"),
    fileSize: document.getElementById("file-size"),
    btnInspect: document.getElementById("btn-inspect"),
    btnRemove: document.getElementById("btn-remove"),
    stepItems: document.querySelectorAll(".step-item"),
    panels: {
      upload: document.getElementById("panel-upload"),
      checking: document.getElementById("panel-checking"),
      report: document.getElementById("panel-report"),
      generating: document.getElementById("panel-generating"),
      done: document.getElementById("panel-done"),
      failed: document.getElementById("panel-failed"),
      error: document.getElementById("panel-error"),
    },
    reportContent: document.getElementById("report-content"),
    btnStart: document.getElementById("btn-start-generation"),
    reviewWarningsCheckbox: document.getElementById("review-warnings-checkbox"),
    genMessage: document.getElementById("gen-message"),
    genProgressFill: document.getElementById("gen-progress-fill"),
    genMeta: document.getElementById("gen-meta"),
    btnCancelGen: document.getElementById("btn-cancel-generation"),
    doneSummary: document.getElementById("done-summary"),
    btnDownloadMain: document.getElementById("btn-download-main"),
    previewContainer: document.getElementById("preview-container"),
    batchDownloadsList: document.getElementById("batch-downloads-list"),
    failedErrorsList: document.getElementById("failed-errors-list"),
    errorMessage: document.getElementById("error-panel-message"),
    liveAnnouncer: document.getElementById("aria-live-announcer"),
  };

  function announce(text) {
    if (el.liveAnnouncer) {
      el.liveAnnouncer.textContent = text;
    }
  }

  function setStepActive(stepNum) {
    el.stepItems.forEach((item) => {
      const num = parseInt(item.dataset.step, 10);
      item.classList.remove("active", "completed");
      if (num < stepNum) item.classList.add("completed");
      if (num === stepNum) item.classList.add("active");
    });
  }

  function showPanel(panelKey) {
    Object.values(el.panels).forEach((p) => {
      if (p) p.classList.remove("active");
    });
    if (el.panels[panelKey]) {
      el.panels[panelKey].classList.add("active");
    }
  }

  function formatBytes(bytes) {
    if (bytes === 0) return "0 Bytes";
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
  }

  // --- STATE 1 & 2: FILE SELECTION & DROP ---

  function handleFileSelected(file) {
    if (!file) return;
    const name = file.name.toLowerCase();
    if (!name.endsWith(".xlsx") && !name.endsWith(".csv")) {
      if (name.endsWith(".xls")) {
        showError(STRINGS.errors.wrongType);
      } else {
        showError(STRINGS.errors.wrongType);
      }
      return;
    }

    state.selectedFile = file;
    el.fileName.textContent = file.name;
    el.fileSize.textContent = formatBytes(file.size);
    el.fileStrip.style.display = "flex";
    el.dropZone.style.display = "none";
    el.btnInspect.style.display = "inline-flex";
    setStepActive(1);
    announce(`File ${file.name} selected. Ready to inspect.`);
  }

  el.dropZone.addEventListener("click", () => el.fileInput.click());
  el.dropZone.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      el.fileInput.click();
    }
  });

  el.fileInput.addEventListener("change", (e) => {
    if (e.target.files && e.target.files[0]) {
      handleFileSelected(e.target.files[0]);
    }
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    el.dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      el.dropZone.classList.add("dragover");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    el.dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      el.dropZone.classList.remove("dragover");
    });
  });

  el.dropZone.addEventListener("drop", (e) => {
    if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFileSelected(e.dataTransfer.files[0]);
    }
  });

  el.btnRemove.addEventListener("click", () => {
    state.selectedFile = null;
    state.jobId = null;
    state.reportData = null;
    el.fileInput.value = "";
    el.fileStrip.style.display = "none";
    el.dropZone.style.display = "block";
    el.btnInspect.style.display = "none";
    showPanel("upload");
    setStepActive(1);
  });

  // --- STATE 3 & 4: UPLOAD & PREFLIGHT CHECK ---

  el.btnInspect.addEventListener("click", () => {
    if (!state.selectedFile) return;

    showPanel("checking");
    setStepActive(2);
    announce(STRINGS.state3_checking.desc);

    const formData = new FormData();
    formData.append("file", state.selectedFile);

    fetch("/api/upload", {
      method: "POST",
      body: formData,
    })
      .then((res) => {
        if (!res.ok) {
          return res.json().then((d) => {
            throw new Error(d.message || d.error || "Inspection failed.");
          });
        }
        return res.json();
      })
      .then((data) => {
        state.jobId = data.job_id;
        state.reportData = data;
        renderReport(data);
      })
      .catch((err) => {
        showError(err.message || STRINGS.errors.general);
      });
  });

  function renderReport(data) {
    showPanel("report");
    setStepActive(2);

    const isOdd = data.total_records % 2 !== 0;
    const batches = Math.ceil(data.total_records / 1000);

    let html = "";

    // Headline summary
    if (data.file_type === "xlsx") {
      html += `<p class="outcome-sentence"><strong>${STRINGS.state4_report.foundNotice(
        data.total_records,
        data.sheet_name,
        data.start_row,
        data.start_row + data.total_records - 1
      )}</strong></p>`;
    } else {
      html += `<p class="outcome-sentence"><strong>${STRINGS.state4_report.foundNoticeCsv(
        data.total_records
      )}</strong></p>`;
    }

    // Outcome sentence
    let outcomeText =
      batches === 1
        ? STRINGS.state4_report.outcomeSingle
        : STRINGS.state4_report.outcomeMulti(batches, 1000);
    if (isOdd) {
      outcomeText += ` ${STRINGS.state4_report.outcomeOdd}`;
    }
    html += `<div class="banner banner-caution" style="margin-bottom: var(--space-4);">${outcomeText}</div>`;

    // Blocking Errors
    if (data.blocking_count > 0) {
      html += `<div class="banner banner-error">
        <strong>${STRINGS.state4_report.blockingTitle}</strong>: ${data.blocking_count} error(s) found.<br>
        <span style="font-size: var(--font-size-xs);">${STRINGS.state4_report.blockingHelp}</span>
      </div>`;
      html += `<ul class="flagged-list">`;
      data.blocking_errors.forEach((e) => {
        html += `<li class="flagged-item">Row ${e.row} [${e.field}]: ${e.reason} ${e.value ? "— " + e.value : ""}</li>`;
      });
      html += `</ul>`;
    }

    // Warnings
    if (data.warnings_count > 0) {
      html += `<div class="banner banner-caution">
        <strong>${STRINGS.state4_report.warningsTitle}</strong>: ${data.warnings_count} item(s) flagged for review.
      </div>`;
      html += `<ul class="flagged-list">`;
      data.warnings.forEach((w) => {
        html += `<li class="flagged-item">Row ${w.row} [${w.field}]: ${w.reason} ${w.value ? "— " + w.value : ""}</li>`;
      });
      html += `</ul>`;
    }

    // Download Flagged Rows link
    if (data.has_problems_file) {
      html += `<p style="margin-bottom: var(--space-4); font-size: var(--font-size-sm);">
        <a href="/api/jobs/${data.job_id}/problems" target="_blank">${STRINGS.state4_report.downloadProblems}</a>
      </p>`;
    }

    // Record Preview Table
    html += `<h3 style="font-size: var(--font-size-sm); margin-bottom: var(--space-2);">${STRINGS.state4_report.previewTitle}</h3>`;
    html += `<div class="table-wrap"><table class="mono-table"><thead><tr>
      <th>Serial</th><th>Challan</th><th>Vehicle</th><th>Date</th><th>Name</th><th>Fine</th><th>Section</th>
    </tr></thead><tbody>`;

    const sample = [...data.first_5];
    if (data.total_records > 5) {
      sample.push(...data.last_5);
    }
    sample.forEach((r) => {
      html += `<tr>
        <td>${r.SERIAL_NUMBER || ""}</td>
        <td>${r.E_CHALLAN_NUMBER || ""}</td>
        <td>${r.VEHICLE_NUMBER || ""}</td>
        <td>${r.VIOLATION_DATE || ""}</td>
        <td>${r.NAME || ""}</td>
        <td>${r.FINE_AMOUNT || ""}</td>
        <td>${r.SECTION || ""}</td>
      </tr>`;
    });
    html += `</tbody></table></div>`;

    el.reportContent.innerHTML = html;

    // Enable / disable start button
    if (data.blocking_count > 0) {
      el.btnStart.disabled = true;
      el.reviewWarningsCheckbox.parentElement.style.display = "none";
    } else if (data.warnings_count > 0) {
      el.reviewWarningsCheckbox.parentElement.style.display = "flex";
      el.reviewWarningsCheckbox.checked = false;
      el.btnStart.disabled = true;
      el.reviewWarningsCheckbox.onchange = (e) => {
        el.btnStart.disabled = !e.target.checked;
      };
    } else {
      el.reviewWarningsCheckbox.parentElement.style.display = "none";
      el.btnStart.disabled = false;
    }

    announce("Preflight check report loaded.");
  }

  // --- STATE 5 & 6: GENERATING & VERIFYING ---

  el.btnStart.addEventListener("click", () => {
    if (!state.jobId) return;

    fetch(`/api/jobs/${state.jobId}/start`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({}),
    })
      .then((res) => {
        if (!res.ok) {
          return res.json().then((d) => {
            throw new Error(d.message || "Could not start generation.");
          });
        }
        return res.json();
      })
      .then(() => {
        showPanel("generating");
        setStepActive(3);
        startPollingStatus();
      })
      .catch((err) => {
        showError(err.message || STRINGS.errors.general);
      });
  });

  el.btnCancelGen.addEventListener("click", () => {
    if (!state.jobId) return;
    if (!confirm("Are you sure you want to cancel the generation job?")) return;

    fetch(`/api/jobs/${state.jobId}/cancel`, { method: "POST" })
      .then(() => {
        stopPolling();
        showError(STRINGS.errors.cancelled);
      })
      .catch(() => stopPolling());
  });

  function startPollingStatus() {
    stopPolling();
    state.statusPollTimer = setInterval(pollJobStatus, 2000);
    pollJobStatus(); // Immediate first poll
  }

  function stopPolling() {
    if (state.statusPollTimer) {
      clearInterval(state.statusPollTimer);
      state.statusPollTimer = null;
    }
  }

  function pollJobStatus() {
    if (!state.jobId) return;

    fetch(`/api/jobs/${state.jobId}`)
      .then((res) => {
        if (res.status === 401) {
          stopPolling();
          showError(STRINGS.errors.sessionExpired);
          return null;
        }
        return res.json();
      })
      .then((data) => {
        if (!data) return;

        if (data.state === "generating") {
          setStepActive(3);
          el.genMessage.textContent = data.message || STRINGS.state5_generating.singleBatchNotice;

          if (data.total_batches > 1) {
            const pct = Math.round((data.completed_batches / data.total_batches) * 100);
            el.genProgressFill.style.width = `${pct}%`;
            el.genMeta.textContent = `${STRINGS.state5_generating.batchProgress(
              data.completed_batches,
              data.total_batches
            )} | Elapsed: ${data.elapsed_seconds}s`;
          } else {
            // Single batch honest message: no percentage until batch finishes
            el.genProgressFill.style.width = "100%";
            el.genProgressFill.style.opacity = "0.5";
            el.genMeta.textContent = `Elapsed: ${data.elapsed_seconds}s`;
          }
        } else if (data.state === "verifying") {
          setStepActive(3);
          el.genMessage.textContent = STRINGS.state6_verifying.desc;
          el.genProgressFill.style.width = "90%";
          el.genMeta.textContent = "Verifying word coordinates and layout...";
        } else if (data.state === "completed") {
          stopPolling();
          renderDone(data);
        } else if (data.state === "verification_failed") {
          stopPolling();
          renderFailed(data);
        } else if (data.state === "failed") {
          stopPolling();
          showError(data.message || data.error || STRINGS.errors.general);
        } else if (data.state === "cancelled") {
          stopPolling();
          showError(STRINGS.errors.cancelled);
        }
      })
      .catch((err) => {
        console.warn("Polling error:", err);
      });
  }

  // --- STATE 7: DONE ---

  function renderDone(data) {
    showPanel("done");
    setStepActive(4);

    el.doneSummary.textContent = STRINGS.state7_done.verifiedSuccess(data.total_challans);
    el.btnDownloadMain.href = `/api/jobs/${state.jobId}/download`;

    if (data.download_type === "zip") {
      el.btnDownloadMain.textContent = STRINGS.state7_done.btnDownloadZip;
    } else {
      el.btnDownloadMain.textContent = STRINGS.state7_done.btnDownloadPdf;
    }

    // Render preview image
    el.previewContainer.innerHTML = `<img src="/api/jobs/${state.jobId}/preview" alt="Notice Page 1 Preview" class="preview-img">`;

    // Individual batch list
    if (data.batches && data.batches.length > 1) {
      let bHtml = `<p style="font-size: var(--font-size-xs); font-weight:600; margin-bottom: var(--space-2);">${STRINGS.state7_done.individualBatches}:</p><ul style="list-style:none;">`;
      data.batches.forEach((b) => {
        if (b.ok && b.pdf_name) {
          bHtml += `<li style="font-size: var(--font-size-xs); font-family: var(--font-mono); margin-bottom: 4px;">
            <a href="/api/jobs/${state.jobId}/download/${encodeURIComponent(b.pdf_name)}">${b.pdf_name}</a> (${b.challans} notices)
          </li>`;
        }
      });
      bHtml += `</ul>`;
      el.batchDownloadsList.innerHTML = bHtml;
    } else {
      el.batchDownloadsList.innerHTML = "";
    }

    announce("Notices generated and verified. Ready for download.");
  }

  // --- STATE 8: VERIFICATION FAILED ---

  function renderFailed(data) {
    showPanel("failed");
    setStepActive(4);

    let errHtml = "";
    const errors = (data.verification_result && data.verification_result.errors) || [data.message];
    errors.forEach((err) => {
      errHtml += `<li class="flagged-item" style="color: var(--status-error);">${err}</li>`;
    });
    el.failedErrorsList.innerHTML = errHtml;
    announce("Verification failed. Download disabled.");
  }

  // --- STATE 9: ERRORS ---

  function showError(msg) {
    el.errorMessage.textContent = msg;
    showPanel("error");
    announce(`Error: ${msg}`);
  }

  document.querySelectorAll(".btn-restart").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.selectedFile = null;
      state.jobId = null;
      state.reportData = null;
      el.fileInput.value = "";
      el.fileStrip.style.display = "none";
      el.dropZone.style.display = "block";
      el.btnInspect.style.display = "none";
      showPanel("upload");
      setStepActive(1);
    });
  });

})();
