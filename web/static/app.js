(() => {
  const SUCCESS = new Set(["downloaded", "already_exists"]);
  const BLOCKED = new Set(["access_required"]);
  const NO_OA = new Set(["no_open_access"]);
  const ERROR = new Set(["failed", "invalid_doi"]);

  const STORAGE_JOB_KEY = "paperfetch_last_job";
  const STORAGE_EMAIL_KEY = "paperfetch_email";

  const els = {
    brandHome: document.getElementById("brand-home"),
    uploadView: document.getElementById("view-upload"),
    progressView: document.getElementById("view-progress"),
    resultsView: document.getElementById("view-results"),
    form: document.getElementById("upload-form"),
    fileInput: document.getElementById("file-input"),
    dropzone: document.getElementById("dropzone"),
    dropTitle: document.getElementById("dropzone-title"),
    dropSub: document.getElementById("dropzone-sub"),
    email: document.getElementById("email-input"),
    emailHint: document.getElementById("email-hint"),
    limit: document.getElementById("limit-input"),
    dryRun: document.getElementById("dry-run-input"),
    startBtn: document.getElementById("start-btn"),
    formError: document.getElementById("form-error"),
    btnCancel: document.getElementById("btn-cancel"),
    progressFile: document.getElementById("progress-file"),
    progressBar: document.getElementById("progress-bar"),
    progressLabel: document.getElementById("progress-label"),
    progressTiming: document.getElementById("progress-timing"),
    progressPercent: document.getElementById("progress-percent"),
    currentTitle: document.getElementById("current-title"),
    currentDoi: document.getElementById("current-doi"),
    liveKpis: document.getElementById("live-kpis"),
    liveBreakdown: document.getElementById("live-breakdown"),
    liveFeedCount: document.getElementById("live-feed-count"),
    liveFeedBody: document.getElementById("live-feed-body"),
    resultsEyebrow: document.getElementById("results-eyebrow"),
    resultsSubtitle: document.getElementById("results-subtitle"),
    resultsKpis: document.getElementById("results-kpis"),
    resultsBreakdown: document.getElementById("results-breakdown"),
    resultsBody: document.getElementById("results-body"),
    emptyFilter: document.getElementById("empty-filter"),
    search: document.getElementById("search-input"),
    pageSize: document.getElementById("page-size"),
    pagination: document.getElementById("pagination"),
    pageInfo: document.getElementById("page-info"),
    pagePrev: document.getElementById("page-prev"),
    pageNext: document.getElementById("page-next"),
    btnXlsx: document.getElementById("btn-xlsx"),
    btnZip: document.getElementById("btn-zip"),
    btnReport: document.getElementById("btn-report"),
    btnNew: document.getElementById("btn-new"),
    chips: [...document.querySelectorAll(".chip")],
  };

  let pollTimer = null;
  let searchTimer = null;
  let currentJob = null;
  let activeFilter = "all";
  let currentPage = 1;
  let selectedFile = null;
  let jobStartedMs = null;

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;");
  }

  function showView(name) {
    els.uploadView.classList.toggle("is-active", name === "upload");
    els.progressView.classList.toggle("is-active", name === "progress");
    els.resultsView.classList.toggle("is-active", name === "results");
  }

  function setJobUrl(jobId) {
    const url = new URL(window.location.href);
    if (jobId) {
      url.searchParams.set("job", jobId);
      localStorage.setItem(STORAGE_JOB_KEY, jobId);
    } else {
      url.searchParams.delete("job");
      localStorage.removeItem(STORAGE_JOB_KEY);
    }
    window.history.replaceState({}, "", url.toString());
  }

  function setFile(file) {
    selectedFile = file || null;
    if (!selectedFile) {
      els.dropzone.classList.remove("has-file");
      els.dropTitle.textContent = "Drop your Excel or CSV file here";
      els.dropSub.textContent = "or click to browse · .xlsx, .xlsm, .xls, .csv";
      return;
    }
    els.dropzone.classList.add("has-file");
    els.dropTitle.textContent = selectedFile.name;
    const kb = Math.max(1, Math.round(selectedFile.size / 1024));
    els.dropSub.textContent = `${kb.toLocaleString()} KB ready to upload`;
  }

  function showFormError(message) {
    if (!message) {
      els.formError.hidden = true;
      els.formError.textContent = "";
      return;
    }
    els.formError.hidden = false;
    els.formError.textContent = message;
  }

  function updateEmailHint() {
    const val = (els.email.value || "").trim().toLowerCase();
    if (val.endsWith("@example.com") || val.endsWith("@example.org")) {
      els.emailHint.textContent = "Note: Unpaywall rejects @example.com emails — use a real email address for best results.";
      els.emailHint.classList.add("is-warn");
    } else {
      els.emailHint.textContent = "Required by OpenAlex / Unpaywall polite pool (avoid @example.com)";
      els.emailHint.classList.remove("is-warn");
    }
  }

  function computeStats(summary = {}, results = []) {
    const newlyDownloaded = Number(summary.downloaded ?? 0);
    const alreadyExists = Number(summary.already_exists ?? 0);
    const downloadedTotal = newlyDownloaded + alreadyExists;
    const noOa = Number(summary.no_oa ?? 0);
    const accessRequired = Number(summary.access_required ?? 0);
    const invalidDoi = Number(summary.invalid_doi ?? 0);
    const failed = Number(summary.failed ?? 0);
    const notFound =
      summary.not_found !== undefined
        ? Number(summary.not_found)
        : noOa + accessRequired + invalidDoi + failed;
    const total = Number(summary.total ?? results.length);
    return {
      total,
      downloadedTotal,
      newlyDownloaded,
      alreadyExists,
      notFound,
      noOa,
      accessRequired,
      invalidDoi,
      failed,
    };
  }

  function renderKpiHtml(stats, isDryRun) {
    const dlLabel = isDryRun ? "OA Available" : "Downloaded";
    const dlSub = isDryRun
      ? "Dry run (PDF links found)"
      : `${stats.newlyDownloaded} new · ${stats.alreadyExists} existing`;
    return `
      <div class="kpi-card">
        <span class="kpi-label">Total Papers</span>
        <strong class="kpi-value">${stats.total.toLocaleString()}</strong>
        <span class="kpi-sub">In this run</span>
      </div>
      <div class="kpi-card kpi-ok">
        <span class="kpi-label">${dlLabel}</span>
        <strong class="kpi-value">${stats.downloadedTotal.toLocaleString()}</strong>
        <span class="kpi-sub">${dlSub}</span>
      </div>
      <div class="kpi-card kpi-warn">
        <span class="kpi-label">Not Found / Unavailable</span>
        <strong class="kpi-value">${stats.notFound.toLocaleString()}</strong>
        <span class="kpi-sub">Reason included in Excel</span>
      </div>
    `;
  }

  function renderBreakdownHtml(stats) {
    return `
      <span class="mini-pill">No Open Access: <strong>${stats.noOa}</strong></span>
      <span class="mini-pill">Blocked / Non-PDF: <strong>${stats.accessRequired}</strong></span>
      <span class="mini-pill">Missing / Invalid DOI: <strong>${stats.invalidDoi}</strong></span>
      <span class="mini-pill">Errors: <strong>${stats.failed}</strong></span>
    `;
  }

  function statusBadge(status) {
    const labelMap = {
      downloaded: "Downloaded",
      already_exists: "Already existed",
      no_open_access: "No Open Access",
      access_required: "Blocked / Non-PDF",
      invalid_doi: "Missing DOI",
      failed: "Error",
    };
    const label = labelMap[status] || (status || "unknown").replaceAll("_", " ");
    let cls = "badge-muted";
    if (SUCCESS.has(status)) cls = "badge-ok";
    else if (NO_OA.has(status) || BLOCKED.has(status)) cls = "badge-warn";
    else if (ERROR.has(status)) cls = "badge-err";
    return `<span class="badge ${cls}">${escapeHtml(label)}</span>`;
  }

  function formatReason(row) {
    if (SUCCESS.has(row.status)) {
      if (row.error === "dry-run") {
        return row.oa_source ? `OA found (${row.oa_source})` : "OA URL found (dry run)";
      }
      return row.filename || "Saved";
    }
    return row.error || "Not available";
  }

  function formatDuration(ms) {
    const totalSec = Math.max(0, Math.round(ms / 1000));
    const m = Math.floor(totalSec / 60);
    const s = totalSec % 60;
    return m > 0 ? `${m}m ${s}s` : `${s}s`;
  }

  function updateProgress(job) {
    const progress = job.progress || {};
    const percent = Math.min(100, Number(progress.percent || 0));
    const of = Number(progress.of || 0);
    const index = Number(progress.index || 0);

    els.progressBar.style.width = `${percent}%`;
    els.progressBar.parentElement.setAttribute("aria-valuenow", String(percent));
    els.progressPercent.textContent = `${Math.round(percent)}%`;
    els.progressFile.textContent = `${job.filename}${job.dry_run ? " · dry run" : ""}`;

    if (job.current) {
      els.progressLabel.textContent = `Paper ${job.current.index} of ${job.current.of}`;
      els.currentTitle.textContent = job.current.title || "Untitled paper";
      els.currentDoi.textContent = job.current.doi || "Searching by title…";
    } else if (job.status === "queued") {
      els.progressLabel.textContent = "Queued — starting soon";
      els.currentTitle.textContent = "Preparing job…";
      els.currentDoi.textContent = "";
    } else {
      els.progressLabel.textContent = of > 0 ? `Processed ${index} of ${of}` : "Working…";
    }

    if (jobStartedMs && index > 0 && of > index) {
      const elapsed = Date.now() - jobStartedMs;
      const perPaper = elapsed / index;
      const eta = perPaper * (of - index);
      els.progressTiming.textContent = `Elapsed ${formatDuration(elapsed)} · ETA ~${formatDuration(eta)}`;
    } else if (jobStartedMs && index > 0) {
      els.progressTiming.textContent = `Elapsed ${formatDuration(Date.now() - jobStartedMs)}`;
    } else {
      els.progressTiming.textContent = "";
    }

    els.btnCancel.disabled = Boolean(job.cancel_requested);
    els.btnCancel.textContent = job.cancel_requested ? "Stopping…" : "Stop & Keep Results";

    const stats = computeStats(job.summary || {}, job.results || []);
    els.liveKpis.innerHTML = renderKpiHtml(stats, Boolean(job.dry_run));
    els.liveBreakdown.innerHTML = renderBreakdownHtml(stats);

    const results = job.results || [];
    els.liveFeedCount.textContent = `${results.length} processed`;
    const recent = results.slice(-10).reverse();
    els.liveFeedBody.innerHTML = recent
      .map(
        (row) => `
        <tr>
          <td>${escapeHtml(row.paper_number)}</td>
          <td class="title-cell">${escapeHtml(row.title || "—")}</td>
          <td>${statusBadge(row.status)}</td>
          <td class="error-cell">${escapeHtml(formatReason(row))}</td>
        </tr>`
      )
      .join("");
  }

  function updateFilterCounts(results) {
    const counts = {
      all: results.length,
      success: 0,
      not_found: 0,
      no_oa: 0,
      blocked: 0,
      invalid_doi: 0,
    };
    for (const row of results) {
      const s = row.status || "";
      if (SUCCESS.has(s)) {
        counts.success += 1;
      } else {
        counts.not_found += 1;
        if (NO_OA.has(s)) counts.no_oa += 1;
        else if (BLOCKED.has(s)) counts.blocked += 1;
        else if (ERROR.has(s)) counts.invalid_doi += 1;
      }
    }
    els.chips.forEach((chip) => {
      const key = chip.dataset.filter || "all";
      const span = chip.querySelector(".chip-count");
      if (span) span.textContent = String(counts[key] ?? 0);
    });
  }

  function matchesFilter(row) {
    const status = row.status || "";
    if (activeFilter === "success") return SUCCESS.has(status);
    if (activeFilter === "not_found") return !SUCCESS.has(status);
    if (activeFilter === "no_oa") return NO_OA.has(status);
    if (activeFilter === "blocked") return BLOCKED.has(status);
    if (activeFilter === "invalid_doi") return ERROR.has(status);
    return true;
  }

  function matchesSearch(row, query) {
    if (!query) return true;
    const hay = [
      row.paper_number,
      row.title,
      row.doi,
      row.publisher,
      row.status,
      row.error,
      row.destination_folder,
      row.filename,
    ]
      .join(" ")
      .toLowerCase();
    return hay.includes(query);
  }

  function renderResults(job) {
    const results = job.results || [];
    const stats = computeStats(job.summary || {}, results);
    const query = (els.search.value || "").trim().toLowerCase();
    const filtered = results.filter(
      (row) => matchesFilter(row) && matchesSearch(row, query)
    );

    const perPage = Math.max(10, Number(els.pageSize.value || 50));
    const totalPages = Math.max(1, Math.ceil(filtered.length / perPage));
    if (currentPage > totalPages) currentPage = totalPages;
    const start = (currentPage - 1) * perPage;
    const pageRows = filtered.slice(start, start + perPage);

    els.resultsEyebrow.textContent =
      job.status === "cancelled" ? "Stopped early (partial results saved)" : "Run complete";
    els.resultsSubtitle.textContent = `${job.filename} · ${results.length} processed · job ${job.id}`;

    els.resultsKpis.innerHTML = renderKpiHtml(stats, Boolean(job.dry_run));
    els.resultsBreakdown.innerHTML = renderBreakdownHtml(stats);
    updateFilterCounts(results);

    els.btnXlsx.href = `/api/jobs/${job.id}/report.xlsx`;
    els.btnReport.href = `/api/jobs/${job.id}/report.csv`;
    els.btnZip.href = `/api/jobs/${job.id}/papers.zip`;
    const hasSavedPdfs = !job.dry_run && stats.downloadedTotal > 0;
    els.btnZip.hidden = !hasSavedPdfs;

    const html = pageRows
      .map((row) => {
        const doiText = row.doi || "";
        const doiCell = doiText.startsWith("10.")
          ? `<a href="https://doi.org/${encodeURIComponent(doiText)}" target="_blank" rel="noopener">${escapeHtml(doiText)}</a>`
          : escapeHtml(doiText || "—");

        let reasonHtml = escapeHtml(formatReason(row));
        if (!SUCCESS.has(row.status) && row.pdf_url && row.pdf_url.startsWith("http")) {
          reasonHtml += ` <a href="${escapeHtml(row.pdf_url)}" target="_blank" rel="noopener">[Open link]</a>`;
        }

        return `
          <tr>
            <td>${escapeHtml(row.paper_number)}</td>
            <td class="title-cell">
              ${escapeHtml(row.title || "—")}
              ${row.publisher ? `<span class="pub-sub">${escapeHtml(row.publisher)}</span>` : ""}
            </td>
            <td class="doi-cell">${doiCell}</td>
            <td>${statusBadge(row.status)}</td>
            <td>${escapeHtml(row.destination_folder || "—")}</td>
            <td class="error-cell">${reasonHtml}</td>
          </tr>`;
      })
      .join("");

    els.resultsBody.innerHTML = html;
    els.emptyFilter.hidden = filtered.length > 0;

    if (filtered.length > perPage) {
      els.pagination.hidden = false;
      els.pageInfo.textContent = `Showing ${start + 1}–${Math.min(start + perPage, filtered.length)} of ${filtered.length} papers (page ${currentPage} of ${totalPages})`;
      els.pagePrev.disabled = currentPage <= 1;
      els.pageNext.disabled = currentPage >= totalPages;
    } else {
      els.pagination.hidden = true;
    }
  }

  async function pollJob(jobId) {
    try {
      const since = currentJob && currentJob.id === jobId ? (currentJob.results || []).length : 0;
      const res = await fetch(`/api/jobs/${jobId}?since=${since}`);
      if (!res.ok) throw new Error("Could not load job status.");
      const patch = await res.json();

      if (currentJob && currentJob.id === jobId && patch.results_offset > 0) {
        const merged = (currentJob.results || []).slice(0, patch.results_offset).concat(patch.results || []);
        patch.results = merged;
      }
      currentJob = patch;

      if (patch.status === "completed" || patch.status === "cancelled") {
        clearInterval(pollTimer);
        pollTimer = null;
        els.startBtn.disabled = false;
        els.startBtn.textContent = "Run & Download Papers";
        renderResults(patch);
        showView("results");
        return;
      }

      if (patch.status === "failed") {
        clearInterval(pollTimer);
        pollTimer = null;
        setJobUrl(null);
        showView("upload");
        showFormError(patch.error || "Download job failed.");
        els.startBtn.disabled = false;
        els.startBtn.textContent = "Run & Download Papers";
        return;
      }

      showView("progress");
      updateProgress(patch);
    } catch (err) {
      clearInterval(pollTimer);
      pollTimer = null;
      showView("upload");
      showFormError(err.message || "Lost connection to the server.");
      els.startBtn.disabled = false;
      els.startBtn.textContent = "Run & Download Papers";
    }
  }

  function startPolling(jobId) {
    if (pollTimer) clearInterval(pollTimer);
    pollJob(jobId);
    pollTimer = setInterval(() => pollJob(jobId), 1200);
  }

  function resetToUpload() {
    if (pollTimer) clearInterval(pollTimer);
    pollTimer = null;
    currentJob = null;
    activeFilter = "all";
    currentPage = 1;
    jobStartedMs = null;
    setJobUrl(null);
    els.chips.forEach((c) => c.classList.toggle("is-active", c.dataset.filter === "all"));
    els.search.value = "";
    els.form.reset();
    const savedEmail = localStorage.getItem(STORAGE_EMAIL_KEY);
    if (savedEmail) els.email.value = savedEmail;
    updateEmailHint();
    setFile(null);
    els.startBtn.disabled = false;
    els.startBtn.textContent = "Run & Download Papers";
    showFormError("");
    showView("upload");
  }

  els.fileInput.addEventListener("change", () => {
    setFile(els.fileInput.files?.[0] || null);
    showFormError("");
  });

  els.email.addEventListener("input", updateEmailHint);

  ["dragenter", "dragover"].forEach((eventName) => {
    els.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      els.dropzone.classList.add("is-dragover");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    els.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      els.dropzone.classList.remove("is-dragover");
    });
  });

  els.dropzone.addEventListener("drop", (event) => {
    const file = event.dataTransfer?.files?.[0];
    if (!file) return;
    const transfer = new DataTransfer();
    transfer.items.add(file);
    els.fileInput.files = transfer.files;
    setFile(file);
    showFormError("");
  });

  els.form.addEventListener("submit", async (event) => {
    event.preventDefault();
    showFormError("");

    const file = selectedFile || els.fileInput.files?.[0];
    if (!file) {
      showFormError("Please choose an Excel (.xlsx/.xls) or CSV file first.");
      return;
    }

    const email = els.email.value.trim();
    if (!email || !email.includes("@") || !email.split("@")[1]?.includes(".")) {
      showFormError("Enter a valid contact email (e.g. you@gmail.com) for OpenAlex / Unpaywall.");
      return;
    }
    localStorage.setItem(STORAGE_EMAIL_KEY, email);

    const body = new FormData();
    body.append("file", file);
    body.append("email", email);
    body.append("dry_run", els.dryRun.checked ? "true" : "false");
    body.append("limit", String(Math.max(0, Number(els.limit.value || 0))));

    els.startBtn.disabled = true;
    els.startBtn.textContent = "Uploading…";
    jobStartedMs = Date.now();
    showView("progress");
    updateProgress({
      filename: file.name,
      dry_run: els.dryRun.checked,
      progress: { percent: 0, index: 0, of: 0 },
      summary: {},
      results: [],
      status: "queued",
      current: null,
    });

    try {
      const res = await fetch("/api/jobs", { method: "POST", body });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) {
        const detail = data.detail;
        const message = Array.isArray(detail)
          ? detail.map((item) => item.msg || JSON.stringify(item)).join("; ")
          : detail || "Upload failed.";
        throw new Error(message);
      }
      currentJob = data;
      setJobUrl(data.id);
      els.startBtn.textContent = "Run & Download Papers";
      startPolling(data.id);
    } catch (err) {
      showView("upload");
      showFormError(err.message || "Could not start the job.");
      els.startBtn.disabled = false;
      els.startBtn.textContent = "Run & Download Papers";
    }
  });

  els.btnCancel.addEventListener("click", async () => {
    if (!currentJob) return;
    els.btnCancel.disabled = true;
    els.btnCancel.textContent = "Stopping…";
    try {
      await fetch(`/api/jobs/${currentJob.id}/cancel`, { method: "POST" });
    } catch {
      // Next poll picks up state.
    }
  });

  els.chips.forEach((chip) => {
    chip.addEventListener("click", () => {
      activeFilter = chip.dataset.filter || "all";
      currentPage = 1;
      els.chips.forEach((c) => c.classList.toggle("is-active", c === chip));
      if (currentJob) renderResults(currentJob);
    });
  });

  els.search.addEventListener("input", () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      currentPage = 1;
      if (currentJob) renderResults(currentJob);
    }, 120);
  });

  els.pageSize.addEventListener("change", () => {
    currentPage = 1;
    if (currentJob) renderResults(currentJob);
  });

  els.pagePrev.addEventListener("click", () => {
    if (currentPage > 1) {
      currentPage -= 1;
      if (currentJob) renderResults(currentJob);
    }
  });

  els.pageNext.addEventListener("click", () => {
    currentPage += 1;
    if (currentJob) renderResults(currentJob);
  });

  els.btnNew.addEventListener("click", resetToUpload);
  els.brandHome.addEventListener("click", (event) => {
    event.preventDefault();
    resetToUpload();
  });

  // Restore saved email and reconnect to ?job=<id> if present on reload.
  const savedEmail = localStorage.getItem(STORAGE_EMAIL_KEY);
  if (savedEmail) els.email.value = savedEmail;
  updateEmailHint();

  const urlJobId = new URL(window.location.href).searchParams.get("job");
  if (urlJobId) {
    jobStartedMs = Date.now();
    startPolling(urlJobId);
  }
})();
