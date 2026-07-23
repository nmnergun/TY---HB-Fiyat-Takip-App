const $ = (id) => document.getElementById(id);

const views = {
  login: $("loginView"),
  dashboard: $("dashboard"),
};

let currentJobId = null;
let pollTimer = null;

const statusLabels = {
  queued: "Sırada",
  running: "Çalışıyor",
  completed: "Tamamlandı",
  failed: "Hata",
  cancelled: "İptal edildi",
};

async function api(path, options = {}) {
  const response = await fetch(path, {
    credentials: "same-origin",
    headers: { "Content-Type": "application/json", ...(options.headers || {}) },
    ...options,
  });
  let payload = {};
  try { payload = await response.json(); } catch (_) {}
  if (!response.ok) {
    const error = new Error(payload.detail || "İşlem tamamlanamadı.");
    error.status = response.status;
    throw error;
  }
  return payload;
}

function showToast(message) {
  const toast = $("toast");
  toast.textContent = message;
  toast.classList.remove("hidden");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => toast.classList.add("hidden"), 3500);
}

function showDashboard() {
  views.login.classList.add("hidden");
  views.dashboard.classList.remove("hidden");
}

function showLogin() {
  views.dashboard.classList.add("hidden");
  views.login.classList.remove("hidden");
  $("password").focus();
}

function renderJob(job) {
  const running = job && ["queued", "running"].includes(job.status);
  currentJobId = job?.id || null;
  $("startButton").disabled = running;
  $("startButton").innerHTML = running
    ? '<span class="button-icon">●</span> Rapor hazırlanıyor'
    : '<span class="button-icon">▶</span> Raporu başlat';
  $("cancelButton").classList.toggle("hidden", !running);

  if (!job) return;
  const progress = Number(job.progress || 0);
  $("runTitle").textContent = job.current_model
    ? `${job.current_model} kontrol ediliyor`
    : "Hepsiburada fiyat taraması";
  $("statusBadge").textContent = statusLabels[job.status] || job.status;
  $("statusBadge").className = `status-badge ${job.status}`;
  $("progressBar").style.width = `${progress}%`;
  $("progressText").textContent = `${progress.toFixed(progress % 1 ? 1 : 0)}%`;
  $("runMessage").textContent = job.message || "İşleniyor";
  $("completedMetric").textContent = `${job.completed} / ${job.total}`;
  $("foundMetric").textContent = job.found;
  $("notFoundMetric").textContent = job.not_found;
  $("errorMetric").textContent = job.errors;
  $("logOutput").textContent = job.logs?.length
    ? job.logs.join("\n")
    : "Tarama başlatılıyor…";
  $("logOutput").scrollTop = $("logOutput").scrollHeight;

  const ready = Boolean(job.download_ready);
  $("downloadArea").classList.toggle("hidden", !ready);
  if (ready) $("downloadButton").href = `/api/jobs/${job.id}/download`;
}

async function refreshJob() {
  try {
    const payload = await api("/api/jobs/current");
    renderJob(payload.job);
    if (payload.job && ["queued", "running"].includes(payload.job.status)) {
      schedulePoll();
    } else {
      window.clearTimeout(pollTimer);
      pollTimer = null;
    }
  } catch (error) {
    if (error.status === 401) showLogin();
    else showToast(error.message);
  }
}

function schedulePoll() {
  window.clearTimeout(pollTimer);
  pollTimer = window.setTimeout(refreshJob, 1000);
}

$("startButton").addEventListener("click", async () => {
  try {
    $("startButton").disabled = true;
    const job = await api("/api/jobs", { method: "POST", body: "{}" });
    renderJob(job);
    schedulePoll();
  } catch (error) {
    $("startButton").disabled = false;
    showToast(error.message);
    refreshJob();
  }
});

$("cancelButton").addEventListener("click", async () => {
  if (!currentJobId || !window.confirm("Devam eden taramayı iptal etmek istiyor musunuz?")) return;
  try {
    const job = await api(`/api/jobs/${currentJobId}/cancel`, { method: "POST", body: "{}" });
    renderJob(job);
  } catch (error) {
    showToast(error.message);
  }
});

$("loginForm").addEventListener("submit", async (event) => {
  event.preventDefault();
  $("loginError").textContent = "";
  try {
    await api("/api/login", {
      method: "POST",
      body: JSON.stringify({ password: $("password").value }),
    });
    $("password").value = "";
    showDashboard();
    refreshJob();
  } catch (error) {
    $("loginError").textContent = error.message;
  }
});

async function init() {
  try {
    const config = await api("/api/config");
    $("productCount").textContent = config.product_count;
    if (config.requires_auth && !config.authenticated) {
      showLogin();
      return;
    }
    showDashboard();
    await refreshJob();
  } catch (error) {
    showToast("Sunucuya bağlanılamadı.");
  }
}

init();
