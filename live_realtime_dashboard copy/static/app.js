const STATE_REFRESH_INTERVAL_MS = 8000;
let lastState = null;
let configDirty = false;
const globalStatus = document.querySelector("#globalStatus");

function configPanel() {
  return $("#configPanel");
}

function $on(el, ev, fn) {
  if (el) el.addEventListener(ev, fn);
}

function $(selector) {
  return document.querySelector(selector);
}

function $all(selector) {
  return Array.from(document.querySelectorAll(selector));
}

function fmtNum(v) {
  const n = Number(v || 0);
  return Number.isFinite(n) ? n.toLocaleString() : "0";
}

function escapeHtml(s) {
  const d = document.createElement("div");
  d.textContent = s == null ? "" : String(s);
  return d.innerHTML;
}

function safeHttpUrl(u) {
  const s = String(u || "").trim();
  if (s.startsWith("http://") || s.startsWith("https://")) return s;
  return "";
}

function linkCell(url, label) {
  const u = safeHttpUrl(url);
  if (!u) return "—";
  return `<a href="${escapeHtml(u)}" target="_blank" rel="noopener noreferrer">${escapeHtml(label || "进入")}</a>`;
}

function deltaHtml(v) {
  const n = Number(v || 0);
  if (n > 0) return `<span class="num-up">+${fmtNum(n)}</span>`;
  if (n < 0) return `<span class="num-down">${fmtNum(n)}</span>`;
  return "0";
}

function selectedPlatforms() {
  return $all('[data-role="platformItem"]:checked').map((el) => el.value);
}

function setSelectedPlatforms(platforms) {
  const selected = new Set(Array.isArray(platforms) ? platforms : []);
  $all('[data-role="platformItem"]').forEach((el) => {
    el.checked = selected.has(el.value);
  });
}

function readConfig() {
  return {
    platforms: selectedPlatforms(),
    keywords: $("#keywords") ? $("#keywords").value.trim() : "",
    interval_minutes: Number($("#intervalMinutes") ? $("#intervalMinutes").value : 120),
    summary_time: $("#summaryTime") ? $("#summaryTime").value : "",
    summary_range_start: $("#summaryRangeStart") ? $("#summaryRangeStart").value : "",
    summary_range_end: $("#summaryRangeEnd") ? $("#summaryRangeEnd").value : "",
    max_pages: Number($("#maxPages") ? $("#maxPages").value : 2),
    enable_ws: $("#enableWs") ? $("#enableWs").checked : true,
    ws_workers: Number($("#wsWorkers") ? $("#wsWorkers").value : 16),
    wecom_webhook_url: $("#wecomWebhook") ? $("#wecomWebhook").value.trim() : "",
    wecom_auto_push: $("#wecomAutoPush") ? $("#wecomAutoPush").checked : true,
    blacklist_authors: $("#blacklistAuthors") ? $("#blacklistAuthors").value : "",
  };
}

function fillConfig(config) {
  if (configDirty) return;
  const data = config || {};
  if ($("#keywords") && document.activeElement !== $("#keywords")) {
    $("#keywords").value = Array.isArray(data.keywords) ? data.keywords.join(", ") : "";
  }
  if ($("#intervalMinutes") && document.activeElement !== $("#intervalMinutes")) {
    $("#intervalMinutes").value = data.interval_minutes || 120;
  }
  if ($("#summaryTime") && document.activeElement !== $("#summaryTime")) {
    $("#summaryTime").value = data.summary_time || "17:00";
  }
  if ($("#summaryRangeStart") && document.activeElement !== $("#summaryRangeStart")) {
    $("#summaryRangeStart").value = data.summary_range_start || "00:00";
  }
  if ($("#summaryRangeEnd") && document.activeElement !== $("#summaryRangeEnd")) {
    $("#summaryRangeEnd").value = data.summary_range_end || "23:59";
  }
  if ($("#maxPages") && document.activeElement !== $("#maxPages")) {
    $("#maxPages").value = data.max_pages || 2;
  }
  if ($("#wsWorkers") && document.activeElement !== $("#wsWorkers")) {
    $("#wsWorkers").value = data.ws_workers || 16;
  }
  if ($("#enableWs")) $("#enableWs").checked = Boolean(data.enable_ws);
  if ($("#wecomWebhook") && document.activeElement !== $("#wecomWebhook")) {
    $("#wecomWebhook").value = data.wecom_webhook_url || "";
  }
  if ($("#wecomAutoPush")) $("#wecomAutoPush").checked = Boolean(data.wecom_auto_push);
  if ($("#blacklistAuthors") && document.activeElement !== $("#blacklistAuthors")) {
    $("#blacklistAuthors").value = Array.isArray(data.blacklist_authors) ? data.blacklist_authors.join("\n") : "";
  }
  setSelectedPlatforms(data.platforms || []);
}

function setConfigLocked(locked) {
  const panel = configPanel();
  if (!panel) return;
  panel.classList.toggle("config-locked", Boolean(locked));
  const controls = [
    "#keywords",
    "#intervalMinutes",
    "#summaryTime",
    "#summaryRangeStart",
    "#summaryRangeEnd",
    "#maxPages",
    "#enableWs",
    "#wsWorkers",
    "#wecomWebhook",
    "#wecomAutoPush",
    "#blacklistAuthors",
    '[data-role="platformItem"]',
  ];
  controls.forEach((selector) => {
    $all(selector).forEach((el) => {
      el.disabled = Boolean(locked);
    });
  });
  if ($("#saveConfigBtn")) $("#saveConfigBtn").disabled = Boolean(locked);
  if ($("#unlockConfigBtn")) $("#unlockConfigBtn").disabled = !locked;
}

function setSelectOptions(selectEl, items, placeholder) {
  if (!selectEl) return;
  const current = selectEl.value;
  selectEl.innerHTML = "";
  const empty = document.createElement("option");
  empty.value = "";
  empty.textContent = placeholder;
  selectEl.appendChild(empty);
  for (const item of items || []) {
    const opt = document.createElement("option");
    opt.value = String(item.value || "");
    opt.textContent = String(item.label || item.value || "");
    selectEl.appendChild(opt);
  }
  if ([...selectEl.options].some((opt) => opt.value === current)) {
    selectEl.value = current;
  }
}

function unionNodeOptions(platformStates, platforms) {
  const merged = new Map();
  for (const name of platforms) {
    const list = (platformStates[name] && platformStates[name].available_nodes) || [];
    for (const item of list) {
      if (item && item.value && !merged.has(item.value)) {
        merged.set(item.value, item);
      }
    }
  }
  return [...merged.values()].sort((a, b) => String(b.value).localeCompare(String(a.value)));
}

function renderNodeSelectors(data) {
  const platforms = selectedPlatforms();
  const nodeOptions = unionNodeOptions(data.platforms || {}, platforms);
  setSelectOptions($("#summaryStartNode"), nodeOptions, "请选择起始节点");
  setSelectOptions($("#summaryEndNode"), nodeOptions, "请选择结束节点");
  if ($("#summaryRangeStatus")) {
    $("#summaryRangeStatus").textContent = nodeOptions.length
      ? `节点汇总: 已加载 ${nodeOptions.length} 个节点`
      : "节点汇总: 当前没有可用节点";
  }
}

function renderOverviewCards(data) {
  const box = $("#overviewCards");
  if (!box) return;
  box.innerHTML = "";
  for (const name of data.platform_order || []) {
    const platformData = data.platforms && data.platforms[name];
    if (!platformData) continue;
    const s = platformData.summary || {};
    const card = document.createElement("div");
    card.className = "overview-card";
    card.innerHTML = `
      <div class="overview-title">${escapeHtml(s.label || name)}</div>
      <div class="overview-line">状态: <b>${s.running ? "运行中" : "已停止"}</b></div>
      <div class="overview-line">节点数: <b>${fmtNum(s.node_count)}</b></div>
      <div class="overview-line">当前房间: <b>${fmtNum(s.total_rooms)}</b></div>
      <div class="overview-line">累计不同直播间: <b>${fmtNum(s.unique_rooms_total)}</b></div>
      <div class="overview-line">黑名单作者: <b>${fmtNum(s.blacklist_total)}</b></div>
      <div class="overview-line">最后节点: <b>${escapeHtml(s.last_run_at || "—")}</b></div>
      <div class="overview-line">最近推送: <b>${escapeHtml(s.last_push_message || "—")}</b></div>
      <div class="overview-line">最近错误: <b>${escapeHtml(s.last_error || "无")}</b></div>
    `;
    box.appendChild(card);
  }
}

function extraInfo(platform, row) {
  if (platform === "douyin") {
    const extra = row.extra || {};
    return `粉丝 ${fmtNum(extra.follower_count)} / 粉丝团 ${fmtNum(extra.fans_club_count)}`;
  }
  return escapeHtml(row.category || "—");
}

function formatSummaryRangeText(start, end) {
  const startText = String(start || "").trim() || "—";
  const endText = String(end || "").trim() || "—";
  if (startText === "—" || endText === "—") {
    return `${startText} ~ ${endText}`;
  }
  if (endText <= startText) {
    return `${startText}（前一天）~ ${endText}（当天）`;
  }
  return `${startText} ~ ${endText}`;
}

function renderPlatformTables(data) {
  const selected = new Set(selectedPlatforms());
  $all(".platform-data").forEach((section) => {
    const platform = section.dataset.platform;
    const label = section.dataset.label || platform;
    const platformData = data.platforms && data.platforms[platform];
    const tbody = section.querySelector('[data-role="roomsTableBody"]');
    const meta = section.querySelector('[data-role="platformMeta"]');
    const summary = platformData ? platformData.summary || {} : {};
    const rows = platformData ? platformData.current_rooms || [] : [];
    const shouldShow = selected.has(platform) || Boolean(summary.running) || rows.length > 0;
    section.style.display = shouldShow ? "block" : "none";
    if (!shouldShow || !tbody) return;
    if (meta) {
      meta.textContent = `${label} · ${summary.running ? "运行中" : "已停止"} · 节点 ${summary.node_count || 0} · 最后抓取 ${summary.last_run_at || "—"}`;
    }
    tbody.innerHTML = "";
    if (!rows.length) {
      const tr = document.createElement("tr");
      tr.innerHTML = '<td colspan="14" class="empty-cell">当前没有节点数据。</td>';
      tbody.appendChild(tr);
      return;
    }
    for (const row of rows) {
      const tr = document.createElement("tr");
      const home = safeHttpUrl(row.author_home);
      tr.innerHTML = `
        <td>${escapeHtml(row.search_keyword || "")}</td>
        <td>${escapeHtml(row.fetched_at || "")}</td>
        <td>${home ? `<a href="${escapeHtml(home)}" target="_blank" rel="noopener noreferrer">主页</a>` : "—"}</td>
        <td>${linkCell(row.live_url, "进入")}</td>
        <td>${escapeHtml(row.anchor_name || "")}</td>
        <td class="cell-title">${escapeHtml(row.title || "")}</td>
        <td>${escapeHtml(row.live_start_time || "")}</td>
        <td>${fmtNum(row.watched_count)}</td>
        <td>${fmtNum(row.ws_like_count)}</td>
        <td>${escapeHtml(row.like_source || "—")}</td>
        <td>${fmtNum(row.online_hot)}</td>
        <td>${deltaHtml(row.delta_watched)}</td>
        <td>${deltaHtml(row.delta_like)}</td>
        <td>${extraInfo(platform, row)}</td>
      `;
      tbody.appendChild(tr);
    }
  });
}

function renderConfigStatus(config, data) {
  if (!$("#configStatus")) return;
  const runningPlatforms = (data.selected_platforms || []).map((name) => (data.labels && data.labels[name]) || name);
  const parts = [];
  parts.push(config.saved ? "统一配置已保存" : "统一配置未保存");
  parts.push(config.locked ? "页面已锁定" : "页面可编辑");
  if (config.updated_at) parts.push(`保存于 ${config.updated_at}`);
  parts.push(`当前平台 ${runningPlatforms.length ? runningPlatforms.join("、") : "未启动"}`);
  parts.push(`汇总时间 ${config.summary_time || "—"}`);
  parts.push(`汇总区间 ${formatSummaryRangeText(config.summary_range_start, config.summary_range_end)}`);
  $("#configStatus").textContent = `统一配置: ${parts.join(" · ")}`;
}

function renderDouyinCookieStatus(info) {
  const el = $("#cookieStatus");
  if (!el) return;
  if (!info) {
    el.textContent = "Cookie 状态: 未知";
    return;
  }
  if (!info.configured) {
    el.innerHTML = '<span class="num-down">Cookie 状态: 未配置</span>';
    return;
  }
  const flags = [];
  flags.push(info.has_uifid ? "UIFID ✓" : "UIFID ✗");
  flags.push(info.has_session ? "sessionid ✓" : "sessionid ✗");
  flags.push(info.has_ttwid ? "ttwid ✓" : "ttwid ✗");
  const savedAt = info.saved_at_text ? ` · 更新时间 ${info.saved_at_text}` : "";
  el.textContent = `Cookie 状态: 已配置 · 字段数 ${info.fields || 0}${savedAt} · ${flags.join(" / ")}`;
}

function renderDouyinCookieAutoStatus(info) {
  const el = $("#cookieAutoStatus");
  if (!el) return;
  if (!info) {
    el.textContent = "自动化状态: 未知";
    return;
  }
  if (!info.available) {
    el.innerHTML = '<span class="num-down">自动化状态: 当前环境缺少 DrissionPage</span>';
    return;
  }
  const parts = [];
  parts.push(info.browser_opened ? "浏览器已打开" : "浏览器未打开");
  parts.push(info.auto_refresh_running ? "持续同步中" : "未同步");
  parts.push(info.logged_in ? "已识别登录态" : "未识别登录态");
  if (info.last_message) parts.push(info.last_message);
  if (info.last_error) parts.push(`错误: ${info.last_error}`);
  el.textContent = `自动化状态: ${parts.join(" · ")}`;
}

async function postJson(url, payload) {
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload || {}),
  });
  if (!res.ok) {
    let message = "";
    try {
      const data = await res.json();
      message = data && data.message ? String(data.message) : "";
    } catch (err) {
      try {
        message = await res.text();
      } catch (textErr) {
        message = "";
      }
    }
    throw new Error(message || `请求失败: ${res.status}`);
  }
  return res;
}

async function downloadExport(payload, fallbackName) {
  const res = await fetch("/api/export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload || {}),
  });
  if (!res.ok) {
    let message = "";
    try {
      const data = await res.json();
      message = data && data.message ? String(data.message) : "";
    } catch (_err) {
      message = `请求失败: ${res.status}`;
    }
    throw new Error(message || `请求失败: ${res.status}`);
  }
  const blob = await res.blob();
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const cd = res.headers.get("Content-Disposition") || "";
  const m = cd.match(/filename\*=UTF-8''([^;]+)/) || cd.match(/filename="?([^";]+)"?/);
  a.download = m ? decodeURIComponent(m[1]) : fallbackName || "export.xlsx";
  document.body.appendChild(a);
  a.click();
  a.remove();
  window.URL.revokeObjectURL(url);
}

async function saveConfig() {
  const res = await postJson("/api/config/save", readConfig());
  const json = await res.json();
  configDirty = false;
  setConfigLocked(true);
  globalStatus.textContent = `状态: ${json.message}`;
  await fetchState();
}

async function startMonitor() {
  const res = await postJson("/api/start", readConfig());
  const json = await res.json();
  configDirty = false;
  globalStatus.textContent = `状态: ${json.message}`;
  await fetchState();
}

async function stopMonitor() {
  const res = await postJson("/api/stop", { platforms: selectedPlatforms() });
  const json = await res.json();
  globalStatus.textContent = `状态: ${json.message}`;
  await fetchState();
}

async function unlockConfig() {
  const res = await postJson("/api/config/unlock", {});
  const json = await res.json();
  setConfigLocked(false);
  configDirty = false;
  globalStatus.textContent = `状态: ${json.message}`;
  await fetchState();
}

function firstKeyword() {
  const raw = $("#keywords") ? $("#keywords").value : "";
  return String(raw || "")
    .replaceAll("，", ",")
    .replaceAll("\n", ",")
    .split(",")
    .map((v) => v.trim())
    .find(Boolean) || "";
}

async function fetchState() {
  try {
    const res = await fetch("/api/state");
    const json = await res.json();
    if (!json.ok) return;
    const data = json.data || {};
    lastState = data;
    fillConfig(data.unified_config || {});
    setConfigLocked(Boolean((data.unified_config || {}).locked));
    renderConfigStatus(data.unified_config || {}, data);
    renderOverviewCards(data);
    renderPlatformTables(data);
    renderNodeSelectors(data);
    renderDouyinCookieStatus(data.douyin_cookie);
    renderDouyinCookieAutoStatus(data.douyin_cookie_auto);
    const summary = [];
    for (const name of data.platform_order || []) {
      const s = ((data.platforms || {})[name] || {}).summary || {};
      const label = (data.labels && data.labels[name]) || name;
      summary.push(`${label}: ${s.running ? "运行中" : "已停止"} · 节点 ${s.node_count || 0} · 当前 ${s.total_rooms || 0}`);
    }
    globalStatus.textContent = `状态: ${summary.join(" | ")}`;
  } catch (err) {
    globalStatus.textContent = `状态异常: ${err.message}`;
  }
}

async function saveDouyinCookie() {
  const raw = $("#douyinCookie") ? $("#douyinCookie").value : "";
  if (!raw || !raw.trim()) {
    alert("请粘贴抖音 Cookie 字符串或 JSON");
    return;
  }
  const res = await postJson("/api/douyin/cookie", { cookie: raw });
  const json = await res.json();
  renderDouyinCookieStatus(json.data);
  if ($("#cookieStatus") && json.message) $("#cookieStatus").textContent = json.message;
  if ($("#douyinCookie") && json.ok) $("#douyinCookie").value = "";
}

async function reloadDouyinCookie() {
  const res = await fetch("/api/douyin/cookie");
  const json = await res.json();
  renderDouyinCookieStatus(json.data);
}

async function startDouyinCookieAuto() {
  const keyword = firstKeyword();
  const res = await postJson("/api/douyin/cookie/auto/start", { keyword });
  const json = await res.json();
  renderDouyinCookieAutoStatus(json.data);
  globalStatus.textContent = `状态: ${json.message}`;
  await fetchState();
}

async function refreshDouyinCookieFromAuto() {
  const res = await postJson("/api/douyin/cookie/auto/refresh", {});
  const json = await res.json();
  if (json.data && json.data.cookie) renderDouyinCookieStatus(json.data.cookie);
  if (json.data && json.data.auto) renderDouyinCookieAutoStatus(json.data.auto);
  globalStatus.textContent = `状态: ${json.message}`;
  await fetchState();
}

async function pushCurrentNode() {
  const payload = readConfig();
  const res = await postJson("/api/wecom/push", payload);
  const json = await res.json();
  globalStatus.textContent = `状态: ${json.message}`;
  await fetchState();
}

async function pushSummaryNow() {
  const payload = readConfig();
  payload.start_at = $("#summaryStartNode") ? $("#summaryStartNode").value : "";
  payload.end_at = $("#summaryEndNode") ? $("#summaryEndNode").value : "";
  const res = await postJson("/api/wecom/summary/push", payload);
  const json = await res.json();
  globalStatus.textContent = `状态: ${json.message}`;
  await fetchState();
}

async function exportSelected() {
  const platforms = selectedPlatforms();
  await downloadExport({ platforms, mode: "full" }, "直播监控数据.zip");
}

function bindConfigDirtyTracking() {
  const selectors = [
    "#keywords",
    "#intervalMinutes",
    "#summaryTime",
    "#summaryRangeStart",
    "#summaryRangeEnd",
    "#maxPages",
    "#enableWs",
    "#wsWorkers",
    "#wecomWebhook",
    "#wecomAutoPush",
    "#blacklistAuthors",
    '[data-role="platformItem"]',
  ];
  selectors.forEach((selector) => {
    $all(selector).forEach((el) => {
      $on(el, "input", () => {
        configDirty = true;
      });
      $on(el, "change", () => {
        configDirty = true;
      });
    });
  });
}

$on($("#saveConfigBtn"), "click", () => saveConfig().catch((e) => alert(e.message)));
$on($("#unlockConfigBtn"), "click", () => unlockConfig().catch((e) => alert(e.message)));
$on($("#startBtn"), "click", () => startMonitor().catch((e) => alert(e.message)));
$on($("#stopBtn"), "click", () => stopMonitor().catch((e) => alert(e.message)));
$on($("#pushCurrentBtn"), "click", () => pushCurrentNode().catch((e) => alert(e.message)));
$on($("#pushSummaryBtn"), "click", () => pushSummaryNow().catch((e) => alert(e.message)));
$on($("#exportBtn"), "click", () => exportSelected().catch((e) => alert(e.message)));
$on($("#refreshBtn"), "click", () => fetchState().catch((e) => alert(e.message)));
$on($("#autoCookieStartBtn"), "click", () => startDouyinCookieAuto().catch((e) => alert(e.message)));
$on($("#autoCookieRefreshBtn"), "click", () => refreshDouyinCookieFromAuto().catch((e) => alert(e.message)));
$on($("#saveCookieBtn"), "click", () => saveDouyinCookie().catch((e) => alert(e.message)));
$on($("#reloadCookieBtn"), "click", () => reloadDouyinCookie().catch((e) => alert(e.message)));

bindConfigDirtyTracking();
fetchState();
setInterval(fetchState, STATE_REFRESH_INTERVAL_MS);
