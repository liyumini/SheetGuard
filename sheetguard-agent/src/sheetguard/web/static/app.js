/* SheetGuard web 前端：fetch /api/*，无框架，浅色精致工具风。 */
const state = { name: null, round: null, items: [], verdicts: {}, max_review_rounds: null };
const $ = (id) => document.getElementById(id);
// 动态值进 innerHTML 前统一转义（防 XSS，保持最小实现）
const esc = (s) => String(s ?? "").replace(/[&<>"']/g,
  (c) => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

const api = (path) => `/api/workbooks/${encodeURIComponent(state.name)}/${path}`;

// 终态中文映射（与 graph _STAGE_NAMES 对齐扩展；不再裸吐英文）
const STATUS_TEXT = {
  success: "修复成功", partial_success: "部分成功",
  completed_without_repairs: "完成（无需修复）", no_candidates: "无候选",
  failed: "修复失败", unknown: "未知状态",
};

function showTab(name) {
  for (const tab of ["home", "repair", "review", "result"]) {
    const section = $("tab-" + tab);
    const btn = document.querySelector(`nav button[data-tab="${tab}"]`);
    section.hidden = tab !== name;
    btn.classList.toggle("active", tab === name);
    // 已到达过的 Tab 允许通过导航按钮返回（home 始终可用）
    if (tab === name || tab === "home") btn.disabled = false;
  }
}

// 启动时一次性拿到模型配置与审查上限；顶栏状态灯 + 黄条共用此载荷
async function refreshModelStatus() {
  try {
    const s = await (await fetch("/api/model-status")).json();
    state.max_review_rounds = s.max_review_rounds;
    $("lamp-dot").classList.toggle("ok", s.configured);
    $("lamp-dot").classList.toggle("bad", !s.configured);
    $("lamp-text").textContent = s.configured ? "模型已配置" : "模型未配置";
    $("model-banner").classList.toggle("hidden", s.configured);
  } catch {
    $("lamp-text").textContent = "状态未知";
  }
}

async function refreshWorkbooks() {
  const { workbooks } = await (await fetch("/api/workbooks")).json();
  const box = $("workbook-list");
  box.innerHTML = "";
  if (!workbooks.length) {
    box.innerHTML = `<div class="card round-empty">还没有工作簿——先上传一个 .xlsx。</div>`;
    return;
  }
  for (const wb of workbooks) {
    const card = document.createElement("div");
    card.className = "card wb-card";
    const statusText = {idle: "空闲", reviewable: "待审查", repairing: "修复中…"}[wb.status] || wb.status;
    const badgeClass = {idle: "", reviewable: "badge-accent", repairing: "badge-info"}[wb.status] || "";
    const INDEX_TEXT = {fully_certified: "全认证", partial: "部分认证",
                        unknown: "待补全", contested: "有争议"};
    const idxText = wb.index_status ? (INDEX_TEXT[wb.index_status] || wb.index_status) : null;
    card.innerHTML = `
      <div class="rc-head"><b class="mono-chip">${esc(wb.name)}</b>
        <span class="badge ${badgeClass}" data-s>${esc(statusText)}</span>
        <span class="wb-enter">进入 →</span></div>
      ${idxText ? `<div class="rc-line">数据集状态：${esc(idxText)}</div>` : ""}
      <div class="stepper"><div class="round-empty">轮次历史加载中…</div></div>`;
    card.onclick = () => openWorkbook(wb.name, wb.latest_round);
    box.appendChild(card);
    loadRoundTimeline(wb.name, card.querySelector(".stepper"));
  }
}

// 轮次历史 stepper：每轮一个节点（绿=已反馈提交 / 蓝=待审查 / 灰=已提交无待审）
async function loadRoundTimeline(name, box) {
  if (!box) return;
  try {
    const { rounds } = await (await fetch(
      `/api/workbooks/${encodeURIComponent(name)}/rounds`)).json();
    if (!rounds || !rounds.length) {
      box.innerHTML = `<div class="round-empty">尚未运行修复。</div>`;
      return;
    }
    box.innerHTML = rounds.map((r) => {
      const cls = r.has_feedback ? "dot-done" : "dot-open";
      const arts = [];
      if (r.has_audit) arts.push(`<a href="/api/workbooks/${encodeURIComponent(name)}/rounds/${r.round}/report">报告</a>`);
      if (r.has_repaired) arts.push(`<a href="/api/workbooks/${encodeURIComponent(name)}/rounds/${r.round}/repaired">修复副本</a>`);
      return `<div class="step"><span class="${cls} dot"></span><div>
        <div class="step-title">round-${esc(r.round)} <span class="mono-chip">${esc(STATUS_TEXT[r.status] || r.status || "—")}</span></div>
        <div class="step-arts">${arts.length ? arts.join("") : "无产物"}</div>
      </div></div>`;
    }).join("");
    box.onclick = (e) => e.stopPropagation();
  } catch {
    box.innerHTML = `<div class="round-empty">轮次历史加载失败。</div>`;
  }
}

async function openWorkbook(name, round) {
  state.name = name;
  $("repair-name").textContent = name;
  resetRepairView();
  updateRoundBadges(round);
  showTab("repair");
}

// 单参数：审查上限统一从 state.max_review_rounds 读取（model-status 启动时注入，
// 打开审查时以 review-items 载荷刷新）
function updateRoundBadges(round) {
  state.round = round;
  const cap = state.max_review_rounds ?? "?";
  const text = round ? `第 ${round} 轮 · 审查上限 ${cap} 轮` : "尚未开始";
  $("round-badge").textContent = text;
  $("review-round-badge").textContent = text;
}

function resetRepairView() {
  $("inspect-overview").hidden = true;
  $("progress-box").hidden = true;
  $("report-box").hidden = true;
  $("action-bar").hidden = true;
  $("stage-badge").hidden = true;
  $("bar-fill").style.width = "0%";
  $("progress-text").textContent = "";
  $("stage-badge").textContent = "";
}

// 统一任务跟踪：进度回调 + 终态回调，fetch 异常兜底，计时器统一清理
// （startRepair 第 1 轮与 submitReview 第 N 轮共用，进度 UX 完全一致）
// 回调均为 async：统一 await 并观测其异常，避免 onDone 内 renderReport
// 静默失败导致进度条冻结在 100%。
function trackJob(jobId, {onProgress, onDone, onError}) {
  let settled = false; // 防重叠 tick 竞争：终态后不再触发回调
  const timer = setInterval(async () => {
    try {
      const snap = await (await fetch(`/api/jobs/${jobId}`)).json();
      try { onProgress && await onProgress(snap); } catch { /* 进度回调失败不阻断轮询 */ }
      if (snap.status !== "running" && !settled) {
        settled = true;
        clearInterval(timer);
        if (snap.status === "done") {
          try { await (onDone && onDone(snap)); }
          catch (exc) { onError && onError(`报告渲染失败：${exc}`); }
        } else {
          onError && onError(msgSafe(snap.error));
        }
      }
    } catch (exc) {
      if (settled) return;
      settled = true;
      clearInterval(timer);
      onError && onError(`轮询中断：${exc}`);
    }
  }, 1000);
  return timer;
}

// 错误消息统一入口：null/对象兜底为字符串，再交给 esc 转义
function msgSafe(s) { return String(s ?? "未知错误"); }

async function startRepair() {
  resetRepairView();
  $("progress-box").hidden = false;
  $("progress-text").textContent = "准备中…";
  const resp = await fetch(`/api/repair/${encodeURIComponent(state.name)}`, {method: "POST"});
  if (resp.status === 409) {
    const d = await resp.json();
    $("progress-text").textContent =
      d.detail.error_code === "model_not_configured"
        ? "请先在 .env 配置 OPENAI_API_KEY / OPENAI_MODEL。" : "已有任务在运行。";
    return;
  }
  if (!resp.ok) { $("progress-text").textContent = "启动修复失败，请重试。"; return; }
  const {job_id} = await resp.json();
  trackJob(job_id, {
    onProgress: renderProgress,
    onDone: async (snap) => {
      // 完成即回写轮号：轮次徽章不能再停留在「尚未开始」
      updateRoundBadges(snap.result.round);
      await renderReport(state.name, snap.result.round);
    },
    onError: (err) => {
      $("progress-text").innerHTML = `修复失败：<span class="ck-fail">${esc(msgSafe(err))}</span>`;
    },
  });
}

function renderProgress(snap) {
  const p = snap.progress || {};
  $("stage-badge").hidden = !p.stage;
  $("stage-badge").textContent = p.stage || "";
  $("bar-fill").style.width =
    p.total ? `${100 * (p.done || 0) / p.total}%` : "0%";
  const cur = p.current
    ? `正在修复 <span class="mono-chip">${esc(p.current.cell)}</span> · 尝试 ${esc(p.current.attempt)}/${esc(p.current.max_attempts)}`
    : "";
  $("progress-text").innerHTML =
    p.total ? `已完成 ${p.done || 0}/${p.total} 格 ${cur}` : "准备中…";
}

// 六项验证压缩串：`syntax ✓ pattern ✗ recalc —`（三态字符 + 语义色 span）
function checksText(verification) {
  const checks = ((verification || {}).checks) || {};
  return Object.entries(checks).map(([name, c]) => {
    const mark = c && c.status === "passed" ? "✓"
               : c && c.status === "failed" ? "✗" : "—";
    const cls = mark === "✓" ? "ck-pass" : mark === "✗" ? "ck-fail" : "";
    return `<span class="${cls}">${esc(name)} ${mark}</span>`;
  }).join(" ");
}

async function renderReport(name, round) {
  const audit = await (await fetch(
    `/api/workbooks/${encodeURIComponent(name)}/rounds/${round}/audit`)).json();
  const rows = (audit.fixed || []).map((e) => `
    <tr><td class="mono-chip">${esc(e.target)}</td>
    <td><span class="formula-old">${esc(e.old_formula) || "—"}</span></td>
    <td><span class="formula-new">${esc(e.new_formula) || "—"}</span></td>
    <td>${esc((e.hypothesis || {}).error_type)}</td>
    <td>${esc(e.score ?? "—")}</td>
    <td>${checksText(e.verification) || "—"}</td>
    <td class="rc-line">${esc((e.hypothesis || {}).hypothesis)}</td></tr>`).join("");
  const failedLines = (audit.failed || []).map((f) =>
    `<div class="failed-line">${esc(f.target)} — 失败原因 ${esc(f.reason)} · 尝试 ${esc(f.attempts ?? "?")} 次</div>`).join("");
  const dismissedLines = (audit.dismissed || []).map((d) =>
    `<div class="dismissed-line">${esc(d.target)}：agent 结论 ${esc(d.reason)}</div>`).join("");
  const st = esc(STATUS_TEXT[audit.status] || audit.status || "—");
  $("report-box").hidden = false;
  $("report-box").innerHTML = `
    <h3>第 ${esc(round)} 轮 · 终态 <span class="badge badge-ok">${st}</span></h3>
    <p class="rc-line">修复成功 ${(audit.fixed || []).length} · 失败 ${(audit.failed || []).length}
       · dismissed ${(audit.dismissed || []).length} · 延期 ${audit.deferred_count || 0}</p>
    ${rows ? `<table class="report"><tr><th>格子</th><th>原公式</th><th>新公式</th><th>错误类型</th>
      <th>可疑分</th><th>六项验证</th><th>诊断文字</th></tr>${rows}</table>` : "<p class='rc-line'>本轮无修复提案。</p>"}
    ${failedLines ? `<h4>失败候选</h4>${failedLines}` : ""}
    ${dismissedLines ? `<h4>agent 结论（dismissed）</h4>${dismissedLines}` : ""}`;
  $("download-repaired").href =
    `/api/workbooks/${encodeURIComponent(name)}/rounds/${round}/repaired`;
  $("download-report").href =
    `/api/workbooks/${encodeURIComponent(name)}/rounds/${round}/report`;
  $("action-bar").hidden = false;
}

async function openReview() {
  const resp = await fetch(api("review-items"));
  if (resp.status === 404) { alert("没有待审查的轮次，请先运行修复。"); return; }
  const data = await resp.json();
  state.items = data.items;
  state.max_review_rounds = data.max_review_rounds;
  state.verdicts = Object.fromEntries(Object.entries(data.verdicts).map(
    ([t, [v, m]]) => [t, {verdict: v, remark: m}]));
  updateRoundBadges(data.round);
  renderReviewCards();
  showTab("review");
}

function renderReviewCards() {
  const box = $("review-cards");
  box.innerHTML = "";
  for (const item of state.items) {
    // 提案对仅在有提案可确认时出现：failed/dismissed 无新提案，一律隐藏
    const hasProposal = item.kind !== "failed" && item.kind !== "dismissed" && item.new_formula;
    const card = document.createElement("div");
    card.className = "card review-card";
    card.dataset.target = item.target;
    const diag = item.hypothesis
      ? `<div class="rc-diag">诊断：${esc(typeof item.hypothesis === "string"
          ? item.hypothesis
          : ((item.hypothesis.hypothesis || item.hypothesis.error_type) || JSON.stringify(item.hypothesis)))}</div>`
      : "";
    card.innerHTML = `
      <div class="rc-head"><b class="mono-chip">${esc(item.target)}</b>
        ${item.src_round ? `<span class="rc-line">来自第 ${esc(item.src_round)} 轮提案</span>` : ""}</div>
      <span class="badge ${item.kind === "failed" ? "badge-warn" : item.kind === "dismissed" ? "" : "badge-accent"} kind-badge">${esc(kindText(item.kind))}</span>
      <div class="rc-line">原公式：<span class="formula-old">${esc(item.old_formula) || "—"}</span></div>
      <div class="rc-line">${item.kind === "failed" ? "最后尝试" : "提案"}：<span class="formula-new">${esc(item.new_formula) || "—"}</span></div>
      ${diag}
      ${item.reason ? `<div class="rc-line">agent 结论：${esc(item.reason)}</div>` : ""}
      <div class="segmented">
        ${hasProposal ? `<button data-v="correct">提案对</button>` : ""}
        <button data-v="incorrect">没修对</button>
        ${item.kind === "dismissed" ? "" : `<button data-v="false_positive">本来没坏</button>`}
        <button data-v="skipped">跳过</button>
      </div>
      <input class="remark" placeholder="备注（选填，作为下一轮修复线索）" style="display:none">`;
    card.querySelectorAll(".segmented button").forEach((btn) => {
      btn.onclick = () => judge(card, item, btn.dataset.v);
    });
    // 备注在判定后输入：实时更新草稿（等价 CLI q-draft 的备注部分）
    const remarkInput = card.querySelector(".remark");
    remarkInput.addEventListener("input", () => {
      const v = state.verdicts[item.target];
      if (!v) return;
      v.remark = remarkInput.value.trim() || null;
      fetch(api("review-draft"), {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({target: item.target, verdict: v.verdict, remark: v.remark}),
      });
    });
    const v = state.verdicts[item.target];
    if (v) {
      markSelected(card, v.verdict);
      if ((v.verdict === "incorrect" || v.verdict === "false_positive") && v.remark) {
        remarkInput.style.display = "block";
        remarkInput.value = v.remark;
      }
    }
    box.appendChild(card);
  }
  updateReviewProgress();
}

function kindText(kind) {
  return {fixed: "本轮修复", failed: "修复失败", skipped: "历史跳过",
          dismissed: "agent 认为没坏"}[kind] || kind;
}

function judge(card, item, verdict) {
  const remarkInput = card.querySelector(".remark");
  let remark = null;
  if (verdict === "incorrect" || verdict === "false_positive") {
    remarkInput.style.display = "block";
    remark = remarkInput.value.trim() || null;
  }
  state.verdicts[item.target] = {verdict, remark};
  markSelected(card, verdict);
  updateReviewProgress();
  fetch(api("review-draft"), {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify({target: item.target, verdict, remark}),
  });
}

function markSelected(card, verdict) {
  card.querySelectorAll(".segmented button").forEach((b) =>
    b.classList.toggle("selected", b.dataset.v === verdict));
}

function updateReviewProgress() {
  const counts = {fixed: 0, failed: 0, skipped: 0, dismissed: 0};
  for (const item of state.items) counts[item.kind] = (counts[item.kind] || 0) + 1;
  const line1 = `本轮共 ${state.items.length} 条（成功 ${counts.fixed} · 失败 ${counts.failed}`
    + ` · 历史跳过 ${counts.skipped} · 争议 ${counts.dismissed}）`
    + ` · 剩余审查上限 ${state.max_review_rounds ?? "?"} 轮`;
  const judged = Object.keys(state.verdicts).length;
  const line2 = `已判定 ${judged}/${state.items.length} 条（未判提交时等价「跳过」）`;
  $("review-progress").innerHTML = `<div>${esc(line1)}</div><div>${esc(line2)}</div>`;
}

function renderInspectOverview() {
  const box = $("inspect-overview");
  const ins = state.inspect;
  if (!ins) { box.hidden = true; return; }
  box.innerHTML = `<h3>检测概览</h3>
    <div class="result-num">
      <div><b>${ins.sheets ?? 0}</b><span>工作表</span></div>
      <div><b>${ins.formula_cells ?? 0}</b><span>公式格</span></div>
      <div><b>${ins.cross_sheet_dependencies ?? 0}</b><span>跨表依赖</span></div>
      <div><b>${(ins.candidates || []).length}</b><span>可疑候选</span></div>
    </div>`;
  box.hidden = false;
}

async function submitReview() {
  // 防双击/竞态：提交期间禁用按钮。所有退出路径统一由 finally 恢复按钮；
  // 唯一例外是下一轮重修路径（deferred 标记）：恢复权移交 trackJob 的
  // onDone/onError，避免轮次运行期间按钮被提前解禁。
  $("submit-review").disabled = true;
  $("goto-review").disabled = true;
  const reenable = () => {
    $("submit-review").disabled = false;
    $("goto-review").disabled = false;
  };
  let deferred = false;
  try {
    // 只提交已判定条目；未判定的由服务端按「跳过」处理
    const verdicts = Object.entries(state.verdicts).map(([target, v]) => (
      {target, verdict: v.verdict, remark: v.remark}));
    const resp = await fetch(api("review-submit"), {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({verdicts}),
    });
    let body;
    try { body = await resp.json(); }
    catch { alert("提交失败：网络错误"); return; }
    if (resp.status === 409) { alert((body.detail && body.detail.message) || "任务冲突"); return; }
    if (!resp.ok) { alert(body.detail || "提交失败"); return; }
    const box = $("result-card");
    const nums = [
      ["correct", "对", body.counts], ["incorrect", "错", body.counts],
      ["false_positive", "误修", body.counts], ["skipped", "跳过", body.counts],
    ];
    box.innerHTML = `
      <h3>第 ${esc(body.round)} 轮裁决</h3>
      <div class="result-num">${nums.map(([k, label, c]) =>
        `<div><b>${esc(c[k] ?? 0)}</b><span>${label}</span></div>`).join("")}</div>
      ${body.next_round
        ? `<p class="rc-line">已启动第 ${esc(body.next_round.round)} 轮重修，正在修复页实时展示进度。</p>`
        : body.exhausted
          ? `<div class="banner banner-inline">已达轮次上限，被否条目记为 rejected_exhausted。</div>`
          : `<div class="banner banner-inline">没有被否条目，流程结束。</div>`}`;
    if (body.next_round) {
      // 立即切到修复页：第 1/N 轮同一条 trackJob 进度通道
      deferred = true; // 按钮恢复交给 trackJob 终态回调
      state.verdicts = {};
      updateRoundBadges(body.next_round.round);
      resetRepairView();
      $("progress-box").hidden = false;
      $("progress-text").innerHTML = `已完成 0/? 格 <span class="mono-chip">轮次重修中…</span>`;
      trackJob(body.next_round.job_id, {
        onProgress: renderProgress,
        onDone: async (snap) => {
          reenable();
          updateRoundBadges(snap.result.round);
          await renderReport(state.name, snap.result.round);
        },
        onError: (err) => {
          reenable();
          $("progress-text").innerHTML = `下一轮重修失败：<span class="ck-fail">${esc(msgSafe(err))}</span>`;
        },
      });
      showTab("repair");
    } else {
      showTab("result");
    }
  } finally {
    if (!deferred) reenable();
  }
}

// 上传工作簿（file-input 与拖拽共用同一流程）
async function uploadWorkbook(file) {
  const fd = new FormData();
  fd.append("file", file);
  const resp = await fetch("/api/workbooks", {method: "POST", body: fd});
  if (!resp.ok) { alert((await resp.json()).detail); return; }
  const body = await resp.json();
  state.name = body.name;
  state.inspect = body.inspect;
  $("repair-name").textContent = body.name;
  resetRepairView();
  renderInspectOverview();
  updateRoundBadges(null);
  await refreshWorkbooks();
  showTab("repair");
}

$("file-input").onchange = async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  await uploadWorkbook(file);
  e.target.value = "";
};

const dropzone = $("dropzone");
dropzone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropzone.classList.add("dragover");
});
dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropzone.classList.remove("dragover");
  const file = e.dataTransfer.files && e.dataTransfer.files[0];
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".xlsx")) { alert("仅支持 .xlsx"); return; }
  uploadWorkbook(file);
});

$("model-status-lamp").onclick = () =>
  $("model-banner").classList.toggle("hidden");
$("start-repair").onclick = startRepair;
$("goto-review").onclick = openReview;
$("submit-review").onclick = submitReview;
document.querySelectorAll("nav button").forEach((b) =>
  b.onclick = () => showTab(b.dataset.tab));
refreshModelStatus();
refreshWorkbooks();
showTab("home");
