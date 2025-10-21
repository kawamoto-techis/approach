/* =========================================================
   script.js  ― フラット&クリーン版（2025-10）
   - ローカル保存＋GAS 同期（履歴＝即時、リスト＝自動保存+ボタン）
   - 履歴一覧（登録ページ & 閲覧専用ページ）
   - クイックメモ、検索/絞込/並べ替え、折りたたみ
   - 詳細モーダルで編集/削除
   ======================================================= */

/* ---- 設定（あなたの GAS デプロイ URL に置き換え可） ------------- */
const GAS_BASE = "https://script.google.com/macros/s/AKfycbx1u3qfMh7GxCZ6jMa2h3m2Q296w9ZgV3V8pKuWdXyop4r8TVocDS4eAP_lUKP16Jnq6A/exec";

/* ---- 便利関数 --------------------------------------------------- */
const $  = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const uuid = () => (crypto?.randomUUID?.() || `id-${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`);

const debounce = (fn, ms=800) => { let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>fn(...a), ms); }; };

// トースト
const toast = (() => {
  let el;
  return (msg, type="error") => {
    if (!el) {
      el = document.createElement("div");
      el.style.cssText = `position:fixed; right:16px; top:16px; z-index:9999; padding:10px 14px; border-radius:10px; color:#fff; box-shadow:0 8px 24px rgba(0,0,0,.18); max-width:420px`;
      document.body.appendChild(el);
    }
    el.style.background = (type==="ok") ? "#00abae" : "#e74c3c";
    el.textContent = msg;
    el.style.opacity = "1";
    setTimeout(()=>{ el.style.transition="opacity .4s"; el.style.opacity="0"; }, 2800);
  };
})();

// 日付
const parseToDate = (v) => {
  if (!v) return null;
  const d1 = new Date(v);
  if (!isNaN(d1.getTime())) return d1;
  const m = String(v).match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (m) {
    const [ , y, mo, da, hh='0', mm='0' ] = m;
    return new Date(+y, +mo-1, +da, +hh, +mm, 0);
  }
  return null;
};
const ymKey = (v) => { const d = parseToDate(v); if (!d) return ""; return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`; };
const fmtDisplay = (iso) => {
  const d = parseToDate(iso); if (!d) return iso || "";
  return `${d.getFullYear()}/${d.getMonth()+1}/${d.getDate()} ${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`;
};
const norm = (s) => String(s || "").toLowerCase().replace(/\s+/g, "").trim();

/* ---- GAS I/O（厳格） -------------------------------------------- */
const parseJsonStrict = async (res) => {
  const txt = await res.text();
  let json;
  try { json = JSON.parse(txt); }
  catch { throw new Error(`GASがJSONを返しませんでした（${res.status}）。応答: ${txt.slice(0,120)}...`); }
  if (json && json.ok === false) throw new Error(json.error || "GAS returned ok:false");
  return json;
};
const apiGet  = async (params) => {
  const url = GAS_BASE + "?" + new URLSearchParams(params);
  const res = await fetch(url, { method: "GET", mode: "cors", redirect: "follow" });
  if (!res.ok) throw new Error(`GET ${res.status}`);
  return parseJsonStrict(res);
};
const apiPost = async (body) => {
  const res = await fetch(GAS_BASE, { method: "POST", mode: "cors", redirect: "follow", body: JSON.stringify(body) });
  if (!res.ok) throw new Error(`POST ${res.status}`);
  return parseJsonStrict(res);
};

/* ---- ステート ---------------------------------------------------- */
let kigyouWords = JSON.parse(localStorage.getItem("kigyouList")) || [];
let mikomiWords = JSON.parse(localStorage.getItem("mikomiList")) || [];
let botsuWords  = JSON.parse(localStorage.getItem("botsuList"))  || [];
let historyData = JSON.parse(localStorage.getItem("historyData")) || [];
let contactsData = JSON.parse(localStorage.getItem("contactsData")) || [];
let kigyouSet = new Set(kigyouWords.map(norm));
let mikomiSet = new Set(mikomiWords.map(norm));
let botsuSet  = new Set(botsuWords.map(norm));
const refreshSets = () => { kigyouSet = new Set(kigyouWords.map(norm)); mikomiSet = new Set(mikomiWords.map(norm)); botsuSet  = new Set(botsuWords.map(norm)); };
const saveLocal = () => {
  localStorage.setItem("kigyouList", JSON.stringify(kigyouWords));
  localStorage.setItem("mikomiList", JSON.stringify(mikomiWords));
  localStorage.setItem("botsuList",  JSON.stringify(botsuWords));
  localStorage.setItem("historyData", JSON.stringify(historyData));
  localStorage.setItem("contactsData", JSON.stringify(contactsData));
};

/* ---- DOM -------------------------------------------------------- */
const tabBtns = $$(".tab-btn");
const pages   = $$(".page-content");
const quickLinks = $("#top-quick-links");

const historyForm = $("#history-form");
const companyNameEl = $("#company-name");
const mediaSelectEl = $("#media-select");
const historyNoteEl = $("#history-note");
const searchBoxMini = $("#search-box");
const tbodyMini  = $("#history-table-body");
const historyImportBtn = $("#history-import-btn");
const historyExportBtn = $("#history-export-btn");
const historyImportFile= $("#history-import-file");

const memoTextarea = $("#quick-memo");
const memoSaveBtn  = $("#quick-memo-save");
const memoClearBtn = $("#quick-memo-clear");
const memoStatusEl = $("#memo-status");

const monthFilter2   = $("#month-filter-2");
const sortOrder2     = $("#sort-order-2");
const clearFilters2  = $("#clear-filters-2");
const searchBox2     = $("#search-box-2");
const collapseAllBtn = $("#collapse-all");
const tbodyFull      = $("#history-table-body-2");

const kigyouTA = $("#kigyou-list");
const mikomiTA = $("#mikomi-list");
const botsuTA  = $("#botsu-list");
const kigyouImportBtn = $("#kigyou-import-btn");
const kigyouExportBtn = $("#kigyou-export-btn");
const mikomiImportBtn = $("#mikomi-import-btn");
const mikomiExportBtn = $("#mikomi-export-btn");
const botsuImportBtn  = $("#botsu-import-btn");
const botsuExportBtn  = $("#botsu-export-btn");
const kigyouImportFile= $("#kigyou-import-file");
const mikomiImportFile= $("#mikomi-import-file");
const botsuImportFile = $("#botsu-import-file");

const kigyouSaveBtn  = $("#kigyou-save-btn");
const kigyouReloadBtn= $("#kigyou-reload-btn");
const mikomiSaveBtn  = $("#mikomi-save-btn");
const mikomiReloadBtn= $("#mikomi-reload-btn");
const botsuSaveBtn   = $("#botsu-save-btn");
const botsuReloadBtn = $("#botsu-reload-btn");

const contactsForm = $("#contacts-form");
const contactCompanyEl = $("#contact-company");
const contactNameEl = $("#contact-name");
const contactEmailEl = $("#contact-email");
const contactTelEl = $("#contact-tel");
const contactMemoEl = $("#contact-memo");
const contactsSearchBox = $("#contacts-search-box");
const contactsTbody = $("#contacts-table-body");
const contactsImportBtn = $("#contacts-import-btn");
const contactsExportBtn = $("#contacts-export-btn");
const contactsImportFile = $("#contacts-import-file");

const modal           = $("#detail-modal");
const modalCloseBtn   = modal?.querySelector(".close-btn");
const detailForm      = $("#detail-form");
const detailIdEl      = $("#detail-id");
const detailCreatedEl = $("#detail-created");
const detailCompanyEl = $("#detail-company");
const detailMediaEl   = $("#detail-media");
const detailNoteEl    = $("#detail-note");
const detailDeleteBtn = $("#detail-delete");
const detailCancelBtn = $("#detail-cancel");

const approachModal = $("#approach-modal");
const approachModalTitle = $("#approach-modal-title");
const approachForm = $("#approach-form");
const approachCompanyNameEl = $("#approach-company-name");
const approachMediaSelectEl = $("#approach-media-select");
const approachNoteEl = $("#approach-note");
const approachHistoryTbody = $("#approach-history-table-body");
const approachModalCloseBtn = approachModal?.querySelector(".close-btn");

const openApproachModal = (companyName) => {
  if (!approachModal) return;
  approachModalTitle.textContent = `${companyName} へのアプローチ`;
  approachCompanyNameEl.value = companyName;

  const companyHistory = historyData
    .filter(item => item.company === companyName)
    .sort((a, b) => (parseToDate(b.createdAt)?.getTime() ?? 0) - (parseToDate(a.createdAt)?.getTime() ?? 0));

  approachHistoryTbody.innerHTML = "";
  companyHistory.forEach(item => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${fmtDisplay(item.createdAt)}</td>
      <td>${escapeCsv(item.media)}</td>
      <td>${escapeCsv(item.note)}</td>
    `;
    approachHistoryTbody.appendChild(tr);
  });

  approachModal.classList.add("active");
  approachModal.setAttribute("aria-hidden", "false");
};

const closeApproachModal = () => {
  if (!approachModal) return;
  approachModal.classList.remove("active");
  approachModal.setAttribute("aria-hidden", "true");
  approachForm.reset();
};

approachModalCloseBtn?.addEventListener("click", closeApproachModal);
window.addEventListener("click", (e) => { if (e.target === approachModal) closeApproachModal(); });
document.addEventListener("keydown", (e) => { if (e.key === "Escape" && approachModal?.classList.contains("active")) closeApproachModal(); });


/* ---- タブ切替 & 右上ショートカット ------------------------------ */
const showPage = (id) => {
  pages.forEach(p => p.classList.remove("active"));
  $(id)?.classList.add("active");
  tabBtns.forEach(b => b.classList.remove("active"));
  tabBtns.find(b => b.dataset.target === id)?.classList.add("active");
  quickLinks.style.display = (id === "#page-history-list") ? "flex" : "none";
};
tabBtns.forEach(btn => btn.addEventListener("click", () => showPage(btn.dataset.target)));

/* ---- 行ステータス ------------------------------------------------ */
const getStatusClass = (company) => {
  const k = norm(company);
  if (botsuSet.has(k))  return "status-botsu";
  if (kigyouSet.has(k)) return "status-kigyou";
  if (mikomiSet.has(k)) return "status-mikomi";
  return "";
};

/* ---- 描画 -------------------------------------------------------- */
const createNoteCell = (text, collapsed=false) => {
  const td = document.createElement("td");
  td.className = "note-cell";
  td.style.maxWidth = "680px";
  td.style.whiteSpace = "pre-wrap";
  td.style.wordBreak  = "break-word";
  if (collapsed) td.dataset.collapsed = "1";
  td.textContent = text || "";
  return td;
};

const renderMini = () => {
  if (!tbodyMini) return;
  const q = (searchBoxMini?.value || "").toLowerCase();
  tbodyMini.innerHTML = "";
  historyData
    .filter(it => it.company.toLowerCase().includes(q) || it.media.toLowerCase().includes(q) || it.note.toLowerCase().includes(q))
    .slice()
    .sort((a,b)=> (parseToDate(b.createdAt)?.getTime() ?? 0) - (parseToDate(a.createdAt)?.getTime() ?? 0))
    .forEach(item => {
      const tr = document.createElement("tr");
      const st = getStatusClass(item.company);
      if (st) tr.classList.add(st);
      const tdDate = document.createElement("td"); tdDate.textContent = fmtDisplay(item.createdAt);
      const tdCom  = document.createElement("td"); tdCom.textContent = item.company;
      const tdMed  = document.createElement("td"); tdMed.textContent = item.media;
      const tdNote = createNoteCell(item.note, true);
      tr.setAttribute("role", "button");
      tr.addEventListener("click", () => openDetail(item.id));
      tr.append(tdDate, tdCom, tdMed, tdNote);
      tbodyMini.appendChild(tr);
    });
};

const renderFull = () => {
  if (!tbodyFull) return;
  const kw = (searchBox2?.value || "").toLowerCase();
  const ym = monthFilter2?.value || "";
  const sortAsc = (sortOrder2?.value || "desc") === "asc";

  let list = historyData.slice();
  if (ym) list = list.filter(it => ymKey(it.createdAt) === ym);
  if (kw) list = list.filter(it => it.company.toLowerCase().includes(kw) || it.media.toLowerCase().includes(kw) || it.note.toLowerCase().includes(kw));
  list.sort((a,b) => {
    const diff = (parseToDate(a.createdAt)?.getTime() ?? 0) - (parseToDate(b.createdAt)?.getTime() ?? 0);
    return sortAsc ? diff : -diff;
  });

  tbodyFull.innerHTML = "";
  list.forEach(item => {
    const tr = document.createElement("tr");
    const st = getStatusClass(item.company);
    if (st) tr.classList.add(st);
    const tdDate = document.createElement("td"); tdDate.textContent = fmtDisplay(item.createdAt);
    const tdCom  = document.createElement("td"); tdCom.textContent = item.company;
    const tdMed  = document.createElement("td"); tdMed.textContent = item.media;
    const tdNote = createNoteCell(item.note, true);
    tr.setAttribute("role", "button");
    tr.addEventListener("click", () => openDetail(item.id));
    tr.append(tdDate, tdCom, tdMed, tdNote);
    tbodyFull.appendChild(tr);
  });
};

const renderAll = () => { renderMini(); renderFull(); renderContacts(); };

const renderContacts = () => {
  if (!contactsTbody) return;
  const q = (contactsSearchBox?.value || "").toLowerCase();
  contactsTbody.innerHTML = "";
  contactsData
    .filter(it => it.company.toLowerCase().includes(q) || it.name.toLowerCase().includes(q) || it.email.toLowerCase().includes(q) || it.tel.toLowerCase().includes(q) || it.memo.toLowerCase().includes(q))
    .forEach(item => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeCsv(item.company)}</td>
        <td>${escapeCsv(item.name)}</td>
        <td>${escapeCsv(item.email)}</td>
        <td>${escapeCsv(item.tel)}</td>
        <td>${escapeCsv(item.memo)}</td>
        <td>
          <button class="primary-btn approach-contact-btn" data-id="${item.id}" data-company="${escapeCsv(item.company)}">アプローチ</button>
          <button class="secondary-btn edit-contact-btn" data-id="${item.id}">編集</button>
          <button class="danger-btn delete-contact-btn" data-id="${item.id}">削除</button>
        </td>
      `;
      contactsTbody.appendChild(tr);
    });
};

/* ---- 検索/フィルタ ----------------------------------------------- */
searchBoxMini?.addEventListener("input", renderMini);
searchBox2?.addEventListener("input", renderFull);
monthFilter2?.addEventListener("change", renderFull);
sortOrder2?.addEventListener("change", renderFull);
clearFilters2?.addEventListener("click", () => {
  monthFilter2 && (monthFilter2.value = "");
  sortOrder2   && (sortOrder2.value   = "desc");
  searchBox2   && (searchBox2.value   = "");
  renderFull();
});
contactsSearchBox?.addEventListener("input", renderContacts);

/* ---- 折りたたみ ------------------------------------------------- */
const setCollapsedAll = (collapse) => { $$(".note-cell", $("#page-history-list")).forEach(td => collapse ? td.dataset.collapsed="1" : delete td.dataset.collapsed); };
collapseAllBtn?.addEventListener("click", () => {
  const collapsedExists = $$(".note-cell", $("#page-history-list")).some(td => td.dataset.collapsed === "1");
  if (collapsedExists) { setCollapsedAll(false); collapseAllBtn.textContent = "折りたたむ"; }
  else { setCollapsedAll(true); collapseAllBtn.textContent = "展開する"; }
});

/* ---- クイックメモ ------------------------------------------------ */
const loadMemo = () => { const t = localStorage.getItem("quickMemo") || ""; memoTextarea && (memoTextarea.value = t); memoStatusEl && (memoStatusEl.textContent = t ? "ローカル保存済み" : ""); };
const saveMemo = () => { localStorage.setItem("quickMemo", memoTextarea.value || ""); memoStatusEl && (memoStatusEl.textContent = "ローカル保存済み"); };
memoSaveBtn?.addEventListener("click", saveMemo);
memoClearBtn?.addEventListener("click", () => { memoTextarea.value = ""; saveMemo(); });

/* ---- 履歴：追加/更新/削除 --------------------------------------- */
historyForm?.addEventListener("submit", async (e) => {
  e.preventDefault();
  const item = { id: uuid(), createdAt: new Date().toISOString(), company: companyNameEl.value.trim(), media: mediaSelectEl.value.trim(), note: historyNoteEl.value.trim() };
  historyData.push(item); saveLocal(); renderAll(); historyForm.reset(); companyNameEl.focus();
  try {
    await apiPost({ action: "upsertHistory", item });
    const hist = await apiGet({ action: "history" });
    if (hist?.history) { historyData = hist.history; saveLocal(); renderAll(); }
    toast("GASへ保存しました。", "ok");
  } catch (err) { toast("GAS保存に失敗。公開設定とexec URLを確認。", "error"); console.error(err); }
});

const openDetail = (id) => {
  const it = historyData.find(x => x.id === id); if (!it) return;
  detailIdEl.value = it.id; detailCreatedEl.value = fmtDisplay(it.createdAt);
  detailCompanyEl.value = it.company; detailMediaEl.value = it.media; detailNoteEl.value = it.note;
  modal.classList.add("active"); modal.setAttribute("aria-hidden", "false");
};
const closeDetail = () => { modal.classList.remove("active"); modal.setAttribute("aria-hidden", "true"); };
modalCloseBtn?.addEventListener("click", closeDetail);
detailCancelBtn?.addEventListener("click", closeDetail);
window.addEventListener("click", (e)=>{ if (e.target === modal) closeDetail(); });
document.addEventListener("keydown", (e)=>{ if (e.key === "Escape" && modal?.classList.contains("active")) closeDetail(); });

detailForm?.addEventListener("submit", async (e) => {
  e.preventDefault();
  const id = detailIdEl.value;
  const idx = historyData.findIndex(x => x.id === id); if (idx === -1) return;
  historyData[idx].company = detailCompanyEl.value.trim();
  historyData[idx].media   = detailMediaEl.value.trim();
  historyData[idx].note    = detailNoteEl.value.trim();
  saveLocal(); renderAll(); closeDetail();
  try {
    await apiPost({ action: "upsertHistory", item: historyData[idx] });
    const hist = await apiGet({ action: "history" });
    if (hist?.history) { historyData = hist.history; saveLocal(); renderAll(); }
    toast("更新しました。", "ok");
  } catch (err) { toast("更新に失敗：GASを確認してください。", "error"); console.error(err); }
});

detailDeleteBtn?.addEventListener("click", async () => {
  const id = detailIdEl.value;
  if (!confirm("この履歴を削除しますか？")) return;
  historyData = historyData.filter(x => x.id !== id);
  saveLocal(); renderAll(); closeDetail();
  try {
    await apiPost({ action: "deleteHistory", id });
    const hist = await apiGet({ action: "history" });
    if (hist?.history) { historyData = hist.history; saveLocal(); renderAll(); }
    toast("削除しました。", "ok");
  } catch (err) { toast("削除に失敗：GASを確認してください。", "error"); console.error(err); }
});

/* ---- 連絡先：追加/更新/削除 ----------------------------------- */
contactsForm?.addEventListener("submit", (e) => {
  e.preventDefault();
  const item = {
    id: uuid(),
    company: contactCompanyEl.value.trim(),
    name: contactNameEl.value.trim(),
    email: contactEmailEl.value.trim(),
    tel: contactTelEl.value.trim(),
    memo: contactMemoEl.value.trim(),
  };
  contactsData.push(item);
  saveLocal();
  renderContacts();
  contactsForm.reset();
  contactCompanyEl.focus();
});

contactsTbody?.addEventListener("click", (e) => {
  const target = e.target;
  const id = target.dataset.id;
  if (!id) return;

  if (target.classList.contains("approach-contact-btn")) {
    const companyName = target.dataset.company;
    openApproachModal(companyName);
  } else if (target.classList.contains("delete-contact-btn")) {
    if (confirm("この連絡先を削除しますか？")) {
      contactsData = contactsData.filter(item => item.id !== id);
      saveLocal();
      renderContacts();
    }
  } else if (target.classList.contains("edit-contact-btn")) {
    const item = contactsData.find(item => item.id === id);
    if (!item) return;

    const tr = target.closest("tr");
    tr.innerHTML = `
      <td><input type="text" value="${escapeCsv(item.company)}" class="edit-company"></td>
      <td><input type="text" value="${escapeCsv(item.name)}" class="edit-name"></td>
      <td><input type="email" value="${escapeCsv(item.email)}" class="edit-email"></td>
      <td><input type="tel" value="${escapeCsv(item.tel)}" class="edit-tel"></td>
      <td><textarea class="edit-memo">${escapeCsv(item.memo)}</textarea></td>
      <td>
        <button class="primary-btn save-contact-btn" data-id="${id}">保存</button>
        <button class="secondary-btn cancel-edit-btn" data-id="${id}">キャンセル</button>
      </td>
    `;
  } else if (target.classList.contains("save-contact-btn")) {
    const tr = target.closest("tr");
    const updatedItem = {
      id,
      company: tr.querySelector(".edit-company").value.trim(),
      name: tr.querySelector(".edit-name").value.trim(),
      email: tr.querySelector(".edit-email").value.trim(),
      tel: tr.querySelector(".edit-tel").value.trim(),
      memo: tr.querySelector(".edit-memo").value.trim(),
    };
    const index = contactsData.findIndex(item => item.id === id);
    if (index !== -1) {
      contactsData[index] = updatedItem;
      saveLocal();
    }
    renderContacts();
  } else if (target.classList.contains("cancel-edit-btn")) {
    renderContacts();
  }
});

approachForm?.addEventListener("submit", async (e) => {
  e.preventDefault();
  const companyName = approachCompanyNameEl.value;
  if (!companyName) return;

  const item = {
    id: uuid(),
    createdAt: new Date().toISOString(),
    company: companyName,
    media: approachMediaSelectEl.value.trim(),
    note: approachNoteEl.value.trim(),
  };

  historyData.push(item);
  saveLocal();
  renderAll(); // メインの履歴テーブルも更新

  // モーダル内の履歴を再描画し、フォームをリセット
  openApproachModal(companyName);
  approachMediaSelectEl.value = "";
  approachNoteEl.value = "";
  approachNoteEl.focus();

  try {
    await apiPost({ action: "upsertHistory", item });
    toast("GASへ保存しました。", "ok");
  } catch (err) {
    toast("GAS保存に失敗。公開設定とexec URLを確認。", "error");
    console.error(err);
  }
});

/* ---- CSV（履歴/各リスト） -------------------------------------- */
const escapeCsv = (v) => { const s = String(v ?? ""); return (/[",\n]/.test(s)) ? `"${s.replace(/"/g, '""')}"` : s; };
const downloadCsv = (text, filename) => { const blob=new Blob([text],{type:"text/csv;charset=utf-8;"}); const url=URL.createObjectURL(blob); const a=document.createElement("a"); a.href=url; a.download=filename; a.click(); URL.revokeObjectURL(url); };

const exportHistoryCsv = () => {
  const headers = ["id","createdAt","company","media","note"];
  let csv = headers.join(",") + "\n";
  historyData.forEach(it => { csv += [escapeCsv(it.id),escapeCsv(it.createdAt),escapeCsv(it.company),escapeCsv(it.media),escapeCsv(it.note)].join(",") + "\n"; });
  downloadCsv(csv, "アプローチ履歴.csv");
};
historyExportBtn?.addEventListener("click", exportHistoryCsv);

const importHistoryCsv = (file) => {
  const reader = new FileReader();
  reader.onload = () => {
    let text = String(reader.result); if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    const lines = text.split(/\r?\n/).filter(l => l.trim().length > 0); if (!lines.length) return;
    const header = lines[0].split(",").map(h => h.trim().replace(/^\uFEFF/, "")); const hasId = header.includes("id");
    const dataLines = (header.length >= 4) ? lines.slice(1) : lines;
    const imported = dataLines.map(line => {
      const cols = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(v => v.replace(/^"|"$/g,"").replace(/""/g,'"'));
      let id, createdAt, company, media, note;
      if (hasId) { [id, createdAt, company, media, note] = cols; } else { [createdAt, company, media, note] = cols; id = uuid(); }
      if (!createdAt) createdAt = new Date().toISOString();
      return { id, createdAt, company, media, note };
    });
    if (confirm("現在の履歴を上書きしてインポートしますか？")) { historyData = imported; saveLocal(); renderAll(); }
  };
  reader.readAsText(file);
};
historyImportBtn?.addEventListener("click", () => historyImportFile.click());
historyImportFile?.addEventListener("change", (e)=>{ if (e.target.files?.length) importHistoryCsv(e.target.files[0]); });

const exportListCsv = (arr, title) => downloadCsv(arr.join("\n"), `${title}.csv`);
const importListCsv = (file, target) => {
  const reader = new FileReader();
  reader.onload = () => {
    let text = String(reader.result); if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    const arr = text.split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
    if (target === "kigyou") { kigyouWords = arr; kigyouTA.value = arr.join("\n"); }
    if (target === "mikomi") { mikomiWords = arr; mikomiTA.value = arr.join("\n"); }
    if (target === "botsu")  { botsuWords  = arr; botsuTA.value  = arr.join("\n"); }
    saveLocal(); refreshSets(); renderAll(); pushListsDebounced();
  };
  reader.readAsText(file);
};
kigyouExportBtn?.addEventListener("click", () => exportListCsv(kigyouWords, "企業リスト"));
mikomiExportBtn?.addEventListener("click", () => exportListCsv(mikomiWords, "見込みリスト"));
botsuExportBtn ?.addEventListener("click", () => exportListCsv(botsuWords,  "没リスト"));
kigyouImportBtn?.addEventListener("click", () => kigyouImportFile.click());
mikomiImportBtn?.addEventListener("click", () => mikomiImportFile.click());
botsuImportBtn ?.addEventListener("click", () => botsuImportFile.click());
kigyouImportFile?.addEventListener("change", (e)=> e.target.files?.length && importListCsv(e.target.files[0], "kigyou"));
mikomiImportFile?.addEventListener("change", (e)=> e.target.files?.length && importListCsv(e.target.files[0], "mikomi"));
botsuImportFile ?.addEventListener("change", (e)=> e.target.files?.length && importListCsv(e.target.files[0], "botsu"));

const exportContactsCsv = () => {
  const headers = ["id", "company", "name", "email", "tel", "memo"];
  let csv = headers.join(",") + "\n";
  contactsData.forEach(it => {
    csv += [
      escapeCsv(it.id),
      escapeCsv(it.company),
      escapeCsv(it.name),
      escapeCsv(it.email),
      escapeCsv(it.tel),
      escapeCsv(it.memo)
    ].join(",") + "\n";
  });
  downloadCsv(csv, "連絡先リスト.csv");
};
contactsExportBtn?.addEventListener("click", exportContactsCsv);

const importContactsCsv = (file) => {
  const reader = new FileReader();
  reader.onload = () => {
    let text = String(reader.result);
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    const lines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
    if (!lines.length) return;
    const header = lines[0].split(",").map(h => h.trim().replace(/^\uFEFF/, ""));
    const hasId = header.includes("id");
    const dataLines = (header.length >= 5) ? lines.slice(1) : lines;
    const imported = dataLines.map(line => {
      const cols = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(v => v.replace(/^"|"$/g, "").replace(/""/g, '"'));
      let id, company, name, email, tel, memo;
      if (hasId) {
        [id, company, name, email, tel, memo] = cols;
      } else {
        [company, name, email, tel, memo] = cols;
        id = uuid();
      }
      return { id, company, name, email, tel, memo };
    });
    if (confirm("現在の連絡先を上書きしてインポートしますか？")) {
      contactsData = imported;
      saveLocal();
      renderContacts();
    }
  };
  reader.readAsText(file);
};
contactsImportBtn?.addEventListener("click", () => contactsImportFile.click());
contactsImportFile?.addEventListener("change", (e) => { if (e.target.files?.length) importContactsCsv(e.target.files[0]); });

/* ---- リスト 双方向API ------------------------------------------ */
const pushListsToGAS = async () => {
  try{
    await apiPost({ action:"saveLists", kigyou:kigyouWords, mikomi:mikomiWords, botsu:botsuWords });
    toast("リストをGASへ保存しました。", "ok");
  }catch(err){ toast("リスト保存に失敗：GAS公開設定/URLを確認。", "error"); console.error(err); }
};
const pullListsFromGAS = async () => {
  try{
    const lists = await apiGet({ action:"lists" });
    kigyouWords = lists.kigyou || []; mikomiWords = lists.mikomi || []; botsuWords = lists.botsu || [];
    kigyouTA && (kigyouTA.value = kigyouWords.join("\n"));
    mikomiTA && (mikomiTA.value = mikomiWords.join("\n"));
    botsuTA  && (botsuTA.value  = botsuWords.join("\n"));
    refreshSets(); saveLocal(); renderAll();
    toast("GASからリストを再読み込みしました。", "ok");
  }catch(err){ toast("リストの再読み込みに失敗。", "error"); console.error(err); }
};
const pushListsDebounced = debounce(pushListsToGAS, 1200);

kigyouSaveBtn ?.addEventListener("click", pushListsToGAS);
mikomiSaveBtn ?.addEventListener("click", pushListsToGAS);
botsuSaveBtn  ?.addEventListener("click", pushListsToGAS);
kigyouReloadBtn?.addEventListener("click", pullListsFromGAS);
mikomiReloadBtn?.addEventListener("click", pullListsFromGAS);
botsuReloadBtn ?.addEventListener("click", pullListsFromGAS);

/* ---- URLクエリ（軽量同期） ------------------------------------- */
const applyQuery = () => { const sp=new URLSearchParams(location.search); monthFilter2 && sp.has("m") && (monthFilter2.value=sp.get("m")); sortOrder2 && sp.has("sort") && (sortOrder2.value=sp.get("sort")); searchBox2 && sp.has("q") && (searchBox2.value=sp.get("q")); };
const pushQuery = () => { const sp=new URLSearchParams(location.search);
  monthFilter2?.value ? sp.set("m",monthFilter2.value):sp.delete("m");
  sortOrder2?.value   ? sp.set("sort",sortOrder2.value):sp.delete("sort");
  searchBox2?.value   ? sp.set("q",searchBox2.value):sp.delete("q");
  history.replaceState(null,"",`${location.pathname}${sp.toString()?`?${sp.toString()}`:""}`);
};

/* ---- 初期化 ----------------------------------------------------- */
const init = async () => {
  kigyouTA && (kigyouTA.value = kigyouWords.join("\n"));
  mikomiTA && (mikomiTA.value = mikomiWords.join("\n"));
  botsuTA  && (botsuTA.value  = botsuWords.join("\n"));
  refreshSets(); loadMemo(); applyQuery(); renderAll();

  try {
    const [lists, hist] = await Promise.all([ apiGet({ action: "lists" }), apiGet({ action: "history" }) ]);
    if (lists) {
      kigyouWords = lists.kigyou || []; mikomiWords = lists.mikomi || []; botsuWords = lists.botsu || [];
      kigyouTA && (kigyouTA.value = kigyouWords.join("\n"));
      mikomiTA && (mikomiTA.value = mikomiWords.join("\n"));
      botsuTA  && (botsuTA.value  = botsuWords.join("\n"));
      refreshSets();
    }
    if (Array.isArray(hist?.history)) {
      historyData = hist.history.map(it => ({ id:it.id || "", createdAt: it.createdAt || it.date || "", company:it.company||"", media:it.media||"", note:it.note||"" }));
      const need = [];
      historyData.forEach(it => { if (!it.id) { it.id = uuid(); need.push({...it}); } });
      if (need.length) {
        try { await apiPost({ action: "bulkUpsertHistory", items: need }); }
        catch { for (const item of need) { try { await apiPost({ action:"upsertHistory", item }); } catch(_){} } }
      }
    }
    saveLocal(); renderAll();
  } catch (err) { console.warn("GAS読み込み失敗。ローカル使用:", err); toast("GAS読み込みに失敗（ローカル表示中）。", "error"); }
};

/* ---- リスト編集：ローカル反映＋自動GAS保存 ---------------------- */
kigyouTA?.addEventListener("input", () => { kigyouWords = kigyouTA.value.split(/\r?\n/).map(s=>s.trim()).filter(Boolean); saveLocal(); refreshSets(); renderAll(); pushListsDebounced(); });
mikomiTA?.addEventListener("input", () => { mikomiWords = mikomiTA.value.split(/\r?\n/).map(s=>s.trim()).filter(Boolean); saveLocal(); refreshSets(); renderAll(); pushListsDebounced(); });
botsuTA ?.addEventListener("input", () => { botsuWords  = botsuTA.value.split(/\r?\n/).map(s=>s.trim()).filter(Boolean); saveLocal(); refreshSets(); renderAll(); pushListsDebounced(); });

/* ---- DOMReady --------------------------------------------------- */
document.addEventListener("DOMContentLoaded", () => {
  showPage("#page-history-list");
  collapseAllBtn && (collapseAllBtn.textContent = "展開する");
  init();
  [monthFilter2, sortOrder2, searchBox2].forEach(el => { el?.addEventListener("change", pushQuery); el?.addEventListener("input", pushQuery); });
});
