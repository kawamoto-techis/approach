/* =========================================================
   script.js - Approach System 2.0
   ======================================================= */

/* ---- Config & Utils ------------------------------------- */
const GAS_BASE = "https://script.google.com/macros/s/AKfycbx1u3qfMh7GxCZ6jMa2h3m2Q296w9ZgV3V8pKuWdXyop4r8TVocDS4eAP_lUKP16Jnq6A/exec";

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const uuid = () => (crypto?.randomUUID?.() || `id-${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`);

const toast = (msg, type = "error") => {
  const el = document.createElement("div");
  el.style.cssText = `position:fixed; right:20px; top:20px; z-index:9999; padding:12px 20px; border-radius:8px; color:#fff; box-shadow:0 10px 30px rgba(0,0,0,0.2); font-weight:600; animation: fadeIn 0.3s ease;`;
  el.style.background = (type === "ok") ? "#00abae" : "#e74c3c";
  el.textContent = msg;
  document.body.appendChild(el);
  setTimeout(() => { el.style.opacity = "0"; setTimeout(() => el.remove(), 300); }, 3000);
};

const parseToDate = (v) => {
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
};
const fmtDisplay = (iso) => {
  const d = parseToDate(iso); if (!d) return iso || "";
  return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()} ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
};
const norm = (s) => String(s || "").toLowerCase().replace(/\s+/g, "").trim();
const escapeCsv = (v) => { const s = String(v ?? ""); return (/[",\n]/.test(s)) ? `"${s.replace(/"/g, '""')}"` : s; };

/* ---- State ---------------------------------------------- */
let historyData = JSON.parse(localStorage.getItem("historyData")) || [];
let contactsData = JSON.parse(localStorage.getItem("contactsData")) || [];
let chartInstance = null;

const saveLocal = () => {
  localStorage.setItem("historyData", JSON.stringify(historyData));
  localStorage.setItem("contactsData", JSON.stringify(contactsData));
  updateDashboard(); // データ変更時にダッシュボード更新

  // Sync to GAS (Debounced or immediate)
  // For simplicity, we sync contacts immediately on save for now, or we could do it separately.
  // Let's do it in the specific action handlers to avoid too many requests.
};

const syncContactsToGAS = async () => {
  try {
    await apiPost({ action: "saveContacts", contacts: contactsData });
    console.log("Contacts synced to GAS");
  } catch (e) {
    console.error("Failed to sync contacts:", e);
  }
};

/* ---- Navigation ----------------------------------------- */
const navBtns = $$(".nav-btn");
const pages = $$(".page-section");

const showPage = (targetId) => {
  pages.forEach(p => p.classList.remove("active"));
  $(targetId)?.classList.add("active");

  navBtns.forEach(b => b.classList.remove("active"));
  navBtns.find(b => b.dataset.target === targetId)?.classList.add("active");

  if (targetId === "#page-dashboard") updateDashboard();
  if (targetId === "#page-kanban") renderKanban();
};

navBtns.forEach(btn => btn.addEventListener("click", () => showPage(btn.dataset.target)));

/* ---- Dashboard ------------------------------------------ */
const updateDashboard = () => {
  // 1. Stats
  $("#stat-total-contacts").textContent = contactsData.length;
  $("#stat-user-count").textContent = contactsData.filter(c => c.status === "ユーザー").length;

  const now = new Date();
  const thisMonthKey = `${now.getFullYear()}-${now.getMonth()}`;
  const monthlyApproaches = historyData.filter(h => {
    const d = parseToDate(h.createdAt);
    return d && `${d.getFullYear()}-${d.getMonth()}` === thisMonthKey;
  }).length;
  $("#stat-monthly-approaches").textContent = monthlyApproaches;

  // 2. Chart
  const statusCounts = {
    "ユーザー": 0, "見込": 0, "没": 0, "未分類": 0
  };
  contactsData.forEach(c => {
    if (statusCounts[c.status] !== undefined) statusCounts[c.status]++;
    else statusCounts["未分類"]++;
  });

  const ctx = $("#statusChart").getContext("2d");
  if (chartInstance) chartInstance.destroy();

  chartInstance = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ["ユーザー", "見込", "没", "未分類"],
      datasets: [{
        data: [statusCounts["ユーザー"], statusCounts["見込"], statusCounts["没"], statusCounts["未分類"]],
        backgroundColor: ["#c6f6d5", "#fefcbf", "#fed7d7", "#e2e8f0"],
        borderColor: ["#9ae6b4", "#faf089", "#feb2b2", "#cbd5e0"],
        borderWidth: 1
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'right' } }
    }
  });

  // 3. Recent Activity
  const recentList = $("#recent-activity-list");
  recentList.innerHTML = "";
  const sortedHistory = [...historyData].sort((a, b) => (parseToDate(b.createdAt) || 0) - (parseToDate(a.createdAt) || 0)).slice(0, 5);

  sortedHistory.forEach(h => {
    const li = document.createElement("li");
    li.innerHTML = `
      <span class="activity-time">${fmtDisplay(h.createdAt)}</span>
      <span class="activity-company">${h.company}</span>
      <span style="color:#666; margin-left:8px;">${h.media}</span>
    `;
    recentList.appendChild(li);
  });
};

/* ---- Kanban Board --------------------------------------- */
const renderKanban = () => {
  const cols = {
    "": $("#kanban-none"),
    "見込": $("#kanban-mikomi"),
    "ユーザー": $("#kanban-user"),
    "没": $("#kanban-botsu")
  };
  const counts = {
    "": $("#count-none"),
    "見込": $("#count-mikomi"),
    "ユーザー": $("#count-user"),
    "没": $("#count-botsu")
  };

  // Clear
  Object.values(cols).forEach(el => el.innerHTML = "");

  // Distribute
  contactsData.forEach(c => {
    const st = (c.status && cols[c.status]) ? c.status : "";
    const card = createKanbanCard(c);
    cols[st].appendChild(card);
  });

  // Update Counts
  Object.keys(cols).forEach(k => {
    counts[k].textContent = cols[k].children.length;
  });
};

const createKanbanCard = (contact) => {
  const div = document.createElement("div");
  div.className = "kanban-card";
  div.draggable = true;
  div.dataset.id = contact.id;

  div.innerHTML = `
    <div class="k-card-title">${contact.company}</div>
    <div class="k-card-info">${contact.name || "担当なし"}</div>
    <div class="k-card-actions">
      <button class="k-btn-action approach-btn">記録</button>
    </div>
  `;

  div.querySelector(".approach-btn").addEventListener("click", (e) => {
    e.stopPropagation();
    openApproachModal(contact);
  });

  // Drag Events
  div.addEventListener("dragstart", (e) => {
    div.classList.add("dragging");
    e.dataTransfer.setData("text/plain", contact.id);
  });
  div.addEventListener("dragend", () => {
    div.classList.remove("dragging");
  });

  return div;
};

// Drop Zones
$$(".kanban-column").forEach(col => {
  col.addEventListener("dragover", (e) => {
    e.preventDefault(); // Allow drop
    const afterElement = getDragAfterElement(col, e.clientY);
    const draggable = $(".dragging");
    if (afterElement == null) {
      col.querySelector(".kanban-items").appendChild(draggable);
    } else {
      col.querySelector(".kanban-items").insertBefore(draggable, afterElement);
    }
  });

  col.addEventListener("drop", async (e) => {
    e.preventDefault();
    const id = e.dataTransfer.getData("text/plain");
    const newStatus = col.dataset.status;

    const contact = contactsData.find(c => c.id === id);
    if (contact && contact.status !== newStatus) {
      contact.status = newStatus;
      saveLocal();
      renderKanban(); // Re-render to update counts and sort if needed
      toast(`ステータスを「${newStatus || "未分類"}」に変更しました`, "ok");

      // Sync to GAS
      syncContactsToGAS();
    }
  });
});

const getDragAfterElement = (container, y) => {
  const draggableElements = [...container.querySelectorAll('.kanban-card:not(.dragging)')];
  return draggableElements.reduce((closest, child) => {
    const box = child.getBoundingClientRect();
    const offset = y - box.top - box.height / 2;
    if (offset < 0 && offset > closest.offset) {
      return { offset: offset, element: child };
    } else {
      return closest;
    }
  }, { offset: Number.NEGATIVE_INFINITY }).element;
};

/* ---- History & Contacts Logic (Existing adapted) -------- */

// History Render
const renderHistory = () => {
  const tbody = $("#history-table-body");
  const tbody2 = $("#history-table-body-2");
  if (!tbody) return;

  // Mini Table (Input Page)
  const miniQ = ($("#search-box")?.value || "").toLowerCase();
  tbody.innerHTML = "";
  historyData
    .filter(it => it.company.toLowerCase().includes(miniQ) || it.note.toLowerCase().includes(miniQ))
    .sort((a, b) => (parseToDate(b.createdAt) || 0) - (parseToDate(a.createdAt) || 0))
    .slice(0, 20)
    .forEach(item => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${fmtDisplay(item.createdAt)}</td><td>${item.company}</td><td>${item.note}</td>`;
      tr.style.cursor = "pointer";
      tr.addEventListener("click", () => openDetail(item.id));
      tbody.appendChild(tr);
    });

  // Full Table (List Page)
  if (!tbody2) return;
  const fullQ = ($("#search-box-2")?.value || "").toLowerCase();
  const ym = $("#month-filter-2")?.value;
  const sortAsc = $("#sort-order-2")?.value === "asc";

  let list = historyData.slice();
  if (ym) list = list.filter(it => it.createdAt.startsWith(ym));
  if (fullQ) list = list.filter(it => it.company.toLowerCase().includes(fullQ) || it.note.toLowerCase().includes(fullQ));

  list.sort((a, b) => {
    const diff = (parseToDate(a.createdAt) || 0) - (parseToDate(b.createdAt) || 0);
    return sortAsc ? diff : -diff;
  });

  tbody2.innerHTML = "";
  list.forEach(item => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${fmtDisplay(item.createdAt)}</td><td>${item.company}</td><td>${item.media}</td><td>${item.note}</td>`;
    tr.style.cursor = "pointer";
    tr.addEventListener("click", () => openDetail(item.id));
    tbody2.appendChild(tr);
  });
};

// Contacts Render
let contactFilter = "all";
const renderContacts = () => {
  const tbody = $("#contacts-table-body");
  if (!tbody) return;

  const q = ($("#contacts-search-box")?.value || "").toLowerCase();
  tbody.innerHTML = "";

  let list = contactsData;
  if (contactFilter === "user") list = list.filter(c => c.status === "ユーザー");
  if (contactFilter === "prospect") list = list.filter(c => c.status === "見込");
  if (contactFilter === "none") list = list.filter(c => !c.status);

  list = list.filter(c => c.company.toLowerCase().includes(q) || c.name.toLowerCase().includes(q));

  list.forEach(c => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${c.company}</td>
      <td>${c.name}</td>
      <td>${c.tel}</td>
      <td><span class="status-badge ${getStatusBadgeClass(c.status)}">${c.status || "-"}</span></td>
      <td>
        <button class="icon-btn approach-contact-btn" title="アプローチ">✏️</button>
        <button class="icon-btn edit-contact-btn" title="編集">📝</button>
        <button class="icon-btn delete-contact-btn" title="削除" style="color:var(--danger)">🗑️</button>
      </td>
    `;

    tr.querySelector(".approach-contact-btn").addEventListener("click", () => openApproachModal(c));
    tr.querySelector(".edit-contact-btn").addEventListener("click", () => openContactEdit(c));
    tr.querySelector(".delete-contact-btn").addEventListener("click", () => deleteContact(c.id));

    tbody.appendChild(tr);
  });
};

const getStatusBadgeClass = (status) => {
  if (status === "ユーザー") return "status-user";
  if (status === "見込") return "status-mikomi";
  if (status === "没") return "status-botsu";
  return "";
};

/* ---- Modals & Actions ----------------------------------- */
// Detail Modal
const openDetail = (id) => {
  const it = historyData.find(x => x.id === id); if (!it) return;
  $("#detail-id").value = it.id;
  $("#detail-created").value = fmtDisplay(it.createdAt);
  $("#detail-company").value = it.company;
  $("#detail-media").value = it.media;
  $("#detail-note").value = it.note;
  $("#detail-modal").classList.add("active");
};
$("#detail-form")?.addEventListener("submit", (e) => {
  e.preventDefault();
  const id = $("#detail-id").value;
  const idx = historyData.findIndex(x => x.id === id);
  if (idx >= 0) {
    historyData[idx].company = $("#detail-company").value;
    historyData[idx].media = $("#detail-media").value;
    historyData[idx].note = $("#detail-note").value;
    saveLocal(); renderHistory(); $("#detail-modal").classList.remove("active");
    toast("更新しました", "ok");
  }
});
$("#detail-delete")?.addEventListener("click", () => {
  if (confirm("削除しますか？")) {
    historyData = historyData.filter(x => x.id !== $("#detail-id").value);
    saveLocal(); renderHistory(); $("#detail-modal").classList.remove("active");
    toast("削除しました", "ok");
  }
});

// Approach Modal
const openApproachModal = (contact) => {
  $("#approach-modal-title").textContent = `${contact.company} へのアプローチ`;
  $("#approach-company-name").value = contact.company;

  $("#approach-contact-details").innerHTML = `
    <strong>担当:</strong> ${contact.name}<br>
    <strong>Tel:</strong> ${contact.tel}<br>
    <strong>Email:</strong> ${contact.email}<br>
    <strong>Memo:</strong> ${contact.memo}
  `;

  const hist = historyData.filter(h => h.company === contact.company).sort((a, b) => (parseToDate(b.createdAt) || 0) - (parseToDate(a.createdAt) || 0));
  const tbody = $("#approach-history-table-body");
  tbody.innerHTML = "";
  hist.forEach(h => {
    tbody.innerHTML += `<tr><td>${fmtDisplay(h.createdAt)}</td><td>${h.media}</td><td>${h.note}</td></tr>`;
  });

  $("#approach-modal").classList.add("active");
};

$("#approach-form")?.addEventListener("submit", async (e) => {
  e.preventDefault();
  const item = {
    id: uuid(),
    createdAt: new Date().toISOString(),
    company: $("#approach-company-name").value,
    media: $("#approach-media-select").value,
    note: $("#approach-note").value
  };
  historyData.push(item);
  saveLocal();
  renderHistory();
  updateDashboard();
  $("#approach-modal").classList.remove("active");
  $("#approach-form").reset();
  toast("記録しました", "ok");

  // GAS Sync
  try { await apiPost({ action: "upsertHistory", item }); } catch (e) { console.error(e); }
});

// Contact Add/Edit
const openContactEdit = (contact) => {
  // Reuse the add modal for editing? Or create a separate one.
  // For simplicity in this version, let's use the same modal but populate it.
  // Ideally we should have a hidden ID field.
  // Adding a hidden ID field to the contact form dynamically.
  let idInput = $("#contact-id-hidden");
  if (!idInput) {
    idInput = document.createElement("input");
    idInput.type = "hidden";
    idInput.id = "contact-id-hidden";
    $("#contacts-form").appendChild(idInput);
  }

  idInput.value = contact.id;
  $("#contact-company").value = contact.company;
  $("#contact-name").value = contact.name;
  $("#contact-email").value = contact.email;
  $("#contact-tel").value = contact.tel;
  $("#contact-memo").value = contact.memo;
  $("#contact-status-init").value = contact.status;

  $("#contact-modal").classList.add("active");
};

$("#btn-add-contact-modal")?.addEventListener("click", () => {
  $("#contacts-form").reset();
  let idInput = $("#contact-id-hidden");
  if (idInput) idInput.value = ""; // Clear ID for new
  $("#contact-modal").classList.add("active");
});

$("#contacts-form")?.addEventListener("submit", (e) => {
  e.preventDefault();
  const idInput = $("#contact-id-hidden");
  const id = idInput?.value;

  const newItem = {
    id: id || uuid(),
    company: $("#contact-company").value,
    name: $("#contact-name").value,
    email: $("#contact-email").value,
    tel: $("#contact-tel").value,
    memo: $("#contact-memo").value,
    status: $("#contact-status-init").value
  };

  if (id) {
    const idx = contactsData.findIndex(c => c.id === id);
    if (idx >= 0) contactsData[idx] = newItem;
  } else {
    contactsData.push(newItem);
  }

  saveLocal();
  renderContacts();
  renderKanban();
  updateDashboard();
  $("#contact-modal").classList.remove("active");
  toast(id ? "更新しました" : "追加しました", "ok");
  syncContactsToGAS();
});

const deleteContact = (id) => {
  if (confirm("本当に削除しますか？")) {
    contactsData = contactsData.filter(c => c.id !== id);
    saveLocal();
    renderContacts();
    renderKanban();
    updateDashboard();
    updateDashboard();
    toast("削除しました", "ok");
    syncContactsToGAS();
  }
};

/* ---- Event Listeners (Global) --------------------------- */
$$(".close-btn").forEach(b => b.addEventListener("click", () => $$(".modal").forEach(m => m.classList.remove("active"))));
window.addEventListener("click", (e) => { if (e.target.classList.contains("modal")) e.target.classList.remove("active"); });

$("#history-form")?.addEventListener("submit", async (e) => {
  e.preventDefault();
  const item = {
    id: uuid(),
    createdAt: new Date().toISOString(),
    company: $("#company-name").value,
    media: $("#media-select").value,
    note: $("#history-note").value
  };
  historyData.push(item);
  saveLocal();
  renderHistory();
  updateDashboard();
  $("#history-form").reset();
  toast("追加しました", "ok");
  try { await apiPost({ action: "upsertHistory", item }); } catch (e) { console.error(e); }
});

// Filters
$("#search-box")?.addEventListener("input", renderHistory);
$("#search-box-2")?.addEventListener("input", renderHistory);
$("#month-filter-2")?.addEventListener("change", renderHistory);
$("#sort-order-2")?.addEventListener("change", renderHistory);
$("#contacts-search-box")?.addEventListener("input", renderContacts);

$$(".filter-chip").forEach(btn => {
  btn.addEventListener("click", () => {
    $$(".filter-chip").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    if (btn.id === "filter-all-btn") contactFilter = "all";
    if (btn.id === "filter-user-btn") contactFilter = "user";
    if (btn.id === "filter-prospect-btn") contactFilter = "prospect";
    if (btn.id === "filter-none-btn") contactFilter = "none";
    renderContacts();
  });
});

/* ---- Import / Export ------------------------------------ */
// Export Contacts
$("#contacts-export-btn")?.addEventListener("click", () => {
  if (contactsData.length === 0) return toast("エクスポートするデータがありません", "error");

  const header = ["ID", "Company", "Name", "Email", "Tel", "Memo", "Status"];
  const rows = contactsData.map(c => [
    escapeCsv(c.id),
    escapeCsv(c.company),
    escapeCsv(c.name),
    escapeCsv(c.email),
    escapeCsv(c.tel),
    escapeCsv(c.memo),
    escapeCsv(c.status)
  ]);

  const csvContent = [header.join(","), ...rows.map(r => r.join(","))].join("\n");
  const blob = new Blob([new Uint8Array([0xEF, 0xBB, 0xBF]), csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.setAttribute("href", url);
  link.setAttribute("download", `contacts_${new Date().toISOString().slice(0, 10)}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
});

// Import Contacts
$("#contacts-import-btn")?.addEventListener("click", () => $("#contacts-import-file").click());
$("#contacts-import-file")?.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    const text = evt.target.result;
    const lines = text.split(/\r\n|\n/).filter(line => line.trim() !== "");
    if (lines.length < 2) return toast("有効なデータがありません", "error");

    // Simple CSV parser (assumes standard CSV)
    // Skip header
    let addedCount = 0;
    for (let i = 1; i < lines.length; i++) {
      // Handle quotes if needed, but for now simple split if simple CSV
      // For better CSV parsing, we'd need a proper parser. 
      // Let's assume simple comma separation for MVP or use a regex for quoted fields.
      const row = parseCSVLine(lines[i]);
      if (row.length < 2) continue; // At least Company needed

      // Map columns: ID, Company, Name, Email, Tel, Memo, Status
      // If ID exists and matches, update? Or just add new?
      // Let's assume import adds new or updates by ID if present.

      const id = row[0] || uuid();
      const newItem = {
        id: id,
        company: row[1] || "",
        name: row[2] || "",
        email: row[3] || "",
        tel: row[4] || "",
        memo: row[5] || "",
        status: row[6] || ""
      };

      const idx = contactsData.findIndex(c => c.id === newItem.id);
      if (idx >= 0) {
        contactsData[idx] = newItem;
      } else {
        contactsData.push(newItem);
      }
      addedCount++;
    }

    saveLocal();
    renderContacts();
    renderKanban();
    updateDashboard();
    syncContactsToGAS();
    toast(`${addedCount}件インポートしました`, "ok");
    e.target.value = ""; // Reset
  };
  reader.readAsText(file);
});

const parseCSVLine = (text) => {
  // Simple regex to handle quoted fields
  const re_value = /(?!\s*$)\s*(?:'([^']*)'|"([^"]*)"|([^,'"]*))\s*(?:,|$)/g;
  const a = [];
  text.replace(re_value, (m0, m1, m2, m3) => {
    if (m1 !== undefined) a.push(m1.replace(/''/g, "'"));
    else if (m2 !== undefined) a.push(m2.replace(/""/g, '"'));
    else if (m3 !== undefined) a.push(m3);
    return "";
  });
  // Handle last empty field if comma is at end
  if (/,\s*$/.test(text)) a.push("");
  return a;
};

/* ---- GAS API -------------------------------------------- */
const apiPost = async (body) => {
  try {
    const res = await fetch(GAS_BASE, {
      method: "POST",
      body: JSON.stringify(body),
      // mode: "no-cors" // REMOVED to allow reading response
    });
    // If we want to read response, we need the GAS script to return JSON and handle CORS (or use simple POST)
    // With 'no-cors' removed, if GAS doesn't send CORS headers, this might fail in browser.
    // However, GAS ContentService usually works if we follow redirects.
    // Let's try to parse if possible, or just ignore if it fails.
    const text = await res.text();
    return text ? JSON.parse(text) : {};
  } catch (e) {
    console.warn("API Request failed (might be CORS if GAS not updated):", e);
    return null;
  }
};

const fetchDataFromGAS = async () => {
  toast("データを同期中...", "ok");
  try {
    // Modified to use GET, as POST {action:'getData'} is failing.
    // Added timestamp to prevent caching.
    const url = `${GAS_BASE}?t=${Date.now()}`;
    const res = await fetch(url);
    const json = await res.json();

    // Handle potential wrapper: { ok: true, history: [...] } or { data: { ... } }
    const data = json.data || json;

    if (data && (data.history || data.contacts)) {
      historyData = data.history || [];
      contactsData = data.contacts || [];
      saveLocal();
      renderHistory();
      renderContacts();
      renderKanban();
      updateDashboard();
      toast("データ同期完了", "ok");
    } else {
      console.log("No data returned from GAS or invalid format", data);
      toast("データ形式が不正です", "error");
    }
  } catch (e) {
    console.error("Fetch error:", e);
    toast("同期に失敗しました", "error");
  }
};

/* ---- Init ----------------------------------------------- */
document.addEventListener("DOMContentLoaded", () => {
  renderHistory();
  renderContacts();
  renderKanban();
  updateDashboard();

  // Fetch latest data
  fetchDataFromGAS();

  // Quick Memo
  $("#quick-memo").value = localStorage.getItem("quickMemo") || "";
  $("#quick-memo-save").addEventListener("click", () => {
    localStorage.setItem("quickMemo", $("#quick-memo").value);
    toast("メモを保存しました", "ok");
  });
});
