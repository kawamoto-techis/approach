/* =========================================================
   script.js - Approach System 3.0
   ======================================================= */

/* ---- Config ---------------------------------------------- */
const GAS_BASE = "https://script.google.com/macros/s/AKfycbx1u3qfMh7GxCZ6jMa2h3m2Q296w9ZgV3V8pKuWdXyop4r8TVocDS4eAP_lUKP16Jnq6A/exec";

/* ---- Utils ----------------------------------------------- */
const $  = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const uuid = () => crypto?.randomUUID?.() || `id-${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
const norm = s => String(s || '').toLowerCase().replace(/\s+/g, '').trim();

const parseDate = v => {
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
};

const fmtDate = iso => {
  const d = parseDate(iso);
  if (!d) return iso || '';
  return `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
};

const fmtDateShort = iso => {
  const d = parseDate(iso);
  if (!d) return iso || '';
  return `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}`;
};

const escapeCsv = v => {
  const s = String(v ?? '');
  return /[",\n]/.test(s) ? `"${s.replace(/"/g,'""')}"` : s;
};

let _toastTimer;
const toast = (msg, type = 'info') => {
  clearTimeout(_toastTimer);
  const old = $('.toast');
  if (old) old.remove();
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.textContent = msg;
  document.body.appendChild(el);
  _toastTimer = setTimeout(() => { el.style.opacity = '0'; el.style.transition = 'opacity 0.3s'; setTimeout(() => el.remove(), 300); }, 3000);
};

/* ---- State ----------------------------------------------- */
let historyData  = JSON.parse(localStorage.getItem('historyData'))  || [];
let contactsData = JSON.parse(localStorage.getItem('contactsData')) || [];
let chartInst    = null;
let contactFilter = 'all';
let historyCollapsed = true;

const saveLocal = () => {
  localStorage.setItem('historyData',  JSON.stringify(historyData));
  localStorage.setItem('contactsData', JSON.stringify(contactsData));
  updateNavBadges();
  updateDashboard();
  updateCompanyDatalist();
};

/* ---- Navigation ------------------------------------------ */
const PAGE_META = {
  'page-dashboard':    { title: 'ダッシュボード',  sub: '現在の状況サマリー' },
  'page-contacts':     { title: '連絡先',          sub: '企業リスト管理' },
  'page-kanban':       { title: 'カンバンボード',  sub: 'ステータス管理' },
  'page-history':      { title: '履歴入力',        sub: 'アプローチ記録' },
  'page-history-list': { title: '履歴一覧',        sub: '過去のアプローチ' }
};

const showPage = id => {
  $$('.page-section').forEach(p => p.classList.remove('active'));
  $(`#${id}`)?.classList.add('active');
  $$('.nav-item').forEach(b => b.classList.remove('active'));
  $(`#nav-${id.replace('page-', '')}`)?.classList.add('active');
  const meta = PAGE_META[id] || {};
  if ($('#topbar-title'))  $('#topbar-title').textContent  = meta.title || '';
  if ($('#topbar-sub'))    $('#topbar-sub').textContent    = meta.sub || '';

  if (id === 'page-dashboard')    updateDashboard();
  if (id === 'page-kanban')       renderKanban();
  if (id === 'page-contacts')     renderContacts();
  if (id === 'page-history')      renderHistory();
  if (id === 'page-history-list') renderHistoryList();
};

$$('.nav-item').forEach(btn => btn.addEventListener('click', () => showPage(btn.dataset.page)));

const updateNavBadges = () => {
  const b = $('#nav-badge-contacts'); if (b) b.textContent = contactsData.length;
  const h = $('#nav-badge-history');  if (h) h.textContent = historyData.length;
};

/* ---- Company Datalist ------------------------------------ */
const updateCompanyDatalist = () => {
  const dl = $('#company-datalist');
  if (!dl) return;
  const companies = [...new Set([
    ...contactsData.map(c => c.company),
    ...historyData.map(h => h.company)
  ])].sort();
  dl.innerHTML = companies.map(c => `<option value="${c}">`).join('');
};

/* ---- Dashboard ------------------------------------------- */
const updateDashboard = () => {
  const now = new Date();
  const todayKey  = `${now.getFullYear()}-${now.getMonth()}-${now.getDate()}`;
  const monthKey  = `${now.getFullYear()}-${now.getMonth()}`;
  const lastMonth = new Date(now.getFullYear(), now.getMonth()-1, 1);
  const lastKey   = `${lastMonth.getFullYear()}-${lastMonth.getMonth()}`;

  const callToday = historyData.filter(h => {
    const d = parseDate(h.createdAt);
    return d && `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}` === todayKey;
  }).length;

  const callMonth = historyData.filter(h => {
    const d = parseDate(h.createdAt);
    return d && `${d.getFullYear()}-${d.getMonth()}` === monthKey;
  }).length;

  const callLast = historyData.filter(h => {
    const d = parseDate(h.createdAt);
    return d && `${d.getFullYear()}-${d.getMonth()}` === lastKey;
  }).length;

  const setEl = (id, v) => { const el = $(id); if (el) el.textContent = v; };
  setEl('#stat-call-today',   callToday);
  setEl('#stat-call-month',   callMonth);
  setEl('#stat-total-contacts', contactsData.length);
  setEl('#stat-users', contactsData.filter(c => c.status === 'ユーザー').length);
  setEl('#mini-call-today',  callToday);
  setEl('#mini-call-month',  callMonth);
  setEl('#mini-call-last',   callLast);
  setEl('#mini-call-total',  historyData.length);

  // Chart
  const counts = { 'ユーザー': 0, '見込': 0, '没': 0, '未分類': 0 };
  contactsData.forEach(c => {
    const k = counts[c.status] !== undefined ? c.status : '未分類';
    counts[k]++;
  });
  const ctx = $('#statusChart')?.getContext('2d');
  if (ctx) {
    if (chartInst) chartInst.destroy();
    chartInst = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: ['ユーザー', '見込', '没', '未分類'],
        datasets: [{
          data: [counts['ユーザー'], counts['見込'], counts['没'], counts['未分類']],
          backgroundColor: ['rgba(0,212,160,0.8)', 'rgba(245,184,0,0.8)', 'rgba(239,68,68,0.8)', 'rgba(100,116,139,0.5)'],
          borderColor: ['rgba(0,212,160,1)', 'rgba(245,184,0,1)', 'rgba(239,68,68,1)', 'rgba(100,116,139,1)'],
          borderWidth: 1
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'right',
            labels: { color: '#8899a8', font: { size: 12 }, boxWidth: 12, padding: 12 }
          }
        },
        cutout: '65%'
      }
    });
  }

  // Recent activity
  const list = $('#recent-activity-list');
  if (list) {
    list.innerHTML = '';
    [...historyData]
      .sort((a,b) => (parseDate(b.createdAt)||0) - (parseDate(a.createdAt)||0))
      .slice(0, 10)
      .forEach(h => {
        const li = document.createElement('li');
        li.className = 'activity-item';
        li.innerHTML = `
          <div class="activity-dot"></div>
          <div style="flex:1; min-width:0;">
            <div style="display:flex; justify-content:space-between; align-items:center; gap:8px;">
              <div class="activity-company">${h.company}</div>
              <div class="activity-time">${fmtDateShort(h.createdAt)}</div>
            </div>
            <div class="activity-note">${h.note}</div>
          </div>
        `;
        list.appendChild(li);
      });
    if (list.children.length === 0) {
      list.innerHTML = '<li class="empty-state" style="padding:24px;"><div class="empty-title">履歴がありません</div></li>';
    }
  }

  // Sync memos
  syncMemos();
};

/* ---- Memo ------------------------------------------------ */
const syncMemos = () => {
  const saved = localStorage.getItem('quickMemo') || '';
  const m1 = $('#quick-memo');
  const m2 = $('#quick-memo-2');
  if (m1 && m1 !== document.activeElement) m1.value = saved;
  if (m2 && m2 !== document.activeElement) m2.value = saved;

  const ts = localStorage.getItem('quickMemoTime');
  const el = $('#memo-saved-time');
  if (el && ts) el.textContent = `保存済: ${new Date(ts).toLocaleTimeString('ja-JP')}`;
};

const saveMemo = (val) => {
  localStorage.setItem('quickMemo', val);
  localStorage.setItem('quickMemoTime', new Date().toISOString());
  syncMemos();
  toast('メモを保存しました', 'ok');
};

$('#quick-memo-save')?.addEventListener('click', () => saveMemo($('#quick-memo').value));
$('#quick-memo-clear')?.addEventListener('click', () => { $('#quick-memo').value = ''; saveMemo(''); });
$('#quick-memo-2-save')?.addEventListener('click', () => saveMemo($('#quick-memo-2').value));
$('#quick-memo-2-clear')?.addEventListener('click', () => { $('#quick-memo-2').value = ''; saveMemo(''); });

/* ---- Contacts -------------------------------------------- */
const statusBadge = status => {
  if (status === 'ユーザー') return `<span class="badge badge-user">ユーザー</span>`;
  if (status === '見込')     return `<span class="badge badge-mikomi">見込</span>`;
  if (status === '没')       return `<span class="badge badge-botsu">没</span>`;
  return `<span class="badge badge-none">未分類</span>`;
};

const renderContacts = () => {
  const tbody = $('#contacts-table-body');
  if (!tbody) return;
  const q = ($('#contacts-search-box')?.value || '').toLowerCase();
  tbody.innerHTML = '';

  let list = contactsData.slice();
  if (contactFilter === 'user')    list = list.filter(c => c.status === 'ユーザー');
  if (contactFilter === 'prospect')list = list.filter(c => c.status === '見込');
  if (contactFilter === 'botsu')   list = list.filter(c => c.status === '没');
  if (contactFilter === 'none')    list = list.filter(c => !c.status);
  if (q) list = list.filter(c => norm(c.company).includes(norm(q)) || norm(c.name).includes(norm(q)));

  if (list.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6"><div class="empty-state"><div class="empty-icon">◉</div><div class="empty-title">データがありません</div><div class="empty-desc">連絡先を追加してください</div></div></td></tr>`;
    return;
  }

  list.forEach(c => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><span style="font-weight:600;">${c.company}</span></td>
      <td class="td-muted">${c.name || '—'}</td>
      <td class="td-muted">${c.tel || '—'}</td>
      <td>${statusBadge(c.status)}</td>
      <td class="note-cell">${c.memo || '—'}</td>
      <td>
        <div class="flex gap-4">
          <button class="btn btn-ghost btn-xs approach-btn" data-id="${c.id}" title="アプローチ記録">✎</button>
          <button class="btn btn-ghost btn-xs edit-btn" data-id="${c.id}" title="編集">✏</button>
          <button class="btn btn-danger btn-xs del-btn" data-id="${c.id}" title="削除">✕</button>
        </div>
      </td>
    `;
    tr.querySelector('.approach-btn').addEventListener('click', () => openApproachModal(c));
    tr.querySelector('.edit-btn').addEventListener('click', () => loadContactEdit(c));
    tr.querySelector('.del-btn').addEventListener('click', () => deleteContact(c.id));
    tbody.appendChild(tr);
  });
};

// contact filter tabs
$('#contact-filter-tabs')?.addEventListener('click', e => {
  const btn = e.target.closest('.filter-tab');
  if (!btn) return;
  $$('#contact-filter-tabs .filter-tab').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  contactFilter = btn.dataset.filter;
  renderContacts();
});

$('#contacts-search-box')?.addEventListener('input', renderContacts);

const loadContactEdit = c => {
  $('#contact-edit-id').value = c.id;
  $('#contact-company').value = c.company;
  $('#contact-status').value  = c.status || '';
  $('#contact-name').value    = c.name || '';
  $('#contact-tel').value     = c.tel || '';
  $('#contact-email').value   = c.email || '';
  $('#contact-address').value = c.address || '';
  $('#contact-memo').value    = c.memo || '';
  $('#contact-form-title').textContent = '編集中';
  $('#contact-company').focus();
};

$('#contact-reset-btn')?.addEventListener('click', () => {
  ['contact-edit-id','contact-company','contact-status','contact-name','contact-tel','contact-email','contact-address','contact-memo'].forEach(id => {
    const el = $(`#${id}`); if (el) el.value = el.tagName === 'SELECT' ? '' : '';
  });
  $('#contact-form-title').textContent = '新規登録';
});

$('#contact-submit-btn')?.addEventListener('click', async () => {
  const company = $('#contact-company').value.trim();
  if (!company) { toast('企業名は必須です', 'error'); return; }

  const id = $('#contact-edit-id').value || uuid();
  const newItem = {
    id,
    company,
    status:  $('#contact-status').value,
    name:    $('#contact-name').value.trim(),
    tel:     $('#contact-tel').value.trim(),
    email:   $('#contact-email').value.trim(),
    address: $('#contact-address').value.trim(),
    memo:    $('#contact-memo').value.trim()
  };

  const idx = contactsData.findIndex(c => c.id === id);
  if (idx >= 0) { contactsData[idx] = newItem; } else { contactsData.push(newItem); }

  saveLocal();
  renderContacts();
  renderKanban();
  toast(idx >= 0 ? '更新しました' : '登録しました', 'ok');
  $('#contact-reset-btn').click();

  await syncContactsToGAS();
});

const deleteContact = async id => {
  if (!confirm('この連絡先を削除しますか？')) return;
  contactsData = contactsData.filter(c => c.id !== id);
  saveLocal();
  renderContacts();
  renderKanban();
  toast('削除しました', 'ok');
  await syncContactsToGAS();
};

/* ---- Approach Modal -------------------------------------- */
const openApproachModal = c => {
  $('#approach-modal-title').textContent = `${c.company} へのアプローチ`;
  $('#approach-company-name').value = c.company;
  $('#approach-contact-details').innerHTML = `
    <strong>企業</strong> ${c.company}<br>
    <strong>担当</strong> ${c.name || '—'}<br>
    <strong>Tel</strong> ${c.tel || '—'}<br>
    <strong>Email</strong> ${c.email || '—'}
  `;
  const hist = historyData.filter(h => h.company === c.company).sort((a,b) => (parseDate(b.createdAt)||0) - (parseDate(a.createdAt)||0));
  const tbody = $('#approach-history-table-body');
  tbody.innerHTML = hist.length === 0
    ? `<tr><td colspan="3"><div class="empty-state" style="padding:20px;"><div class="empty-title">履歴なし</div></div></td></tr>`
    : hist.map(h => `<tr><td class="td-muted" style="font-size:11px;">${fmtDate(h.createdAt)}</td><td class="td-muted">${h.media}</td><td style="font-size:12px;">${h.note}</td></tr>`).join('');
  $('#approach-modal').classList.add('active');
};

$('#approach-submit-btn')?.addEventListener('click', async () => {
  const company = $('#approach-company-name').value;
  const media   = $('#approach-media-select').value;
  const note    = $('#approach-note').value.trim();
  if (!media || !note) { toast('媒体と内容を入力してください', 'error'); return; }

  const item = { id: uuid(), createdAt: new Date().toISOString(), company, media, note };
  historyData.push(item);
  saveLocal();
  renderHistory();
  renderHistoryList();
  openApproachModal(contactsData.find(c => c.company === company) || { company, name: '', tel: '', email: '' });
  $('#approach-note').value = '';
  toast('記録しました', 'ok');
  try { await apiPost({ action: 'upsertHistory', item }); } catch(e) { console.error(e); }
});

/* ---- Kanban ---------------------------------------------- */
const renderKanban = () => {
  const cols = { '': $('#kanban-none'), '見込': $('#kanban-mikomi'), 'ユーザー': $('#kanban-user'), '没': $('#kanban-botsu') };
  const counts = { '': $('#count-none'), '見込': $('#count-mikomi'), 'ユーザー': $('#count-user'), '没': $('#count-botsu') };
  Object.values(cols).forEach(c => { if (c) c.innerHTML = ''; });

  contactsData.forEach(c => {
    const st = cols[c.status] ? c.status : '';
    const card = document.createElement('div');
    card.className = 'kanban-card';
    card.draggable = true;
    card.dataset.id = c.id;
    card.innerHTML = `
      <div class="k-company">${c.company}</div>
      <div class="k-name">${c.name || '担当未登録'}</div>
      <div class="k-actions"><button class="btn btn-ghost btn-xs approach-k-btn">✎ 記録</button></div>
    `;
    card.querySelector('.approach-k-btn').addEventListener('click', e => { e.stopPropagation(); openApproachModal(c); });
    card.addEventListener('dragstart', e => { card.classList.add('dragging'); e.dataTransfer.setData('text/plain', c.id); });
    card.addEventListener('dragend', () => card.classList.remove('dragging'));
    cols[st].appendChild(card);
  });

  Object.keys(cols).forEach(k => { const el = counts[k]; if (el) el.textContent = cols[k]?.children.length || 0; });
};

$$('.kanban-column').forEach(col => {
  col.addEventListener('dragover', e => {
    e.preventDefault();
    const after = getDragAfterEl(col, e.clientY);
    const drag = $('.dragging');
    if (!drag) return;
    if (after == null) col.querySelector('.kanban-items').appendChild(drag);
    else col.querySelector('.kanban-items').insertBefore(drag, after);
  });

  col.addEventListener('drop', async e => {
    e.preventDefault();
    const id = e.dataTransfer.getData('text/plain');
    const newStatus = col.dataset.status;
    const c = contactsData.find(x => x.id === id);
    if (c && c.status !== newStatus) {
      c.status = newStatus;
      saveLocal();
      renderKanban();
      renderContacts();
      toast(`「${newStatus || '未分類'}」に移動しました`, 'ok');
      await syncContactsToGAS();
    }
  });
});

const getDragAfterEl = (container, y) => {
  const els = [...container.querySelectorAll('.kanban-card:not(.dragging)')];
  return els.reduce((closest, el) => {
    const box = el.getBoundingClientRect();
    const offset = y - box.top - box.height / 2;
    return offset < 0 && offset > closest.offset ? { offset, element: el } : closest;
  }, { offset: Number.NEGATIVE_INFINITY }).element;
};

/* ---- History --------------------------------------------- */
const renderHistory = () => {
  const tbody = $('#history-table-body');
  if (!tbody) return;
  const q = ($('#search-box')?.value || '').toLowerCase();
  tbody.innerHTML = '';
  [...historyData]
    .filter(h => !q || norm(h.company).includes(norm(q)) || norm(h.note).includes(norm(q)))
    .sort((a,b) => (parseDate(b.createdAt)||0) - (parseDate(a.createdAt)||0))
    .slice(0, 30)
    .forEach(h => {
      const tr = document.createElement('tr');
      tr.setAttribute('data-clickable', '');
      tr.innerHTML = `
        <td class="td-muted" style="font-size:11px;">${fmtDate(h.createdAt)}</td>
        <td style="font-weight:600;">${h.company}</td>
        <td class="td-muted">${h.media}</td>
        <td class="note-cell">${h.note}</td>
      `;
      tr.addEventListener('click', () => openDetail(h.id));
      tbody.appendChild(tr);
    });

  if (tbody.children.length === 0) {
    tbody.innerHTML = `<tr><td colspan="4"><div class="empty-state"><div class="empty-title">履歴がありません</div></div></td></tr>`;
  }
};

const renderHistoryList = () => {
  const tbody = $('#history-table-body-2');
  if (!tbody) return;
  const q   = ($('#search-box-2')?.value || '').toLowerCase();
  const ym  = $('#month-filter-2')?.value;
  const asc = $('#sort-order-2')?.value === 'asc';

  let list = historyData.slice();
  if (ym) list = list.filter(h => (h.createdAt || '').startsWith(ym) || fmtDate(h.createdAt).startsWith(ym));
  if (q)  list = list.filter(h => norm(h.company).includes(norm(q)) || norm(h.note).includes(norm(q)));
  list.sort((a,b) => { const d = (parseDate(a.createdAt)||0) - (parseDate(b.createdAt)||0); return asc ? d : -d; });

  const lbl = $('#history-count-label');
  if (lbl) lbl.textContent = `${list.length}件`;

  tbody.innerHTML = '';
  if (list.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5"><div class="empty-state"><div class="empty-title">履歴がありません</div></div></td></tr>`;
    return;
  }

  list.forEach(h => {
    const tr = document.createElement('tr');
    tr.setAttribute('data-clickable', '');
    tr.innerHTML = `
      <td class="td-muted" style="font-size:11px;">${fmtDate(h.createdAt)}</td>
      <td style="font-weight:600;">${h.company}</td>
      <td class="td-muted">${h.media}</td>
      <td class="note-cell ${historyCollapsed ? '' : 'expanded'}">${h.note}</td>
      <td><button class="btn btn-ghost btn-xs expand-btn" data-id="${h.id}">▼</button></td>
    `;
    const noteCell = tr.querySelector('.note-cell');
    tr.querySelector('.expand-btn').addEventListener('click', e => {
      e.stopPropagation();
      noteCell.classList.toggle('expanded');
    });
    tr.addEventListener('click', () => openDetail(h.id));
    tbody.appendChild(tr);
  });
};

$('#search-box')?.addEventListener('input', renderHistory);
$('#search-box-2')?.addEventListener('input', renderHistoryList);
$('#month-filter-2')?.addEventListener('change', renderHistoryList);
$('#sort-order-2')?.addEventListener('change', renderHistoryList);
$('#clear-filters-2')?.addEventListener('click', () => {
  $('#search-box-2').value = '';
  $('#month-filter-2').value = '';
  $('#sort-order-2').value = 'desc';
  renderHistoryList();
});

$('#collapse-all')?.addEventListener('click', () => {
  historyCollapsed = !historyCollapsed;
  $('#collapse-all').textContent = historyCollapsed ? '折りたたむ' : '展開する';
  renderHistoryList();
});

/* ---- History Form ---------------------------------------- */
$('#history-submit-btn')?.addEventListener('click', async () => {
  const company = $('#company-name').value.trim();
  const media   = $('#media-select').value;
  const note    = $('#history-note').value.trim();
  if (!company || !media || !note) { toast('全項目を入力してください', 'error'); return; }

  const item = { id: uuid(), createdAt: new Date().toISOString(), company, media, note };
  historyData.push(item);
  saveLocal();
  renderHistory();
  $('#company-name').value = '';
  $('#media-select').value = '';
  $('#history-note').value = '';
  toast('記録しました', 'ok');
  try { await apiPost({ action: 'upsertHistory', item }); } catch(e) { console.error(e); }
});

/* ---- Detail Modal ---------------------------------------- */
const openDetail = id => {
  const h = historyData.find(x => x.id === id);
  if (!h) return;
  $('#detail-id').value      = h.id;
  $('#detail-created').value = fmtDate(h.createdAt);
  $('#detail-company').value = h.company;
  $('#detail-media').value   = h.media;
  $('#detail-note').value    = h.note;
  $('#detail-modal').classList.add('active');
};

$('#detail-save')?.addEventListener('click', async () => {
  const id  = $('#detail-id').value;
  const idx = historyData.findIndex(x => x.id === id);
  if (idx < 0) return;
  historyData[idx] = {
    ...historyData[idx],
    company: $('#detail-company').value,
    media:   $('#detail-media').value,
    note:    $('#detail-note').value
  };
  saveLocal();
  renderHistory();
  renderHistoryList();
  $('#detail-modal').classList.remove('active');
  toast('更新しました', 'ok');
  try { await apiPost({ action: 'upsertHistory', item: historyData[idx] }); } catch(e) { console.error(e); }
});

$('#detail-delete')?.addEventListener('click', async () => {
  const id = $('#detail-id').value;
  if (!confirm('この履歴を削除しますか？')) return;
  historyData = historyData.filter(x => x.id !== id);
  saveLocal();
  renderHistory();
  renderHistoryList();
  $('#detail-modal').classList.remove('active');
  toast('削除しました', 'ok');
  try { await apiPost({ action: 'deleteHistory', id }); } catch(e) { console.error(e); }
});

/* ---- Modal Close ----------------------------------------- */
$$('.close-btn').forEach(btn => btn.addEventListener('click', closeAllModals));
window.addEventListener('click', e => { if (e.target.classList.contains('modal')) closeAllModals(); });
const closeAllModals = () => $$('.modal').forEach(m => m.classList.remove('active'));

/* ---- GAS Sync -------------------------------------------- */
const setSyncStatus = (state, label) => {
  const dot = $('#sync-dot');
  const lbl = $('#sync-label');
  if (dot) { dot.className = 'sync-dot'; if (state !== 'ok') dot.classList.add(state); }
  if (lbl) lbl.textContent = label;
};

const apiPost = async body => {
  const res = await fetch(GAS_BASE, {
    method: 'POST',
    body: JSON.stringify(body)
  });
  const text = await res.text();
  return text ? JSON.parse(text) : {};
};

const apiGet = async action => {
  const res = await fetch(`${GAS_BASE}?action=${action}&t=${Date.now()}`);
  return res.json();
};

const syncContactsToGAS = async () => {
  const kigyou = contactsData.filter(c => c.status === 'ユーザー').map(c => c.company);
  const mikomi = contactsData.filter(c => c.status === '見込').map(c => c.company);
  const botsu  = contactsData.filter(c => c.status === '没').map(c => c.company);
  try {
    await apiPost({ action: 'saveLists', kigyou, mikomi, botsu });
    console.log('GAS lists synced');
  } catch(e) {
    console.error('GAS sync failed:', e);
  }
};

const fetchDataFromGAS = async () => {
  setSyncStatus('syncing', '同期中...');
  try {
    const [listsData, histData] = await Promise.all([
      apiGet('lists'),
      apiGet('history')
    ]);

    // Merge lists into contacts (kigyou=ユーザー, mikomi=見込, botsu=没)
    // Only import if local contactsData is EMPTY (fresh start) OR merge from GAS
    const kigyou = listsData?.kigyou || [];
    const mikomi = listsData?.mikomi || [];
    const botsu  = listsData?.botsu  || [];

    // Build a map of company -> status from GAS
    const gasStatusMap = {};
    kigyou.forEach(c => gasStatusMap[c] = 'ユーザー');
    mikomi.forEach(c => gasStatusMap[c] = '見込');
    botsu.forEach(c  => gasStatusMap[c] = '没');

    // Upsert contacts from GAS lists
    let changed = false;
    Object.entries(gasStatusMap).forEach(([company, status]) => {
      if (!contactsData.find(c => norm(c.company) === norm(company))) {
        contactsData.push({ id: uuid(), company, status, name: '', tel: '', email: '', address: '', memo: '' });
        changed = true;
      }
    });

    // Update statuses for existing contacts
    contactsData.forEach(c => {
      const gasStatus = gasStatusMap[c.company];
      if (gasStatus && c.status !== gasStatus) {
        c.status = gasStatus;
        changed = true;
      }
    });

    // Load history
    if (Array.isArray(histData?.history) && histData.history.length > 0) {
      // Merge: keep local-only records and add GAS records
      const localIds = new Set(historyData.map(h => h.id));
      const gasItems = histData.history.map(it => ({
        id: it.id || uuid(),
        createdAt: it.createdAt || new Date().toISOString(),
        company: it.company || '',
        media: it.media || '',
        note: it.note || ''
      }));
      // Overwrite with GAS data (GAS is source of truth for history)
      historyData = gasItems;
      changed = true;
    }

    if (changed) {
      saveLocal();
      renderContacts();
      renderKanban();
      renderHistory();
      renderHistoryList();
      updateDashboard();
    }

    setSyncStatus('ok', `同期済 ${new Date().toLocaleTimeString('ja-JP')}`);
    toast('同期完了', 'ok');
  } catch(e) {
    console.error('Fetch error:', e);
    setSyncStatus('error', '同期失敗');
    toast('GAS同期に失敗しました', 'error');
  }
};

$('#btn-sync')?.addEventListener('click', fetchDataFromGAS);

/* ---- Import / Export ------------------------------------- */
$('#contacts-import-btn')?.addEventListener('click', () => $('#contacts-import-file').click());
$('#contacts-import-file')?.addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = async evt => {
    let text = evt.target.result;
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    const lines = text.split(/\r?\n/).filter(l => l.trim());
    if (lines.length < 2) return toast('データが空です', 'error');
    let count = 0;
    for (let i = 1; i < lines.length; i++) {
      const row = parseCSVLine(lines[i]);
      if (row.length < 2) continue;
      const id = row[0] || uuid();
      const item = { id, company: row[1]||'', status: row[2]||'', name: row[3]||'', tel: row[4]||'', email: row[5]||'', address: row[6]||'', memo: row[7]||'' };
      const idx = contactsData.findIndex(c => c.id === id);
      if (idx >= 0) contactsData[idx] = item; else contactsData.push(item);
      count++;
    }
    saveLocal(); renderContacts(); renderKanban();
    toast(`${count}件をインポートしました`, 'ok');
    await syncContactsToGAS();
  };
  reader.readAsText(file);
  e.target.value = '';
});

$('#contacts-export-btn')?.addEventListener('click', () => {
  if (!contactsData.length) return toast('データがありません', 'error');
  const headers = ['id','company','status','name','tel','email','address','memo'];
  const rows = contactsData.map(c => headers.map(k => escapeCsv(c[k]||'')));
  downloadCsv([headers.join(','), ...rows.map(r => r.join(','))].join('\n'), `contacts_${new Date().toISOString().slice(0,10)}.csv`);
});

$('#history-import-btn')?.addEventListener('click', () => $('#history-import-file').click());
$('#history-import-file')?.addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = async evt => {
    let text = evt.target.result;
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    const lines = text.split(/\r?\n/).filter(l => l.trim());
    if (lines.length < 2) return toast('データが空です', 'error');
    const header = lines[0].split(',').map(h => h.trim());
    const hasId = header.includes('id');
    const imported = lines.slice(1).map(line => {
      const cols = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(v => v.replace(/^"|"$/g,'').replace(/""/g,'"'));
      let id, createdAt, company, media, note;
      if (hasId) [id,createdAt,company,media,note] = cols;
      else [createdAt,company,media,note] = cols;
      return { id: id || uuid(), createdAt: createdAt || new Date().toISOString(), company: company||'', media: media||'', note: note||'' };
    });
    if (confirm(`${imported.length}件をインポートします（現在の履歴を上書き）。よろしいですか？`)) {
      historyData = imported;
      saveLocal(); renderHistory(); renderHistoryList();
      toast(`${imported.length}件をインポートしました`, 'ok');
    }
  };
  reader.readAsText(file);
  e.target.value = '';
});

$('#history-export-btn')?.addEventListener('click', () => {
  const headers = ['id','createdAt','company','media','note'];
  const rows = historyData.map(h => headers.map(k => escapeCsv(h[k]||'')));
  downloadCsv([headers.join(','), ...rows.map(r => r.join(','))].join('\n'), `history_${new Date().toISOString().slice(0,10)}.csv`);
});

const downloadCsv = (text, filename) => {
  const blob = new Blob([new Uint8Array([0xEF,0xBB,0xBF]), text], { type: 'text/csv;charset=utf-8' });
  const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: filename, style: 'display:none' });
  document.body.appendChild(a); a.click(); URL.revokeObjectURL(a.href); a.remove();
};

const parseCSVLine = text => {
  const re = /(?!\s*$)\s*(?:'([^']*)'|"([^"]*)"|([^,'"]*))\s*(?:,|$)/g;
  const a = [];
  text.replace(re, (m,m1,m2,m3) => {
    if (m1 !== undefined) a.push(m1.replace(/''/g,"'"));
    else if (m2 !== undefined) a.push(m2.replace(/""/g,'"'));
    else if (m3 !== undefined) a.push(m3);
    return '';
  });
  if (/,\s*$/.test(text)) a.push('');
  return a;
};

/* ---- Init ------------------------------------------------ */
document.addEventListener('DOMContentLoaded', () => {
  syncMemos();
  updateNavBadges();
  updateCompanyDatalist();
  renderContacts();
  renderKanban();
  renderHistory();
  renderHistoryList();
  updateDashboard();
  fetchDataFromGAS();
});
