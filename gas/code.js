/**
 * Personal TODO — backend (コンテナバインド / 全機能ぶんのスキーマ対応版)
 *
 * 既存シート(id/text/done/created)を開くと、足りない列を自動追加して移行します。
 * スコープは spreadsheets.currentonly のまま(=非機密)。
 *
 * 列: id, text, done, created, priority, due, category, sort, memo, parent, status, doneAt
 *   priority : 0=なし 1=高 2=中 3=低
 *   due      : 'YYYY-MM-DD' 文字列
 *   sort     : 手動並び替え用の数値
 *   parent   : サブタスクの親id(空=トップレベル)
 *   status   : カンバン用 'todo' / 'doing' / 'done'
 *   doneAt   : 完了日時 (Dateオブジェクト、未完了は空)
 */

const SHEET_NAME = 'todos';
const COLUMNS = ['id','text','done','created','priority','due','category','sort','memo','parent','status','doneAt'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('タスク整理くん')
    .addMetaTag('viewport',
      'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover');
}

/* ---------- sheet / schema ---------- */

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  ensureColumns_(sh);
  return sh;
}

function ensureColumns_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) {
    sh.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    return;
  }
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = COLUMNS.filter(function (c) { return header.indexOf(c) === -1; });
  if (missing.length) {
    sh.getRange(1, header.length + 1, 1, missing.length).setValues([missing]);
  }
}

function colMap_(sh) {
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const m = {};
  header.forEach(function (name, i) { m[name] = i; }); // 0-based
  return m;
}

function rowToObj_(r, m) {
  function dueStr(v) {
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return v ? String(v) : '';
  }
  return {
    id: String(r[m.id]),
    text: String(r[m.text]),
    done: r[m.done] === true || r[m.done] === 'TRUE' || r[m.done] === 'true',
    created: r[m.created] ? new Date(r[m.created]).getTime() : 0,
    priority: Number(r[m.priority]) || 0,
    due: dueStr(r[m.due]),
    category: r[m.category] ? String(r[m.category]) : '',
    sort: Number(r[m.sort]) || 0,
    memo: r[m.memo] ? String(r[m.memo]) : '',
    parent: r[m.parent] ? String(r[m.parent]) : '',
    status: r[m.status] ? String(r[m.status]) : 'todo',
    doneAt: (m.doneAt !== undefined && r[m.doneAt]) ? new Date(r[m.doneAt]).getTime() : 0
  };
}

/* ---------- read ---------- */

function getTodos() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const m = colMap_(sh);
  const rows = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  return rows.map(function (r) { return rowToObj_(r, m); });
}

/* ---------- helpers ---------- */

function findRow_(sh, id, m) {
  m = m || colMap_(sh);
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, m.id + 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function setField_(id, field, value) {
  const sh = getSheet_();
  const m = colMap_(sh);
  const row = findRow_(sh, id, m);
  if (row > 0 && m[field] !== undefined) sh.getRange(row, m[field] + 1).setValue(value);
  return getTodos();
}

function appendRow_(sh, obj) {
  const m = colMap_(sh);
  const row = new Array(sh.getLastColumn()).fill('');
  Object.keys(obj).forEach(function (k) { if (m[k] !== undefined) row[m[k]] = obj[k]; });
  sh.appendRow(row);
}

/* ---------- write ---------- */

function addTodo(text) {
  return addTodoFull(text, '', 'todo');
}

// カテゴリ・ステータス込みで1回の呼び出しで追加する(追加後のsetStatus/setCategory往復を無くす)
function addTodoFull(text, category, status) {
  text = String(text || '').trim();
  if (!text) return getTodos();
  const sh = getSheet_();
  appendRow_(sh, {
    id: Utilities.getUuid(), text: text, done: false, created: new Date(),
    priority: 0, due: '', category: category ? String(category) : '', sort: Date.now(),
    memo: '', parent: '', status: status ? String(status) : 'todo'
  });
  return getTodos();
}

// 複数フィールドを1回の呼び出しでまとめて更新する
function updateTodo(id, fields) {
  const sh = getSheet_();
  const m = colMap_(sh);
  const row = findRow_(sh, id, m);
  if (row > 0 && fields) {
    const allowed = ['text','priority','due','category','memo','status'];
    allowed.forEach(function (k) {
      if (fields[k] !== undefined && m[k] !== undefined) {
        sh.getRange(row, m[k] + 1).setValue(fields[k]);
      }
    });
  }
  return getTodos();
}

// タスク移動専用: ステータス・カテゴリ・並び順を1回の呼び出しで適用する
// status/category は null なら変更なし。sortUpdates は [{id, sort}] または null
function moveTask(id, status, category, sortUpdates) {
  const sh = getSheet_();
  const m = colMap_(sh);
  const last = sh.getLastRow();
  if (last >= 2) {
    const ids = sh.getRange(2, m.id + 1, last - 1, 1).getValues();
    const idx = {};
    ids.forEach(function (r, i) { idx[String(r[0])] = i; });
    const ri = idx[String(id)];
    if (ri !== undefined) {
      if (status != null)   sh.getRange(ri + 2, m.status + 1).setValue(String(status));
      if (category != null) sh.getRange(ri + 2, m.category + 1).setValue(String(category));
    }
    if (sortUpdates && sortUpdates.length) {
      const rng = sh.getRange(2, m.sort + 1, last - 1, 1);
      const sorts = rng.getValues();
      let changed = false;
      sortUpdates.forEach(function (u) {
        const j = idx[String(u.id)];
        if (j !== undefined) {
          const v = Number(u.sort) || 0;
          if (Number(sorts[j][0]) !== v) { sorts[j][0] = v; changed = true; }
        }
      });
      if (changed) rng.setValues(sorts);
    }
  }
  return getTodos();
}

function setDone(id, done) {
  const sh = getSheet_();
  const m = colMap_(sh);
  const row = findRow_(sh, id, m);
  if (row > 0) {
    sh.getRange(row, m['done'] + 1).setValue(done === true);
    if (m['doneAt'] !== undefined) {
      sh.getRange(row, m['doneAt'] + 1).setValue(done === true ? new Date() : '');
    }
  }
  return getTodos();
}
function editTodo(id, text)     { text = String(text || '').trim(); return text ? setField_(id, 'text', text) : getTodos(); }
function setPriority(id, p)     { return setField_(id, 'priority', Number(p) || 0); }
function setDue(id, due)        { return setField_(id, 'due', due ? String(due) : ''); }
function setCategory(id, cat)   { return setField_(id, 'category', cat ? String(cat) : ''); }
function setMemo(id, memo)      { return setField_(id, 'memo', memo ? String(memo) : ''); }
function setStatus(id, s)       { return setField_(id, 'status', s ? String(s) : 'todo'); }

function renameCategory(oldCat, newCat) {
  oldCat = String(oldCat || '');
  newCat = String(newCat || '').trim();
  if (!newCat || oldCat === newCat) return getTodos();
  const sh = getSheet_();
  const m = colMap_(sh);
  const last = sh.getLastRow();
  if (last < 2) return getTodos();
  const cats = sh.getRange(2, m.category + 1, last - 1, 1).getValues();
  cats.forEach(function(row, i) {
    if (String(row[0]) === oldCat) {
      sh.getRange(i + 2, m.category + 1).setValue(newCat);
    }
  });
  return getTodos();
}

function deleteTodo(id) {
  const sh = getSheet_();
  const m = colMap_(sh);
  const row = findRow_(sh, id, m);
  if (row > 0) sh.deleteRow(row);
  return getTodos();
}

function updateSort(updates) {
  if (!updates || !updates.length) return getTodos();
  const sh = getSheet_();
  const m = colMap_(sh);
  const last = sh.getLastRow();
  if (last < 2) return getTodos();
  // id列を1回だけ読み、sort列を1回のsetValuesでまとめて書く(1件ずつの検索+書き込みを廃止)
  const ids = sh.getRange(2, m.id + 1, last - 1, 1).getValues();
  const idx = {};
  ids.forEach(function (r, i) { idx[String(r[0])] = i; });
  const rng = sh.getRange(2, m.sort + 1, last - 1, 1);
  const sorts = rng.getValues();
  let changed = false;
  updates.forEach(function (u) {
    const j = idx[String(u.id)];
    if (j !== undefined) {
      const v = Number(u.sort) || 0;
      if (Number(sorts[j][0]) !== v) { sorts[j][0] = v; changed = true; }
    }
  });
  if (changed) rng.setValues(sorts);
  return getTodos();
}

function clearDone() {
  const sh = getSheet_();
  const m = colMap_(sh);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, m.done + 1, last - 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    const d = vals[i][0] === true || vals[i][0] === 'TRUE' || vals[i][0] === 'true';
    if (d) sh.deleteRow(i + 2);
  }
  return getTodos();
}
