// app.js (module)
// Full corrected JS for Sched Master
// - Single, clean initialization
// - All console logs retained
// - Fixed paste -> table rendering issue
// - Exports: loadState, updateButtonStates, recheckConflicts

// If you have firebase.js that exports getFirestore/auth, keep it as-is.
// This file imports Firestore helpers from ./firebase.js
import { getFirestore, collection, getDocs, onSnapshot, addDoc, updateDoc, deleteDoc, doc } from "./firebase.js";
// If your firebase.js also exports `auth`, you can import it there. We check for global fallback below.

console.log("🔧 app.js module loaded");

///////////////////////
// STATE & CONSTANTS //
///////////////////////
let workScheduleData = [];
let restDayData = [];
let monitoringData = []; // Firestore-managed
let currentlyEditing = { type: null, index: null };
let unsubscribeMonitoring = () => {};
const undoStack = { work: [], rest: [] };
const redoStack = { work: [], rest: [] };
const LEADERSHIP_POSITIONS = ['Branch Head', 'Site Supervisor', 'OIC'];

///////////////////////
// DOM REFERENCES    //
///////////////////////
// Grab once on module load — but DOM might not be ready yet, so we assign null and re-query in init
let workInput = null;
let restInput = null;
let generateWorkFileBtn = null;
let generateRestFileBtn = null;
let clearWorkBtn = null;
let clearRestBtn = null;
let undoWorkBtn = null;
let redoWorkBtn = null;
let undoRestBtn = null;
let redoRestBtn = null;

let tabSchedule = null;
let tabMonitoring = null;
let viewSchedule = null;
let viewMonitoring = null;

let addMonitoringRowBtn = null;

let workTableBody = null;
let restTableBody = null;
let monitoringTableBody = null;

let summaryEl = null;
let warningBanner = null;
let successMsg = null;
let backToTopBtn = null;

let editModal = null;
let closeEditModalBtn = null;
let cancelEditBtn = null;
let editForm = null;
let editShiftCodeWrapper = null;

let monitoringProgressBar = null;
let monitoringProgressText = null;
let monitoringCollectionRef = null;

//////////////////////////
// FIRESTORE / DB SETUP //
//////////////////////////
const db = (() => {
  try {
    return getFirestore();
  } catch (err) {
    console.warn("Firestore not available from import:", err);
    return null;
  }
})();

if (db) {
  monitoringCollectionRef = collection(db, "monitoring");
  console.log("✅ Firestore collection reference created");
} else {
  console.warn("⚠️ Firestore DB not initialized — realtime features will be disabled");
}

//////////////////////
// INITIALIZATION   //
//////////////////////
function initDOMRefs() {
  // Query DOM elements (safe to call after DOMContentLoaded)
  workInput = document.getElementById("workScheduleInput");
  restInput = document.getElementById("restScheduleInput");
  generateWorkFileBtn = document.getElementById("generateWorkFile");
  generateRestFileBtn = document.getElementById("generateRestFile");
  clearWorkBtn = document.getElementById("clearWorkData");
  clearRestBtn = document.getElementById("clearRestData");
  undoWorkBtn = document.getElementById("undoWork");
  redoWorkBtn = document.getElementById("redoWork");
  undoRestBtn = document.getElementById("undoRest");
  redoRestBtn = document.getElementById("redoRest");

  tabSchedule = document.getElementById("tab-schedule");
  tabMonitoring = document.getElementById("tab-monitoring");
  viewSchedule = document.getElementById("view-schedule");
  viewMonitoring = document.getElementById("view-monitoring");

  addMonitoringRowBtn = document.getElementById("addMonitoringRowBtn");

  workTableBody = document.getElementById("workTableBody");
  restTableBody = document.getElementById("restTableBody");
  monitoringTableBody = document.getElementById("monitoringTableBody");

  summaryEl = document.getElementById("summary");
  warningBanner = document.getElementById("warning-banner");
  successMsg = document.getElementById("success-message");
  backToTopBtn = document.getElementById("backToTopBtn");

  editModal = document.getElementById("editModal");
  closeEditModalBtn = document.getElementById("closeEditModalBtn");
  cancelEditBtn = document.getElementById("cancelEditBtn");
  editForm = document.getElementById("editForm");
  editShiftCodeWrapper = document.getElementById("editShiftCodeWrapper");

  monitoringProgressBar = document.getElementById("monitoringProgressBar");
  monitoringProgressText = document.getElementById("monitoringProgressText");

  console.log("✅ DOM refs initialized");
}

function attachListeners() {
  if (!generateWorkFileBtn || !generateRestFileBtn) {
    console.warn("⚠️ Important buttons not found — listeners may be incomplete.");
  }

  // Schedule controls
  generateWorkFileBtn?.addEventListener('click', () => {
    const branchName = document.getElementById('branchNameInput')?.value.trim();
    if (!branchName) {
      alert('⚠️ Please enter the Branch Name before generating the Work File.');
      return;
    }
    generateFile(workScheduleData, 'WorkSchedule');
  });

  generateRestFileBtn?.addEventListener('click', () => {
    const branchName = document.getElementById('branchNameInput')?.value.trim();
    if (!branchName) {
      alert('⚠️ Please enter the Branch Name before generating the Rest Day File.');
      return;
    }
    generateFile(restDayData, 'RestDaySchedule');
  });

  clearWorkBtn?.addEventListener('click', () => clearData('work'));
  clearRestBtn?.addEventListener('click', () => clearData('rest'));

  undoWorkBtn?.addEventListener('click', () => undo('work'));
  redoWorkBtn?.addEventListener('click', () => redo('work'));
  undoRestBtn?.addEventListener('click', () => undo('rest'));
  redoRestBtn?.addEventListener('click', () => redo('rest'));

  tabSchedule?.addEventListener('click', () => switchTab('schedule'));
  tabMonitoring?.addEventListener('click', () => switchTab('monitoring'));

  addMonitoringRowBtn?.addEventListener('click', addMonitoringBranch);

  closeEditModalBtn?.addEventListener('click', hideEditModal);
  cancelEditBtn?.addEventListener('click', hideEditModal);
  editForm?.addEventListener('submit', handleSaveEdit);

  // close modal clicking overlay
  editModal?.addEventListener('click', (e) => {
    if (e.target.classList.contains('modal-overlay')) {
      hideEditModal();
    }
  });

  // paste listeners for inputs
  workInput?.addEventListener("paste", handlePaste);
  restInput?.addEventListener("paste", handlePaste);

  console.log("✅ Event listeners attached");
}

window.addEventListener("DOMContentLoaded", () => {
  initDOMRefs();
  attachListeners();
  loadState();
  updateButtonStates();
  recheckConflicts();

  // Auto-select current month/year if those elements exist
  const now = new Date();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const year = now.getFullYear();
  const monthPicker = document.getElementById("monthPicker");
  const yearPicker = document.getElementById("yearPicker");
  if (monthPicker && yearPicker) {
    monthPicker.value = month;
    yearPicker.value = year;
    document.getElementById("loadMonthData")?.click();
  }

  // Firestore realtime listener (if DB initialized)
  if (monitoringCollectionRef && typeof onSnapshot === 'function') {
    listenForMonitoringUpdates();
  } else {
    console.warn("Realtime monitoring disabled (no DB reference)");
  }

  // Back to top button
  if (backToTopBtn) {
    window.addEventListener('scroll', () => {
      window.scrollY > 400 ? backToTopBtn.classList.add('show') : backToTopBtn.classList.remove('show');
    });
    backToTopBtn.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));
  }

  // particles init (safe if container present)
  initParticles();
});

//////////////////////
// CORE: Column map //
//////////////////////
function detectColumnMapping(rows, isWork) {
  const headerRow = 0;
  const headers = (rows[headerRow] || []).map(h => (h || '').trim().toLowerCase());

  const mapping = {
    employeeNo: headers.findIndex(h =>
      h.includes('emp') && (h.includes('no') || h.includes('#'))
    ),
    name: headers.findIndex(h => h.includes('name')),
    position: headers.findIndex(h => h.includes('pos')),
    date: headers.findIndex(h => h.includes('date')),
    dayOfWeek: headers.findIndex(h => h.includes('day')),
    shiftCode: isWork ? headers.findIndex(h => h.includes('shift')) : null
  };

  // fallback if missing headers
  if (mapping.employeeNo === -1) mapping.employeeNo = 0;
  if (mapping.name === -1) mapping.name = 1;
  if (mapping.position === -1) mapping.position = 2;
  if (mapping.date === -1) mapping.date = 3;

  return { mapping, headerRow };
}

//////////////////////
// HANDLE PASTE     //
//////////////////////
function handlePaste(event) {
  event.preventDefault();

  const branchName = document.getElementById("branchNameInput")?.value.trim();
  if (!branchName) {
    alert("⚠️ Please enter the Branch Name before pasting schedule data.");
    return;
  }

  const rawText = (event.clipboardData || window.clipboardData).getData("text") || "";
  const text = rawText.trim();
  const isWork = event.target && event.target.id === "workScheduleInput";
  const type = isWork ? "work" : "rest";

  console.log("📋 Paste detected (isWork:", isWork, ")");
  saveUndoState(type);

  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  if (lines.length === 0) {
    console.warn("Paste contained no text.");
    return;
  }

  const parsedRows = lines.map((line) => {
    if (line.includes("\t")) return line.split("\t").map(c => c.trim());
    if (/\s{2,}/.test(line)) return line.split(/\s{2,}/).map(c => c.trim());
    if (line.includes(",")) return line.split(",").map(c => c.trim());
    return line.split(/\s+/).map(c => c.trim());
  });

  const { mapping, headerRow } = detectColumnMapping(parsedRows, isWork);

  const data = parsedRows
    .slice(headerRow + 1)
    .map((row) => {
      const entry = {};
      for (const key in mapping) {
        if (mapping[key] !== null && row[mapping[key]] !== undefined) {
          entry[key] = (row[mapping[key]] || "").trim();
        } else {
          entry[key] = "";
        }
      }
      return entry;
    })
    .filter((entry) => entry.employeeNo && entry.name && entry.date);

  if (isWork) {
    workScheduleData = data.map((d) => ({
      ...d,
      date: excelDateToJS(d.date),
      conflict: false,
      conflictReason: ''
    }));
  } else {
    restDayData = data.map((d) => ({
      ...d,
      date: excelDateToJS(d.date),
      conflict: false,
      conflictReason: ''
    }));
  }

  recheckConflicts();
  updateButtonStates();
  renderWorkTable();
  renderRestTable();
  saveState();

  // smooth scroll to the table
  setTimeout(() => {
    const target = isWork ? workTableBody : restTableBody;
    if (target) {
      target.scrollIntoView({ behavior: "smooth", block: "center" });
    }
  }, 300);
}

//////////////////////
// CONFLICT CHECK   //
//////////////////////
function recheckConflicts() {
  const scheduleMap = new Map();

  workScheduleData.forEach(d => { d.conflict = false; d.conflictReason = ''; });
  restDayData.forEach(d => { d.conflict = false; d.conflictReason = ''; });

  workScheduleData.forEach((item, index) => {
    const key = `${item.employeeNo}-${item.date}`;
    if (!scheduleMap.has(key)) scheduleMap.set(key, { type: 'work', rowNum: index + 1 });
  });

  restDayData.forEach(item => {
    const key = `${item.employeeNo}-${item.date}`;
    if (scheduleMap.has(key)) {
      const workEntry = scheduleMap.get(key);
      item.conflict = true;
      item.conflictReason = `vs. Work Sched Row #${workEntry.rowNum}`;

      const workItem = workScheduleData.find(d => d.employeeNo === item.employeeNo && d.date === item.date);
      if (workItem) workItem.conflict = true;
    }
  });

  const conflictCount = restDayData.filter(d => d.conflict).length;
  const leadershipConflict = [...workScheduleData, ...restDayData]
    .some(d => d.conflict && LEADERSHIP_POSITIONS.includes((d.position || '').trim()));

  renderSummary(conflictCount, leadershipConflict);
  renderWorkTable();
  renderRestTable();
  console.log("🔍 Conflicts rechecked", { conflictCount, leadershipConflict });
}

//////////////////////
// FILE GENERATION  //
//////////////////////
function generateFile(data, fileNamePrefix) {
  if (!Array.isArray(data) || data.length === 0) {
    showWarning('No data to generate file.');
    return;
  }

  const branchName = document.getElementById('branchNameInput')?.value.trim();
  if (!branchName) {
    alert('⚠️ Please enter the Branch Name before generating the file.');
    return;
  }

  const now = new Date();
  const month = now.toLocaleString('default', { month: 'short' });
  const year = now.getFullYear();

  let formattedData;

  if (fileNamePrefix === 'WorkSchedule') {
    formattedData = data.map(row => ({
      'Employee Number': row.employeeNo,
      'Work Date': new Date(row.date),
      'Shift Code': row.shiftCode,
    }));
  } else if (fileNamePrefix === 'RestDaySchedule') {
    formattedData = data.map(row => ({
      'Employee No': row.employeeNo,
      'Rest Day Date': new Date(row.date),
    }));
  } else {
    formattedData = data.map(row => {
      const newRow = { ...row };
      delete newRow.conflict;
      delete newRow.conflictReason;
      newRow.date = new Date(newRow.date);
      return newRow;
    });
  }

  const worksheet = XLSX.utils.json_to_sheet(formattedData);
  Object.keys(worksheet).forEach(cell => {
    if (cell[0] === "!" || !worksheet[cell].v) return;
    const val = worksheet[cell].v;
    if (val instanceof Date || /^\d{4}-\d{2}-\d{2}/.test(val)) {
      worksheet[cell].t = "d";
      worksheet[cell].z = "mm/dd/yy";
    }
  });

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

  const safeBranch = branchName.replace(/[^\w\s-]/g, '').replace(/\s+/g, '_');
  const fileName = `${safeBranch}_${fileNamePrefix}_${month}${year}.xlsx`;

  XLSX.writeFile(workbook, fileName);
  showSuccess(`File generated successfully: ${fileName}`);
}

//////////////////////
// RENDERING TABLES //
//////////////////////
function renderTable(tbody, data, columns, type) {
  if (!tbody) {
    console.warn("Table body not found for", type);
    return;
  }

  tbody.innerHTML = '';
  if (!data || data.length === 0) {
    const tr = tbody.insertRow();
    const cell = tr.insertCell();
    let colspan = columns.length + 2; // # and Actions
    if (type === 'rest') colspan++; // Conflict column
    cell.colSpan = colspan;
    cell.textContent = 'No data available.';
    cell.className = 'text-center text-slate-500 py-4';
    return;
  }

  data.forEach((item, index) => {
    const tr = tbody.insertRow();
    if (item.conflict) tr.classList.add('conflict');

    if (type === 'rest') {
      const conflictCell = tr.insertCell();
      conflictCell.className = 'text-center';
      if (item.conflict) {
        conflictCell.innerHTML = `⚠️ <span class="conflict-reason">${item.conflictReason || ''}</span>`;
      } else {
        conflictCell.textContent = '';
      }
    }

    const numCell = tr.insertCell();
    numCell.textContent = index + 1;
    numCell.className = 'text-sm text-slate-600 text-center';

    columns.forEach(col => {
      const cell = tr.insertCell();
      // If it's a date string, display it as-is; editing will use edit form
      cell.textContent = (item[col] !== undefined && item[col] !== null) ? item[col] : '';
    });

    const actionsCell = tr.insertCell();
    actionsCell.className = 'flex items-center justify-center space-x-3';
    actionsCell.innerHTML = `
      <ion-icon name="create-outline" class="action-icon edit-icon text-xl" data-type="${type}" data-index="${index}"></ion-icon>
      <ion-icon name="trash-outline" class="action-icon delete-icon text-xl" data-type="${type}" data-index="${index}"></ion-icon>
    `;
  });

  // attach icons listeners (delegation might be better, but keep current approach)
  tbody.querySelectorAll('.edit-icon').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const t = e.currentTarget.dataset.type;
      const i = Number(e.currentTarget.dataset.index);
      handleEditRow(t, i);
    });
  });
  tbody.querySelectorAll('.delete-icon').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const t = e.currentTarget.dataset.type;
      const i = Number(e.currentTarget.dataset.index);
      handleDeleteRow(t, i);
    });
  });
}

function renderWorkTable() {
  renderTable(workTableBody, workScheduleData, ['employeeNo', 'name', 'position', 'date', 'shiftCode', 'dayOfWeek'], 'work');
}

function renderRestTable() {
  renderTable(restTableBody, restDayData, ['employeeNo', 'name', 'position', 'date', 'dayOfWeek'], 'rest');
}

//////////////////////
// SUMMARY / BUTTONS //
//////////////////////
function renderSummary(conflictCount, leadershipConflict) {
  if (!summaryEl) return;
  if (workScheduleData.length === 0 && restDayData.length === 0) {
    summaryEl.innerHTML = `<h2 class="text-2xl font-bold mb-4 text-slate-800">Summary & Conflicts</h2><p id="summaryText">No data pasted yet.</p>`;
    return;
  }
  let summaryHTML = `<h2 class="text-2xl font-bold mb-4 text-slate-800">Summary & Conflicts</h2>`;
  if (conflictCount > 0) {
    summaryHTML += `<div class="bg-red-100 border-l-4 border-red-500 text-red-700 p-4 rounded-lg" role="alert">
      <p class="font-bold">Conflicts Detected!</p>
      <p>${conflictCount} employee(s) have a rest day scheduled on a work day. Please review the highlighted rows.</p>
      ${leadershipConflict ? `<p class="mt-2"><strong>Warning:</strong> A conflict involves leadership.</p>` : ''}
    </div>`;
  } else {
    summaryHTML += `<div class="bg-green-100 border-l-4 border-green-500 text-green-700 p-4 rounded-lg" role="alert">
      <p class="font-bold">No Conflicts Found!</p>
      <p>All schedules appear to be in order.</p>
    </div>`;
  }
  summaryEl.innerHTML = summaryHTML;
}

function updateButtonStates() {
  generateWorkFileBtn && (generateWorkFileBtn.disabled = workScheduleData.length === 0);
  clearWorkBtn && (clearWorkBtn.disabled = workScheduleData.length === 0);
  undoWorkBtn && (undoWorkBtn.disabled = undoStack.work.length === 0);
  redoWorkBtn && (redoWorkBtn.disabled = redoStack.work.length === 0);

  generateRestFileBtn && (generateRestFileBtn.disabled = restDayData.length === 0);
  clearRestBtn && (clearRestBtn.disabled = restDayData.length === 0);
  undoRestBtn && (undoRestBtn.disabled = undoStack.rest.length === 0);
  redoRestBtn && (redoRestBtn.disabled = redoStack.rest.length === 0);
}

function switchTab(tab) {
  if (!tabSchedule || !tabMonitoring || !viewSchedule || !viewMonitoring) return;
  if (tab === 'schedule') {
    tabSchedule.classList.add('active');
    tabMonitoring.classList.remove('active');
    viewSchedule.classList.remove('hidden');
    viewMonitoring.classList.add('hidden');
  } else {
    tabSchedule.classList.remove('active');
    tabMonitoring.classList.add('active');
    viewSchedule.classList.add('hidden');
    viewMonitoring.classList.remove('hidden');
  }
}

//////////////////////
// EDIT / DELETE    //
//////////////////////
function handleDeleteRow(type, index) {
  if (!confirm('Are you sure you want to delete this entry?')) return;
  const dataArray = type === 'work' ? workScheduleData : restDayData;
  saveUndoState(type);
  dataArray.splice(index, 1);
  recheckConflicts();
  saveState();
}

function handleEditRow(type, index) {
  currentlyEditing = { type, index };
  const dataArray = type === 'work' ? workScheduleData : restDayData;
  const item = dataArray[index] || {};

  document.getElementById('editEmployeeNo').value = item.employeeNo || '';
  document.getElementById('editName').value = item.name || '';
  document.getElementById('editPosition').value = item.position || '';
  document.getElementById('editDate').value = item.date || '';
  document.getElementById('editDayOfWeek').value = item.dayOfWeek || '';

  if (type === 'work') {
    document.getElementById('editShiftCode').value = item.shiftCode || '';
    if (editShiftCodeWrapper) editShiftCodeWrapper.style.display = 'block';
  } else {
    if (editShiftCodeWrapper) editShiftCodeWrapper.style.display = 'none';
  }

  showEditModal();
}

function handleSaveEdit(event) {
  event.preventDefault();
  const { type, index } = currentlyEditing;
  if (type === null || index === null) return;

  const dataArray = type === 'work' ? workScheduleData : restDayData;
  saveUndoState(type);

  const updatedItem = {
    ...dataArray[index],
    employeeNo: document.getElementById('editEmployeeNo').value,
    name: document.getElementById('editName').value,
    position: document.getElementById('editPosition').value,
    date: document.getElementById('editDate').value,
    dayOfWeek: document.getElementById('editDayOfWeek').value,
  };

  if (type === 'work') {
    updatedItem.shiftCode = document.getElementById('editShiftCode').value;
  }

  dataArray[index] = updatedItem;

  hideEditModal();
  recheckConflicts();
  saveState();
}

function showEditModal() {
  if (!editModal) return;
  editModal.classList.remove('hidden');
  editModal.style.display = 'flex';
}
function hideEditModal() {
  if (!editModal) return;
  editModal.classList.add('hidden');
  editModal.style.display = 'none';
  currentlyEditing = { type: null, index: null };
}

//////////////////////
// UNDO / REDO      //
//////////////////////
function saveUndoState(type) {
  const data = type === 'work' ? workScheduleData : restDayData;
  undoStack[type].push(JSON.parse(JSON.stringify(data)));
  redoStack[type] = [];
  if (undoStack[type].length > 10) undoStack[type].shift();
  updateButtonStates();
}

function undo(type) {
  if (!undoStack[type] || undoStack[type].length === 0) return;
  const previousState = undoStack[type].pop();
  const currentState = (type === 'work') ? workScheduleData : restDayData;
  redoStack[type].push(JSON.parse(JSON.stringify(currentState)));

  if (type === 'work') workScheduleData = previousState;
  else restDayData = previousState;

  recheckConflicts();
  updateButtonStates();
  saveState();
}

function redo(type) {
  if (!redoStack[type] || redoStack[type].length === 0) return;
  const nextState = redoStack[type].pop();
  const currentState = (type === 'work') ? workScheduleData : restDayData;
  undoStack[type].push(JSON.parse(JSON.stringify(currentState)));

  if (type === 'work') workScheduleData = nextState;
  else restDayData = nextState;

  recheckConflicts();
  updateButtonStates();
  saveState();
}

//////////////////////
// CLEAR DATA       //
//////////////////////
function clearData(type) {
  if (!confirm(`Are you sure you want to clear all ${type} data?`)) return;
  saveUndoState(type);
  if (type === 'work') {
    workScheduleData = [];
    if (workInput) workInput.value = '';
  } else {
    restDayData = [];
    if (restInput) restInput.value = '';
  }
  recheckConflicts();
  updateButtonStates();
  saveState();
}

//////////////////////
// MONITORING UI    //
//////////////////////
document.getElementById("loadMonthData")?.addEventListener("click", async () => {
  const month = document.getElementById("monthPicker")?.value;
  const year = document.getElementById("yearPicker")?.value;

  if (!db) {
    console.warn("Firestore not available — cannot load monitoring data");
    return;
  }

  try {
    const monitoringRef = collection(db, "monitoring");
    const snapshot = await getDocs(monitoringRef);
    const tableBody = document.getElementById("monitoringTableBody");
    if (!tableBody) return;
    tableBody.innerHTML = "";

    snapshot.forEach((d) => {
      const docData = d.data();
      if (!docData) return;
      if (String(docData.month) === String(month) && String(docData.year) === String(year)) {
        const row = `
          <tr>
            <td>${docData.posCode || ""}</td>
            <td>${docData.branchName || ""}</td>
            <td>${docData.sapCode || ""}</td>
            <td class="text-center">${docData.uploaded ? "✅" : "❌"}</td>
            <td>${docData.uploadedBy || ""}</td>
            <td>${docData.uploadedDate || ""}</td>
            <td><button class="action-btn danger" data-id="${d.id}">🗑️ Delete</button></td>
          </tr>
        `;
        tableBody.insertAdjacentHTML("beforeend", row);
      }
    });

    updateMonitoringProgress();
    console.log("✅ Monitoring month data loaded");
  } catch (err) {
    console.error("Error loading monitoring data:", err);
  }
});

document.getElementById("exportExcel")?.addEventListener("click", () => {
  const tableBody = document.getElementById("monitoringTableBody");
  const rows = tableBody?.rows;
  if (!rows || !rows.length) {
    alert("No data to export!");
    return;
  }

  const wb = XLSX.utils.book_new();
  const data = [["POS Code", "Branch Name", "SAP Code", "Uploaded", "Uploaded By", "Uploaded Date"]];
  Array.from(rows).forEach((row) => {
    const cells = Array.from(row.cells).map((c) => c.innerText);
    data.push(cells);
  });

  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, "Monitoring");
  XLSX.writeFile(wb, "MonitoringData.xlsx");
  console.log("✅ Monitoring exported to Excel");
});

function updateMonitoringProgress() {
  const rows = document.querySelectorAll("#monitoringTableBody tr");
  if (!rows.length) return;
  const total = rows.length;
  const completed = [...rows].filter((r) => r.cells[3].innerText === "✅").length;
  const percent = Math.round((completed / total) * 100);
  if (monitoringProgressBar) monitoringProgressBar.style.width = `${percent}%`;
  if (monitoringProgressText) monitoringProgressText.textContent = `${percent}%`;
  console.log("📊 Monitoring progress updated:", percent);
}

//////////////////////
// REAL-TIME MONITOR //
//////////////////////
function listenForMonitoringUpdates() {
  if (!monitoringCollectionRef || typeof onSnapshot !== 'function') return;
  try {
    unsubscribeMonitoring(); // detach previous if any
  } catch(e) { /* ignore */ }

  unsubscribeMonitoring = onSnapshot(monitoringCollectionRef, (snapshot) => {
    const serverData = [];
    snapshot.forEach((d) => {
      serverData.push({ id: d.id, ...d.data() });
    });
    monitoringData = serverData.sort((a,b) => (a.posCode || '').localeCompare(b.posCode || ''));
    renderMonitoringDashboard();
    console.log("🔁 Monitoring realtime snapshot received, rows:", monitoringData.length);
  }, (error) => {
    console.error("Error listening to monitoring updates:", error);
    showWarning("Real-time connection lost. Please refresh.");
  });
}

function renderMonitoringDashboard() {
  if (!monitoringTableBody) {
    monitoringTableBody = document.getElementById("monitoringTableBody");
    if (!monitoringTableBody) { console.warn("monitoringTableBody not found"); return; }
  }
  monitoringTableBody.innerHTML = '';
  monitoringData.forEach((branch) => {
    const tr = monitoringTableBody.insertRow();
    tr.innerHTML = `
      <td><input class="monitoring-table-input" type="text" data-id="${branch.id}" data-field="posCode" value="${branch.posCode || ''}"></td>
      <td><input class="monitoring-table-input" type="text" data-id="${branch.id}" data-field="branchName" value="${branch.branchName || ''}"></td>
      <td><input class="monitoring-table-input" type="text" data-id="${branch.id}" data-field="sapCode" value="${branch.sapCode || ''}"></td>
      <td class="text-center"><input type="checkbox" data-id="${branch.id}" data-field="isUploaded" ${branch.isUploaded ? 'checked' : ''} class="h-5 w-5 rounded border-slate-300 text-cyan-600 focus:ring-cyan-500"></td>
      <td><input class="monitoring-table-input" type="text" data-id="${branch.id}" data-field="uploadedBy" value="${branch.uploadedBy || ''}"></td>
      <td><input class="monitoring-table-input" type="text" data-id="${branch.id}" data-field="uploadedDate" value="${branch.uploadedDate || ''}"></td>
      <td class="text-center"><ion-icon name="trash-outline" class="action-icon delete-icon text-xl" data-id="${branch.id}"></ion-icon></td>
    `;
  });

  // wire inputs
  monitoringTableBody.querySelectorAll('input[type="text"]').forEach(input => {
    input.addEventListener('change', (e) => updateMonitoringBranch(e.target.dataset.id, e.target.dataset.field, e.target.value));
  });
  monitoringTableBody.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
    checkbox.addEventListener('change', (e) => updateMonitoringBranch(e.target.dataset.id, e.target.dataset.field, e.target.checked));
  });
  monitoringTableBody.querySelectorAll('.delete-icon').forEach(button => {
    button.addEventListener('click', (e) => deleteMonitoringBranch(e.currentTarget.dataset.id));
  });

  updateMonitoringProgress();
}

async function addMonitoringBranch() {
  if (!monitoringCollectionRef) {
    showWarning("Monitoring not available (no DB).");
    return;
  }
  const newBranch = { posCode: '', branchName: '', sapCode: '', isUploaded: false, uploadedBy: '', uploadedDate: '' };
  try {
    await addDoc(monitoringCollectionRef, newBranch);
    console.log("➕ Monitoring branch added");
  } catch (error) {
    console.error("Error adding branch:", error);
    showWarning("Could not add branch.");
  }
}

async function deleteMonitoringBranch(id) {
  if (!confirm('Are you sure you want to remove this branch from the tracker?')) return;
  if (!db) return showWarning("Firestore not available");
  try {
    await deleteDoc(doc(db, monitoringCollectionRef.path, id));
    console.log("🗑️ Monitoring branch deleted", id);
  } catch (error) {
    console.error("Error deleting branch:", error);
    showWarning("Could not delete branch.");
  }
}

async function updateMonitoringBranch(id, field, value) {
  if (!db) return;
  // If `auth` is not exported from firebase.js, attempt to use global `auth`
  const authObj = (typeof auth !== 'undefined') ? auth : (window.auth || null);
  const updateData = { [field]: value };

  if (field === 'isUploaded') {
    updateData.uploadedBy = value ? (authObj?.currentUser?.email || 'User') : '';
    updateData.uploadedDate = value ? new Date().toLocaleDateString() : '';
  }

  try {
    await updateDoc(doc(db, monitoringCollectionRef.path, id), updateData);
    console.log("✅ Monitoring branch updated", id, field, value);
  } catch (error) {
    console.error("Error updating branch:", error);
    showWarning("Could not save changes.");
  }
}

//////////////////////
// UTIL / HELPERS   //
//////////////////////
function excelDateToJS(excelDate) {
  if (typeof excelDate === 'string' && (excelDate.includes('/') || excelDate.includes('-'))) {
    return new Date(excelDate).toLocaleDateString();
  }
  const a = parseFloat(excelDate);
  if (isNaN(a)) return excelDate;
  return new Date((a - 25569) * 86400 * 1000).toLocaleDateString();
}
function jsDateToExcel(jsDateStr) {
  const date = new Date(jsDateStr);
  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  return (date - excelEpoch) / (24 * 60 * 60 * 1000);
}

function showWarning(message) {
  if (!warningBanner) {
    console.warn("Warning banner missing:", message);
    return;
  }
  warningBanner.textContent = message;
  warningBanner.classList.remove('hidden');
  gsap.to(warningBanner, { opacity: 1, duration: 0.5 });
  setTimeout(() => {
    gsap.to(warningBanner, { opacity: 0, duration: 0.5, onComplete: () => warningBanner.classList.add('hidden') });
  }, 4000);
}

function showSuccess(message) {
  if (!successMsg) { console.log("Success:", message); return; }
  successMsg.textContent = message;
  successMsg.classList.remove('hidden');
  gsap.to(successMsg, { opacity: 1, duration: 0.5 });
  setTimeout(() => {
    gsap.to(successMsg, { opacity: 0, duration: 0.5, onComplete: () => successMsg.classList.add('hidden') });
  }, 3000);
}

function initParticles() {
  const header = document.querySelector('header');
  if (header && window.matchMedia('(min-width: 1024px)').matches) {
    document.addEventListener('mousemove', (e) => {
      const x = (e.clientX / window.innerWidth - 0.5) * -3;
      const y = (e.clientY / window.innerHeight - 0.5) * -3;
      gsap.to(header, { duration: 1.5, backgroundPosition: `${50 + x}% ${50 + y}%`, ease: "power2.out" });
    });
  }

  const particleContainer = document.getElementById('particle-container');
  if (!particleContainer) return;
  const particleCount = 50;
  const colors = ['#a7f3d0', '#67e8f9', '#5eead4', '#99f6e4'];
  for (let i = 0; i < particleCount; i++) {
    const particle = document.createElement('div');
    particle.classList.add('particle');
    const size = Math.random() * 6 + 2;
    particle.style.cssText = `width:${size}px; height:${size}px; background:${colors[Math.floor(Math.random()*colors.length)]}; border-radius:50%; position:absolute; top:${Math.random()*100}%; left:${Math.random()*100}%; opacity:${Math.random()*0.5+0.1};`;
    particleContainer.appendChild(particle);
    gsap.to(particle, { x: (Math.random()-0.5)*200, y: (Math.random()-0.5)*200, duration: Math.random()*20+15, repeat: -1, yoyo: true, ease: 'sine.inOut' });
  }
}

//////////////////////
// LOCAL STORAGE    //
//////////////////////
function saveState() {
  try {
    localStorage.setItem('workScheduleData_v3', JSON.stringify(workScheduleData));
    localStorage.setItem('restDayData_v3', JSON.stringify(restDayData));
    console.log("💾 State saved to localStorage");
  } catch (e) { console.warn("Could not save schedule state.", e); }
}
function loadState() {
  try {
    const w = JSON.parse(localStorage.getItem('workScheduleData_v3') || '[]');
    const r = JSON.parse(localStorage.getItem('restDayData_v3') || '[]');
    workScheduleData = Array.isArray(w) ? w : [];
    restDayData = Array.isArray(r) ? r : [];
    console.log("📥 State loaded from localStorage", { work: workScheduleData.length, rest: restDayData.length });
  } catch (e) { console.error("Could not load schedule state.", e); }
}

//////////////////////
// Exports (module) //
//////////////////////
export { loadState, updateButtonStates, recheckConflicts };
