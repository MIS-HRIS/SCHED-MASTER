/***** State *****/
      let workScheduleData = [];
      let restDayData = [];
      let monitoringData = []; // This will be managed by Firestore
      let currentlyEditing = { type: null, index: null };
      let unsubscribeMonitoring = () => {};
      const undoStack = { work: [], rest: [] };
      const redoStack = { work: [], rest: [] };
      const LEADERSHIP_POSITIONS = ['Branch Head', 'Site Supervisor', 'OIC'];

      /***** Element refs *****/
      const workInput = document.getElementById('workScheduleInput');
      const restInput = document.getElementById('restScheduleInput');
      const workTableBody = document.getElementById('workTableBody');
      const restTableBody = document.getElementById('restTableBody');
      const summaryEl = document.getElementById('summary');
      const warningBanner = document.getElementById('warning-banner');
      const successMsg = document.getElementById('success-message');
      const generateWorkFileBtn = document.getElementById('generateWorkFile');
      const generateRestFileBtn = document.getElementById('generateRestFile');
      const clearWorkBtn = document.getElementById('clearWorkData');
      const clearRestBtn = document.getElementById('clearRestData');
      const undoWorkBtn = document.getElementById('undoWork');
      const redoWorkBtn = document.getElementById('redoWork');
      const undoRestBtn = document.getElementById('undoRest');
      const redoRestBtn = document.getElementById('redoRest');
      const backToTopBtn = document.getElementById('backToTopBtn');
      
      const tabSchedule = document.getElementById('tab-schedule');
      const tabMonitoring = document.getElementById('tab-monitoring');
      const viewSchedule = document.getElementById('view-schedule');
      const viewMonitoring = document.getElementById('view-monitoring');
      
      const monitoringTableBody = document.getElementById('monitoringTableBody');
      const addMonitoringRowBtn = document.getElementById('addMonitoringRowBtn');
      const monitoringProgressBar = document.getElementById('monitoringProgressBar');
      const monitoringProgressText = document.getElementById('monitoringProgressText');
      
      const editModal = document.getElementById('editModal');
      const closeEditModalBtn = document.getElementById('closeEditModalBtn');
      const cancelEditBtn = document.getElementById('cancelEditBtn');
      const editForm = document.getElementById('editForm');
      const editShiftCodeWrapper = document.getElementById('editShiftCodeWrapper');


      /*************************\
       * EVENT LISTENERS     *
      \*************************/
      workInput.addEventListener('paste', handlePaste);
      restInput.addEventListener('paste', handlePaste);

      generateWorkFileBtn.addEventListener('click', () => generateFile(workScheduleData, 'WorkSchedule'));
      generateRestFileBtn.addEventListener('click', () => generateFile(restDayData, 'RestDaySchedule'));
      
      clearWorkBtn.addEventListener('click', () => clearData('work'));
      clearRestBtn.addEventListener('click', () => clearData('rest'));

      undoWorkBtn.addEventListener('click', () => undo('work'));
      redoWorkBtn.addEventListener('click', () => redo('work'));
      undoRestBtn.addEventListener('click', () => undo('rest'));
      redoRestBtn.addEventListener('click', () => redo('rest'));

      tabSchedule.addEventListener('click', () => switchTab('schedule'));
      tabMonitoring.addEventListener('click', () => switchTab('monitoring'));
      
      addMonitoringRowBtn.addEventListener('click', addMonitoringBranch);
      
      closeEditModalBtn.addEventListener('click', hideEditModal);
      cancelEditBtn.addEventListener('click', hideEditModal);
      editForm.addEventListener('submit', handleSaveEdit);
      editModal.addEventListener('click', (e) => {
        if (e.target.classList.contains('modal-overlay')) {
            hideEditModal();
        }
      });


      /*************************\
       * CORE FUNCTIONS      *
      \*************************/

function handlePaste(event) {
  event.preventDefault();
  const text = (event.clipboardData || window.clipboardData).getData('text');
  const isWork = event.target.id === 'workScheduleInput';
  const type = isWork ? 'work' : 'rest';

  saveUndoState(type);

  const rows = text
    .split('\n')
    .map(row => row.trim())
    .filter(row => row.length > 0)
    .map(row => (row.includes('\t') ? row.split('\t') : row.split(/\s{2,}|\t/)));

  const { mapping, headerRow } = detectColumnMapping(rows, isWork);
  let data;

  if (headerRow !== -1 && Object.values(mapping).some(v => v !== null)) {
    // ✅ Structured header input
    data = rows.slice(headerRow + 1).map(row => {
      const entry = {};
      for (const key in mapping) {
        entry[key] =
          mapping[key] !== null && row[mapping[key]] !== undefined
            ? row[mapping[key]].trim()
            : '';
      }
      return entry;
    });
  } else {
    // ✅ Flexible parsing (no headers)
    const classifyValue = (v) => {
      v = v.trim();

      if (/^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/.test(v) || /^\d{4}-\d{1,2}-\d{1,2}$/.test(v))
        return "date"; // Date

      if (/^(sun(day)?|mon(day)?|tue(s|sday)?|wed(nesday)?|thu(rs|rsday)?|fri(day)?|sat(urday)?)$/i.test(v))
        return "dayOfWeek"; // Day

      if (/^\d{1,6}$/.test(v))
        return "employeeNo"; // Employee number

      if (/^[A-Za-z]{3}-\d{3}$/.test(v))
        return "shiftCode"; // Shift code

      if (/^(cashier|manager|supervisor|assistant|oic|head|lead|ia|mac|expert|branch\s*head)$/i.test(v))
        return "position"; // Known position words

      if (/^[A-Za-z]+$/.test(v))
        return "namePart"; // Name fragment

      return "unknown";
    };

    data = rows.map(rawRow => {
      const tokens = rawRow.join(' ').split(/\s+/).filter(Boolean);
      const entry = { nameParts: [], positionParts: [] };

      tokens.forEach(t => {
        const type = classifyValue(t);
        if (type === "namePart") {
          entry.nameParts.push(t);
        } else if (type === "position") {
          entry.positionParts.push(t);
        } else if (type !== "unknown" && !entry[type]) {
          entry[type] = t;
        }
      });

      if (entry.nameParts.length > 0) entry.name = entry.nameParts.join(' ');
      if (entry.positionParts.length > 0) entry.position = entry.positionParts.join(' ');
      delete entry.nameParts;
      delete entry.positionParts;

      return entry;
    });
  }

  // ✅ Ensure data goes into table
  if (isWork) {
    workScheduleData = data.map(d => ({
      ...d,
      date: excelDateToJS(d.date),
    }));
  } else {
    restDayData = data.map(d => ({
      ...d,
      date: excelDateToJS(d.date),
    }));
  }

  // ✅ Always refresh UI
  recheckConflicts();
  updateButtonStates();

  console.log('Parsed data:', data);
}

function detectColumnMapping(rows, isWork) {
  let headerRow = -1;
  let mapping = {};
  const MAX_ROWS_TO_CHECK = 20;

  // ✅ Step 1: Try detecting header row first
  const potentialHeaders = {
    employeeNo: ['employee no.', 'emp no', 'employee number', 'id'],
    name: ['name', 'employee name', 'fullname'],
    position: ['position', 'designation', 'role', 'title'],
    date: isWork ? ['work date', 'date'] : ['rest day date', 'date', 'rest day'],
    shiftCode: ['shift code', 'shift', 'scode'],
    dayOfWeek: ['day of week', 'day']
  };

  for (let i = 0; i < Math.min(rows.length, MAX_ROWS_TO_CHECK); i++) {
    const row = rows[i].map(h => h.toLowerCase().trim().replace(':', ''));
    let tempMapping = {};
    for (const key in potentialHeaders) {
      const index = row.findIndex(header => potentialHeaders[key].includes(header));
      tempMapping[key] = index !== -1 ? index : null;
    }
    if (tempMapping.employeeNo !== null && tempMapping.name !== null && tempMapping.date !== null) {
      headerRow = i;
      mapping = tempMapping;
      break;
    }
  }

  // ✅ Step 2: If no header row found → pattern recognition
  if (headerRow === -1) {
    const classifyValue = (v) => {
      v = v.trim();
      if (/^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/.test(v)) return "date";
      if (/^(sun(day)?|mon(day)?|tue(s|sday)?|wed(nesday)?|thu(rs|rsday)?|fri(day)?|sat(urday)?)$/i.test(v)) return "dayOfWeek";
      if (/^\d{1,6}$/.test(v)) return "employeeNo";
      if (/^[A-Za-z]{3}-\d{3}$/.test(v)) return "shiftCode";
      if (/^(branch\s*head|oic|ia|supervisor|manager|assistant|lead)$/i.test(v)) return "position";
      if (/^[A-Za-z\s]+$/.test(v)) return "name";
      return "unknown";
    };

    const sampleRows = rows.slice(0, 5);
    const colCount = Math.max(...sampleRows.map(r => r.length));
    const colTypeScores = Array.from({ length: colCount }, (_, i) => {
      const values = sampleRows.map(r => r[i]?.trim()).filter(Boolean);
      const typeCounts = {};
      for (const val of values) {
        const type = classifyValue(val);
        if (type !== "unknown") typeCounts[type] = (typeCounts[type] || 0) + 1;
      }
      const bestType = Object.entries(typeCounts).sort((a, b) => b[1] - a[1])[0]?.[0] || "unknown";
      return { index: i, bestType };
    });

    const findCol = (type) => colTypeScores.find(c => c.bestType === type)?.index ?? null;

    mapping = {
      employeeNo: findCol("employeeNo"),
      name: findCol("name"),
      position: findCol("position"),
      date: findCol("date"),
      shiftCode: isWork ? findCol("shiftCode") : null,
      dayOfWeek: findCol("dayOfWeek")
    };

    // ✅ Fallback trigger for short 1–2 column pastes
    const nonNullCount = Object.values(mapping).filter(v => v !== null).length;
    const avgCols = rows.reduce((a, r) => a + r.length, 0) / rows.length;
    if (nonNullCount <= 1 || avgCols < 3) {
      mapping = {}; // force fallback mode
      headerRow = -1;
    }
  }

  return { mapping, headerRow };
}

function recheckConflicts() {
  if (!Array.isArray(workScheduleData) || !Array.isArray(restDayData)) return;

  const LEADERSHIP_POSITIONS = ['Branch Head', 'Site Supervisor', 'OIC'];
  const scheduleMap = new Map();

  // Reset all conflicts
  [...workScheduleData, ...restDayData].forEach(d => {
    d.conflict = false;
    d.conflictReasons = [];
    d.conflictReason = '';
    d.conflictType = '';
  });

  // --- 1️⃣ Same employee/date across WS and RD ---
  workScheduleData.forEach((w, wIdx) => {
    if (!w.employeeNo || !w.date) return;
    const key = `${w.employeeNo}-${w.date}`;
    scheduleMap.set(key, { rowNum: wIdx + 1, ...w });
  });

  restDayData.forEach((r, rIdx) => {
    if (!r.employeeNo || !r.date) return;
    const key = `${r.employeeNo}-${r.date}`;
    if (scheduleMap.has(key)) {
      const w = scheduleMap.get(key);
      r.conflict = true;
      r.conflictType = 'sameDate';
      r.conflictReasons.push(`Has rest day on same date as Work Schedule (WS row #${w.rowNum})`);
      const match = workScheduleData.find(d => d.employeeNo === r.employeeNo && d.date === r.date);
      if (match) {
        match.conflict = true;
        match.conflictType = 'sameDate';
        match.conflictReasons.push(`Has work schedule on same date as Rest Day (RD row #${rIdx + 1})`);
      }
    }
  });

  // --- 2️⃣ Duplicate dates within WS or RD ---
  const detectDuplicates = (data, label) => {
    const seen = new Map();
    data.forEach((d, idx) => {
      if (!d.employeeNo || !d.date) return;
      const key = `${d.employeeNo}-${d.date}`;
      if (seen.has(key)) {
        const other = seen.get(key);
        d.conflict = true;
        d.conflictType = 'duplicate';
        other.conflict = true;
        other.conflictType = 'duplicate';
        d.conflictReasons.push(`Duplicate date in ${label} (row #${other._rowNum})`);
        other.conflictReasons.push(`Duplicate date in ${label} (row #${idx + 1})`);
      } else {
        d._rowNum = idx + 1;
        seen.set(key, d);
      }
    });
  };
  detectDuplicates(workScheduleData, 'Work Schedule');
  detectDuplicates(restDayData, 'Rest Day Schedule');

  // --- 3️⃣ RD employee not found in WS ---
  const workNos = new Set(workScheduleData.map(d => d.employeeNo));
  restDayData.forEach(r => {
    if (r.employeeNo && !workNos.has(r.employeeNo)) {
      r.conflict = true;
      r.conflictType ||= 'missing';
      r.conflictReasons.push(`Employee not found in Work Schedule`);
    }
  });

  // --- 4️⃣ Multiple leaders resting same day ---
  const byDate = {};
  restDayData.forEach(r => {
    if (!r.date || !r.position) return;
    if (!byDate[r.date]) byDate[r.date] = [];
    if (LEADERSHIP_POSITIONS.includes(r.position)) byDate[r.date].push(r);
  });
  Object.entries(byDate).forEach(([date, list]) => {
    if (list.length > 1) {
      list.forEach(ld => {
        ld.conflict = true;
        ld.conflictType ||= 'leadership';
        ld.conflictReasons.push(`Multiple leaders resting on ${date}`);
      });
    }
  });

// --- 5️⃣ Smarter Weekend RD Counting ---
const weekendDays = ['Friday', 'Saturday', 'Sunday'];
const weekendAdjacent = ['Thursday', 'Monday'];
const validWeekendDays = [...weekendDays, ...weekendAdjacent];

const weekendGroups = {}; // { empNo: { weekKey: [dates] } }

// Helper: get ISO week key
function getWeekKey(dateStr) {
  const date = new Date(dateStr);
  const year = date.getFullYear();
  const firstThursday = new Date(date);
  firstThursday.setDate(date.getDate() + 4 - (date.getDay() || 7));
  const weekNo = Math.ceil((((firstThursday - new Date(year, 0, 1)) / 86400000) + 1) / 7);
  return `${year}-W${weekNo}`;
}

// Helper: get day index
function getDayIndex(dayName) {
  const map = {
    Sunday: 0, Monday: 1, Tuesday: 2, Wednesday: 3,
    Thursday: 4, Friday: 5, Saturday: 6
  };
  return map[dayName] ?? -1;
}

// Group rest days per employee per week
restDayData.forEach(r => {
  if (!r.employeeNo || !r.date) return;
  const date = new Date(r.date);
  if (isNaN(date)) return;
  const empNo = r.employeeNo;
  const weekKey = getWeekKey(r.date);
  const dayName = date.toLocaleDateString('en-US', { weekday: 'long' });

  if (!weekendGroups[empNo]) weekendGroups[empNo] = {};
  if (!weekendGroups[empNo][weekKey]) weekendGroups[empNo][weekKey] = [];

  if (validWeekendDays.includes(dayName)) {
    weekendGroups[empNo][weekKey].push({
      day: dayName,
      date: r.date,
      index: getDayIndex(dayName)
    });
  }
});

// Count weekend rest groups
const weekendCount = {};

Object.entries(weekendGroups).forEach(([empNo, weeks]) => {
  let totalGroups = 0;

  Object.values(weeks).forEach(days => {
    if (days.length === 0) return;

    // Sort by weekday order
    days.sort((a, b) => a.index - b.index);

    // Merge consecutive or overlapping days into a single "weekend group"
    let currentGroup = [days[0]];

    for (let i = 1; i < days.length; i++) {
      const prev = days[i - 1];
      const curr = days[i];
      const diff = curr.index - prev.index;

      // Consecutive day or Fri→Mon type overlap
      if (diff === 1 || (prev.day === 'Friday' && curr.day === 'Monday')) {
        currentGroup.push(curr);
      } else {
        // Evaluate previous group before starting new one
        if (currentGroup.some(d => weekendDays.includes(d.day))) {
          totalGroups++;
        }
        currentGroup = [curr];
      }
    }

    // Count the last group if it contains a weekend day
    if (currentGroup.some(d => weekendDays.includes(d.day))) {
      totalGroups++;
    }
  });

  weekendCount[empNo] = totalGroups;
});

// Apply conflicts only if > 2 weekend RD groups total
Object.entries(weekendCount).forEach(([empNo, count]) => {
  if (count > 2) {
    restDayData
      .filter(r => r.employeeNo === empNo)
      .forEach(r => {
        const dayName = new Date(r.date).toLocaleDateString('en-US', { weekday: 'long' });
        if (validWeekendDays.includes(dayName)) {
          r.conflict = true;
          r.conflictType ||= 'weekend';
          r.conflictReason = '⚠️ Too many weekends';
          r.conflictReasons.push(`Exceeded limit: ${count} weekend rest groups (max 2 allowed)`);
        }
      });
  }
});

  // --- 6️⃣ Row short messages (only for display) ---
  const shortMessage = {
    sameDate: 'Same date conflict',
    duplicate: 'Duplicate date',
    missing: 'Not in WS',
    leadership: 'Leadership overlap',
    weekend: 'Too many weekends'
  };
  [...workScheduleData, ...restDayData].forEach(d => {
    if (d.conflict) {
      d.conflictReason = shortMessage[d.conflictType] || 'Conflict detected';
    }
  });

  // --- 7️⃣ Summary table (full detail) ---
  const allConflicts = [...workScheduleData, ...restDayData].filter(d => d.conflict);
  const summaryLines = allConflicts.map(d => {
    const type = workScheduleData.includes(d) ? 'WS' : 'RD';
    return `${type} | ${d.employeeNo || '(No EmpNo)'} ${d.name || ''} — ${d.conflictReasons.join('; ')}`;
  });

  try {
    renderSummary(allConflicts.length, summaryLines);
    renderWorkTable();
    renderRestTable();
  } catch (err) {
    console.error('⚠️ Render error:', err);
  }
}


function generateFile(data, fileNamePrefix) {
    if (data.length === 0) {
        showWarning('No data to generate file.');
        return;
    }

    // 🏷️ Identify type
    const isWorkSchedule = fileNamePrefix === 'WorkSchedule';
    const schedType = isWorkSchedule ? 'WS' : 'RD';

    // ✅ Use correct branch input field
    let branchInput;
    if (isWorkSchedule) {
        branchInput = document.getElementById('branchNameInput'); // for WS
    } else {
        branchInput = document.getElementById('branchNameRestInput'); // for RD
    }

    let branchName = '';
if (isWorkSchedule) {
  branchName = document.getElementById('branchNameInput').value || 'UnnamedBranch';
} else {
  branchName = document.getElementById('branchNameRestInput').value || 'UnnamedBranch';
}

    // ✅ Headers
    const headers = isWorkSchedule
        ? ['Employee Number', 'Work Date', 'Shift Code']
        : ['Employee No', 'Rest Day Date'];

    // ✅ Data formatting
    const formattedData = data.map(row => {
        if (isWorkSchedule) {
            return {
                'Employee Number': row.employeeNo,
                'Work Date': new Date(row.date),
                'Shift Code': row.shiftCode,
            };
        } else {
            return {
                'Employee No': row.employeeNo,
                'Rest Day Date': new Date(row.date),
            };
        }
    });

    // ✅ Convert to sheet (no header row like “UnnamedBranch_RD_October”)
    const sheet = XLSX.utils.json_to_sheet(formattedData);

    // ✅ Date formatting
    Object.keys(sheet).forEach(cell => {
        if (cell[0] === "!" || !sheet[cell].v) return;
        const val = sheet[cell].v;
        if (val instanceof Date || /^\d{4}-\d{2}-\d{2}/.test(val)) {
            sheet[cell].t = "d";
            sheet[cell].z = "mm/dd/yy";
        }
    });

    // 📘 Create workbook
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "Sheet1");

    // 💾 File naming
    const month = 'October'; // optional: can make this auto-detect later
    const filename = isWorkSchedule
        ? `${branchName}_WS - Work schedule.xlsx`
        : `${branchName}_RD - Rest day.xlsx`;

    // 💾 Export
    XLSX.writeFile(workbook, filename);
    showSuccess(`File "${filename}" generated successfully!`);
}

      /*************************\
       * UI & RENDERING     *
      \*************************/
       function renderTable(tbody, data, columns, type) {
           tbody.innerHTML = '';
           if (!data || data.length === 0) {
               const tr = tbody.insertRow();
               const cell = tr.insertCell();
               let colspan = columns.length + 2; // # and Actions
               if (type === 'rest') colspan++; // Add one for Conflict column
               cell.colSpan = colspan;
               cell.textContent = 'No data available.';
               cell.className = 'text-center text-slate-500 py-4';
               return;
           }
           data.forEach((item, index) => {
const tr = tbody.insertRow();

// 🟢 Apply conflict highlight color only once
let rowClass = '';
if (item.conflict) {
  switch (item.conflictType) {
    case 'sameDate':
      rowClass = 'conflict-samedate';
      break;
    case 'duplicate':
      rowClass = 'conflict-duplicate';
      break;
    case 'leadership':
      rowClass = 'conflict-leadership';
      break;
    case 'weekend':
      rowClass = 'conflict-weekend';
      break;
    case 'missing':
      rowClass = 'conflict-missing';
      break;
    default:
      rowClass = 'conflict';
  }
}
tr.className = rowClass;

               if (type === 'rest') {
                   const conflictCell = tr.insertCell();
                   if (item.conflict) {
                       conflictCell.innerHTML = `⚠️ <span class="conflict-reason">${item.conflictReason}</span>`;
                   }
                   conflictCell.className = 'text-center';
               }

               const numCell = tr.insertCell();
               numCell.textContent = index + 1;
               numCell.className = 'text-sm text-slate-600 text-center';

               columns.forEach(col => {
                   const cell = tr.insertCell();
                   cell.textContent = item[col] || '';
               });

               const actionsCell = tr.insertCell();
               actionsCell.className = 'flex items-center justify-center space-x-3';
               actionsCell.innerHTML = `
                <ion-icon name="create-outline" class="action-icon edit-icon text-xl" data-type="${type}" data-index="${index}"></ion-icon>
                <ion-icon name="trash-outline" class="action-icon delete-icon text-xl" data-type="${type}" data-index="${index}"></ion-icon>
               `;
           });

            tbody.querySelectorAll('.edit-icon').forEach(btn => btn.addEventListener('click', (e) => handleEditRow(e.target.dataset.type, e.target.dataset.index)));
            tbody.querySelectorAll('.delete-icon').forEach(btn => btn.addEventListener('click', (e) => handleDeleteRow(e.target.dataset.type, e.target.dataset.index)));
       }

       function renderWorkTable() {
           renderTable(workTableBody, workScheduleData, ['employeeNo', 'name', 'position', 'date', 'shiftCode', 'dayOfWeek'], 'work');
       }

       function renderRestTable() {
           renderTable(restTableBody, restDayData, ['employeeNo', 'name', 'position', 'date', 'dayOfWeek'], 'rest');
       }
      
      function renderSummary(conflictCount, summaryLines) {
  if (workScheduleData.length === 0 && restDayData.length === 0) {
    summaryEl.innerHTML = `
      <h2 class="text-2xl font-bold mb-4 text-slate-800">Summary & Conflicts</h2>
      <p id="summaryText">No data pasted yet.</p>`;
    return;
  }

  let html = `<h2 class="text-2xl font-bold mb-4 text-slate-800">Summary & Conflicts</h2>`;

  if (conflictCount > 0) {
    // Limit visible conflicts to 6
    const visibleItems = summaryLines.slice(0, 6);
    const hiddenItems = summaryLines.slice(6);

    html += `
      <div class="bg-red-100 border-l-4 border-red-500 text-red-700 p-4 rounded-lg mb-4">
        <p class="font-bold mb-1">⚠️ Conflicts Detected (${conflictCount})</p>
        <p class="mb-2">Review the highlighted rows for issues. Detailed breakdown below:</p>

        <div id="summaryListWrapper" class="max-h-60 overflow-y-auto transition-all duration-300">
          <ul id="summaryList" class="list-disc ml-6 text-sm leading-snug space-y-1">
            ${visibleItems.map(line => `<li>${line}</li>`).join('')}
            ${
              hiddenItems.length > 0
                ? `<div id="hiddenSummary" class="hidden">
                    ${hiddenItems.map(line => `<li>${line}</li>`).join('')}
                   </div>`
                : ''
            }
          </ul>
        </div>

        ${
          hiddenItems.length > 0
            ? `<button id="toggleSummaryBtn"
                class="mt-3 text-sm font-medium text-cyan-700 hover:text-cyan-900 transition">
                See more ▼
              </button>`
            : ''
        }
      </div>`;
  } else {
    html += `
      <div class="bg-green-100 border-l-4 border-green-500 text-green-700 p-4 rounded-lg">
        <p class="font-bold">✅ No Conflicts Found</p>
        <p>All schedules appear valid.</p>
      </div>`;
  }

  summaryEl.innerHTML = html;

  // Handle See More / See Less toggle
  const toggleBtn = document.getElementById('toggleSummaryBtn');
  const hiddenSummary = document.getElementById('hiddenSummary');

  if (toggleBtn && hiddenSummary) {
    let expanded = false;
    toggleBtn.addEventListener('click', () => {
      expanded = !expanded;
      hiddenSummary.classList.toggle('hidden', !expanded);
      toggleBtn.textContent = expanded ? 'See less ▲' : 'See more ▼';
    });
  }
}

      function updateButtonStates() {
        generateWorkFileBtn.disabled = workScheduleData.length === 0;
        clearWorkBtn.disabled = workScheduleData.length === 0;
        undoWorkBtn.disabled = undoStack.work.length === 0;
        redoWorkBtn.disabled = redoStack.work.length === 0;

        generateRestFileBtn.disabled = restDayData.length === 0;
        clearRestBtn.disabled = restDayData.length === 0;
        undoRestBtn.disabled = undoStack.rest.length === 0;
        redoRestBtn.disabled = redoStack.rest.length === 0;
      }

      function switchTab(tab) {
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
           const item = dataArray[index];

           document.getElementById('editEmployeeNo').value = item.employeeNo || '';
           document.getElementById('editName').value = item.name || '';
           document.getElementById('editPosition').value = item.position || '';
           document.getElementById('editDate').value = item.date || '';
           document.getElementById('editDayOfWeek').value = item.dayOfWeek || '';
           
           if(type === 'work') {
               document.getElementById('editShiftCode').value = item.shiftCode || '';
               editShiftCodeWrapper.style.display = 'block';
           } else {
               editShiftCodeWrapper.style.display = 'none';
           }
           
           showEditModal();
       }

       function handleSaveEdit(event) {
            event.preventDefault();
            const {type, index} = currentlyEditing;
            if(type === null || index === null) return;
            
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

            if(type === 'work') {
                updatedItem.shiftCode = document.getElementById('editShiftCode').value;
            }

            dataArray[index] = updatedItem;

            hideEditModal();
            recheckConflicts();
            saveState();
       }
       
       function showEditModal() { editModal.classList.remove('hidden'); editModal.style.display = 'flex'; }
       function hideEditModal() { editModal.classList.add('hidden'); editModal.style.display = 'none'; currentlyEditing = {type: null, index: null};}

       function saveUndoState(type) {
           const data = type === 'work' ? workScheduleData : restDayData;
           undoStack[type].push(JSON.parse(JSON.stringify(data)));
           redoStack[type] = [];
           if(undoStack[type].length > 10) undoStack[type].shift();
           updateButtonStates();
       }

       function undo(type) {
           if (undoStack[type].length === 0) return;
           const previousState = undoStack[type].pop();
           const currentState = type === 'work' ? workScheduleData : restDayData;
           redoStack[type].push(JSON.parse(JSON.stringify(currentState)));
           
           if(type === 'work') { workScheduleData = previousState; } 
           else { restDayData = previousState; }
           recheckConflicts();
           updateButtonStates();
       }

       function redo(type) {
           if (redoStack[type].length === 0) return;
           const nextState = redoStack[type].pop();
           const currentState = type === 'work' ? workScheduleData : restDayData;
           undoStack[type].push(JSON.parse(JSON.stringify(currentState)));

           if(type === 'work') { workScheduleData = nextState; } 
           else { restDayData = nextState; }
           recheckConflicts();
           updateButtonStates();
       }
       
       function clearData(type) {
           if (!confirm(`Are you sure you want to clear all ${type} data?`)) return;
           saveUndoState(type);
           if (type === 'work') {
               workScheduleData = [];
               workInput.value = '';
           } else {
               restDayData = [];
               restInput.value = '';
           }
           recheckConflicts();
           updateButtonStates();
       }

        /*********************************\
        * MONITORING DASHBOARD (REALTIME) *
        \*********************************/
        function listenForMonitoringUpdates() {
            if (!monitoringCollectionRef) return;
            unsubscribeMonitoring(); // Detach any old listener
            unsubscribeMonitoring = onSnapshot(monitoringCollectionRef, (snapshot) => {
                const serverData = [];
                snapshot.forEach((doc) => {
                    serverData.push({ id: doc.id, ...doc.data() });
                });
                monitoringData = serverData.sort((a,b) => (a.posCode || '').localeCompare(b.posCode || ''));
                renderMonitoringDashboard();
            }, (error) => {
                console.error("Error listening to monitoring updates:", error);
                showWarning("Real-time connection lost. Please refresh.");
            });
        }

        function renderMonitoringDashboard() {
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
            
            monitoringTableBody.querySelectorAll('input[type="text"]').forEach(input => {
                input.addEventListener('change', (e) => updateMonitoringBranch(e.target.dataset.id, e.target.dataset.field, e.target.value));
            });
            monitoringTableBody.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
                checkbox.addEventListener('change', (e) => updateMonitoringBranch(e.target.dataset.id, e.target.dataset.field, e.target.checked));
            });
            monitoringTableBody.querySelectorAll('.delete-icon').forEach(button => {
                 button.addEventListener('click', (e) => deleteMonitoringBranch(e.target.dataset.id));
            });
            updateMonitoringProgress();
        }

        async function addMonitoringBranch() {
            const newBranch = {posCode: '', branchName: '', sapCode: '', isUploaded: false, uploadedBy: '', uploadedDate: ''};
            try {
                await addDoc(monitoringCollectionRef, newBranch);
            } catch (error) {
                console.error("Error adding branch:", error);
                showWarning("Could not add branch.");
            }
        }

        async function deleteMonitoringBranch(id) {
            if(!confirm('Are you sure you want to remove this branch from the tracker?')) return;
            try {
                await deleteDoc(doc(db, monitoringCollectionRef.path, id));
            } catch (error) {
                console.error("Error deleting branch:", error);
                showWarning("Could not delete branch.");
            }
        }

        async function updateMonitoringBranch(id, field, value) {
            const updateData = { [field]: value };
            
            if(field === 'isUploaded') {
                updateData.uploadedBy = value ? (auth.currentUser?.email || 'User') : '';
                updateData.uploadedDate = value ? new Date().toLocaleDateString() : '';
            }

            try {
                await updateDoc(doc(db, monitoringCollectionRef.path, id), updateData);
            } catch (error) {
                console.error("Error updating branch:", error);
                showWarning("Could not save changes.");
            }
        }
        
        function updateMonitoringProgress() {
            const total = monitoringData.length;
            if(total === 0) {
                 monitoringProgressBar.style.width = '0%';
                 monitoringProgressText.textContent = 'N/A';
                 return;
            }
            const uploadedCount = monitoringData.filter(b => b.isUploaded).length;
            const percentage = Math.round((uploadedCount / total) * 100);
            monitoringProgressBar.style.width = `${percentage}%`;
            monitoringProgressText.textContent = `${percentage}% (${uploadedCount}/${total})`;
        }



      /*************************\
       * UTILITY FUNCTIONS   *
      \*************************/
      function excelDateToJS(excelDate) {
  if (!excelDate) return '';

  // ✅ 1. Already a Date object
  if (excelDate instanceof Date && !isNaN(excelDate)) {
    return excelDate.toLocaleDateString();
  }

  // ✅ 2. Numeric Excel serial date (e.g., 45678)
  const a = parseFloat(excelDate);
  if (!isNaN(a) && a > 20000) {
    const date = new Date((a - 25569) * 86400 * 1000);
    return date.toLocaleDateString();
  }

  // ✅ 3. Text-based formats (handles "22-Nov", "Nov 22", "11/22/25", etc.)
  let parsedDate = null;
  const str = excelDate.toString().trim();

  // Match patterns like "22-Nov" or "22 Nov"
  const shortMonthMatch = str.match(/^(\d{1,2})[-\s]([A-Za-z]{3,})$/);
  if (shortMonthMatch) {
    const [_, day, mon] = shortMonthMatch;
    const currentYear = new Date().getFullYear();
    parsedDate = new Date(`${mon} ${day}, ${currentYear}`);
  }

  // Match patterns like "Nov 22, 25" or "Nov 22 2025"
  if (!parsedDate || isNaN(parsedDate)) {
    const flexibleDate = Date.parse(str.replace(/(\d{1,2})-(\w{3,})/, '$2 $1'));
    if (!isNaN(flexibleDate)) parsedDate = new Date(flexibleDate);
  }

  // Match common numeric formats
  if ((!parsedDate || isNaN(parsedDate)) && /[-/]/.test(str)) {
    const normalized = str.replace(/-/g, '/');
    parsedDate = new Date(normalized);
  }

  // Fallback: try native Date()
  if (!parsedDate || isNaN(parsedDate)) {
    parsedDate = new Date(str);
  }

  // ✅ If still invalid, just return raw text (so it doesn’t break parsing)
  if (!parsedDate || isNaN(parsedDate)) {
    console.warn('⚠️ Could not parse date:', excelDate);
    return excelDate;
  }

  return parsedDate.toLocaleDateString();
}
      
      function jsDateToExcel(jsDateStr) {
        const date = new Date(jsDateStr);
        const excelEpoch = new Date(Date.UTC(1899, 11, 30));
        return (date - excelEpoch) / (24 * 60 * 60 * 1000);
      }
      function showWarning(message) {
        warningBanner.textContent = message;
        warningBanner.classList.remove('hidden');
        gsap.to(warningBanner, { opacity: 1, duration: 0.5 });
        setTimeout(() => {
          gsap.to(warningBanner, { opacity: 0, duration: 0.5, onComplete: () => warningBanner.classList.add('hidden') });
        }, 4000);
      }
      function showSuccess(message) {
        successMsg.textContent = message;
        successMsg.classList.remove('hidden');
        gsap.to(successMsg, { opacity: 1, duration: 0.5 });
        setTimeout(() => {
          gsap.to(successMsg, { opacity: 0, duration: 0.5, onComplete: () => successMsg.classList.add('hidden') });
        }, 3000);
      }
      
      if (backToTopBtn) {
        window.addEventListener('scroll', () => { window.scrollY > 400 ? backToTopBtn.classList.add('show') : backToTopBtn.classList.remove('show'); });
        backToTopBtn.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));
      }

      /***** LOCAL STORAGE FOR SCHEDULES *****/
      function saveState() {
        try {
          localStorage.setItem('workScheduleData_v3', JSON.stringify(workScheduleData));
          localStorage.setItem('restDayData_v3', JSON.stringify(restDayData));
        } catch (e) { console.warn("Could not save schedule state.", e); }
      }
      function loadState() {
        try {
          const w = JSON.parse(localStorage.getItem('workScheduleData_v3') || '[]');
          const r = JSON.parse(localStorage.getItem('restDayData_v3') || '[]');
          workScheduleData = Array.isArray(w) ? w : [];
          restDayData = Array.isArray(r) ? r : [];
        } catch (e) { console.error("Could not load schedule state.", e); }
      }
      
      ['paste', 'click'].forEach(evt => {
        workInput.addEventListener(evt, saveState);
        restInput.addEventListener(evt, saveState);
        clearWorkBtn.addEventListener(evt, saveState);
        clearRestBtn.addEventListener(evt, saveState);
      });

      /***** Parallax Header & Particles *****/
      const header = document.querySelector('header');
      if (header && window.matchMedia('(min-width: 1024px)').matches) {
        document.addEventListener('mousemove', (e) => {
          const x = (e.clientX / window.innerWidth - 0.5) * -3;
          const y = (e.clientY / window.innerHeight - 0.5) * -3;
          gsap.to(header, { duration: 1.5, backgroundPosition: `${50 + x}% ${50 + y}%`, ease: "power2.out" });
        });
      }

      const particleContainer = document.getElementById('particle-container');
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
// Function to create a reference for the selected month and year
const progressRef = (month, year) => {
    return collection(db, "monitoring", "_meta", "progress", `${year}-${month}`, "entries");
};

// Function to load data for the selected month and year
async function loadMonthData() {
    const month = document.getElementById("monthPicker").value;
    const year = document.getElementById("yearPicker").value;

    // Log selected month and year
    console.log("Fetching data for:", month, year);

    // Fetch data from Firestore based on selected month and year
    const progressDataRef = progressRef(month, year); // Firestore reference
    const snapshot = await getDocs(progressDataRef); // Fetch data
    const progressData = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));

    // Render the data in the table
    renderMonitoringTable(progressData);

    // If no data is found, show an alert
    if (progressData.length === 0) {
        alert("No data found for this month/year");
    }
}

// Event listener to trigger loadMonthData() when the "Load Data" button is clicked
document.getElementById("loadMonthData").addEventListener("click", loadMonthData);

// Function to export data to Excel
function exportToExcel(progressData) {
    const month = document.getElementById("monthPicker").value;
    const year = document.getElementById("yearPicker").value;

    // Log exporting data
    console.log('Exporting data:', progressData);

    // Create a new workbook and add a worksheet
    const ws = XLSX.utils.json_to_sheet(progressData);
    const wb = XLSX.utils.book_new();

    // Set a header with the month and year
    const header = `Branch Upload Monitoring Report\nMonth: ${month} ${year}\n\n`;

    // Add the header to the worksheet
    const wsHeader = XLSX.utils.aoa_to_sheet([[header]]);
    XLSX.utils.book_append_sheet(wb, wsHeader, "Header");
    XLSX.utils.book_append_sheet(wb, ws, "Progress Data");

    // Write the Excel file
    XLSX.writeFile(wb, `Monitoring_Report_${year}_${month}.xlsx`);
}

// Event listener for the export button
document.getElementById("exportExcel").addEventListener("click", function() {
    // Load the progress data from Firestore for the selected month/year before exporting
    const month = document.getElementById("monthPicker").value;
    const year = document.getElementById("yearPicker").value;
    const progressDataRef = progressRef(month, year); // Get Firestore reference

    // Fetch data from Firestore and pass it to exportToExcel
    getDocs(progressDataRef).then(snapshot => {
        const progressData = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        exportToExcel(progressData); // Export the fetched data to Excel
    });
});