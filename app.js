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
      const importScheduleFiles = document.getElementById('importScheduleFiles');
      const addScheduleFilesBtn = document.getElementById('addScheduleFilesBtn');
const addScheduleFilesInput = document.getElementById('addScheduleFilesInput');
const importScheduleBtn = document.getElementById('importScheduleBtn');
const generateImportedBtn = document.getElementById('generateImportedBtn');
const importSummaryPanel = document.getElementById('importSummaryPanel');
const importSummaryText = document.getElementById('importSummaryText');
const importSummaryBadge = document.getElementById('importSummaryBadge');
const importSummaryList = document.getElementById('importSummaryList');
const importConflictActions = document.getElementById('importConflictActions');
const removeConflictFilesBtn = document.getElementById('removeConflictFilesBtn');
const continueConflictFilesBtn = document.getElementById('continueConflictFilesBtn');
const removeAllImportedFilesBtn = document.getElementById('removeAllImportedFilesBtn');

let importedFiles = [];
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

importScheduleBtn.addEventListener('click', function () {
  importScheduleFiles.value = '';
  importScheduleFiles.click();
});

addScheduleFilesBtn.addEventListener('click', function () {
  addScheduleFilesInput.value = '';
  addScheduleFilesInput.click();
});

importScheduleFiles.addEventListener('change', async (event) => {
  await handleImportFiles(event, false);
});

addScheduleFilesInput.addEventListener('change', async (event) => {
  await handleImportFiles(event, true);
});

generateImportedBtn.addEventListener('click', generateImportedData);

removeConflictFilesBtn.addEventListener('click', () => {
  importedFiles = importedFiles.filter(file => file.conflicts.length === 0);
  generateImportedBtn.disabled = importedFiles.length === 0;

if (importedFiles.length === 0) {
  generateImportedBtn.classList.add('hidden');
}
if (importedFiles.length === 0) {
  addScheduleFilesBtn.classList.add('hidden');
}
  renderImportSummaryDashboard();
  showWarning('Conflicted file(s) removed.');
});

removeAllImportedFilesBtn.addEventListener('click', () => {
  importedFiles = [];
  renderImportSummaryDashboard();
  showWarning('All imported file(s) removed.');
});

continueConflictFilesBtn.addEventListener('click', () => {
  importConflictActions.classList.add('hidden');
  showSuccess('Conflicted file(s) kept. You may now generate.');
});
      
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

function detectImportedFileConflicts(files, isWork) {
  const conflictFiles = [];

  files.forEach(file => {
    const reasons = [];
    const data = parseImportedRows(file.rows, isWork).map(d => ({
      ...d,
      date: excelDateToJS(d.date)
    }));

    const seen = new Set();

    data.forEach(row => {
      if (!row.employeeNo || !row.date) return;

      const key = `${row.employeeNo}-${row.date}`;

      if (seen.has(key)) {
        reasons.push(`Duplicate employee/date found: ${row.employeeNo} - ${row.date}`);
      }

      seen.add(key);

      if (isWork && !row.shiftCode) {
        reasons.push(`Missing shift code: ${row.employeeNo} - ${row.date}`);
      }
    });

    if (reasons.length > 0) {
      conflictFiles.push({
        fileName: file.fileName,
        reasons: [...new Set(reasons)]
      });
    }
  });

  return conflictFiles;
}

function parseImportedRows(rows, isWork) {
  const { mapping, headerRow } = detectColumnMapping(rows, isWork);

  let data;

  if (headerRow !== -1 && Object.values(mapping).some(v => v !== null)) {
    data = rows.slice(headerRow + 1).map(row => {
      const entry = {};

      for (const key in mapping) {
        entry[key] =
          mapping[key] !== null && row[mapping[key]] !== undefined
            ? String(row[mapping[key]]).trim()
            : '';
      }

      return entry;
    });
  } else {
    data = rows.map(rawRow => {
      const tokens = rawRow.join(' ').split(/\s+/).filter(Boolean);

      const entry = {
        nameParts: [],
        positionParts: []
      };

      tokens.forEach(t => {
        const value = t.trim();
        const upperValue = value.toUpperCase();

        if (/^\d{1,6}$/.test(value)) {
          if (!entry.employeeNo) entry.employeeNo = value;
        } else if (/^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/.test(value) || /^\d{4}-\d{1,2}-\d{1,2}$/.test(value)) {
          if (!entry.date) entry.date = value;
        } else if (/^(sun(day)?|mon(day)?|tue(s|sday)?|wed(nesday)?|thu(rs|rsday)?|fri(day)?|sat(urday)?)$/i.test(value)) {
          if (!entry.dayOfWeek) entry.dayOfWeek = value;
        } else if (/^(?:AASP|RBG|RBT|WHSE|CHAR)-\d{3}(?:\s*\([^)]*\))?$/i.test(upperValue)) {
          if (!entry.shiftCode) entry.shiftCode = upperValue;
        } else if (/^(cashier|manager|supervisor|assistant|oic|head|lead|ia|mac|expert|branch\s*head|site\s*supervisor)$/i.test(value)) {
          entry.positionParts.push(value);
        } else if (/^[A-Za-z]+$/.test(value)) {
          entry.nameParts.push(value);
        }
      });

      if (entry.nameParts.length > 0) entry.name = entry.nameParts.join(' ');
      if (entry.positionParts.length > 0) entry.position = entry.positionParts.join(' ');

      delete entry.nameParts;
      delete entry.positionParts;

      return entry;
    });
  }

  return data;
}

function detectDominantMonthYearFromRows(rows) {
  const counts = {};

  rows.flat().forEach(cell => {
    const value = String(cell || '').trim();

    const match = value.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](20\d{2})$/);
    if (!match) return;

    const [, month, , year] = match;
    const key = `${Number(month)}-${year}`;

    counts[key] = (counts[key] || 0) + 1;
  });

  const top = Object.entries(counts).sort((a, b) => b[1] - a[1])[0];

  if (!top) return null;

  const [month, year] = top[0].split('-');

  return {
    month: Number(month),
    year: Number(year)
  };
}

function parseMixedScheduleRows(rows, dateContext = null) {
  const parsed = [];

  let workHeader = null;
  let restHeader = null;

  const normalize = (v) =>
    String(v || '').toUpperCase().replace(/\s+/g, ' ').trim();

  const isShiftCode = (v) =>
    /^(?:AASP|RBG|RBT|WHSE|CHAR)-\d{3}/i.test(String(v || '').trim());

  const isDate = (v) => {
    const value = String(v || '').trim();
    if (!value) return false;

    if (/^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/.test(value)) return true;
    if (/^[A-Za-z]+\s+\d{1,2},?\s+\d{4}$/i.test(value)) return true;
    if (!isNaN(parseFloat(value)) && parseFloat(value) > 20000) return true;

    return !isNaN(Date.parse(value)) && /[A-Za-z]/.test(value);
  };

  const isDay = (v) =>
    /^(sun(day)?|mon(day)?|tue(s|sday)?|wed(nesday)?|thu(rs|rsday)?|fri(day)?|sat(urday)?)$/i.test(String(v || '').trim());

const findHeaderIndexes = (row) => {
  const workHeaders = {
    name: null,
    employeeNo: null,
    date: null,
    shiftCode: null,
    dayOfWeek: null,
    position: null
  };

  const restHeaders = {
    name: null,
    employeeNo: null,
    date: null,
    dayOfWeek: null,
    position: null
  };

  row.forEach((cell, idx) => {
    const h = normalize(cell);

    if (h === 'NAME' || h === 'EMPLOYEE NAME') {
      if (workHeaders.name === null) workHeaders.name = idx;
      else if (restHeaders.name === null) restHeaders.name = idx;
    }

    else if (
      h.includes('EMPLOYEE NO') ||
      h.includes('EMPLOYEE NUMBER') ||
      h.includes('EMP NO')
    ) {
      if (workHeaders.employeeNo === null) workHeaders.employeeNo = idx;
      else if (restHeaders.employeeNo === null) restHeaders.employeeNo = idx;
    }

    else if (h.includes('WORK DATE')) {
      workHeaders.date = idx;
    }

    else if (
      h.includes('REST DAY DATE') ||
      h.includes('RD DATE')
    ) {
      restHeaders.date = idx;
    }

    else if (
      h.includes('SHIFT CODE') ||
      h.includes('SHIFTCODE') ||
      h.includes('SHIFT CODES') ||
      h.includes('SHIFTCODES') ||
      h.includes('SCHED CODE')
    ) {
      workHeaders.shiftCode = idx;
    }

    else if (h.includes('DAY OF WEEK')) {
      if (workHeaders.dayOfWeek === null) workHeaders.dayOfWeek = idx;
      else if (restHeaders.dayOfWeek === null) restHeaders.dayOfWeek = idx;
    }

    else if (h === 'POSITION' || h.includes('BRANCH ASSIGNMENT')) {
      if (workHeaders.position === null) workHeaders.position = idx;
      else if (restHeaders.position === null) restHeaders.position = idx;
    }
  });

  return {
    work: workHeaders,
    rest: restHeaders
  };
}

  const hasUsableHeader = (header) =>
    header.employeeNo !== null && header.date !== null;

  rows.forEach((row, index) => {
    const joined = normalize(row.join(' '));

const headers = findHeaderIndexes(row);
const workDetectedHeader = headers.work;
const restDetectedHeader = headers.rest;

if (hasUsableHeader(workDetectedHeader)) {
  workHeader = workDetectedHeader;
}

if (hasUsableHeader(restDetectedHeader)) {
  restHeader = restDetectedHeader;
}

const isHeaderRow =
  hasUsableHeader(workDetectedHeader) ||
  hasUsableHeader(restDetectedHeader);

if (isHeaderRow) {
  return;
}

    // If row contains two tables side-by-side, split by big blank gap
    const blankIndexes = row
      .map((cell, idx) => ({ cell: String(cell || '').trim(), idx }))
      .filter(x => x.cell === '')
      .map(x => x.idx);

const blocks = [];

if (workHeader) {
  blocks.push({
    type: 'work',
    offset: 0,
    row
  });
}

if (restHeader) {
  const restOffset = restHeader.employeeNo ?? restHeader.name ?? restHeader.date ?? 0;

  blocks.push({
    type: 'rest',
    offset: restOffset,
    row: row.slice(restOffset)
  });
}

if (!workHeader && !restHeader) {
  blocks.push({
    type: null,
    offset: 0,
    row
  });
}
blocks.forEach(block => {
      const activeHeader =
        block.type === 'work'
          ? workHeader
          : block.type === 'rest'
            ? restHeader
            : workHeader || restHeader;

      let entry = {
        rowNumber: index + 1,
        employeeNo: '',
        name: '',
        position: '',
        date: '',
        shiftCode: '',
        dayOfWeek: '',
        type: block.type || ''
      };

      if (activeHeader) {
        const getVal = (key) => {
          if (activeHeader[key] === null || activeHeader[key] === undefined) return '';
          const localIndex = activeHeader[key] - block.offset;
          return String(block.row[localIndex] || '').trim();
        };

        entry.name = getVal('name');
        entry.employeeNo = getVal('employeeNo');
        const rawDate = getVal('date');
entry.date = rawDate ? excelDateToJS(rawDate, dateContext) : '';
        entry.shiftCode = getVal('shiftCode').toUpperCase();
        entry.dayOfWeek = getVal('dayOfWeek');
        entry.position = getVal('position');
      } else {
        block.row.forEach(value => {
          value = String(value || '').trim();
          const upper = value.toUpperCase();

          if (/^\d{1,6}$/.test(value)) entry.employeeNo ||= value;
          else if (isShiftCode(upper)) entry.shiftCode ||= upper;
          else if (isDate(value)) entry.date ||= excelDateToJS(value, dateContext);
          else if (isDay(value)) entry.dayOfWeek ||= value;
          else if (/^(branch\s*head|oic|officer\s*in\s*charge|cashier|manager|assistant|lead|mac\s*expert|site\s*supervisor)$/i.test(value)) entry.position ||= value;
          else if (/^[A-Za-z,\s.-]+$/.test(value)) entry.name ||= value;
        });
      }

      if (!entry.type) {
        if (entry.shiftCode && entry.date) entry.type = 'work';
        else if (!entry.shiftCode && entry.date) entry.type = 'rest';
      }

      const currentYear = new Date().getFullYear();
const rowYear = new Date(entry.date).getFullYear();

const allowedYears = [
  currentYear,
  currentYear + 1
];

if (entry.date && !allowedYears.includes(rowYear)) {
  return;
}

      const isSampleRow =
  String(entry.employeeNo || '').trim() === '1010' ||
  String(entry.name || '').toUpperCase().includes('JUAN DELA CRUZ');

if (isSampleRow) {
  return;
}

// Clean the data before checking - SAFE VERSION
// Fallback to empty string || '' prevents crashes on empty cells
const cleanName = String(entry.name || '').trim();
const cleanEmpNo = String(entry.employeeNo || '').trim();
const cleanDate = String(entry.date || '').trim();
const cleanShift = String(entry.shiftCode || '').trim();

let hasUsefulData = false;

// REALISTIC CHECK: Accept row if at least Name and Date exist
if (cleanName !== '' && cleanDate !== '') {
    if (entry.type === 'work' && cleanEmpNo !== '' && cleanShift !== '') {
        hasUsefulData = true;
    } else if (entry.type === 'rest' && cleanEmpNo !== '') {
        hasUsefulData = true;
    }
}

if (hasUsefulData) {
    parsed.push(entry);
}
});
});

return parsed;
}