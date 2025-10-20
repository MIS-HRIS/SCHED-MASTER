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
      workInput.addEventListener('paste', handlePaste);
restInput.addEventListener('paste', handlePaste);
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

generateWorkFileBtn.addEventListener('click', () => {
  const branchName = document.getElementById('branchNameInput')?.value.trim();
  if (!branchName) {
    alert('⚠️ Please enter the Branch Name before generating the Work File.');
    return;
  }
  generateFile(workScheduleData, 'WorkSchedule');
});

generateRestFileBtn.addEventListener('click', () => {
  const branchName = document.getElementById('branchNameInput')?.value.trim();
  if (!branchName) {
    alert('⚠️ Please enter the Branch Name before generating the Rest Day File.');
    return;
  }
  generateFile(restDayData, 'RestDaySchedule');
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

function handlePaste(event) {
  event.preventDefault();

  // ✅ Require branch name before pasting
  const branchName = document.getElementById("branchNameInput")?.value.trim();
  if (!branchName) {
    alert("⚠️ Please enter the Branch Name before pasting schedule data.");
    return;
  }

  const text = (event.clipboardData || window.clipboardData).getData("text");
  const isWork = event.target.id === "workScheduleInput";
  const type = isWork ? "work" : "rest";

  saveUndoState(type);

  const rows = text.split("\n").map((row) => row.split("\t"));
  const { mapping, headerRow } = detectColumnMapping(rows, isWork);

  const data = rows
    .slice(headerRow + 1)
    .map((row) => {
      const entry = {};
      for (const key in mapping) {
        if (mapping[key] !== null && row[mapping[key]] !== undefined) {
          entry[key] = row[mapping[key]].trim();
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
    }));
  } else {
    restDayData = data.map((d) => ({
      ...d,
      date: excelDateToJS(d.date),
    }));
  }

  recheckConflicts();
  updateButtonStates();
  renderWorkTable();
  renderRestTable();

  // ✅ Wait until next frame (DOM updated), then scroll to table
  setTimeout(() => {
    const target = isWork
      ? document.querySelector("#workTableBody")
      : document.querySelector("#restTableBody");
    if (target) {
      target.scrollIntoView({ behavior: "smooth", block: "center" });
    }
  }, 300);
}

      function recheckConflicts() {
          const scheduleMap = new Map();
          
          workScheduleData.forEach(d => d.conflict = false);
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
              .some(d => d.conflict && LEADERSHIP_POSITIONS.includes(d.position));

          renderSummary(conflictCount, leadershipConflict);
          renderWorkTable();
          renderRestTable();
      }
      
function generateFile(data, fileNamePrefix) {
  if (data.length === 0) {
    showWarning('No data to generate file.');
    return;
  }

  // 🏷️ Require branch name first
  const branchName = document.getElementById('branchNameInput')?.value.trim();
  if (!branchName) {
    alert('⚠️ Please enter the Branch Name before generating the file.');
    return;
  }

  // 🗓️ Auto month + year
  const now = new Date();
  const month = now.toLocaleString('default', { month: 'short' }); // e.g., Oct
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

  // 🧾 Create Excel worksheet
  const worksheet = XLSX.utils.json_to_sheet(formattedData);

  // ✅ Format Excel date column
  Object.keys(worksheet).forEach(cell => {
    if (cell[0] === "!" || !worksheet[cell].v) return;
    const val = worksheet[cell].v;
    if (val instanceof Date || /^\d{4}-\d{2}-\d{2}/.test(val)) {
      worksheet[cell].t = "d";
      worksheet[cell].z = "mm/dd/yy";
    }
  });

  // 📘 Finalize workbook
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

  // 💾 Filename format → BranchName_FileType_MonthYear.xlsx
  const safeBranch = branchName.replace(/[^\w\s-]/g, '').replace(/\s+/g, '_'); // sanitize
  const fileName = `${safeBranch}_${fileNamePrefix}_${month}${year}.xlsx`;

  XLSX.writeFile(workbook, fileName);
  showSuccess(`File generated successfully: ${fileName}`);
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
               tr.className = item.conflict ? 'conflict' : '';

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
      
      function renderSummary(conflictCount, leadershipConflict) {
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

       /* ===========================
   MONITORING DASHBOARD FIXES
   — Load, Export, Auto-Month
=========================== */

// ✅ Firestore imports (already available via firebase.js)
import { getFirestore, collection, getDocs, onSnapshot, addDoc, updateDoc, deleteDoc, doc } from "./firebase.js";
const db = getFirestore();

const monitoringCollectionRef = collection(db, "monitoring");


// --- Auto-select current month/year when page opens ---
window.addEventListener("DOMContentLoaded", () => {
  const now = new Date();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const year = now.getFullYear();

  const monthPicker = document.getElementById("monthPicker");
  const yearPicker = document.getElementById("yearPicker");

  if (monthPicker && yearPicker) {
    monthPicker.value = month;
    yearPicker.value = year;
    // Automatically load data for current month
    document.getElementById("loadMonthData")?.click();
  }
});

// --- Load Data button ---
document.getElementById("loadMonthData")?.addEventListener("click", async () => {
  const month = document.getElementById("monthPicker").value;
  const year = document.getElementById("yearPicker").value;

  try {
    const monitoringRef = collection(db, "monitoring");
    const snapshot = await getDocs(monitoringRef);
    const tableBody = document.getElementById("monitoringTableBody");
    tableBody.innerHTML = "";

    snapshot.forEach((doc) => {
      const data = doc.data();
      if (data.month === month && data.year === year) {
        const row = `
          <tr>
            <td>${data.posCode || ""}</td>
            <td>${data.branchName || ""}</td>
            <td>${data.sapCode || ""}</td>
            <td class="text-center">${data.uploaded ? "✅" : "❌"}</td>
            <td>${data.uploadedBy || ""}</td>
            <td>${data.uploadedDate || ""}</td>
            <td><button class="action-btn danger" data-id="${doc.id}">🗑️ Delete</button></td>
          </tr>
        `;
        tableBody.insertAdjacentHTML("beforeend", row);
      }
    });

    updateMonitoringProgress();
  } catch (err) {
    console.error("Error loading monitoring data:", err);
  }
});

// --- Export button ---
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
});

// --- Progress bar update helper ---
function updateMonitoringProgress() {
  const rows = document.querySelectorAll("#monitoringTableBody tr");
  if (!rows.length) return;

  const total = rows.length;
  const completed = [...rows].filter((r) => r.cells[3].innerText === "✅").length;
  const percent = Math.round((completed / total) * 100);

  const bar = document.getElementById("monitoringProgressBar");
  const text = document.getElementById("monitoringProgressText");
  if (bar && text) {
    bar.style.width = `${percent}%`;
    text.textContent = `${percent}%`;
  }
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
            updateMonitoringProgressRealtime();
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
