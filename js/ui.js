import { DAYS } from "./constants.js";
import { getSample, generateSchedule } from "./api.js";
import { downloadExcel } from "./excel.js";

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// App bar / drawer / sheet refs
const appBar = $(".app-bar");
const drawer = $("#drawer");
const drawerBackdrop = $("#drawer-backdrop");
const openDrawerBtn = $("#open-drawer-btn");
const closeDrawerBtn = $("#close-drawer-btn");
const emptyOpenDrawerBtn = $("#empty-open-drawer-btn");
const savedSheet = $("#saved-sheet");
const savedBackdrop = $("#saved-backdrop");
const showSavedBtn = $("#show-saved-btn");
const closeSavedBtn = $("#close-saved-btn");

// Form / input refs
const form = $("#schedule-form");
const studentsInput = $("#students-input");
const addStudentButton = $("#add-student-button");
const sampleButton = $("#sample-button");

// Schedule / action refs
const gridTarget = $("#grid");
const alertBanner = $("#alert-banner");
const emptyHint = $("#empty-hint");
const regenerateButton = $("#regenerate-button");
const confirmButton = $("#confirm-button");
const excelButton = $("#excel-button");
const clearButton = $("#clear-button");
const openEventModalBtn = $("#open-event-modal-btn");
const savedList = $("#saved-list");

// Toast
const toast = $("#toast");
const toastText = $("#toast-text");
const toastSpinner = $("#toast-spinner");

// Modals
const textModal = $("#text-modal");
const textModalLabel = $("#text-modal-label");
const textModalInput = $("#text-modal-input");
const textModalOk = $("#text-modal-ok");
const textModalCancel = $("#text-modal-cancel");

const eventModal = $("#event-modal");
const eventModalTitle = $("#event-modal-title");
const eventModalError = $("#event-modal-error");
const eventModalName = $("#event-modal-name");
const eventModalDay = $("#event-modal-day");
const eventModalStart = $("#event-modal-start");
const eventModalEnd = $("#event-modal-end");
const eventModalDelete = $("#event-modal-delete");
const eventModalSave = $("#event-modal-save");
const eventModalCancel = $("#event-modal-cancel");

let latestAssignments = [];
let fixedAssignments = [];
let confirmedMode = false;
let studentId = 0;
let conditionId = 0;
let toastTimer = null;
const STORAGE_KEY = "auto-scheduler-confirmed-schedules";

// ===== Utils =====

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
function escapeAttr(value) {
  return escapeHtml(value).replaceAll("`", "&#096;");
}

function timeToMinutes(value) {
  const [hour, minute] = value.split(":").map(Number);
  return hour * 60 + minute;
}
function minutesToTime(minutes) {
  const hour = Math.floor(minutes / 60);
  const minute = minutes % 60;
  return `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`;
}
function hourLabel(hour) {
  const h = ((hour % 24) + 24) % 24;
  if (h === 0) return "오전 12시";
  if (h === 12) return "오후 12시";
  if (h < 12) return `오전 ${h}시`;
  return `오후 ${h - 12}시`;
}

// ===== Toast & Alert =====

function showToast(text, { loading = false, duration = 0 } = {}) {
  toast.classList.remove("hidden");
  toastText.textContent = text;
  toastSpinner.classList.toggle("hidden", !loading);
  clearTimeout(toastTimer);
  if (duration > 0) {
    toastTimer = setTimeout(() => toast.classList.add("hidden"), duration);
  }
}
function hideToast() {
  clearTimeout(toastTimer);
  toast.classList.add("hidden");
}

function showAlert(title, errors = []) {
  alertBanner.classList.remove("hidden");
  alertBanner.innerHTML = `
    <button class="icon-button alert-close" type="button" aria-label="닫기">×</button>
    <strong>${escapeHtml(title)}</strong>
    ${errors.length ? `<ul>${errors.map((e) => `<li>${escapeHtml(e)}</li>`).join("")}</ul>` : ""}
  `;
  alertBanner.querySelector(".alert-close").addEventListener("click", hideAlert);
}
function hideAlert() {
  alertBanner.classList.add("hidden");
  alertBanner.innerHTML = "";
}

// ===== Drawer =====

function openDrawer() {
  drawer.classList.add("open");
  drawerBackdrop.classList.remove("hidden");
}
function closeDrawer() {
  drawer.classList.remove("open");
  drawerBackdrop.classList.add("hidden");
}

// ===== Bottom sheet =====

function openSavedSheet() {
  renderSavedSchedules();
  savedSheet.classList.remove("hidden");
  savedBackdrop.classList.remove("hidden");
}
function closeSavedSheet() {
  savedSheet.classList.add("hidden");
  savedBackdrop.classList.add("hidden");
}

// ===== Text modal (for confirmation/title input) =====

function openTextModal(label, initialValue = "") {
  return new Promise((resolve) => {
    textModalLabel.textContent = label;
    textModalInput.value = initialValue;
    textModal.classList.remove("hidden");
    textModalInput.focus();
    textModalInput.select();

    const cleanup = (value) => {
      textModal.classList.add("hidden");
      textModalOk.removeEventListener("click", onOk);
      textModalCancel.removeEventListener("click", onCancel);
      textModal.removeEventListener("click", onBackdrop);
      textModalInput.removeEventListener("keydown", onKeydown);
      resolve(value);
    };
    const onOk = () => cleanup(textModalInput.value.trim());
    const onCancel = () => cleanup(null);
    const onBackdrop = (event) => {
      if (event.target === textModal) cleanup(null);
    };
    const onKeydown = (event) => {
      if (event.key === "Enter") onOk();
      if (event.key === "Escape") onCancel();
    };

    textModalOk.addEventListener("click", onOk);
    textModalCancel.addEventListener("click", onCancel);
    textModal.addEventListener("click", onBackdrop);
    textModalInput.addEventListener("keydown", onKeydown);
  });
}

// ===== Student card =====

function addStudentCard(student = {}) {
  studentId += 1;
  const id = studentId;
  const card = document.createElement("section");
  card.className = "student-card";
  card.dataset.studentId = String(id);
  card.innerHTML = `
    <header class="student-card-title">
      <button class="student-toggle" type="button" aria-label="펼치기/접기"></button>
      <span class="student-number"></span>
      <button class="icon-button danger remove-student" type="button" aria-label="학생 삭제">×</button>
    </header>
    <div class="student-card-body">
      <div class="student-head">
        <div class="field">
          <label for="name-${id}">이름</label>
          <input id="name-${id}" class="student-name" value="${escapeAttr(student.name || "")}" placeholder="이름" />
        </div>
        <div class="field">
          <label for="count-${id}">주 횟수</label>
          <input id="count-${id}" class="weekly-count" type="number" min="1" max="7" step="1" value="${student.weekly_count || 1}" />
        </div>
      </div>
      <div class="conditions"></div>
      <button class="add-condition" type="button">+ 가능 시간 추가</button>
    </div>
  `;

  const conditions = card.querySelector(".conditions");
  const rows = student.conditions || [{ days: [], start_time: "18:00", end_time: "23:00" }];
  rows.forEach((condition) => addConditionRow(conditions, condition));

  card.querySelector(".remove-student").addEventListener("click", () => {
    card.remove();
    if (!studentsInput.children.length) addStudentCard();
  });
  card.querySelector(".add-condition").addEventListener("click", () => addConditionRow(conditions));
  card.querySelector(".student-toggle").addEventListener("click", () => {
    card.classList.toggle("collapsed");
  });

  studentsInput.appendChild(card);
  return card;
}

function addConditionRow(container, condition = {}) {
  conditionId += 1;
  const id = conditionId;
  const row = document.createElement("div");
  row.className = "condition-row";
  row.innerHTML = `
    <button class="icon-button danger remove-condition" type="button" aria-label="시간대 삭제">×</button>
    <div class="field">
      <div class="label">가능 요일</div>
      <div class="days">
        ${DAYS.map((day) => {
          const checked = (condition.days || []).includes(day) ? "checked" : "";
          return `<label class="day"><input type="checkbox" value="${day}" ${checked} /><span>${day}</span></label>`;
        }).join("")}
      </div>
    </div>
    <div class="time-range">
      <div class="field">
        <label for="start-${id}">시작</label>
        <input id="start-${id}" class="start-time" type="time" value="${condition.start_time || "18:00"}" />
      </div>
      <div class="field">
        <label for="end-${id}">종료</label>
        <input id="end-${id}" class="end-time" type="time" value="${condition.end_time || "23:00"}" />
      </div>
    </div>
  `;
  row.querySelector(".remove-condition").addEventListener("click", () => {
    row.remove();
    if (!container.children.length) addConditionRow(container);
  });
  container.appendChild(row);
}

function setStudentCards(students) {
  studentsInput.innerHTML = "";
  studentId = 0;
  conditionId = 0;
  students.forEach((student) => addStudentCard(student));
  if (!students.length) addStudentCard();
}

function groupSampleRows(rows) {
  const grouped = [];
  const byKey = new Map();
  rows.forEach((row) => {
    const key = `${row.name}::${row.weekly_count}`;
    if (!byKey.has(key)) {
      const item = { name: row.name, weekly_count: row.weekly_count, conditions: [] };
      byKey.set(key, item);
      grouped.push(item);
    }
    byKey.get(key).conditions.push({
      days: row.days,
      start_time: row.start_time,
      end_time: row.end_time,
    });
  });
  return grouped;
}

function getPayload() {
  const data = new FormData(form);
  const students = [];
  studentsInput.querySelectorAll(".student-card").forEach((card) => {
    const name = card.querySelector(".student-name").value;
    const weeklyCount = Number(card.querySelector(".weekly-count").value);
    card.querySelectorAll(".condition-row").forEach((row) => {
      students.push({
        name,
        weekly_count: weeklyCount,
        days: [...row.querySelectorAll(".day input:checked")].map((input) => input.value),
        start_time: row.querySelector(".start-time").value,
        end_time: row.querySelector(".end-time").value,
      });
    });
  });

  return {
    students,
    fixed_assignments: fixedAssignments,
    duration_minutes: Number(data.get("duration_minutes")),
    step_minutes: Number(data.get("step_minutes")),
    first_start_time: data.get("first_start_time"),
    last_start_time: data.get("last_start_time"),
    break_minutes: Number(data.get("break_minutes")),
  };
}

function currentStudentNames() {
  return [...studentsInput.querySelectorAll(".student-name")]
    .map((input) => input.value.trim())
    .filter(Boolean);
}

// ===== Grid render =====

function colorForName(name) {
  const palette = [
    ["#dbeafe", "#1e3a8a"],
    ["#dcfce7", "#14532d"],
    ["#fef3c7", "#78350f"],
    ["#fce7f3", "#831843"],
    ["#ede9fe", "#4c1d95"],
    ["#cffafe", "#164e63"],
    ["#ffedd5", "#7c2d12"],
    ["#e0e7ff", "#312e81"],
  ];
  let hash = 0;
  for (const char of name) hash = (hash * 31 + char.charCodeAt(0)) >>> 0;
  return palette[hash % palette.length];
}

function renderGrid(assignments) {
  const firstSetting = timeToMinutes($("#first-start").value || "10:00");
  const lastSetting = timeToMinutes($("#last-start").value || "23:00");
  const duration = Number($("#duration").value) || 50;
  const starts = assignments.length ? assignments.map((item) => timeToMinutes(item.start)) : [firstSetting];
  const ends = assignments.length ? assignments.map((item) => timeToMinutes(item.end)) : [lastSetting + duration];
  const firstHour = Math.floor(Math.min(firstSetting, ...starts) / 60);
  const lastHour = Math.ceil(Math.max(lastSetting + duration, ...ends) / 60);
  const hours = [];
  for (let hour = firstHour; hour < lastHour; hour += 1) hours.push(hour);

  const totalMinutes = (lastHour - firstHour) * 60 || 1;

  const eventsByDay = new Map(DAYS.map((day) => [day, []]));
  assignments.forEach((item, index) => {
    const start = timeToMinutes(item.start);
    const end = timeToMinutes(item.end);
    eventsByDay.get(item.day).push({
      ...item,
      index,
      topPct: ((start - firstHour * 60) / totalMinutes) * 100,
      heightPct: ((end - start) / totalMinutes) * 100,
    });
  });

  const header = [
    '<div class="schedule-header"><div class="schedule-head">시간</div>',
    ...DAYS.map((day) => `<div class="schedule-head">${escapeHtml(day)}</div>`),
    "</div>",
  ].join("");

  const timeColumn = hours.map((hour) => `<div class="schedule-time">${hourLabel(hour)}</div>`).join("");
  const dayColumns = DAYS.map((day) => {
    const events = eventsByDay
      .get(day)
      .map((event) => {
        const [bg, ink] = colorForName(event.name);
        const sourceClass = confirmedMode
          ? "schedule-event--confirmed"
          : event.source === "fixed"
            ? "schedule-event--fixed"
            : "schedule-event--generated";
        const baseStyle = `top:${event.topPct}%;height:${event.heightPct}%`;
        const style = confirmedMode
          ? `${baseStyle};--event-bg:${bg};--event-ink:${ink}`
          : baseStyle;
        return `
          <div class="schedule-event ${sourceClass}" data-index="${event.index}" style="${style}">
            ${escapeHtml(event.name)}
            <small>${escapeHtml(event.start)}-${escapeHtml(event.end)}</small>
          </div>
        `;
      })
      .join("");
    return `<div class="schedule-day" data-day="${day}">${events}</div>`;
  }).join("");
  const body = `<div class="schedule-body"><div class="schedule-time-column">${timeColumn}</div>${dayColumns}</div>`;

  gridTarget.innerHTML = `<div class="schedule" style="--hour-count:${hours.length}">${header}${body}</div>`;
  wireScheduleInteractions();
  updateEmptyState();
}

function wireScheduleInteractions() {
  gridTarget.querySelectorAll(".schedule-event").forEach((eventEl) => {
    eventEl.addEventListener("click", () => {
      const index = Number(eventEl.dataset.index);
      if (Number.isFinite(index)) openEventModal({ existingIndex: index });
    });
  });
}

function updateEmptyState() {
  emptyHint.classList.toggle("hidden", latestAssignments.length > 0);
}

// ===== Assignment ops =====

function assignmentKey(item) {
  return `${item.name}|${item.day}|${item.start}|${item.end}`;
}

function markAssignmentSources(assignments, fixedItems) {
  const fixedKeys = new Set(fixedItems.map(assignmentKey));
  return assignments.map((item) => ({
    ...item,
    source: item.source || (fixedKeys.has(assignmentKey(item)) ? "fixed" : "generated"),
  }));
}

function syncEditedAssignments() {
  fixedAssignments = latestAssignments.map((item) => ({ ...item }));
  const has = latestAssignments.length > 0;
  excelButton.disabled = !has;
  confirmButton.disabled = !has;
  clearButton.disabled = !has;
  updateEmptyState();
}

function removeAssignment(index) {
  const [removed] = latestAssignments.splice(index, 1);
  confirmedMode = false;
  renderGrid(latestAssignments);
  syncEditedAssignments();
  if (removed) {
    showToast(`${removed.day} ${removed.start} ${removed.name} 삭제`, { duration: 1500 });
  }
}

function hasConflict(candidate, ignoreIndex = -1) {
  const cStart = timeToMinutes(candidate.start);
  const cEnd = timeToMinutes(candidate.end);
  return latestAssignments.some((other, i) => {
    if (i === ignoreIndex) return false;
    if (other.day !== candidate.day) return false;
    const oStart = timeToMinutes(other.start);
    const oEnd = timeToMinutes(other.end);
    return cStart < oEnd && oStart < cEnd;
  });
}

// ===== Event modal =====

function showEventModalError(message) {
  if (!message) {
    eventModalError.classList.add("hidden");
    eventModalError.textContent = "";
  } else {
    eventModalError.classList.remove("hidden");
    eventModalError.textContent = message;
  }
}

function defaultEndForStart(startValue) {
  const duration = Number($("#duration").value) || 50;
  return minutesToTime(timeToMinutes(startValue) + duration);
}

function openEventModal({ existingIndex } = {}) {
  const isEdit = existingIndex != null && Number.isFinite(existingIndex);
  const existing = isEdit ? latestAssignments[existingIndex] : null;

  eventModalTitle.textContent = isEdit ? "일정 편집" : "고정 일정 추가";
  eventModalName.value = existing?.name ?? currentStudentNames()[0] ?? "";
  eventModalDay.value = existing?.day ?? "월";
  const startVal = existing?.start ?? "18:00";
  eventModalStart.value = startVal;
  eventModalEnd.value = existing?.end ?? defaultEndForStart(startVal);
  eventModalDelete.classList.toggle("hidden", !isEdit);
  showEventModalError("");

  eventModal.classList.remove("hidden");
  setTimeout(() => eventModalName.focus(), 50);

  const cleanup = () => {
    eventModal.classList.add("hidden");
    eventModalSave.removeEventListener("click", onSave);
    eventModalCancel.removeEventListener("click", onCancel);
    eventModalDelete.removeEventListener("click", onDelete);
    eventModal.removeEventListener("click", onBackdrop);
    eventModalStart.removeEventListener("change", onStartChange);
  };

  const onStartChange = () => {
    const s = timeToMinutes(eventModalStart.value);
    const e = timeToMinutes(eventModalEnd.value);
    if (e <= s) eventModalEnd.value = defaultEndForStart(eventModalStart.value);
  };

  const onSave = () => {
    const name = eventModalName.value.trim();
    const day = eventModalDay.value;
    const start = eventModalStart.value;
    const end = eventModalEnd.value;

    if (!name) return showEventModalError("이름을 입력하세요.");
    if (!start || !end) return showEventModalError("시작/종료 시간을 입력하세요.");
    if (timeToMinutes(end) <= timeToMinutes(start)) {
      return showEventModalError("종료 시간은 시작 시간보다 늦어야 합니다.");
    }
    if (hasConflict({ day, start, end }, isEdit ? existingIndex : -1)) {
      return showEventModalError(`${day} ${start}-${end} 에 이미 다른 일정이 있습니다.`);
    }

    const item = { name, day, start, end, source: "fixed" };
    if (isEdit) {
      latestAssignments[existingIndex] = item;
    } else {
      latestAssignments.push(item);
    }
    confirmedMode = false;
    fixedAssignments = latestAssignments.map((a) => ({ ...a }));
    renderGrid(latestAssignments);
    syncEditedAssignments();
    showToast(isEdit ? "일정 수정됨" : "고정 일정 추가됨", { duration: 1500 });
    cleanup();
  };

  const onDelete = () => {
    if (!isEdit) return;
    removeAssignment(existingIndex);
    cleanup();
  };

  const onCancel = () => cleanup();
  const onBackdrop = (e) => {
    if (e.target === eventModal) cleanup();
  };

  eventModalSave.addEventListener("click", onSave);
  eventModalCancel.addEventListener("click", onCancel);
  eventModalDelete.addEventListener("click", onDelete);
  eventModal.addEventListener("click", onBackdrop);
  eventModalStart.addEventListener("change", onStartChange);
}

// ===== Saved schedules =====

function readSavedSchedules() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
  } catch {
    return [];
  }
}
function writeSavedSchedules(items) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(items));
}

function renderSavedSchedules() {
  const items = readSavedSchedules();
  if (!items.length) {
    savedList.innerHTML = '<div class="saved-list-empty">확정된 시간표가 없습니다.</div>';
    return;
  }

  savedList.innerHTML = items
    .map(
      (item) => `
        <div class="saved-item">
          <div class="saved-title">
            ${escapeHtml(item.title)}
            <small>${escapeHtml(item.created_at)} · ${item.assignments.length}개 수업</small>
          </div>
          <button class="ghost load-schedule" type="button" data-id="${item.id}">불러오기</button>
          <button class="danger delete-schedule" type="button" data-id="${item.id}">삭제</button>
        </div>
      `
    )
    .join("");

  savedList.querySelectorAll(".load-schedule").forEach((button) => {
    button.addEventListener("click", () => loadSavedSchedule(button.dataset.id));
  });
  savedList.querySelectorAll(".delete-schedule").forEach((button) => {
    button.addEventListener("click", () => deleteSavedSchedule(button.dataset.id));
  });
}

async function confirmCurrentSchedule() {
  if (!latestAssignments.length) return;
  const title = await openTextModal("시간표 이름을 입력하세요.", `시간표 ${new Date().toLocaleString("ko-KR")}`);
  if (!title) return;

  confirmedMode = true;
  latestAssignments = latestAssignments.map((item) => ({ ...item, source: item.source || "generated" }));
  fixedAssignments = latestAssignments.map((item) => ({ ...item }));

  const item = {
    id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    title: title.trim(),
    created_at: new Date().toLocaleString("ko-KR"),
    assignments: latestAssignments,
  };
  const items = [item, ...readSavedSchedules()].slice(0, 20);
  writeSavedSchedules(items);
  renderGrid(latestAssignments);
  syncEditedAssignments();
  showToast(`확정 저장: ${item.title}`, { duration: 1800 });
}

function loadSavedSchedule(id) {
  const item = readSavedSchedules().find((saved) => saved.id === id);
  if (!item) return;
  confirmedMode = true;
  latestAssignments = item.assignments.map((assignment) => ({ ...assignment }));
  fixedAssignments = latestAssignments.map((assignment) => ({ ...assignment }));
  excelButton.disabled = false;
  confirmButton.disabled = false;
  clearButton.disabled = false;
  renderGrid(latestAssignments);
  syncEditedAssignments();
  showToast(`불러옴: ${item.title}`, { duration: 1800 });
  closeSavedSheet();
}

function deleteSavedSchedule(id) {
  writeSavedSchedules(readSavedSchedules().filter((item) => item.id !== id));
  renderSavedSchedules();
}

// ===== Result =====

function renderResult(data) {
  if (!data.success) {
    excelButton.disabled = true;
    confirmButton.disabled = true;
    showAlert(data.message, data.errors || []);
    return;
  }

  hideAlert();
  latestAssignments = markAssignmentSources(data.assignments, fixedAssignments);
  confirmedMode = false;
  excelButton.disabled = false;
  confirmButton.disabled = false;
  clearButton.disabled = false;
  renderGrid(latestAssignments);
  syncEditedAssignments();
  showToast(data.message, { duration: 1800 });
}

function applySample() {
  const data = getSample();
  latestAssignments = [];
  fixedAssignments = [];
  confirmedMode = false;
  excelButton.disabled = true;
  confirmButton.disabled = true;
  clearButton.disabled = true;
  setStudentCards(groupSampleRows(data.sample_students || []));
  $("#duration").value = data.duration_minutes;
  $("#step").value = data.step_minutes;
  $("#first-start").value = data.first_start_time;
  $("#last-start").value = data.last_start_time;
  $("#break-minutes").value = data.break_minutes;
  renderGrid([]);
  hideAlert();
  showToast("예시 데이터 적용됨", { duration: 1500 });
}

function clearSchedule() {
  if (!latestAssignments.length) return;
  if (!confirm("시간표를 초기화하시겠습니까? 입력된 학생 정보는 유지됩니다.")) return;
  latestAssignments = [];
  fixedAssignments = [];
  confirmedMode = false;
  renderGrid([]);
  syncEditedAssignments();
  hideAlert();
  showToast("초기화됨", { duration: 1200 });
}

// ===== Form submit =====

function submitGenerate() {
  hideAlert();
  showToast("시간표 생성 중…", { loading: true });

  // 짧은 지연으로 UI가 토스트를 먼저 그리도록
  setTimeout(() => {
    try {
      const data = generateSchedule(getPayload());
      if (data.success) {
        closeDrawer();
      }
      renderResult(data);
      if (data.success) hideToast();
    } catch (exc) {
      hideToast();
      showAlert("오류", [exc.message || String(exc)]);
    }
  }, 30);
}

form.addEventListener("submit", (event) => {
  event.preventDefault();
  submitGenerate();
});

// ===== Wiring =====

openDrawerBtn.addEventListener("click", openDrawer);
closeDrawerBtn.addEventListener("click", closeDrawer);
drawerBackdrop.addEventListener("click", closeDrawer);
emptyOpenDrawerBtn.addEventListener("click", openDrawer);

showSavedBtn.addEventListener("click", openSavedSheet);
closeSavedBtn.addEventListener("click", closeSavedSheet);
savedBackdrop.addEventListener("click", closeSavedSheet);

openEventModalBtn.addEventListener("click", () => openEventModal());

addStudentButton.addEventListener("click", () => {
  const next = addStudentCard();
  next.scrollIntoView({ behavior: "smooth", block: "nearest" });
});
sampleButton.addEventListener("click", applySample);

confirmButton.addEventListener("click", confirmCurrentSchedule);
regenerateButton.addEventListener("click", () => submitGenerate());
clearButton.addEventListener("click", clearSchedule);

excelButton.addEventListener("click", () => {
  if (!latestAssignments.length) return;
  try {
    downloadExcel(latestAssignments);
    showToast("엑셀 다운로드", { duration: 1500 });
  } catch (exc) {
    showAlert("엑셀 저장 실패", [exc.message || String(exc)]);
  }
});

["duration", "step", "first-start", "last-start"].forEach((id) => {
  $(`#${id}`).addEventListener("change", () => renderGrid(latestAssignments));
});

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    if (!eventModal.classList.contains("hidden")) return;
    if (!textModal.classList.contains("hidden")) return;
    if (drawer.classList.contains("open")) closeDrawer();
    else if (!savedSheet.classList.contains("hidden")) closeSavedSheet();
  }
});

// ===== Init =====

setStudentCards([]);
renderGrid([]);
syncEditedAssignments();

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./service-worker.js").catch(() => {});
  });
}
