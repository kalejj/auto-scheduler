import { DAYS } from "./constants.js";
import { getSample, generateSchedule } from "./api.js";
import { downloadExcel } from "./excel.js";

const form = document.querySelector("#schedule-form");
const studentsInput = document.querySelector("#students-input");
const statusBox = document.querySelector("#status");
const assignmentsSection = document.querySelector("#assignments-section");
const gridTarget = document.querySelector("#grid");
const assignmentsTarget = document.querySelector("#assignments");
const addStudentButton = document.querySelector("#add-student-button");
const sampleButton = document.querySelector("#sample-button");
const confirmButton = document.querySelector("#confirm-button");
const excelButton = document.querySelector("#excel-button");
const addFixedButton = document.querySelector("#add-fixed-button");
const regenerateButton = document.querySelector("#regenerate-button");
const savedList = document.querySelector("#saved-list");

const textModal = document.querySelector("#text-modal");
const textModalLabel = document.querySelector("#text-modal-label");
const textModalInput = document.querySelector("#text-modal-input");
const textModalOk = document.querySelector("#text-modal-ok");
const textModalCancel = document.querySelector("#text-modal-cancel");

const eventModal = document.querySelector("#event-modal");
const eventModalTitle = document.querySelector("#event-modal-title");
const eventModalError = document.querySelector("#event-modal-error");
const eventModalName = document.querySelector("#event-modal-name");
const eventModalDay = document.querySelector("#event-modal-day");
const eventModalStart = document.querySelector("#event-modal-start");
const eventModalEnd = document.querySelector("#event-modal-end");
const eventModalDelete = document.querySelector("#event-modal-delete");
const eventModalSave = document.querySelector("#event-modal-save");
const eventModalCancel = document.querySelector("#event-modal-cancel");

let canDownloadExcel = false;
let latestAssignments = [];
let fixedAssignments = [];
let confirmedMode = false;
let studentId = 0;
let conditionId = 0;
const STORAGE_KEY = "auto-scheduler-confirmed-schedules";

const mobileTabs = document.querySelectorAll(".mobile-tabs button");
function setActiveTab(name) {
  document.body.dataset.activeTab = name;
  mobileTabs.forEach((btn) => btn.classList.toggle("active", btn.dataset.tab === name));
}

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

function setStatus(kind, title, detail = "", errors = []) {
  statusBox.className = `panel status ${kind || ""}`;
  statusBox.innerHTML = `
    <strong>${escapeHtml(title)}</strong>
    ${detail ? `<div class="muted">${escapeHtml(detail)}</div>` : ""}
    ${errors.length ? `<ul>${errors.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>` : ""}
  `;
}

function table(headers, rows) {
  const head = headers.map((item) => `<th>${escapeHtml(item.label)}</th>`).join("");
  const body = rows
    .map((row) => {
      const cells = headers
        .map((item) => `<td>${escapeHtml(row[item.key] ?? "")}</td>`)
        .join("");
      return `<tr>${cells}</tr>`;
    })
    .join("");
  return `<table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}

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

function renderGrid(assignments) {
  const firstSetting = timeToMinutes(document.querySelector("#first-start").value || "10:00");
  const lastSetting = timeToMinutes(document.querySelector("#last-start").value || "23:00");
  const duration = Number(document.querySelector("#duration").value) || 50;
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
}

function wireScheduleInteractions() {
  gridTarget.querySelectorAll(".schedule-event").forEach((eventEl) => {
    eventEl.addEventListener("click", () => {
      const index = Number(eventEl.dataset.index);
      if (Number.isFinite(index)) openEventModal({ existingIndex: index });
    });
  });
}

function currentStudentNames() {
  return [...studentsInput.querySelectorAll(".student-name")]
    .map((input) => input.value.trim())
    .filter(Boolean);
}

function syncEditedAssignments() {
  fixedAssignments = latestAssignments.map((item) => ({ ...item }));
  canDownloadExcel = latestAssignments.length > 0;
  excelButton.disabled = !canDownloadExcel;
  confirmButton.disabled = !latestAssignments.length;
  assignmentsTarget.innerHTML = table(
    [
      { key: "day", label: "요일" },
      { key: "start", label: "시작" },
      { key: "end", label: "종료" },
      { key: "name", label: "이름" },
    ],
    latestAssignments
  );
  assignmentsSection.classList.toggle("hidden", !latestAssignments.length);
}

function removeAssignment(index) {
  const [removed] = latestAssignments.splice(index, 1);
  confirmedMode = false;
  renderGrid(latestAssignments);
  syncEditedAssignments();
  if (removed) {
    setStatus("", "수업 삭제됨", `${removed.day} ${removed.start}-${removed.end} ${removed.name}`);
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
  const duration = Number(document.querySelector("#duration").value) || 50;
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
    setStatus("", isEdit ? "일정 수정됨" : "고정 수업 추가됨", `${day} ${start}-${end} ${name}`);
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
    savedList.innerHTML = '<div class="muted">확정된 시간표가 없습니다.</div>';
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
  renderSavedSchedules();
  renderGrid(latestAssignments);
  syncEditedAssignments();
  setStatus("success", "시간표 확정됨", `${item.title} 저장 완료`);
}

function loadSavedSchedule(id) {
  const item = readSavedSchedules().find((saved) => saved.id === id);
  if (!item) return;
  confirmedMode = true;
  latestAssignments = item.assignments.map((assignment) => ({ ...assignment }));
  fixedAssignments = latestAssignments.map((assignment) => ({ ...assignment }));
  canDownloadExcel = true;
  excelButton.disabled = false;
  confirmButton.disabled = false;
  renderGrid(latestAssignments);
  syncEditedAssignments();
  setStatus("success", "시간표 불러옴", item.title);
  setActiveTab("schedule");
}

function deleteSavedSchedule(id) {
  writeSavedSchedules(readSavedSchedules().filter((item) => item.id !== id));
  renderSavedSchedules();
}

function renderResult(data) {
  canDownloadExcel = Boolean(data.success);
  excelButton.disabled = !canDownloadExcel;
  confirmButton.disabled = !data.success;
  confirmedMode = false;

  if (!data.success) {
    canDownloadExcel = false;
    excelButton.disabled = true;
    confirmButton.disabled = true;
    setStatus("error", data.message, "", data.errors || []);
    assignmentsSection.classList.add("hidden");
    return;
  }

  latestAssignments = markAssignmentSources(data.assignments, fixedAssignments);
  setStatus("success", data.message);

  renderGrid(latestAssignments);

  assignmentsTarget.innerHTML = table(
    [
      { key: "day", label: "요일" },
      { key: "start", label: "시작" },
      { key: "end", label: "종료" },
      { key: "name", label: "이름" },
    ],
    latestAssignments
  );
  assignmentsSection.classList.remove("hidden");
}

function applySample() {
  const data = getSample();
  latestAssignments = [];
  fixedAssignments = [];
  confirmedMode = false;
  canDownloadExcel = false;
  excelButton.disabled = true;
  confirmButton.disabled = true;
  setStudentCards(groupSampleRows(data.sample_students || []));
  document.querySelector("#duration").value = data.duration_minutes;
  document.querySelector("#step").value = data.step_minutes;
  document.querySelector("#first-start").value = data.first_start_time;
  document.querySelector("#last-start").value = data.last_start_time;
  document.querySelector("#break-minutes").value = data.break_minutes;
  renderGrid([]);
}

form.addEventListener("submit", (event) => {
  event.preventDefault();
  setStatus(
    "",
    "생성 중",
    fixedAssignments.length
      ? "현재 시간표를 고정한 채 빈 자리를 다시 탐색하고 있습니다."
      : "조건을 탐색하고 있습니다."
  );
  excelButton.disabled = true;

  setTimeout(() => {
    try {
      const data = generateSchedule(getPayload());
      renderResult(data);
      if (data.success) setActiveTab("schedule");
    } catch (exc) {
      setStatus("error", "오류", exc.message || String(exc));
    }
  }, 10);
});

addStudentButton.addEventListener("click", () => {
  const next = addStudentCard();
  next.scrollIntoView({ behavior: "smooth", block: "nearest" });
});
sampleButton.addEventListener("click", applySample);
confirmButton.addEventListener("click", confirmCurrentSchedule);
addFixedButton.addEventListener("click", () => openEventModal());
regenerateButton.addEventListener("click", () => form.requestSubmit());
["duration", "step", "first-start", "last-start"].forEach((id) => {
  document.querySelector(`#${id}`).addEventListener("change", () => renderGrid(latestAssignments));
});
excelButton.addEventListener("click", () => {
  if (!canDownloadExcel) return;
  try {
    downloadExcel(latestAssignments);
  } catch (exc) {
    setStatus("error", "엑셀 저장 실패", exc.message || String(exc));
  }
});

mobileTabs.forEach((btn) => btn.addEventListener("click", () => setActiveTab(btn.dataset.tab)));
setActiveTab("input");

setStudentCards([]);
renderGrid([]);
renderSavedSchedules();

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./service-worker.js").catch(() => {});
  });
}
