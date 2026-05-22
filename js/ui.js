import { DAYS, SAMPLE_STUDENTS } from "./constants.js";
import { generateSchedule } from "./api.js";
import { downloadExcel } from "./excel.js";
import { shareScheduleImage } from "./share.js";
import {
  loadMembers,
  addMember,
  updateMember,
  deleteMember,
  replaceAllMembers,
  membersToStudentsPayload,
} from "./members.js";
import {
  todayISO,
  weekStartISO,
  addWeeks,
  addMonths,
  weekDays,
  dayNameToISO,
  formatDateShort,
  formatWeekRange,
  formatYearMonth,
  monthOfISO,
  monthStartISO,
  calendarGrid,
  pad2,
} from "./dates.js";
import {
  loadSessions,
  loadSessionsForWeek,
  addSession,
  addSessions,
  updateSession,
  deleteSession,
  deleteSessionsForWeekWhere,
  allMonthCounts,
  purgeLegacySavedSchedules,
} from "./sessions.js";

const $ = (sel) => document.querySelector(sel);

// ===== DOM refs =====

const drawer = $("#drawer");
const drawerBackdrop = $("#drawer-backdrop");
const openDrawerBtn = $("#open-drawer-btn");
const closeDrawerBtn = $("#close-drawer-btn");
const openSettingsBtn = $("#open-settings-btn");

const statsSheet = $("#stats-sheet");
const statsBackdrop = $("#stats-backdrop");
const showStatsBtn = $("#show-stats-btn");
const closeStatsBtn = $("#close-stats-btn");
const statsMonthPrev = $("#stats-month-prev");
const statsMonthNext = $("#stats-month-next");
const statsMonthToday = $("#stats-month-today");
const statsMonthText = $("#stats-month-text");
const statsCompleted = $("#stats-completed");
const statsNoShow = $("#stats-noshow");
const statsRevenue = $("#stats-revenue");
const miniCalEl = $("#mini-cal");
const statsMemberTable = $("#stats-member-table");

let statsMonthISO = monthOfISO(todayISO()); // "YYYY-MM"

const weekPrevBtn = $("#week-prev-btn");
const weekNextBtn = $("#week-next-btn");
const weekLabelBtn = $("#week-label-btn");
const weekRangeText = $("#week-range-text");
const weekMonthText = $("#week-month-text");
const emptyGenerateBtn = $("#empty-generate-btn");
const emptyCopyPrevBtn = $("#empty-copy-prev-btn");
const copyPrevButton = $("#copy-prev-button");

const form = $("#schedule-form");
const membersListEl = $("#members-list");
const addMemberButton = $("#add-member-button");
const sampleButton = $("#sample-button");

const gridTarget = $("#grid");
const alertBanner = $("#alert-banner");
// emptyHint 제거 — 대신 empty-state-no-members, empty-state-no-sessions 사용
const regenerateButton = $("#regenerate-button");
const excelButton = $("#excel-button");
const clearButton = $("#clear-button");
const moreButton = $("#more-button");
const moreMenu = $("#more-menu");
const shareButton = $("#share-button");
const openEventModalBtn = $("#open-event-modal-btn");

const emptyStateNoMembers = $("#empty-state-no-members");
const emptyStateNoSessions = $("#empty-state-no-sessions");
const emptyWeekText = $("#empty-week-text");
const emptyAddMemberBtn = $("#empty-add-member-btn");
const emptySampleBtn = $("#empty-sample-btn");

const eventModalLockRow = $("#event-modal-lock-row");
const eventModalLocked = $("#event-modal-locked");
const eventTypeMemberRadio = $("#event-type-member");
const eventTypePersonalRadio = $("#event-type-personal");
const eventModalMemberRow = $("#event-modal-member-row");
const eventModalNameRow = $("#event-modal-name-row");
const eventModalMemberSelect = $("#event-modal-member");

const onboardingModal = $("#onboarding-modal");
const onboardingStartBtn = $("#onboarding-start");

const toast = $("#toast");
const toastText = $("#toast-text");
const toastSpinner = $("#toast-spinner");

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
const eventModalStatusSection = $("#event-modal-status-section");
const eventModalStatusBadge = $("#event-modal-status-badge");
const eventModalMarkCompleted = $("#event-modal-mark-completed");
const eventModalMarkNoShow = $("#event-modal-mark-noshow");
const eventModalMarkCancelled = $("#event-modal-mark-cancelled");
const eventModalMarkReset = $("#event-modal-mark-reset");

const memberModal = $("#member-modal");
const memberModalTitle = $("#member-modal-title");
const memberModalError = $("#member-modal-error");
const memberModalName = $("#member-modal-name");
const memberModalPhone = $("#member-modal-phone");
const memberModalCount = $("#member-modal-count");
const memberModalActive = $("#member-modal-active");
const memberModalConditions = $("#member-modal-conditions");
const memberModalAddCondition = $("#member-modal-add-condition");
const memberModalPlanName = $("#member-modal-plan-name");
const memberModalTotalCount = $("#member-modal-total-count");
const memberModalUsedCount = $("#member-modal-used-count");
const memberModalPrice = $("#member-modal-price");
const memberModalExpires = $("#member-modal-expires");
const memberModalMemo = $("#member-modal-memo");
const memberModalSave = $("#member-modal-save");
const memberModalDelete = $("#member-modal-delete");
const memberModalCancel = $("#member-modal-cancel");

const confirmModal = $("#confirm-modal");
const confirmModalTitle = $("#confirm-modal-title");
const confirmModalMessage = $("#confirm-modal-message");
const confirmModalOk = $("#confirm-modal-ok");
const confirmModalCancel = $("#confirm-modal-cancel");

const settingsModal = $("#settings-modal");
const settingsDuration = $("#settings-duration");
const settingsStep = $("#settings-step");
const settingsFirst = $("#settings-first");
const settingsLast = $("#settings-last");
const settingsBreak = $("#settings-break");
const settingsSave = $("#settings-save");
const settingsCancel = $("#settings-cancel");

const timePickerModal = $("#time-picker-modal");
const timePickerTitle = $("#time-picker-title");
const timePickerOk = $("#time-picker-ok");
const timePickerCancel = $("#time-picker-cancel");

const memberSearchInput = $("#member-search");
const memberSortSelect = $("#member-sort");

// ===== State =====

// 현재 보고 있는 주의 sessions. 진실의 원천은 sessions store.
let currentWeekStart = weekStartISO(todayISO());
let currentWeekSessions = [];
let conditionId = 0;
let toastTimer = null;
let memberModalEditId = null;

const SETTINGS_KEY = "auto-scheduler-settings";
const DEFAULT_SETTINGS = {
  duration_minutes: 50,
  step_minutes: 10,
  first_start_time: "10:00",
  last_start_time: "23:00",
  break_minutes: 0,
};

const WHEEL_ITEM_HEIGHT = 44;
const MINUTE_STEP = 5;

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
  if (!value) return 0;
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
function formatTimeDisplay(hhmm) {
  if (!hhmm) return "";
  const m = timeToMinutes(hhmm);
  const isPm = m >= 720;
  let h12 = Math.floor(m / 60) % 12;
  if (h12 === 0) h12 = 12;
  const min = m % 60;
  return `${isPm ? "오후" : "오전"} ${h12}:${String(min).padStart(2, "0")}`;
}

// ===== Status helpers =====
const STATUS_LABEL = {
  pending: "진행 예정",
  completed: "수업 완료",
  "no-show": "노쇼",
  cancelled: "수업 취소",
};
function statusConsumes(s) {
  return s === "completed" || s === "no-show";
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

// ===== Drawer / Sheets =====

function openDrawer() {
  drawer.classList.add("open");
  drawerBackdrop.classList.remove("hidden");
}
function closeDrawer() {
  drawer.classList.remove("open");
  drawerBackdrop.classList.add("hidden");
}

function openStatsSheet() {
  statsMonthISO = monthOfISO(todayISO());  // 항상 이번 달부터
  renderStats();
  statsSheet.classList.remove("hidden");
  statsBackdrop.classList.remove("hidden");
}
function closeStatsSheet() {
  statsSheet.classList.add("hidden");
  statsBackdrop.classList.add("hidden");
}

// ===== Text modal =====

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
    const onBackdrop = (event) => { if (event.target === textModal) cleanup(null); };
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

// ===== Confirm modal =====

function openConfirmModal(title, message) {
  return new Promise((resolve) => {
    confirmModalTitle.textContent = title;
    confirmModalMessage.textContent = message;
    confirmModal.classList.remove("hidden");

    const cleanup = (value) => {
      confirmModal.classList.add("hidden");
      confirmModalOk.removeEventListener("click", onOk);
      confirmModalCancel.removeEventListener("click", onCancel);
      confirmModal.removeEventListener("click", onBackdrop);
      resolve(value);
    };
    const onOk = () => cleanup(true);
    const onCancel = () => cleanup(false);
    const onBackdrop = (e) => { if (e.target === confirmModal) cleanup(false); };

    confirmModalOk.addEventListener("click", onOk);
    confirmModalCancel.addEventListener("click", onCancel);
    confirmModal.addEventListener("click", onBackdrop);
  });
}

// ===== Wheel time picker =====

let currentTimePickerResolve = null;
const wheelScrollDebouncers = new Map();

function buildWheelItems(wheelName, items, labelFn = (x) => String(x)) {
  const inner = timePickerModal.querySelector(`[data-wheel="${wheelName}"] .wheel-inner`);
  inner.innerHTML = items
    .map((item) => `<div class="wheel-item" data-value="${escapeAttr(String(item))}">${escapeHtml(labelFn(item))}</div>`)
    .join("");
}
function getWheel(name) {
  return timePickerModal.querySelector(`[data-wheel="${name}"]`);
}
function setWheelScroll(wheelName, index) {
  const wheel = getWheel(wheelName);
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      void wheel.offsetHeight;
      wheel.scrollTop = index * WHEEL_ITEM_HEIGHT;
      updateSelectedItem(wheel);
    });
  });
}
function getWheelSelectedIndex(wheelName) {
  const wheel = getWheel(wheelName);
  return Math.max(0, Math.round(wheel.scrollTop / WHEEL_ITEM_HEIGHT));
}
function getWheelSelectedValue(wheelName) {
  const wheel = getWheel(wheelName);
  const items = wheel.querySelectorAll(".wheel-item");
  const idx = Math.min(items.length - 1, getWheelSelectedIndex(wheelName));
  return items[idx]?.dataset.value;
}
function updateSelectedItem(wheel) {
  const items = wheel.querySelectorAll(".wheel-item");
  const idx = Math.min(items.length - 1, Math.max(0, Math.round(wheel.scrollTop / WHEEL_ITEM_HEIGHT)));
  items.forEach((item, i) => item.classList.toggle("selected", i === idx));
}
timePickerModal.querySelectorAll(".wheel").forEach((wheel) => {
  wheel.addEventListener("scroll", () => {
    if (wheelScrollDebouncers.has(wheel)) clearTimeout(wheelScrollDebouncers.get(wheel));
    wheelScrollDebouncers.set(wheel, setTimeout(() => updateSelectedItem(wheel), 60));
  });
});

function openTimePicker(initialHHMM, title = "시간 선택") {
  timePickerTitle.textContent = title;
  const minutes = timeToMinutes(initialHHMM || "18:00");
  const isPm = minutes >= 720;
  let h12 = Math.floor(minutes / 60) % 12;
  if (h12 === 0) h12 = 12;
  const m = minutes % 60;

  buildWheelItems("ampm", ["오전", "오후"]);
  buildWheelItems("hour", Array.from({ length: 12 }, (_, i) => i + 1));
  const minuteOpts = [];
  for (let mm = 0; mm < 60; mm += MINUTE_STEP) minuteOpts.push(mm);
  buildWheelItems("minute", minuteOpts, (mm) => String(mm).padStart(2, "0"));

  timePickerModal.classList.remove("hidden");
  setWheelScroll("ampm", isPm ? 1 : 0);
  setWheelScroll("hour", h12 - 1);
  let minIdx = minuteOpts.indexOf(m);
  if (minIdx < 0) minIdx = Math.round(m / MINUTE_STEP) % minuteOpts.length;
  setWheelScroll("minute", minIdx);

  return new Promise((resolve) => {
    currentTimePickerResolve = resolve;
  });
}
function closeTimePicker(result) {
  timePickerModal.classList.add("hidden");
  const resolve = currentTimePickerResolve;
  currentTimePickerResolve = null;
  if (resolve) resolve(result);
}
timePickerOk.addEventListener("click", () => {
  const isPm = getWheelSelectedIndex("ampm") === 1;
  const h12 = parseInt(getWheelSelectedValue("hour"), 10);
  const m = parseInt(getWheelSelectedValue("minute"), 10);
  let h24 = h12 % 12;
  if (isPm) h24 += 12;
  closeTimePicker(`${String(h24).padStart(2, "0")}:${String(m).padStart(2, "0")}`);
});
timePickerCancel.addEventListener("click", () => closeTimePicker(null));
timePickerModal.addEventListener("click", (e) => {
  if (e.target === timePickerModal) closeTimePicker(null);
});

document.addEventListener("click", async (e) => {
  const btn = e.target.closest(".time-button");
  if (!btn || btn.disabled) return;
  const current = btn.dataset.time || "18:00";
  const newTime = await openTimePicker(current, "시간 선택");
  if (newTime) {
    btn.dataset.time = newTime;
    btn.textContent = formatTimeDisplay(newTime);
    btn.dispatchEvent(new Event("change", { bubbles: true }));
  }
});

// ===== Settings =====

function loadSettings() {
  try {
    const raw = JSON.parse(localStorage.getItem(SETTINGS_KEY) || "{}");
    return { ...DEFAULT_SETTINGS, ...raw };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}
function persistSettings(settings) {
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
}
function applySettingsToForm(s) {
  $("#duration").value = s.duration_minutes;
  $("#step").value = s.step_minutes;
  $("#first-start").value = s.first_start_time;
  $("#last-start").value = s.last_start_time;
  $("#break-minutes").value = s.break_minutes;
}

function openSettingsModal() {
  const s = loadSettings();
  settingsDuration.value = s.duration_minutes;
  settingsStep.value = s.step_minutes;
  settingsFirst.dataset.time = s.first_start_time;
  settingsFirst.textContent = formatTimeDisplay(s.first_start_time);
  settingsLast.dataset.time = s.last_start_time;
  settingsLast.textContent = formatTimeDisplay(s.last_start_time);
  settingsBreak.value = s.break_minutes;
  settingsModal.classList.remove("hidden");
}
function closeSettingsModal() {
  settingsModal.classList.add("hidden");
}
settingsSave.addEventListener("click", () => {
  const s = {
    duration_minutes: Number(settingsDuration.value) || DEFAULT_SETTINGS.duration_minutes,
    step_minutes: Number(settingsStep.value) || DEFAULT_SETTINGS.step_minutes,
    first_start_time: settingsFirst.dataset.time || DEFAULT_SETTINGS.first_start_time,
    last_start_time: settingsLast.dataset.time || DEFAULT_SETTINGS.last_start_time,
    break_minutes: Number(settingsBreak.value) || 0,
  };
  if (timeToMinutes(s.first_start_time) > timeToMinutes(s.last_start_time)) {
    showAlert("설정 오류", ["첫 시작 시간은 마지막 시작 시간보다 같거나 빨라야 합니다."]);
    return;
  }
  persistSettings(s);
  applySettingsToForm(s);
  closeSettingsModal();
  renderGrid(currentWeekSessions);
  showToast("설정 저장됨", { duration: 1500 });
});
settingsCancel.addEventListener("click", closeSettingsModal);
settingsModal.addEventListener("click", (e) => {
  if (e.target === settingsModal) closeSettingsModal();
});
openSettingsBtn.addEventListener("click", openSettingsModal);

// ===== Members list (with search/sort) =====

function getMonthCounts() {
  const now = new Date();
  const monthISO = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}`;
  return allMonthCounts(monthISO);
}

function filterAndSortMembers(members) {
  const term = (memberSearchInput.value || "").trim().toLowerCase();
  let filtered = term
    ? members.filter((m) => (m.name || "").toLowerCase().includes(term))
    : members.slice();

  const sortBy = memberSortSelect.value || "name";
  switch (sortBy) {
    case "recent":
      filtered.sort((a, b) => (b.created_at || "").localeCompare(a.created_at || ""));
      break;
    case "remaining": {
      const rem = (m) => {
        const ms = m.membership;
        if (!ms || !ms.total_count) return Infinity;
        return Math.max(0, ms.total_count - (ms.used_count || 0));
      };
      filtered.sort((a, b) => rem(a) - rem(b));
      break;
    }
    case "active":
      filtered.sort((a, b) => {
        if (a.active !== b.active) return a.active ? -1 : 1;
        return (a.name || "").localeCompare(b.name || "");
      });
      break;
    case "name":
    default:
      filtered.sort((a, b) => (a.name || "").localeCompare(b.name || ""));
  }
  return filtered;
}

function renderMembersList() {
  const all = loadMembers();
  if (all.length === 0) {
    membersListEl.innerHTML = `
      <div class="member-list-empty">
        회원이 없습니다.<br>아래 <strong>+ 회원 추가</strong> 또는 우측 상단 <strong>예시 적용</strong>으로 시작하세요.
      </div>
    `;
    return;
  }
  const members = filterAndSortMembers(all);
  if (members.length === 0) {
    membersListEl.innerHTML = `<div class="member-list-empty">검색 결과가 없습니다.</div>`;
    return;
  }
  const statsMap = getMonthCounts();

  membersListEl.innerHTML = members
    .map((m) => {
      const availSummary = (m.availability || [])
        .map((a) => {
          const dayStr = (a.days || []).join(",");
          return `${dayStr} ${formatTimeDisplay(a.start_time)}~${formatTimeDisplay(a.end_time)}`;
        })
        .join(" / ");

      let ptBadge = "";
      const ms = m.membership;
      if (ms && ms.total_count > 0) {
        const remaining = Math.max(0, ms.total_count - (ms.used_count || 0));
        const cls = remaining === 0 ? "exhausted" : "";
        ptBadge = `<div class="member-card-pt ${cls}">PT ${remaining}/${ms.total_count}</div>`;
      }
      const extras = [];
      if (m.phone) extras.push(`<span>📞 ${escapeHtml(m.phone)}</span>`);
      if (ms?.expires_at) extras.push(`<span>만료 ${escapeHtml(ms.expires_at)}</span>`);
      const stat = statsMap.get(m.id);
      if (stat && (stat.completed > 0 || stat.noShow > 0)) {
        const parts = [];
        if (stat.completed > 0) parts.push(`완료 ${stat.completed}`);
        if (stat.noShow > 0) parts.push(`노쇼 ${stat.noShow}`);
        extras.push(`<span>📅 이번달 ${parts.join(" / ")}</span>`);
      }
      if (m.memo) {
        const oneLine = m.memo.split("\n")[0];
        const trimmed = oneLine.length > 30 ? oneLine.slice(0, 30) + "…" : oneLine;
        extras.push(`<span>📝 ${escapeHtml(trimmed)}</span>`);
      }
      const extraLine = extras.length ? `<div class="member-card-extra">${extras.join("")}</div>` : "";

      return `
        <div class="member-card ${m.active ? "" : "inactive"}" data-member-id="${escapeAttr(m.id)}">
          <div class="member-card-info">
            <div class="member-card-name">${escapeHtml(m.name || "(이름 없음)")}</div>
            <div class="member-card-meta">주 ${m.weekly_count}회${availSummary ? " · " + escapeHtml(availSummary) : ""}</div>
            ${extraLine}
          </div>
          <div class="member-card-badges">
            <div class="member-card-status">${m.active ? "활성" : "휴원"}</div>
            ${ptBadge}
          </div>
        </div>
      `;
    })
    .join("");

  membersListEl.querySelectorAll(".member-card").forEach((card) => {
    card.addEventListener("click", () => openMemberModal({ memberId: card.dataset.memberId }));
  });
}
memberSearchInput.addEventListener("input", () => renderMembersList());
memberSortSelect.addEventListener("change", () => renderMembersList());

// ===== Member modal =====

function addMemberConditionRow(condition = {}) {
  conditionId += 1;
  const id = conditionId;
  const startVal = condition.start_time || "18:00";
  const endVal = condition.end_time || "23:00";
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
        <label>시작</label>
        <button type="button" class="time-button start-time" data-time="${escapeAttr(startVal)}">${escapeHtml(formatTimeDisplay(startVal))}</button>
      </div>
      <div class="field">
        <label>종료</label>
        <button type="button" class="time-button end-time" data-time="${escapeAttr(endVal)}">${escapeHtml(formatTimeDisplay(endVal))}</button>
      </div>
    </div>
  `;
  row.querySelector(".remove-condition").addEventListener("click", () => {
    row.remove();
    if (!memberModalConditions.children.length) addMemberConditionRow();
  });
  memberModalConditions.appendChild(row);
}
function collectMemberConditions() {
  const conditions = [];
  memberModalConditions.querySelectorAll(".condition-row").forEach((row) => {
    const days = [...row.querySelectorAll(".day input:checked")].map((i) => i.value);
    const start_time = row.querySelector(".start-time").dataset.time;
    const end_time = row.querySelector(".end-time").dataset.time;
    conditions.push({ days, start_time, end_time });
  });
  return conditions;
}
function showMemberModalError(message) {
  if (!message) {
    memberModalError.classList.add("hidden");
    memberModalError.textContent = "";
  } else {
    memberModalError.classList.remove("hidden");
    memberModalError.textContent = message;
  }
}

function openMemberModal({ memberId = null } = {}) {
  memberModalEditId = memberId;
  const isEdit = memberId != null;
  let member = null;
  if (isEdit) {
    member = loadMembers().find((m) => m.id === memberId);
    if (!member) {
      memberModalEditId = null;
      return;
    }
  }

  memberModalTitle.textContent = isEdit ? "회원 편집" : "회원 추가";
  memberModalName.value = member?.name || "";
  memberModalPhone.value = member?.phone || "";
  memberModalCount.value = member?.weekly_count || 2;
  memberModalActive.checked = member ? member.active : true;
  memberModalMemo.value = member?.memo || "";

  const ms = member?.membership || null;
  memberModalPlanName.value = ms?.plan_name || "";
  memberModalTotalCount.value = ms?.total_count || "";
  memberModalUsedCount.value = ms?.used_count || "";
  memberModalPrice.value = ms?.price ?? "";
  memberModalExpires.value = ms?.expires_at || "";

  memberModalDelete.classList.toggle("hidden", !isEdit);
  showMemberModalError("");

  memberModalConditions.innerHTML = "";
  conditionId = 0;
  const conds = member?.availability?.length
    ? member.availability
    : [{ days: [], start_time: "18:00", end_time: "23:00" }];
  conds.forEach((c) => addMemberConditionRow(c));

  memberModal.classList.remove("hidden");
  // 자동 포커스 제거 — 키패드 자동 열림 방지
}
function closeMemberModal() {
  memberModal.classList.add("hidden");
  memberModalEditId = null;
}
memberModalCancel.addEventListener("click", closeMemberModal);
memberModal.addEventListener("click", (e) => { if (e.target === memberModal) closeMemberModal(); });
memberModalAddCondition.addEventListener("click", () => addMemberConditionRow());

memberModalSave.addEventListener("click", () => {
  const name = memberModalName.value.trim();
  const phone = memberModalPhone.value.trim();
  const memo = memberModalMemo.value.trim();
  const weekly_count = Number(memberModalCount.value) || 1;
  const active = memberModalActive.checked;
  const availability = collectMemberConditions();

  if (!name) return showMemberModalError("이름을 입력하세요.");
  if (weekly_count < 1 || weekly_count > 7) return showMemberModalError("주 횟수는 1~7 사이여야 합니다.");
  const hasValidCondition = availability.some((a) => (a.days || []).length > 0 && a.start_time && a.end_time);
  if (!hasValidCondition) return showMemberModalError("최소 1개의 가능 시간대 (요일 + 시간) 를 입력하세요.");
  for (const a of availability) {
    if (a.days.length === 0) continue;
    if (timeToMinutes(a.end_time) < timeToMinutes(a.start_time)) {
      return showMemberModalError("종료 시간이 시작 시간보다 빠른 시간대가 있습니다.");
    }
  }

  const planName = memberModalPlanName.value.trim();
  const totalCount = Number(memberModalTotalCount.value) || 0;
  const usedCount = Number(memberModalUsedCount.value) || 0;
  const price = memberModalPrice.value ? Number(memberModalPrice.value) : null;
  const expiresAt = memberModalExpires.value;

  let membership = null;
  if (planName || totalCount > 0) {
    if (usedCount > totalCount) return showMemberModalError("사용 횟수가 총 횟수보다 많습니다.");
    membership = { plan_name: planName, total_count: totalCount, used_count: usedCount, price, started_at: "", expires_at: expiresAt };
  }

  const payload = { name, phone, memo, weekly_count, active, availability, membership };
  if (memberModalEditId) {
    updateMember(memberModalEditId, payload);
    showToast(`${name} 수정됨`, { duration: 1500 });
  } else {
    addMember(payload);
    showToast(`${name} 추가됨`, { duration: 1500 });
  }
  refreshCurrentWeek();
  closeMemberModal();
});

memberModalDelete.addEventListener("click", async () => {
  if (!memberModalEditId) return;
  const member = loadMembers().find((m) => m.id === memberModalEditId);
  const ok = await openConfirmModal("회원 삭제", `${member?.name || "회원"} 을 삭제하시겠습니까? 시간표 데이터는 유지됩니다.`);
  if (!ok) return;
  const name = member?.name;
  deleteMember(memberModalEditId);
  refreshCurrentWeek();
  closeMemberModal();
  if (name) showToast(`${name} 삭제됨`, { duration: 1500 });
});

// ===== Sample =====

function applySample() {
  const grouped = new Map();
  for (const row of SAMPLE_STUDENTS) {
    if (!grouped.has(row.name)) {
      grouped.set(row.name, { name: row.name, weekly_count: row.weekly_count, active: true, availability: [] });
    }
    grouped.get(row.name).availability.push({ days: row.days, start_time: row.start_time, end_time: row.end_time });
  }
  const members = [...grouped.values()].map((m) => ({
    ...m,
    id: `m_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`,
    created_at: new Date().toISOString(),
  }));
  replaceAllMembers(members);
  renderMembersList();
  showToast(`예시 회원 ${members.length}명 적용됨`, { duration: 1500 });
}

// ===== Week navigation & sessions sync =====

function renderWeekNav() {
  weekRangeText.textContent = formatWeekRange(currentWeekStart);
  weekMonthText.textContent = formatYearMonth(currentWeekStart);
}

function refreshCurrentWeek() {
  currentWeekSessions = loadSessionsForWeek(currentWeekStart);
  renderWeekNav();
  renderGrid(currentWeekSessions);
  syncActionButtons();
  renderMembersList();
}

function syncActionButtons() {
  const has = currentWeekSessions.length > 0;
  excelButton.disabled = !has;
  clearButton.disabled = !has;
  shareButton.disabled = !has;
}

function updateEmptyState() {
  const hasMembers = loadMembers().some((m) => m.active && m.name?.trim());
  const hasSessions = currentWeekSessions.length > 0;

  // 상호 배타: 회원 없음 vs 일정 없음 vs 일정 있음
  if (!hasMembers) {
    emptyStateNoMembers.classList.remove("hidden");
    emptyStateNoSessions.classList.add("hidden");
    gridTarget.classList.add("hidden");
  } else if (!hasSessions) {
    emptyStateNoMembers.classList.add("hidden");
    emptyStateNoSessions.classList.remove("hidden");
    emptyWeekText.textContent = `${formatWeekRange(currentWeekStart)} 일정이 없습니다`;
    gridTarget.classList.add("hidden");
  } else {
    emptyStateNoMembers.classList.add("hidden");
    emptyStateNoSessions.classList.add("hidden");
    gridTarget.classList.remove("hidden");
  }
}

weekPrevBtn.addEventListener("click", () => {
  currentWeekStart = addWeeks(currentWeekStart, -1);
  refreshCurrentWeek();
});
weekNextBtn.addEventListener("click", () => {
  currentWeekStart = addWeeks(currentWeekStart, 1);
  refreshCurrentWeek();
});
weekLabelBtn.addEventListener("click", () => {
  currentWeekStart = weekStartISO(todayISO());
  refreshCurrentWeek();
  showToast("이번 주", { duration: 1000 });
});

// ===== Grid =====

function colorForName(name) {
  const palette = [
    ["#dbeafe", "#1e3a8a"], ["#dcfce7", "#14532d"], ["#fef3c7", "#78350f"],
    ["#fce7f3", "#831843"], ["#ede9fe", "#4c1d95"], ["#cffafe", "#164e63"],
    ["#ffedd5", "#7c2d12"], ["#e0e7ff", "#312e81"],
  ];
  let hash = 0;
  for (const char of name) hash = (hash * 31 + char.charCodeAt(0)) >>> 0;
  return palette[hash % palette.length];
}

function renderGrid(sessions) {
  const firstSetting = timeToMinutes($("#first-start").value || "10:00");
  const lastSetting = timeToMinutes($("#last-start").value || "23:00");
  const duration = Number($("#duration").value) || 50;
  const starts = sessions.length ? sessions.map((s) => timeToMinutes(s.start)) : [firstSetting];
  const ends = sessions.length ? sessions.map((s) => timeToMinutes(s.end)) : [lastSetting + duration];
  const firstHour = Math.floor(Math.min(firstSetting, ...starts) / 60);
  const lastHour = Math.ceil(Math.max(lastSetting + duration, ...ends) / 60);
  const hours = [];
  for (let hour = firstHour; hour < lastHour; hour += 1) hours.push(hour);
  const totalMinutes = (lastHour - firstHour) * 60 || 1;

  const dateIsos = weekDays(currentWeekStart); // [iso0..iso6]
  const today = todayISO();

  const eventsByDay = new Map(DAYS.map((day) => [day, []]));
  sessions.forEach((s, index) => {
    const start = timeToMinutes(s.start);
    const end = timeToMinutes(s.end);
    eventsByDay.get(s.day).push({
      ...s,
      index,
      topPct: ((start - firstHour * 60) / totalMinutes) * 100,
      heightPct: ((end - start) / totalMinutes) * 100,
    });
  });

  const header = [
    '<div class="schedule-header"><div class="schedule-head">시간</div>',
    ...DAYS.map((day, i) => {
      const iso = dateIsos[i];
      const todayClass = iso === today ? " is-today" : "";
      return `<div class="schedule-head${todayClass}">
        <span class="schedule-head-dow">${escapeHtml(day)}</span>
        <span class="schedule-head-day">${formatDateShort(iso)}</span>
      </div>`;
    }),
    "</div>",
  ].join("");

  const timeColumn = hours.map((hour) => `<div class="schedule-time">${hourLabel(hour)}</div>`).join("");
  const dayColumns = DAYS.map((day, i) => {
    const iso = dateIsos[i];
    const todayClass = iso === today ? " is-today" : "";
    const events = eventsByDay.get(day).map((event) => {
      const [bg, ink] = colorForName(event.name);
      const lockClass = event.source === "fixed" ? "is-locked" : "";
      const style = `top:${event.topPct}%;height:${event.heightPct}%;--event-bg:${bg};--event-ink:${ink}`;
      const statusAttr = event.status && event.status !== "pending" ? ` data-status="${escapeAttr(event.status)}"` : "";
      return `
        <div class="schedule-event ${lockClass}" data-index="${event.index}" data-session-id="${escapeAttr(event.id)}" style="${style}"${statusAttr}>
          ${escapeHtml(event.name)}
          <small>${escapeHtml(event.start)}-${escapeHtml(event.end)}</small>
        </div>
      `;
    }).join("");
    return `<div class="schedule-day${todayClass}" data-day="${day}">${events}</div>`;
  }).join("");
  const body = `<div class="schedule-body"><div class="schedule-time-column">${timeColumn}</div>${dayColumns}</div>`;

  gridTarget.innerHTML = `<div class="schedule" style="--hour-count:${hours.length}">${header}${body}</div>`;
  wireScheduleInteractions();
  updateEmptyState();
}

function wireScheduleInteractions() {
  gridTarget.querySelectorAll(".schedule-event").forEach((eventEl) => {
    eventEl.addEventListener("click", () => {
      const sessionId = eventEl.dataset.sessionId;
      if (sessionId) openEventModal({ sessionId });
    });
  });
}

// ===== Session ops =====

function hasConflict(candidate, ignoreId = null) {
  const cStart = timeToMinutes(candidate.start);
  const cEnd = timeToMinutes(candidate.end);
  return currentWeekSessions.some((s) => {
    if (s.id === ignoreId) return false;
    if (s.day !== candidate.day) return false;
    const oStart = timeToMinutes(s.start);
    const oEnd = timeToMinutes(s.end);
    return cStart < oEnd && oStart < cEnd;
  });
}

function applyStatusToSession(sessionId, newStatus) {
  const session = currentWeekSessions.find((s) => s.id === sessionId);
  if (!session) return;
  const oldStatus = session.status || "pending";
  if (oldStatus === newStatus) return;

  const oldConsumed = statusConsumes(oldStatus);
  const newConsumed = statusConsumes(newStatus);
  const delta = (newConsumed ? 1 : 0) - (oldConsumed ? 1 : 0);

  if (session.member_id && delta !== 0) {
    const members = loadMembers();
    const member = members.find((m) => m.id === session.member_id);
    if (member?.membership && member.membership.total_count > 0) {
      const newUsed = Math.max(0, Math.min(member.membership.total_count, (member.membership.used_count || 0) + delta));
      updateMember(member.id, { membership: { ...member.membership, used_count: newUsed } });
    }
  }
  updateSession(sessionId, { status: newStatus });
  refreshCurrentWeek();
}

function removeSession(sessionId) {
  const s = currentWeekSessions.find((x) => x.id === sessionId);
  if (!s) return;
  // 만약 완료/노쇼였다면 PT 복원
  if (s.member_id && statusConsumes(s.status)) {
    const member = loadMembers().find((m) => m.id === s.member_id);
    if (member?.membership && member.membership.total_count > 0) {
      const newUsed = Math.max(0, (member.membership.used_count || 0) - 1);
      updateMember(member.id, { membership: { ...member.membership, used_count: newUsed } });
    }
  }
  deleteSession(sessionId);
  refreshCurrentWeek();
  showToast(`${s.day} ${s.start} ${s.name} 삭제`, { duration: 1500 });
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
function setEventModalTime(btn, hhmm) {
  btn.dataset.time = hhmm;
  btn.textContent = formatTimeDisplay(hhmm);
}
function updateEventModalStatusUI(status) {
  const s = status || "pending";
  eventModalStatusBadge.textContent = STATUS_LABEL[s] || STATUS_LABEL.pending;
  eventModalStatusBadge.dataset.status = s;
  eventModalMarkReset.classList.toggle("hidden", s === "pending");
}

function activeMemberNames() {
  return loadMembers().filter((m) => m.active && m.name?.trim()).map((m) => m.name);
}

function populateMemberSelect(selectedId = null) {
  const members = loadMembers().filter((m) => m.active && m.name?.trim());
  if (members.length === 0) {
    eventModalMemberSelect.innerHTML = `<option value="">(등록된 회원 없음)</option>`;
    eventModalMemberSelect.disabled = true;
    return;
  }
  eventModalMemberSelect.disabled = false;
  eventModalMemberSelect.innerHTML = members
    .map((m) => `<option value="${escapeAttr(m.id)}"${m.id === selectedId ? " selected" : ""}>${escapeHtml(m.name)}</option>`)
    .join("");
}

function setEventModalType(type) {
  const isMember = type === "member";
  eventTypeMemberRadio.checked = isMember;
  eventTypePersonalRadio.checked = !isMember;
  eventModalMemberRow.classList.toggle("hidden", !isMember);
  eventModalNameRow.classList.toggle("hidden", isMember);
}

function openEventModal({ sessionId = null } = {}) {
  const session = sessionId ? currentWeekSessions.find((s) => s.id === sessionId) : null;
  const isEdit = !!session;

  eventModalTitle.textContent = isEdit ? "일정 편집" : "일정 추가";

  // 일정 종류 결정
  let type = "member";
  if (isEdit) {
    type = session.member_id ? "member" : "personal";
  } else {
    // 신규: 활성 회원이 있으면 회원 수업 기본, 없으면 개인 일정
    const hasMembers = loadMembers().some((m) => m.active && m.name?.trim());
    type = hasMembers ? "member" : "personal";
  }
  setEventModalType(type);

  // 회원 select 채우기
  populateMemberSelect(session?.member_id || null);

  // 개인 일정 이름
  eventModalName.value = session && !session.member_id ? session.name : "";

  // 날짜/시간
  eventModalDay.value = session?.day ?? "월";
  const startVal = session?.start ?? "18:00";
  setEventModalTime(eventModalStart, startVal);
  setEventModalTime(eventModalEnd, session?.end ?? defaultEndForStart(startVal));
  eventModalDelete.classList.toggle("hidden", !isEdit);
  showEventModalError("");

  const showStatusSection = isEdit && session.member_id;
  eventModalStatusSection.classList.toggle("hidden", !showStatusSection);
  if (showStatusSection) updateEventModalStatusUI(session.status || "pending");

  // 잠금 토글: 편집 모드에만 표시 (새 추가 시는 자동 fixed)
  eventModalLockRow.classList.toggle("hidden", !isEdit);
  eventModalLocked.checked = isEdit ? session.source === "fixed" : true;

  eventModal.classList.remove("hidden");
  // 자동 포커스 제거 — 키패드 자동 열림 방지

  const mark = (newStatus) => {
    if (!isEdit) return;
    applyStatusToSession(session.id, newStatus);
    showToast(`${STATUS_LABEL[newStatus]} 처리됨`, { duration: 1200 });
    cleanup();
  };
  const onMarkCompleted = () => mark("completed");
  const onMarkNoShow = () => mark("no-show");
  const onMarkCancelled = () => mark("cancelled");
  const onMarkReset = () => mark("pending");

  const onTypeChange = () => {
    setEventModalType(eventTypeMemberRadio.checked ? "member" : "personal");
  };

  const cleanup = () => {
    eventModal.classList.add("hidden");
    eventModalSave.removeEventListener("click", onSave);
    eventModalCancel.removeEventListener("click", onCancel);
    eventModalDelete.removeEventListener("click", onDelete);
    eventModal.removeEventListener("click", onBackdrop);
    eventModalStart.removeEventListener("change", onStartChange);
    eventModalMarkCompleted.removeEventListener("click", onMarkCompleted);
    eventModalMarkNoShow.removeEventListener("click", onMarkNoShow);
    eventModalMarkCancelled.removeEventListener("click", onMarkCancelled);
    eventModalMarkReset.removeEventListener("click", onMarkReset);
    eventTypeMemberRadio.removeEventListener("change", onTypeChange);
    eventTypePersonalRadio.removeEventListener("change", onTypeChange);
  };

  const onStartChange = () => {
    const s = timeToMinutes(eventModalStart.dataset.time);
    const e = timeToMinutes(eventModalEnd.dataset.time);
    if (e <= s) setEventModalTime(eventModalEnd, defaultEndForStart(eventModalStart.dataset.time));
  };

  const onSave = () => {
    const isMemberType = eventTypeMemberRadio.checked;
    const day = eventModalDay.value;
    const start = eventModalStart.dataset.time;
    const end = eventModalEnd.dataset.time;

    let name = "";
    let memberId = null;

    if (isMemberType) {
      memberId = eventModalMemberSelect.value;
      if (!memberId) return showEventModalError("회원을 선택하세요.");
      const member = loadMembers().find((m) => m.id === memberId);
      if (!member) return showEventModalError("선택한 회원을 찾을 수 없습니다.");
      name = member.name;
    } else {
      name = eventModalName.value.trim();
      if (!name) return showEventModalError("일정 이름을 입력하세요.");
    }

    if (!start || !end) return showEventModalError("시작/종료 시간을 입력하세요.");
    if (timeToMinutes(end) <= timeToMinutes(start)) {
      return showEventModalError("종료 시간은 시작 시간보다 늦어야 합니다.");
    }
    if (hasConflict({ day, start, end }, isEdit ? session.id : null)) {
      return showEventModalError(`${day} ${start}-${end} 에 이미 다른 일정이 있습니다.`);
    }

    const date = dayNameToISO(currentWeekStart, day);
    const lockedNow = isEdit ? eventModalLocked.checked : true;
    const newSource = lockedNow ? "fixed" : "generated";

    if (isEdit) {
      updateSession(session.id, {
        name, day, date, start, end,
        member_id: memberId,
        source: newSource,
      });
    } else {
      addSession({
        date, day, start, end, name,
        member_id: memberId,
        source: "fixed",
        status: "pending",
      });
    }
    refreshCurrentWeek();
    showToast(isEdit ? "일정 수정됨" : (isMemberType ? `${name} 수업 추가됨` : "일정 추가됨"), { duration: 1500 });
    cleanup();
  };

  const onDelete = () => {
    if (!isEdit) return;
    removeSession(session.id);
    cleanup();
  };
  const onCancel = () => cleanup();
  const onBackdrop = (e) => { if (e.target === eventModal) cleanup(); };

  eventModalSave.addEventListener("click", onSave);
  eventModalCancel.addEventListener("click", onCancel);
  eventModalDelete.addEventListener("click", onDelete);
  eventModal.addEventListener("click", onBackdrop);
  eventModalStart.addEventListener("change", onStartChange);
  eventModalMarkCompleted.addEventListener("click", onMarkCompleted);
  eventModalMarkNoShow.addEventListener("click", onMarkNoShow);
  eventModalMarkCancelled.addEventListener("click", onMarkCancelled);
  eventModalMarkReset.addEventListener("click", onMarkReset);
  eventTypeMemberRadio.addEventListener("change", onTypeChange);
  eventTypePersonalRadio.addEventListener("change", onTypeChange);
}

// ===== Generate =====

function submitGenerate() {
  const members = loadMembers();
  const activeCount = members.filter((m) => m.active && m.name?.trim()).length;
  if (activeCount === 0) {
    showAlert("회원이 없습니다", ["회원 패널에서 회원을 추가하세요."]);
    return;
  }

  const fixedForSolver = currentWeekSessions
    .filter((s) => s.source === "fixed")
    .map((s) => ({ name: s.name, day: s.day, start: s.start, end: s.end }));

  const data2 = new FormData(form);
  const payload = {
    students: membersToStudentsPayload(members),
    fixed_assignments: fixedForSolver,
    duration_minutes: Number(data2.get("duration_minutes")),
    step_minutes: Number(data2.get("step_minutes")),
    first_start_time: data2.get("first_start_time"),
    last_start_time: data2.get("last_start_time"),
    break_minutes: Number(data2.get("break_minutes")),
  };

  hideAlert();
  showToast("시간표 생성 중…", { loading: true });

  setTimeout(() => {
    let data;
    try {
      data = generateSchedule(payload);
    } catch (exc) {
      hideToast();
      showAlert("오류", [exc.message || String(exc)]);
      return;
    }
    hideToast();
    if (!data.success) {
      showAlert(data.message, data.errors || []);
      return;
    }

    // 기존 generated sessions for this week 삭제
    deleteSessionsForWeekWhere(currentWeekStart, (s) => s.source === "generated");

    const byName = new Map(members.map((m) => [m.name, m.id]));
    const newSessions = [];
    for (const a of data.assignments) {
      const isAlreadyFixed = fixedForSolver.some(
        (f) => f.name === a.name && f.day === a.day && f.start === a.start && f.end === a.end
      );
      if (isAlreadyFixed) continue;
      newSessions.push({
        date: dayNameToISO(currentWeekStart, a.day),
        day: a.day,
        start: a.start,
        end: a.end,
        name: a.name,
        member_id: byName.get(a.name) || null,
        source: "generated",
        status: "pending",
      });
    }
    addSessions(newSessions);
    closeDrawer();
    refreshCurrentWeek();
    showToast(data.message, { duration: 1800 });
  }, 30);
}
form.addEventListener("submit", (event) => {
  event.preventDefault();
  submitGenerate();
});

// ===== Clear / Copy previous week =====

async function clearCurrentWeek() {
  if (currentWeekSessions.length === 0) return;
  const ok = await openConfirmModal(
    "이번 주 비우기",
    `${formatWeekRange(currentWeekStart)} 의 모든 일정을 삭제하시겠습니까?\n완료/노쇼 상태였던 일정은 회원의 PT 잔여 횟수가 자동 복원됩니다.`
  );
  if (!ok) return;

  // PT 복원
  for (const s of currentWeekSessions) {
    if (s.member_id && statusConsumes(s.status)) {
      const m = loadMembers().find((mm) => mm.id === s.member_id);
      if (m?.membership && m.membership.total_count > 0) {
        const newUsed = Math.max(0, (m.membership.used_count || 0) - 1);
        updateMember(m.id, { membership: { ...m.membership, used_count: newUsed } });
      }
    }
  }
  deleteSessionsForWeekWhere(currentWeekStart, () => true);
  refreshCurrentWeek();
  hideAlert();
  showToast("이번 주 비웠음", { duration: 1500 });
}

function copyPreviousWeek() {
  if (currentWeekSessions.length > 0) {
    showAlert("이미 일정이 있습니다", ["먼저 '비우기' 로 정리한 뒤 다시 시도하세요."]);
    return;
  }
  const prevStart = addWeeks(currentWeekStart, -1);
  const prevSessions = loadSessionsForWeek(prevStart);
  if (prevSessions.length === 0) {
    showToast("지난 주에 일정이 없습니다", { duration: 1500 });
    return;
  }
  // 복사된 일정은 모두 fixed 처리 → 다음 재생성 시 보호됨
  const newSessions = prevSessions.map((s) => ({
    date: dayNameToISO(currentWeekStart, s.day),
    day: s.day,
    start: s.start,
    end: s.end,
    name: s.name,
    member_id: s.member_id,
    source: "fixed",
    status: "pending",
  }));
  addSessions(newSessions);
  refreshCurrentWeek();
  showToast(`지난 주 일정 ${newSessions.length}개 복사 (모두 고정)`, { duration: 2000 });
}

// (확정 버튼 제거됨 — 잠금은 일정별 모달에서 토글)

// ===== Stats sheet =====

function perSessionRevenue(m) {
  const ms = m.membership;
  if (!ms || !ms.total_count || !ms.price) return 0;
  return ms.price / ms.total_count;
}

function renderStats() {
  const monthISO = statsMonthISO;
  const counts = allMonthCounts(monthISO);
  const members = loadMembers();

  let totalCompleted = 0;
  let totalNoShow = 0;
  let totalRevenue = 0;

  const memberRows = members.map((m) => {
    const c = counts.get(m.id) || { completed: 0, noShow: 0 };
    const perRev = perSessionRevenue(m);
    const consumed = c.completed + c.noShow;
    const revenue = Math.round(perRev * consumed);
    totalCompleted += c.completed;
    totalNoShow += c.noShow;
    totalRevenue += revenue;
    return { id: m.id, name: m.name, ...c, revenue };
  });

  // Month nav text
  statsMonthText.textContent = formatYearMonth(`${monthISO}-01`);

  // Summary
  statsCompleted.textContent = String(totalCompleted);
  statsNoShow.textContent = String(totalNoShow);
  statsRevenue.textContent = totalRevenue > 0
    ? `₩${totalRevenue.toLocaleString()}`
    : "—";

  // Mini calendar
  renderMiniCalendar(monthISO);

  // Member table
  const visibleRows = memberRows.filter((r) => r.completed > 0 || r.noShow > 0 || r.revenue > 0)
    .sort((a, b) => (b.completed + b.noShow) - (a.completed + a.noShow));
  if (visibleRows.length === 0) {
    statsMemberTable.innerHTML = '<div class="saved-list-empty">이번 달 활동 기록이 없습니다.</div>';
  } else {
    statsMemberTable.innerHTML = `
      <table class="stats-table">
        <thead><tr><th>회원</th><th>완료</th><th>노쇼</th><th>매출</th></tr></thead>
        <tbody>
          ${visibleRows.map((r) => `
            <tr>
              <td>${escapeHtml(r.name)}</td>
              <td>${r.completed}</td>
              <td>${r.noShow}</td>
              <td>${r.revenue > 0 ? "₩" + r.revenue.toLocaleString() : "—"}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `;
  }
}

function renderMiniCalendar(monthISO) {
  const grid = calendarGrid(`${monthISO}-01`);
  const sessionsByDate = new Map();
  for (const s of loadSessions()) {
    const arr = sessionsByDate.get(s.date) || [];
    arr.push(s);
    sessionsByDate.set(s.date, arr);
  }

  const headerCells = DAYS.map((d) => `<div class="mini-cal-head">${escapeHtml(d)}</div>`).join("");
  const cells = grid
    .map((c) => {
      const dayCells = sessionsByDate.get(c.iso) || [];
      const count = dayCells.length;
      const dotsHtml = count > 0
        ? `<div class="mini-cal-dots">${
            dayCells.slice(0, 3).map((s) => {
              const [bg] = colorForName(s.name);
              return `<span class="mini-cal-dot" style="background:${bg}"></span>`;
            }).join("")
          }${count > 3 ? `<span class="mini-cal-cell-more">+${count - 3}</span>` : ""}</div>`
        : "";
      const cls = [
        "mini-cal-cell",
        c.isSameMonth ? "" : "is-other-month",
        c.isToday ? "is-today" : "",
      ].filter(Boolean).join(" ");
      return `<button class="${cls}" type="button" data-iso="${c.iso}">
        <div class="mini-cal-cell-num">${c.day}</div>
        ${dotsHtml}
      </button>`;
    })
    .join("");

  miniCalEl.innerHTML = headerCells + cells;
  miniCalEl.querySelectorAll("[data-iso]").forEach((cell) => {
    cell.addEventListener("click", () => {
      currentWeekStart = weekStartISO(cell.dataset.iso);
      refreshCurrentWeek();
      closeStatsSheet();
    });
  });
}

statsMonthPrev.addEventListener("click", () => {
  statsMonthISO = monthOfISO(addMonths(`${statsMonthISO}-01`, -1));
  renderStats();
});
statsMonthNext.addEventListener("click", () => {
  statsMonthISO = monthOfISO(addMonths(`${statsMonthISO}-01`, 1));
  renderStats();
});
statsMonthToday.addEventListener("click", () => {
  statsMonthISO = monthOfISO(todayISO());
  renderStats();
});

// ===== Wiring =====

openDrawerBtn.addEventListener("click", openDrawer);
closeDrawerBtn.addEventListener("click", closeDrawer);
drawerBackdrop.addEventListener("click", closeDrawer);

showStatsBtn.addEventListener("click", openStatsSheet);
closeStatsBtn.addEventListener("click", closeStatsSheet);
statsBackdrop.addEventListener("click", closeStatsSheet);

openEventModalBtn.addEventListener("click", () => openEventModal());

addMemberButton.addEventListener("click", () => openMemberModal());
sampleButton.addEventListener("click", applySample);

regenerateButton.addEventListener("click", () => submitGenerate());
clearButton.addEventListener("click", clearCurrentWeek);
copyPrevButton.addEventListener("click", copyPreviousWeek);
emptyGenerateBtn.addEventListener("click", () => submitGenerate());
emptyCopyPrevBtn.addEventListener("click", copyPreviousWeek);
emptyAddMemberBtn.addEventListener("click", () => { openDrawer(); openMemberModal(); });
emptySampleBtn.addEventListener("click", applySample);

// More menu toggle
function toggleMoreMenu(force) {
  const isHidden = moreMenu.classList.contains("hidden");
  const wantOpen = force === undefined ? isHidden : force;
  moreMenu.classList.toggle("hidden", !wantOpen);
}
moreButton.addEventListener("click", (e) => {
  e.stopPropagation();
  toggleMoreMenu();
});
document.addEventListener("click", (e) => {
  if (moreMenu.classList.contains("hidden")) return;
  if (moreMenu.contains(e.target) || moreButton.contains(e.target)) return;
  toggleMoreMenu(false);
});
// 메뉴 안 버튼 클릭 시 메뉴 닫기
moreMenu.querySelectorAll("button").forEach((btn) => {
  btn.addEventListener("click", () => toggleMoreMenu(false));
});

excelButton.addEventListener("click", () => {
  if (currentWeekSessions.length === 0) return;
  try {
    downloadExcel(currentWeekSessions.map((s) => ({ day: s.day, start: s.start, end: s.end, name: s.name })));
    showToast("엑셀 다운로드", { duration: 1500 });
  } catch (exc) {
    showAlert("엑셀 저장 실패", [exc.message || String(exc)]);
  }
});

shareButton.addEventListener("click", async () => {
  if (currentWeekSessions.length === 0) return;
  const settings = loadSettings();
  showToast("이미지 생성 중…", { loading: true });
  try {
    const res = await shareScheduleImage({
      sessions: currentWeekSessions,
      weekStart: currentWeekStart,
      firstStartTime: settings.first_start_time,
      lastStartTime: settings.last_start_time,
      durationMinutes: settings.duration_minutes,
    });
    hideToast();
    if (res.method === "share") showToast("공유 시작됨", { duration: 1500 });
    else if (res.method === "download") showToast("이미지 다운로드됨", { duration: 1800 });
  } catch (exc) {
    hideToast();
    showAlert("공유 실패", [exc.message || String(exc)]);
  }
});

document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;
  if (!timePickerModal.classList.contains("hidden")) return closeTimePicker(null);
  if (!confirmModal.classList.contains("hidden")) return;
  if (!eventModal.classList.contains("hidden")) return;
  if (!memberModal.classList.contains("hidden")) return closeMemberModal();
  if (!textModal.classList.contains("hidden")) return;
  if (!settingsModal.classList.contains("hidden")) return closeSettingsModal();
  if (drawer.classList.contains("open")) return closeDrawer();
  if (!statsSheet.classList.contains("hidden")) return closeStatsSheet();
});

// ===== Onboarding =====

const ONBOARDING_KEY = "auto-scheduler-onboarding-seen";
function showOnboardingIfNeeded() {
  if (localStorage.getItem(ONBOARDING_KEY)) return;
  onboardingModal.classList.remove("hidden");
}
onboardingStartBtn.addEventListener("click", () => {
  localStorage.setItem(ONBOARDING_KEY, "1");
  onboardingModal.classList.add("hidden");
});
onboardingModal.addEventListener("click", (e) => {
  if (e.target === onboardingModal) {
    localStorage.setItem(ONBOARDING_KEY, "1");
    onboardingModal.classList.add("hidden");
  }
});

// ===== Init =====

purgeLegacySavedSchedules();      // 옛 saved-schedules 정리
applySettingsToForm(loadSettings());
refreshCurrentWeek();
showOnboardingIfNeeded();

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./service-worker.js").catch(() => {});
  });
}

window.addEventListener("error", (e) => {
  showAlert("스크립트 오류", [e.message || String(e.error || e)]);
});
window.addEventListener("unhandledrejection", (e) => {
  showAlert("처리되지 않은 오류", [String(e.reason?.message || e.reason || "")]);
});
