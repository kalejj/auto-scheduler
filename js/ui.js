import { DAYS, SAMPLE_STUDENTS } from "./constants.js";
import { generateSchedule } from "./api.js";
import { downloadExcel } from "./excel.js";
import {
  loadMembers,
  saveMembers,
  addMember,
  updateMember,
  deleteMember,
  toggleMemberActive,
  replaceAllMembers,
  membersToStudentsPayload,
} from "./members.js";

const $ = (sel) => document.querySelector(sel);

const drawer = $("#drawer");
const drawerBackdrop = $("#drawer-backdrop");
const openDrawerBtn = $("#open-drawer-btn");
const closeDrawerBtn = $("#close-drawer-btn");
const emptyOpenDrawerBtn = $("#empty-open-drawer-btn");
const savedSheet = $("#saved-sheet");
const savedBackdrop = $("#saved-backdrop");
const showSavedBtn = $("#show-saved-btn");
const closeSavedBtn = $("#close-saved-btn");
const openSettingsBtn = $("#open-settings-btn");

const form = $("#schedule-form");
const membersListEl = $("#members-list");
const addMemberButton = $("#add-member-button");
const sampleButton = $("#sample-button");

const gridTarget = $("#grid");
const alertBanner = $("#alert-banner");
const emptyHint = $("#empty-hint");
const regenerateButton = $("#regenerate-button");
const confirmButton = $("#confirm-button");
const excelButton = $("#excel-button");
const clearButton = $("#clear-button");
const openEventModalBtn = $("#open-event-modal-btn");
const savedList = $("#saved-list");

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

const memberSearchInput = $("#member-search");
const memberSortSelect = $("#member-sort");

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

let latestAssignments = [];
let fixedAssignments = [];
let confirmedMode = false;
let conditionId = 0;
let toastTimer = null;
let memberModalEditId = null;

const STORAGE_KEY = "auto-scheduler-confirmed-schedules";
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
function findMemberByName(name) {
  if (!name) return null;
  return loadMembers().find((m) => m.name === name) || null;
}
function attachMemberIds(assignments) {
  const members = loadMembers();
  const byName = new Map(members.map((m) => [m.name, m.id]));
  return assignments.map((a) => ({
    ...a,
    member_id: a.member_id || byName.get(a.name) || null,
    status: a.status || "pending",
  }));
}

// ===== Member stats =====
function loadAllAssignmentsForStats() {
  // 저장된 시간표 + 현재 시간표의 status 기반 집계
  const saved = readSavedSchedules();
  const all = [...latestAssignments];
  for (const s of saved) {
    for (const a of s.assignments || []) {
      all.push({ ...a, _saved_at: s.created_at });
    }
  }
  return all;
}
function currentMonthCounts() {
  // 회원 이름 또는 id 기준으로 이번 달 완료/노쇼 카운트
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth();
  const all = loadAllAssignmentsForStats();
  const map = new Map(); // key: member_id || name → {completed, noShow}
  for (const a of all) {
    if (!a.status || a.status === "pending") continue;
    if (a.status !== "completed" && a.status !== "no-show") continue;
    const ts = a._saved_at ? new Date(a._saved_at) : now;
    if (ts.getFullYear() !== year || ts.getMonth() !== month) continue;
    const key = a.member_id || a.name;
    if (!key) continue;
    if (!map.has(key)) map.set(key, { completed: 0, noShow: 0 });
    if (a.status === "completed") map.get(key).completed += 1;
    else map.get(key).noShow += 1;
  }
  return map;
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

// ===== Drawer / Sheet =====

function openDrawer() {
  drawer.classList.add("open");
  drawerBackdrop.classList.remove("hidden");
}
function closeDrawer() {
  drawer.classList.remove("open");
  drawerBackdrop.classList.add("hidden");
}
function openSavedSheet() {
  renderSavedSchedules();
  savedSheet.classList.remove("hidden");
  savedBackdrop.classList.remove("hidden");
}
function closeSavedSheet() {
  savedSheet.classList.add("hidden");
  savedBackdrop.classList.add("hidden");
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
    wheelScrollDebouncers.set(
      wheel,
      setTimeout(() => updateSelectedItem(wheel), 60)
    );
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
  const hhmm = `${String(h24).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
  closeTimePicker(hhmm);
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
  renderGrid(latestAssignments);
  showToast("설정 저장됨", { duration: 1500 });
});
settingsCancel.addEventListener("click", closeSettingsModal);
settingsModal.addEventListener("click", (e) => {
  if (e.target === settingsModal) closeSettingsModal();
});
openSettingsBtn.addEventListener("click", openSettingsModal);

// ===== Members list (in drawer) =====

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
    membersListEl.innerHTML = `
      <div class="member-list-empty">검색 결과가 없습니다.</div>
    `;
    return;
  }

  const statsMap = currentMonthCounts();

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
      const stat = statsMap.get(m.id) || statsMap.get(m.name);
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
      const extraLine = extras.length
        ? `<div class="member-card-extra">${extras.join("")}</div>`
        : "";

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
    card.addEventListener("click", () => {
      openMemberModal({ memberId: card.dataset.memberId });
    });
  });
}

memberSearchInput.addEventListener("input", () => renderMembersList());
memberSortSelect.addEventListener("change", () => renderMembersList());

// ===== Member edit modal =====

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
  setTimeout(() => memberModalName.focus(), 50);
}

function closeMemberModal() {
  memberModal.classList.add("hidden");
  memberModalEditId = null;
}

memberModalCancel.addEventListener("click", closeMemberModal);
memberModal.addEventListener("click", (e) => {
  if (e.target === memberModal) closeMemberModal();
});
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

  const hasValidCondition = availability.some(
    (a) => (a.days || []).length > 0 && a.start_time && a.end_time
  );
  if (!hasValidCondition) {
    return showMemberModalError("최소 1개의 가능 시간대 (요일 + 시간) 를 입력하세요.");
  }

  for (const a of availability) {
    if (a.days.length === 0) continue;
    if (timeToMinutes(a.end_time) < timeToMinutes(a.start_time)) {
      return showMemberModalError("종료 시간이 시작 시간보다 빠른 시간대가 있습니다.");
    }
  }

  // 회원권 정보 수집
  const planName = memberModalPlanName.value.trim();
  const totalCount = Number(memberModalTotalCount.value) || 0;
  const usedCount = Number(memberModalUsedCount.value) || 0;
  const price = memberModalPrice.value ? Number(memberModalPrice.value) : null;
  const expiresAt = memberModalExpires.value;

  let membership = null;
  if (planName || totalCount > 0) {
    if (usedCount > totalCount) {
      return showMemberModalError("사용 횟수가 총 횟수보다 많습니다.");
    }
    membership = {
      plan_name: planName,
      total_count: totalCount,
      used_count: usedCount,
      price,
      started_at: "",
      expires_at: expiresAt,
    };
  }

  const payload = { name, phone, memo, weekly_count, active, availability, membership };

  if (memberModalEditId) {
    updateMember(memberModalEditId, payload);
    showToast(`${name} 수정됨`, { duration: 1500 });
  } else {
    addMember(payload);
    showToast(`${name} 추가됨`, { duration: 1500 });
  }
  renderMembersList();
  closeMemberModal();
});

memberModalDelete.addEventListener("click", async () => {
  if (!memberModalEditId) return;
  const member = loadMembers().find((m) => m.id === memberModalEditId);
  const ok = await openConfirmModal(
    "회원 삭제",
    `${member?.name || "회원"} 을 삭제하시겠습니까? 시간표 데이터는 유지됩니다.`
  );
  if (!ok) return;
  const name = member?.name;
  deleteMember(memberModalEditId);
  renderMembersList();
  closeMemberModal();
  if (name) showToast(`${name} 삭제됨`, { duration: 1500 });
});

// ===== Sample =====

function applySample() {
  // SAMPLE_STUDENTS 의 평탄 데이터를 회원 단위로 묶기
  const grouped = new Map();
  for (const row of SAMPLE_STUDENTS) {
    if (!grouped.has(row.name)) {
      grouped.set(row.name, {
        name: row.name,
        weekly_count: row.weekly_count,
        active: true,
        availability: [],
      });
    }
    grouped.get(row.name).availability.push({
      days: row.days,
      start_time: row.start_time,
      end_time: row.end_time,
    });
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

// ===== Schedule generation =====

function getPayload() {
  const data = new FormData(form);
  const members = loadMembers();
  return {
    students: membersToStudentsPayload(members),
    fixed_assignments: fixedAssignments,
    duration_minutes: Number(data.get("duration_minutes")),
    step_minutes: Number(data.get("step_minutes")),
    first_start_time: data.get("first_start_time"),
    last_start_time: data.get("last_start_time"),
    break_minutes: Number(data.get("break_minutes")),
  };
}

function activeMemberNames() {
  return loadMembers().filter((m) => m.active && m.name?.trim()).map((m) => m.name);
}

// ===== Grid =====

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
        const statusAttr = event.status && event.status !== "pending" ? ` data-status="${escapeAttr(event.status)}"` : "";
        return `
          <div class="schedule-event ${sourceClass}" data-index="${event.index}" style="${style}"${statusAttr}>
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

// ===== Event modal (고정 일정) =====

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

function applyStatusToAssignment(index, newStatus) {
  const a = latestAssignments[index];
  if (!a) return;
  const oldStatus = a.status || "pending";
  if (oldStatus === newStatus) return;

  const oldConsumed = statusConsumes(oldStatus);
  const newConsumed = statusConsumes(newStatus);
  const delta = (newConsumed ? 1 : 0) - (oldConsumed ? 1 : 0);

  // PT 자동 차감
  if (a.member_id && delta !== 0) {
    const members = loadMembers();
    const member = members.find((m) => m.id === a.member_id);
    if (member?.membership && member.membership.total_count > 0) {
      const newUsed = Math.max(
        0,
        Math.min(member.membership.total_count, (member.membership.used_count || 0) + delta)
      );
      updateMember(member.id, {
        membership: { ...member.membership, used_count: newUsed },
      });
    }
  }

  a.status = newStatus;
  latestAssignments[index] = a;
  renderGrid(latestAssignments);
  renderMembersList();
  syncEditedAssignments();
}

function openEventModal({ existingIndex } = {}) {
  const isEdit = existingIndex != null && Number.isFinite(existingIndex);
  const existing = isEdit ? latestAssignments[existingIndex] : null;

  eventModalTitle.textContent = isEdit ? "일정 편집" : "고정 일정 추가";
  eventModalName.value = existing?.name ?? activeMemberNames()[0] ?? "";
  eventModalDay.value = existing?.day ?? "월";
  const startVal = existing?.start ?? "18:00";
  setEventModalTime(eventModalStart, startVal);
  setEventModalTime(eventModalEnd, existing?.end ?? defaultEndForStart(startVal));
  eventModalDelete.classList.toggle("hidden", !isEdit);
  showEventModalError("");

  // 상태 섹션: 회원 연결된 (member_id 있는) 생성된 일정에만 표시
  const showStatusSection = isEdit && existing && existing.source !== "fixed" && existing.member_id;
  eventModalStatusSection.classList.toggle("hidden", !showStatusSection);
  if (showStatusSection) {
    updateEventModalStatusUI(existing.status || "pending");
  }

  eventModal.classList.remove("hidden");
  setTimeout(() => eventModalName.focus(), 50);

  const mark = (newStatus) => {
    if (!isEdit) return;
    applyStatusToAssignment(existingIndex, newStatus);
    updateEventModalStatusUI(newStatus);
    showToast(`${STATUS_LABEL[newStatus]} 처리됨`, { duration: 1200 });
    cleanup();
  };
  const onMarkCompleted = () => mark("completed");
  const onMarkNoShow = () => mark("no-show");
  const onMarkCancelled = () => mark("cancelled");
  const onMarkReset = () => mark("pending");

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
  };

  const onStartChange = () => {
    const s = timeToMinutes(eventModalStart.dataset.time);
    const e = timeToMinutes(eventModalEnd.dataset.time);
    if (e <= s) setEventModalTime(eventModalEnd, defaultEndForStart(eventModalStart.dataset.time));
  };

  const onSave = () => {
    const name = eventModalName.value.trim();
    const day = eventModalDay.value;
    const start = eventModalStart.dataset.time;
    const end = eventModalEnd.dataset.time;

    if (!name) return showEventModalError("이름을 입력하세요.");
    if (!start || !end) return showEventModalError("시작/종료 시간을 입력하세요.");
    if (timeToMinutes(end) <= timeToMinutes(start)) {
      return showEventModalError("종료 시간은 시작 시간보다 늦어야 합니다.");
    }
    if (hasConflict({ day, start, end }, isEdit ? existingIndex : -1)) {
      return showEventModalError(`${day} ${start}-${end} 에 이미 다른 일정이 있습니다.`);
    }

    const item = { name, day, start, end, source: "fixed" };
    if (isEdit) latestAssignments[existingIndex] = item;
    else latestAssignments.push(item);
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
}

// ===== Saved schedules =====

function readSavedSchedules() {
  try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]"); }
  catch { return []; }
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
        </div>`
    )
    .join("");
  savedList.querySelectorAll(".load-schedule").forEach((b) =>
    b.addEventListener("click", () => loadSavedSchedule(b.dataset.id))
  );
  savedList.querySelectorAll(".delete-schedule").forEach((b) =>
    b.addEventListener("click", () => deleteSavedSchedule(b.dataset.id))
  );
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
  latestAssignments = attachMemberIds(item.assignments.map((a) => ({ ...a })));
  fixedAssignments = latestAssignments.map((a) => ({ ...a }));
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
  hideToast();
  if (!data.success) {
    excelButton.disabled = true;
    confirmButton.disabled = true;
    showAlert(data.message, data.errors || []);
    return;
  }
  hideAlert();
  latestAssignments = attachMemberIds(markAssignmentSources(data.assignments, fixedAssignments));
  confirmedMode = false;
  excelButton.disabled = false;
  confirmButton.disabled = false;
  clearButton.disabled = false;
  renderGrid(latestAssignments);
  syncEditedAssignments();
  showToast(data.message, { duration: 1800 });
}

async function clearSchedule() {
  if (!latestAssignments.length) return;
  const ok = await openConfirmModal(
    "시간표 초기화",
    "현재 시간표의 모든 일정을 삭제하시겠습니까? 회원 정보는 유지됩니다."
  );
  if (!ok) return;
  latestAssignments = [];
  fixedAssignments = [];
  confirmedMode = false;
  renderGrid([]);
  syncEditedAssignments();
  hideAlert();
  showToast("초기화됨", { duration: 1200 });
}

// ===== Generate =====

function submitGenerate() {
  const members = loadMembers();
  const activeCount = members.filter((m) => m.active && m.name?.trim()).length;
  if (activeCount === 0) {
    showAlert("회원이 없습니다", ["회원 패널에서 회원을 추가하세요."]);
    return;
  }

  hideAlert();
  showToast("시간표 생성 중…", { loading: true });

  setTimeout(() => {
    let data;
    try {
      data = generateSchedule(getPayload());
    } catch (exc) {
      hideToast();
      showAlert("오류", [exc.message || String(exc)]);
      return;
    }
    if (data.success) closeDrawer();
    renderResult(data);
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

addMemberButton.addEventListener("click", () => openMemberModal());
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

document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;
  if (!timePickerModal.classList.contains("hidden")) return closeTimePicker(null);
  if (!confirmModal.classList.contains("hidden")) return;
  if (!eventModal.classList.contains("hidden")) return;
  if (!memberModal.classList.contains("hidden")) return closeMemberModal();
  if (!textModal.classList.contains("hidden")) return;
  if (!settingsModal.classList.contains("hidden")) return closeSettingsModal();
  if (drawer.classList.contains("open")) return closeDrawer();
  if (!savedSheet.classList.contains("hidden")) return closeSavedSheet();
});

// ===== Init =====

applySettingsToForm(loadSettings());
renderMembersList();
renderGrid([]);
syncEditedAssignments();

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
