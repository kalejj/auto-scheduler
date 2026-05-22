// 날짜/주 유틸. ISO 날짜 (YYYY-MM-DD) 문자열을 주 기본 단위로 사용.
// 주의 시작은 월요일 (한국 일반 관례).

import { DAYS } from "./constants.js";

export function pad2(n) {
  return String(n).padStart(2, "0");
}

export function todayISO() {
  return dateToISO(new Date());
}

export function dateToISO(date) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

export function isoToDate(iso) {
  // YYYY-MM-DD → Date (로컬 자정)
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, m - 1, d);
}

export function addDays(iso, n) {
  const d = isoToDate(iso);
  d.setDate(d.getDate() + n);
  return dateToISO(d);
}

export function addWeeks(iso, n) {
  return addDays(iso, n * 7);
}

/** 해당 날짜가 속한 주의 월요일 ISO. */
export function weekStartISO(iso) {
  const d = isoToDate(iso);
  const dow = d.getDay(); // 0=일 1=월 ... 6=토
  // 월요일까지 거리: 일=-6, 월=0, 화=-1 ...
  const offsetToMonday = dow === 0 ? -6 : 1 - dow;
  d.setDate(d.getDate() + offsetToMonday);
  return dateToISO(d);
}

export function weekEndISO(iso) {
  return addDays(weekStartISO(iso), 6);
}

/** 월요일 ISO 받아서 월~일 7개 ISO 배열 반환. */
export function weekDays(weekStart) {
  return Array.from({ length: 7 }, (_, i) => addDays(weekStart, i));
}

/** DAYS 배열의 인덱스를 받아 해당 주 날짜 ISO. (0=월, 6=일) */
export function dayIndexToISO(weekStart, dayIndex) {
  return addDays(weekStart, dayIndex);
}

/** 요일명 ("월","화",...) → 주 내 날짜 ISO. */
export function dayNameToISO(weekStart, dayName) {
  const idx = DAYS.indexOf(dayName);
  if (idx < 0) return null;
  return addDays(weekStart, idx);
}

/** ISO → "월","화",... */
export function isoToDayName(iso) {
  const d = isoToDate(iso);
  const dow = d.getDay(); // 0=일
  // 일=6, 월=0, 화=1 ...
  const idx = dow === 0 ? 6 : dow - 1;
  return DAYS[idx];
}

/** "12월 16일" */
export function formatDateLong(iso) {
  const d = isoToDate(iso);
  return `${d.getMonth() + 1}월 ${d.getDate()}일`;
}

/** "12/16" */
export function formatDateShort(iso) {
  const d = isoToDate(iso);
  return `${d.getMonth() + 1}/${d.getDate()}`;
}

/** "2025-12-16" 의 일(day) 부분만 → 16 */
export function dayOfMonth(iso) {
  return Number(iso.slice(8, 10));
}

/** "12월 16일 ~ 22일" (월요일 ISO 받음) */
export function formatWeekRange(weekStart) {
  const start = isoToDate(weekStart);
  const end = isoToDate(addDays(weekStart, 6));
  const sm = start.getMonth() + 1;
  const sd = start.getDate();
  const em = end.getMonth() + 1;
  const ed = end.getDate();
  if (sm === em) return `${sm}월 ${sd}일 ~ ${ed}일`;
  return `${sm}월 ${sd}일 ~ ${em}월 ${ed}일`;
}

/** "2025년 12월" 등 (월요일 ISO 받음, 해당 주의 첫 날 기준 월) */
export function formatYearMonth(iso) {
  const d = isoToDate(iso);
  return `${d.getFullYear()}년 ${d.getMonth() + 1}월`;
}

export function isToday(iso) {
  return iso === todayISO();
}

export function isSameMonth(isoA, isoB) {
  return isoA.slice(0, 7) === isoB.slice(0, 7);
}

export function monthOfISO(iso) {
  return iso.slice(0, 7); // "YYYY-MM"
}

export function monthStartISO(iso) {
  const d = isoToDate(iso);
  d.setDate(1);
  return dateToISO(d);
}

export function monthEndISO(iso) {
  const d = isoToDate(iso);
  d.setMonth(d.getMonth() + 1);
  d.setDate(0);
  return dateToISO(d);
}

export function addMonths(iso, n) {
  const d = isoToDate(iso);
  d.setMonth(d.getMonth() + n);
  return dateToISO(d);
}

/**
 * 월간 캘린더 그리드 (6주 × 7일 = 42 셀, 월~일 시작).
 * @param {string} iso "YYYY-MM-DD" (해당 월의 어느 날짜든 OK)
 * @returns 각 셀 { iso, day, isSameMonth, isToday }
 */
export function calendarGrid(iso) {
  const monthStart = monthStartISO(iso);
  const target = monthOfISO(monthStart);
  const start = weekStartISO(monthStart);
  const today = todayISO();
  const cells = [];
  let cur = start;
  for (let i = 0; i < 42; i++) {
    const d = isoToDate(cur);
    cells.push({
      iso: cur,
      day: d.getDate(),
      isSameMonth: monthOfISO(cur) === target,
      isToday: cur === today,
    });
    cur = addDays(cur, 1);
  }
  return cells;
}
