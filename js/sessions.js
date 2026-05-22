// 날짜 기반 세션 영속화.
// "세션" = 특정 날짜의 수업/일정 1개.

import { weekStartISO, weekEndISO, monthOfISO, isoToDayName } from "./dates.js";

const SESSIONS_KEY = "auto-scheduler-sessions";

/**
 * 세션 모델
 * {
 *   id: "s_xxx",
 *   date: "YYYY-MM-DD",
 *   day: "월" ~ "일",            // date 에서 파생, 캐싱
 *   start: "HH:MM",
 *   end: "HH:MM",
 *   name: string,                // 표시용
 *   member_id: string | null,    // 회원 연결 (fixed event 는 null)
 *   source: "generated" | "fixed",
 *   status: "pending" | "completed" | "no-show" | "cancelled",
 *   created_at: ISO,
 *   updated_at: ISO
 * }
 */

function newId() {
  return `s_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`;
}

export function loadSessions() {
  try {
    const raw = JSON.parse(localStorage.getItem(SESSIONS_KEY) || "[]");
    if (!Array.isArray(raw)) return [];
    return raw.map(normalizeSession);
  } catch {
    return [];
  }
}

function normalizeSession(s) {
  const date = s.date || "";
  return {
    id: s.id || newId(),
    date,
    day: s.day || (date ? isoToDayName(date) : ""),
    start: s.start || "",
    end: s.end || "",
    name: s.name || "",
    member_id: s.member_id || null,
    source: s.source === "fixed" ? "fixed" : "generated",
    status: s.status || "pending",
    created_at: s.created_at || new Date().toISOString(),
    updated_at: s.updated_at || s.created_at || new Date().toISOString(),
  };
}

export function saveSessions(sessions) {
  localStorage.setItem(SESSIONS_KEY, JSON.stringify(sessions));
}

export function createSession(input) {
  const now = new Date().toISOString();
  return normalizeSession({
    ...input,
    id: newId(),
    created_at: now,
    updated_at: now,
  });
}

export function addSession(input) {
  const sessions = loadSessions();
  const s = createSession(input);
  sessions.push(s);
  saveSessions(sessions);
  return s;
}

export function addSessions(inputs) {
  const sessions = loadSessions();
  const now = new Date().toISOString();
  const newOnes = inputs.map((input) => normalizeSession({
    ...input,
    id: newId(),
    created_at: now,
    updated_at: now,
  }));
  sessions.push(...newOnes);
  saveSessions(sessions);
  return newOnes;
}

export function updateSession(id, updates) {
  const sessions = loadSessions();
  const idx = sessions.findIndex((s) => s.id === id);
  if (idx < 0) return null;
  sessions[idx] = normalizeSession({
    ...sessions[idx],
    ...updates,
    updated_at: new Date().toISOString(),
  });
  saveSessions(sessions);
  return sessions[idx];
}

export function deleteSession(id) {
  const sessions = loadSessions().filter((s) => s.id !== id);
  saveSessions(sessions);
  return sessions;
}

/** 특정 주 (월요일 ISO 기준) 의 세션들. */
export function loadSessionsForWeek(weekStart) {
  const start = weekStart;
  const end = weekEndISO(weekStart);
  return loadSessions().filter((s) => s.date >= start && s.date <= end);
}

/** 해당 주의 세션 중 조건 일치하는 것만 삭제. */
export function deleteSessionsForWeekWhere(weekStart, predicate) {
  const start = weekStart;
  const end = weekEndISO(weekStart);
  const remaining = loadSessions().filter((s) => {
    if (s.date < start || s.date > end) return true;
    // 주 내부: predicate true 면 삭제
    return !predicate(s);
  });
  saveSessions(remaining);
}

/** 특정 회원의 세션 (정렬: 날짜+시작). */
export function sessionsByMember(memberId) {
  return loadSessions()
    .filter((s) => s.member_id === memberId)
    .sort((a, b) => (a.date.localeCompare(b.date) || a.start.localeCompare(b.start)));
}

/** 특정 회원의 status 카운트 (이번 달). */
export function memberMonthCounts(memberId, monthISO /* "YYYY-MM" */) {
  const counts = { completed: 0, "no-show": 0, cancelled: 0, pending: 0 };
  for (const s of loadSessions()) {
    if (s.member_id !== memberId) continue;
    if (monthOfISO(s.date) !== monthISO) continue;
    counts[s.status] = (counts[s.status] || 0) + 1;
  }
  return counts;
}

/** 회원의 전체 PT 소진 횟수 (status in completed/no-show). */
export function memberConsumedCount(memberId) {
  let n = 0;
  for (const s of loadSessions()) {
    if (s.member_id !== memberId) continue;
    if (s.status === "completed" || s.status === "no-show") n += 1;
  }
  return n;
}

/** 회원별 이번 달 카운트 맵 (id 기준). */
export function allMonthCounts(monthISO) {
  const map = new Map();
  for (const s of loadSessions()) {
    if (monthOfISO(s.date) !== monthISO) continue;
    const key = s.member_id;
    if (!key) continue;
    if (!map.has(key)) map.set(key, { completed: 0, noShow: 0 });
    if (s.status === "completed") map.get(key).completed += 1;
    else if (s.status === "no-show") map.get(key).noShow += 1;
  }
  return map;
}

/** 마이그레이션: 옛 saved-schedules 데이터 정리. */
export function purgeLegacySavedSchedules() {
  localStorage.removeItem("auto-scheduler-confirmed-schedules");
}
