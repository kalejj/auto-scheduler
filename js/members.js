// 회원 데이터 영속화 + CRUD. localStorage 사용.
// 다음 phase 에 클라우드 동기화로 교체될 때 이 모듈만 바꾸면 됨.

const MEMBERS_KEY = "auto-scheduler-members";

/**
 * 회원 데이터 모델
 * {
 *   id: string,                  // 영속 ID
 *   name: string,                // 이름
 *   phone: string,               // 전화번호 (선택, 추후 알림톡)
 *   memo: string,                // 메모 (선택)
 *   weekly_count: number,        // 주 횟수
 *   active: boolean,             // 활성 여부 (시간표 생성에 포함될지)
 *   availability: [              // 가능 시간대 목록
 *     {
 *       days: ["월", "화", ...], // 요일
 *       start_time: "HH:MM",
 *       end_time: "HH:MM"
 *     }
 *   ],
 *   membership: null | {         // 회원권 (선택)
 *     plan_name: string,         // 예: "PT 10회권"
 *     total_count: number,       // 총 횟수
 *     used_count: number,        // 사용한 횟수
 *     price: number,             // 가격 (원) — 선택
 *     started_at: "YYYY-MM-DD",  // 시작일 — 선택
 *     expires_at: "YYYY-MM-DD"   // 만료일 — 선택
 *   },
 *   created_at: ISO string
 * }
 */

function defaultMembership(m) {
  if (!m || typeof m !== "object") return null;
  // 모든 핵심 필드가 비어있으면 null 취급
  const totalCount = Number(m.total_count) || 0;
  if (!m.plan_name && totalCount === 0) return null;
  return {
    plan_name: m.plan_name || "",
    total_count: totalCount,
    used_count: Number(m.used_count) || 0,
    price: m.price ? Number(m.price) : null,
    started_at: m.started_at || "",
    expires_at: m.expires_at || "",
  };
}

function newId() {
  return `m_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`;
}

export function loadMembers() {
  try {
    const raw = JSON.parse(localStorage.getItem(MEMBERS_KEY) || "[]");
    if (!Array.isArray(raw)) return [];
    // 누락 필드 보정 (구버전 호환)
    return raw.map((m) => ({
      id: m.id || newId(),
      name: m.name || "",
      phone: m.phone || "",
      memo: m.memo || "",
      weekly_count: Number(m.weekly_count) || 1,
      active: m.active !== false,
      availability: Array.isArray(m.availability) ? m.availability : [],
      membership: defaultMembership(m.membership),
      created_at: m.created_at || new Date().toISOString(),
    }));
  } catch {
    return [];
  }
}

export function saveMembers(members) {
  localStorage.setItem(MEMBERS_KEY, JSON.stringify(members));
}

export function createMember(input = {}) {
  return {
    id: newId(),
    name: input.name || "",
    phone: input.phone || "",
    memo: input.memo || "",
    weekly_count: Number(input.weekly_count) || 1,
    active: input.active !== false,
    availability: Array.isArray(input.availability) ? input.availability : [],
    membership: defaultMembership(input.membership),
    created_at: new Date().toISOString(),
  };
}

export function addMember(input) {
  const members = loadMembers();
  const member = createMember(input);
  members.push(member);
  saveMembers(members);
  return member;
}

export function updateMember(id, updates) {
  const members = loadMembers();
  const idx = members.findIndex((m) => m.id === id);
  if (idx < 0) return null;
  const merged = { ...members[idx], ...updates };
  // membership 이 업데이트로 넘어왔으면 정규화
  if ("membership" in updates) merged.membership = defaultMembership(updates.membership);
  members[idx] = merged;
  saveMembers(members);
  return members[idx];
}

export function deleteMember(id) {
  const members = loadMembers().filter((m) => m.id !== id);
  saveMembers(members);
  return members;
}

export function toggleMemberActive(id) {
  const members = loadMembers();
  const member = members.find((m) => m.id === id);
  if (!member) return null;
  member.active = !member.active;
  saveMembers(members);
  return member;
}

export function replaceAllMembers(members) {
  saveMembers(members);
}

/**
 * 시간표 생성용 payload 변환:
 * 활성 회원의 availability 를 평탄화하여 generate API 가 받는 students 형태로 변환.
 */
export function membersToStudentsPayload(members) {
  const rows = [];
  for (const m of members) {
    if (!m.active) continue;
    if (!m.name?.trim()) continue;
    for (const a of m.availability || []) {
      rows.push({
        name: m.name,
        weekly_count: m.weekly_count,
        days: a.days || [],
        start_time: a.start_time,
        end_time: a.end_time,
      });
    }
  }
  return rows;
}
