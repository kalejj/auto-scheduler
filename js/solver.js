import { DAYS, DAY_INDEX, TIME_LIMIT_SECONDS_DEFAULT } from "./constants.js";

function intervalOverlaps(start, end, existing, breakMinutes) {
  for (const [otherStart, otherEnd] of existing) {
    if (start < otherEnd + breakMinutes && otherStart < end + breakMinutes) return true;
  }
  return false;
}

function legalCandidates(student, studentDays, dayIntervals, durationMinutes, breakMinutes) {
  const candidates = [];
  for (const day of DAYS) {
    if (studentDays.get(student.name).has(day)) continue;
    const slots = student.availability[day];
    if (!slots || slots.size === 0) continue;
    const sorted = [...slots].sort((a, b) => a - b);
    for (const start of sorted) {
      const end = start + durationMinutes;
      if (!intervalOverlaps(start, end, dayIntervals[day], breakMinutes)) {
        candidates.push([day, start]);
      }
    }
  }
  return candidates;
}

function spacingScoreForCandidate(day, existingDays) {
  if (existingDays.size === 0) return 0;
  const idx = DAY_INDEX[day];
  let minDist = Infinity;
  for (const d of existingDays) {
    const dist = Math.abs(idx - DAY_INDEX[d]);
    if (dist < minDist) minDist = dist;
  }
  return -minDist * 8;
}

function finalScheduleCost(assignments, students) {
  let cost = 0;

  const dayCounts = {};
  for (const d of DAYS) dayCounts[d] = 0;
  for (const item of assignments) dayCounts[item.day] += 1;
  const counts = Object.values(dayCounts);
  const mean = counts.length ? counts.reduce((a, b) => a + b, 0) / counts.length : 0;
  cost += counts.reduce((acc, c) => acc + (c - mean) ** 2, 0) * 4;

  for (const student of students) {
    const idxs = assignments
      .filter((a) => a.name === student.name)
      .map((a) => DAY_INDEX[a.day])
      .sort((a, b) => a - b);
    if (idxs.length <= 1) continue;

    const gaps = [];
    for (let i = 0; i < idxs.length - 1; i++) gaps.push(idxs[i + 1] - idxs[i]);
    const minGap = Math.min(...gaps);
    if (minGap === 1) cost += 30;
    else if (minGap === 2) cost += 8;

    if (idxs.length === 2) {
      cost += Math.max(0, 4 - Math.abs(idxs[1] - idxs[0])) * 3;
    }
  }

  for (const item of assignments) {
    const m = item.start % 60;
    if (m !== 0 && m !== 30) cost += 0.4;
  }

  return cost;
}

export function solveSchedule({
  students,
  durationMinutes,
  breakMinutes,
  fixedAssignments = [],
  timeLimitSeconds = TIME_LIMIT_SECONDS_DEFAULT,
  randomize = true,
}) {
  const remaining = new Map(students.map((s) => [s.name, s.weekly_count]));
  const studentDays = new Map(students.map((s) => [s.name, new Set()]));
  const dayIntervals = {};
  for (const d of DAYS) dayIntervals[d] = [];
  const assignments = [];

  for (const fixed of fixedAssignments) {
    if (remaining.has(fixed.name)) {
      remaining.set(fixed.name, remaining.get(fixed.name) - 1);
      studentDays.get(fixed.name).add(fixed.day);
    }
    dayIntervals[fixed.day].push([fixed.start, fixed.end, fixed.name]);
    assignments.push({ name: fixed.name, day: fixed.day, start: fixed.start, end: fixed.end });
  }

  let bestAssignments = [];
  let bestCost = null;
  let nodes = 0;
  const deadline = performance.now() + timeLimitSeconds * 1000;

  function chooseNextStudent() {
    let bestTuple = null;
    let bestStudent = null;
    let bestCandidates = [];

    for (const student of students) {
      if (remaining.get(student.name) <= 0) continue;

      const candidates = legalCandidates(student, studentDays, dayIntervals, durationMinutes, breakMinutes);
      const candidateDays = new Set(candidates.map(([d]) => d));

      if (candidateDays.size < remaining.get(student.name)) {
        return { student, candidates: [], failReason: `${student.name}: 남은 회차를 채울 수 없습니다.` };
      }

      const key = [candidates.length, -remaining.get(student.name)];
      if (
        bestTuple === null ||
        key[0] < bestTuple[0] ||
        (key[0] === bestTuple[0] && key[1] < bestTuple[1])
      ) {
        bestTuple = key;
        bestCandidates = candidates;
        bestStudent = student;
      }
    }

    if (bestTuple === null) return { student: null, candidates: [], failReason: null };
    return { student: bestStudent, candidates: bestCandidates, failReason: null };
  }

  function candidateOrderKey(student, [day, start]) {
    const dayLoad = dayIntervals[day].length;
    const spacing = spacingScoreForCandidate(day, studentDays.get(student.name));
    const m = start % 60;
    const prettyPenalty = m === 0 || m === 30 ? 0 : 0.2;
    const jitter = randomize ? Math.random() * 2.5 : 0;
    return dayLoad * 3 + spacing + prettyPenalty + jitter;
  }

  function recurse() {
    if (performance.now() > deadline) return;
    nodes += 1;

    let allDone = true;
    for (const v of remaining.values()) if (v !== 0) { allDone = false; break; }
    if (allDone) {
      const cost = finalScheduleCost(assignments, students);
      if (bestCost === null || cost < bestCost) {
        bestCost = cost;
        bestAssignments = assignments.map((a) => ({ ...a }));
      }
      return;
    }

    const { student, candidates, failReason } = chooseNextStudent();
    if (!student || failReason) return;

    const sorted = candidates
      .map((c) => ({ c, key: candidateOrderKey(student, c) }))
      .sort((a, b) => a.key - b.key || a.c[1] - b.c[1])
      .map((x) => x.c);

    for (const [day, start] of sorted) {
      if (performance.now() > deadline) return;
      const end = start + durationMinutes;

      remaining.set(student.name, remaining.get(student.name) - 1);
      studentDays.get(student.name).add(day);
      dayIntervals[day].push([start, end, student.name]);
      assignments.push({ name: student.name, day, start, end });

      recurse();

      assignments.pop();
      dayIntervals[day].pop();
      studentDays.get(student.name).delete(day);
      remaining.set(student.name, remaining.get(student.name) + 1);
    }
  }

  recurse();

  if (bestAssignments.length) {
    bestAssignments.sort(
      (a, b) => DAY_INDEX[a.day] - DAY_INDEX[b.day] || a.start - b.start || a.name.localeCompare(b.name)
    );
    return {
      assignments: bestAssignments,
      success: true,
      message: "시간표를 생성했습니다.",
      nodes,
      cost: bestCost,
    };
  }

  return {
    assignments: [],
    success: false,
    message: "조건을 모두 만족하는 시간표를 찾지 못했습니다. 가능 시간대를 넓히거나 주 횟수를 줄여보세요.",
    nodes,
    cost: null,
  };
}
