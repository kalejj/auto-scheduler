import { DAYS, DAY_INDEX } from "./constants.js";

export class ParseError extends Error {
  constructor(message) {
    super(message);
    this.name = "ParseError";
  }
}

const TIME_RE_SOURCE = "(오전|오후|저녁|밤)?\\s*(\\d{1,2})(?::(\\d{1,2}))?\\s*시?";
const TIME_RE_GLOBAL = new RegExp(TIME_RE_SOURCE, "g");
const TIME_RE_SINGLE = new RegExp(TIME_RE_SOURCE);
const DAY_GROUP_RE_GLOBAL = /([월화수목금토일](?:\s*,\s*[월화수목금토일]|\s*[월화수목금토일])*)/g;

export function minutesToText(minutes) {
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

export function ceilToStep(minutes, step) {
  return Math.ceil(minutes / step) * step;
}

export function floorToStep(minutes, step) {
  return Math.floor(minutes / step) * step;
}

export function parseTimeToken(token) {
  const text = String(token).trim();
  const match = text.match(TIME_RE_SINGLE);
  if (!match) throw new ParseError(`시간을 읽을 수 없습니다: ${token}`);

  const [whole, ampm, hourText, minuteText] = match;
  let hour = parseInt(hourText, 10);
  const minute = parseInt(minuteText || "0", 10);
  const originalHasColon = whole.includes(":");

  if (!(hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59)) {
    throw new ParseError(`시간 범위가 이상합니다: ${token}`);
  }

  if (ampm === "오후" || ampm === "저녁" || ampm === "밤") {
    if (hour < 12) hour += 12;
  } else if (ampm === "오전") {
    if (hour === 12) hour = 0;
  } else if (!originalHasColon && hour >= 1 && hour <= 7) {
    hour += 12;
  }

  return hour * 60 + minute;
}

export function extractTimeTokens(text) {
  const tokens = [];
  for (const match of text.matchAll(TIME_RE_GLOBAL)) {
    tokens.push([match[0], parseTimeToken(match[0])]);
  }
  return tokens;
}

export function extractDays(dayGroup) {
  const days = [];
  for (const ch of dayGroup) {
    if (ch in DAY_INDEX && !days.includes(ch)) days.push(ch);
  }
  return days;
}

export function normalizeText(text) {
  let out = text.replaceAll("~", "-").replaceAll("–", "-").replaceAll("—", "-");
  out = out.replaceAll("부터", "이후");
  out = out.replace(/\s+/g, " ");
  return out.trim();
}

export function parseTimeRange(segmentText, lastStartMin) {
  let text = normalizeText(segmentText);
  text = text.replaceAll("시작", " ").replaceAll("가능", " ").trim();
  text = text.replace(/\s+/g, " ");

  const tokens = extractTimeTokens(text);
  if (tokens.length === 0) {
    throw new ParseError(`요일 뒤에 시간 정보가 없습니다: ${text}`);
  }

  let start;
  let end;
  if (text.includes("이후")) {
    start = tokens[0][1];
    end = lastStartMin;
  } else if (text.includes("-")) {
    if (tokens.length === 1) {
      start = tokens[0][1];
      end = lastStartMin;
    } else {
      start = tokens[0][1];
      end = tokens[1][1];
      if (end < start && end <= 12 * 60) end += 12 * 60;
    }
  } else {
    start = tokens[0][1];
    end = tokens[0][1];
  }

  if (start > lastStartMin) {
    throw new ParseError(`마지막 시작 시간(${minutesToText(lastStartMin)}) 이후입니다: ${text}`);
  }

  end = Math.min(end, lastStartMin);
  if (end < start) {
    throw new ParseError(`시간 범위가 이상합니다: ${text}`);
  }

  return [start, end];
}

export function splitNameAvailabilityCount(line) {
  const rawLine = line.trim();
  if (!rawLine) throw new ParseError("빈 줄입니다.");

  if (rawLine.includes("|")) {
    const parts = rawLine.split("|").map((p) => p.trim()).filter(Boolean);
    if (parts.length < 3) {
      throw new ParseError("| 형식은 '이름 | 가능시간 | 주횟수' 형태로 입력해야 합니다.");
    }
    const name = parts[0];
    const weeklyText = parts[parts.length - 1];
    const availabilityText = parts.slice(1, -1).join(" ");
    const countMatch = weeklyText.match(/(\d+)/);
    if (!countMatch) throw new ParseError(`주 횟수를 읽을 수 없습니다: ${weeklyText}`);
    return [name, availabilityText, parseInt(countMatch[1], 10)];
  }

  const countMatch = rawLine.match(/주\s*(\d+)\s*회/);
  if (!countMatch) throw new ParseError("주 횟수는 '주 2회'처럼 입력해야 합니다.");

  const beforeCount = rawLine.slice(0, countMatch.index).trim();
  const spaceIdx = beforeCount.search(/\s/);
  if (spaceIdx === -1) {
    throw new ParseError("이름과 가능시간을 모두 입력해야 합니다.");
  }
  const name = beforeCount.slice(0, spaceIdx).trim();
  const availabilityText = beforeCount.slice(spaceIdx).trim();
  return [name, availabilityText, parseInt(countMatch[1], 10)];
}

export function parseStudentLine(line, stepMinutes, firstStartMin, lastStartMin) {
  const [name, availabilityText, weeklyCount] = splitNameAvailabilityCount(line);
  if (weeklyCount <= 0) throw new ParseError("주 횟수는 1 이상이어야 합니다.");

  const matches = [...availabilityText.matchAll(DAY_GROUP_RE_GLOBAL)];
  if (matches.length === 0) throw new ParseError("가능 요일을 찾지 못했습니다.");

  const availability = {};
  for (const day of DAYS) availability[day] = new Set();

  for (let i = 0; i < matches.length; i++) {
    const match = matches[i];
    const dayGroup = match[1];
    const days = extractDays(dayGroup);
    const nextStart = i + 1 < matches.length ? matches[i + 1].index : availabilityText.length;
    const segmentEnd = match.index + match[0].length;
    const segmentText = availabilityText.slice(segmentEnd, nextStart).trim();

    if (days.length === 0) continue;
    const [rawStart, rawEnd] = parseTimeRange(segmentText, lastStartMin);
    let startMin = ceilToStep(rawStart, stepMinutes);
    startMin = Math.max(startMin, ceilToStep(firstStartMin, stepMinutes));
    const endMin = floorToStep(rawEnd, stepMinutes);

    if (endMin < startMin) {
      throw new ParseError(`${name}: 시간 단위 보정 후 가능한 시간이 없습니다: ${segmentText}`);
    }

    for (const day of days) {
      for (let start = startMin; start <= endMin; start += stepMinutes) {
        if (start <= lastStartMin) availability[day].add(start);
      }
    }
  }

  let totalSlots = 0;
  let distinctDays = 0;
  for (const day of DAYS) {
    const size = availability[day].size;
    totalSlots += size;
    if (size > 0) distinctDays += 1;
  }
  if (totalSlots === 0) throw new ParseError("가능한 시간 슬롯이 없습니다.");
  if (distinctDays < weeklyCount) {
    throw new ParseError(
      `하루 1회 제한 때문에 주 ${weeklyCount}회를 채울 수 없습니다. 가능한 요일 수: ${distinctDays}`
    );
  }

  return { name, weekly_count: weeklyCount, availability, raw: line };
}

export function parseStudents(rawText, stepMinutes, firstStartMin, lastStartMin) {
  const students = [];
  const errors = [];

  const lines = rawText.split(/\r?\n/);
  lines.forEach((line, idx) => {
    if (!line.trim()) return;
    try {
      students.push(parseStudentLine(line, stepMinutes, firstStartMin, lastStartMin));
    } catch (exc) {
      if (exc instanceof ParseError) errors.push(`${idx + 1}번째 줄: ${exc.message}`);
      else throw exc;
    }
  });

  const nameCounts = new Map();
  for (const s of students) nameCounts.set(s.name, (nameCounts.get(s.name) || 0) + 1);
  const dups = [...nameCounts.entries()].filter(([, c]) => c > 1).map(([n]) => n).sort();
  if (dups.length) errors.push(`이름이 중복되었습니다: ${dups.join(", ")}`);

  return [students, errors];
}

export function parseStructuredStudents(rows, stepMinutes, firstStartMin, lastStartMin) {
  const grouped = new Map();
  const errors = [];

  rows.forEach((row, idx) => {
    const rowNumber = idx + 1;
    const name = (row.name || "").trim();
    if (!name) {
      errors.push(`${rowNumber}번째 행: 이름을 입력하세요.`);
      return;
    }

    const days = (row.days || []).filter((d) => d in DAY_INDEX);
    if (!days.length) {
      errors.push(`${rowNumber}번째 행: 요일을 하나 이상 선택하세요.`);
      return;
    }

    let startMin;
    let endMin;
    try {
      startMin = ceilToStep(parseTimeToken(row.start_time), stepMinutes);
      startMin = Math.max(startMin, ceilToStep(firstStartMin, stepMinutes));
      endMin = floorToStep(parseTimeToken(row.end_time), stepMinutes);
    } catch (exc) {
      if (exc instanceof ParseError) {
        errors.push(`${rowNumber}번째 행: ${exc.message}`);
        return;
      }
      throw exc;
    }

    endMin = Math.min(endMin, lastStartMin);
    if (startMin > lastStartMin) {
      errors.push(`${rowNumber}번째 행: 시작 시간이 마지막 시작 시간 이후입니다.`);
      return;
    }
    if (endMin < startMin) {
      errors.push(`${rowNumber}번째 행: 종료 시간은 시작 시간보다 같거나 늦어야 합니다.`);
      return;
    }

    if (!grouped.has(name)) {
      const availability = {};
      for (const d of DAYS) availability[d] = new Set();
      grouped.set(name, {
        name,
        weekly_count: row.weekly_count,
        availability,
        raw: name,
      });
    } else if (grouped.get(name).weekly_count !== row.weekly_count) {
      errors.push(`${rowNumber}번째 행: ${name}의 주 횟수가 다른 행과 다릅니다.`);
      return;
    }

    const student = grouped.get(name);
    for (const day of days) {
      for (let start = startMin; start <= endMin; start += stepMinutes) {
        student.availability[day].add(start);
      }
    }
  });

  const students = [...grouped.values()];
  for (const student of students) {
    let total = 0;
    let distinct = 0;
    for (const d of DAYS) {
      const size = student.availability[d].size;
      total += size;
      if (size > 0) distinct += 1;
    }
    if (total === 0) errors.push(`${student.name}: 가능한 시간 슬롯이 없습니다.`);
    if (distinct < student.weekly_count) {
      errors.push(
        `${student.name}: 하루 1회 제한 때문에 주 ${student.weekly_count}회를 채울 수 없습니다. ` +
          `가능한 요일 수: ${distinct}`
      );
    }
  }

  return [students, errors];
}

function intervalOverlaps(start, end, existing, breakMinutes) {
  for (const [otherStart, otherEnd] of existing) {
    if (start < otherEnd + breakMinutes && otherStart < end + breakMinutes) return true;
  }
  return false;
}

export function parseFixedAssignments(rows, students, breakMinutes) {
  const studentByName = new Map(students.map((s) => [s.name, s]));
  const fixed = [];
  const errors = [];
  const counts = new Map(students.map((s) => [s.name, 0]));
  const studentDays = new Map(students.map((s) => [s.name, new Set()]));
  const dayIntervals = {};
  for (const d of DAYS) dayIntervals[d] = [];

  rows.forEach((row, idx) => {
    const rowNumber = idx + 1;
    const name = (row.name || "").trim();
    if (!name) return;
    if (!(row.day in DAY_INDEX)) {
      errors.push(`고정 수업 ${rowNumber}번째: 요일을 확인하세요.`);
      return;
    }

    let start;
    let end;
    try {
      start = parseTimeToken(row.start);
      end = parseTimeToken(row.end);
    } catch (exc) {
      if (exc instanceof ParseError) {
        errors.push(`고정 수업 ${rowNumber}번째: ${exc.message}`);
        return;
      }
      throw exc;
    }

    if (end <= start) {
      errors.push(`고정 수업 ${rowNumber}번째: 종료 시간은 시작 시간보다 늦어야 합니다.`);
      return;
    }
    if (intervalOverlaps(start, end, dayIntervals[row.day], breakMinutes)) {
      errors.push(`고정 수업 ${rowNumber}번째: 다른 고정 수업과 겹치거나 쉬는시간을 만족하지 못합니다.`);
      return;
    }

    if (studentByName.has(name)) {
      if (studentDays.get(name).has(row.day)) {
        errors.push(`고정 수업 ${rowNumber}번째: ${name}은 같은 요일에 두 번 배정할 수 없습니다.`);
        return;
      }
      if (counts.get(name) >= studentByName.get(name).weekly_count) {
        errors.push(`고정 수업 ${rowNumber}번째: ${name}의 주 횟수를 초과했습니다.`);
        return;
      }
      counts.set(name, counts.get(name) + 1);
      studentDays.get(name).add(row.day);
    }

    dayIntervals[row.day].push([start, end, name]);
    fixed.push({ name, day: row.day, start, end });
  });

  return [fixed, errors];
}
