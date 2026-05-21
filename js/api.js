import { DAYS, SAMPLE_STUDENTS, SAMPLE_DEFAULTS, TIME_LIMIT_SECONDS_DEFAULT } from "./constants.js";
import {
  ParseError,
  parseTimeToken,
  minutesToText,
  parseStructuredStudents,
  parseFixedAssignments,
} from "./parser.js";
import { solveSchedule } from "./solver.js";

function studentsSummary(students) {
  return students.map((student) => {
    const dayTexts = [];
    let total = 0;
    for (const day of DAYS) {
      const slots = [...student.availability[day]].sort((a, b) => a - b);
      total += slots.length;
      if (slots.length) {
        dayTexts.push(
          `${day} ${minutesToText(slots[0])}~${minutesToText(slots[slots.length - 1])} (${slots.length}개)`
        );
      }
    }
    return {
      name: student.name,
      weekly_count: student.weekly_count,
      availability: dayTexts.join(" / "),
      slot_count: total,
    };
  });
}

function serializeAssignment(a) {
  return {
    day: a.day,
    start: minutesToText(a.start),
    end: minutesToText(a.end),
    name: a.name,
  };
}

export function getSample() {
  return {
    sample_students: SAMPLE_STUDENTS,
    ...SAMPLE_DEFAULTS,
  };
}

export function generateSchedule(payload) {
  let firstStartMin;
  let lastStartMin;
  try {
    firstStartMin = parseTimeToken(payload.first_start_time);
    lastStartMin = parseTimeToken(payload.last_start_time);
  } catch (exc) {
    if (exc instanceof ParseError) {
      return { success: false, message: "첫 시작/마지막 시작 시간은 10:00, 23:00처럼 입력하세요.", errors: [] };
    }
    throw exc;
  }

  if (firstStartMin > lastStartMin) {
    return {
      success: false,
      message: "첫 시작 시간은 마지막 시작 시간보다 같거나 빨라야 합니다.",
      errors: [],
    };
  }

  const [students, errors] = parseStructuredStudents(
    payload.students || [],
    payload.step_minutes,
    firstStartMin,
    lastStartMin
  );
  if (errors.length) return { success: false, message: "입력값을 확인해야 합니다.", errors };
  if (!students.length) return { success: false, message: "입력된 학생이 없습니다.", errors: [] };

  const [fixedAssignments, fixedErrors] = parseFixedAssignments(
    payload.fixed_assignments || [],
    students,
    payload.break_minutes
  );
  if (fixedErrors.length) return { success: false, message: "고정 수업을 확인해야 합니다.", errors: fixedErrors };

  const result = solveSchedule({
    students,
    durationMinutes: payload.duration_minutes,
    breakMinutes: payload.break_minutes,
    fixedAssignments,
    timeLimitSeconds: payload.time_limit_seconds || TIME_LIMIT_SECONDS_DEFAULT,
  });

  return {
    success: result.success,
    message: result.message,
    nodes: result.nodes,
    cost: result.cost,
    total_students: students.length,
    total_lessons: students.reduce((acc, s) => acc + s.weekly_count, 0),
    students: studentsSummary(students),
    assignments: result.assignments.map(serializeAssignment),
    errors: [],
  };
}

export function assignmentsFromDisplay(items) {
  return items.map((item) => ({
    name: item.name,
    day: item.day,
    start: parseTimeToken(item.start),
    end: parseTimeToken(item.end),
  }));
}
