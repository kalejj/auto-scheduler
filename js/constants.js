export const DAYS = ["월", "화", "수", "목", "금", "토", "일"];
export const DAY_INDEX = Object.fromEntries(DAYS.map((d, i) => [d, i]));

export const DURATION_MINUTES_DEFAULT = 50;
export const STEP_MINUTES_DEFAULT = 10;
export const FIRST_START_DEFAULT = "10:00";
export const LAST_START_DEFAULT = "23:00";
export const BREAK_MINUTES_DEFAULT = 0;
export const TIME_LIMIT_SECONDS_DEFAULT = 5.0;

export const SAMPLE_STUDENTS = [
  { name: "이태경", days: ["월", "화", "수", "목", "금"], start_time: "19:40", end_time: "20:00", weekly_count: 2 },
  { name: "김수진", days: ["월", "목"], start_time: "18:30", end_time: "21:00", weekly_count: 1 },
  { name: "전희원", days: ["월", "수"], start_time: "18:30", end_time: "22:10", weekly_count: 2 },
  { name: "전희원", days: ["토"], start_time: "10:00", end_time: "23:00", weekly_count: 2 },
  { name: "송다혜", days: ["화"], start_time: "16:00", end_time: "23:00", weekly_count: 2 },
  { name: "송다혜", days: ["수", "목", "금"], start_time: "19:00", end_time: "22:10", weekly_count: 2 },
  { name: "양선", days: ["월", "화", "수"], start_time: "18:00", end_time: "18:00", weekly_count: 2 },
  { name: "나모나", days: ["수"], start_time: "19:00", end_time: "20:00", weekly_count: 1 },
];

export const SAMPLE_DEFAULTS = {
  duration_minutes: DURATION_MINUTES_DEFAULT,
  step_minutes: STEP_MINUTES_DEFAULT,
  first_start_time: FIRST_START_DEFAULT,
  last_start_time: LAST_START_DEFAULT,
  break_minutes: BREAK_MINUTES_DEFAULT,
};
