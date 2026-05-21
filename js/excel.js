import { minutesToText } from "./parser.js";

function ensureSheetJS() {
  if (typeof XLSX === "undefined") {
    throw new Error("SheetJS(XLSX) 라이브러리가 로드되지 않았습니다. vendor/xlsx.full.min.js 를 확인하세요.");
  }
}

export function assignmentsToExcel(assignments) {
  ensureSheetJS();
  const rows = [["요일", "시작", "종료", "이름"]];
  for (const a of assignments) {
    const start = typeof a.start === "number" ? minutesToText(a.start) : a.start;
    const end = typeof a.end === "number" ? minutesToText(a.end) : a.end;
    rows.push([a.day, start, end, a.name]);
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{ wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 16 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "시간표");

  const binary = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return new Blob([binary], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export function downloadExcel(assignments, filename = "schedule.xlsx") {
  const blob = assignmentsToExcel(assignments);
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
