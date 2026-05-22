// 시간표 이미지 생성 + 공유 (Web Share API)
// Canvas API 만 사용 — 외부 라이브러리 없음.

import { DAYS } from "./constants.js";
import { weekDays, formatDateShort, formatWeekRange, todayISO } from "./dates.js";

function timeToMinutes(value) {
  if (!value) return 0;
  const [h, m] = value.split(":").map(Number);
  return h * 60 + m;
}

function hourLabel(hour) {
  const h = ((hour % 24) + 24) % 24;
  if (h === 0) return "오전 12시";
  if (h === 12) return "오후 12시";
  if (h < 12) return `오전 ${h}시`;
  return `오후 ${h - 12}시`;
}

function colorForName(name) {
  const palette = [
    ["#dbeafe", "#1e3a8a"], ["#dcfce7", "#14532d"], ["#fef3c7", "#78350f"],
    ["#fce7f3", "#831843"], ["#ede9fe", "#4c1d95"], ["#cffafe", "#164e63"],
    ["#ffedd5", "#7c2d12"], ["#e0e7ff", "#312e81"],
  ];
  let hash = 0;
  for (const ch of name) hash = (hash * 31 + ch.charCodeAt(0)) >>> 0;
  return palette[hash % palette.length];
}

function drawRoundRect(ctx, x, y, w, h, r) {
  if (ctx.roundRect) {
    ctx.beginPath();
    ctx.roundRect(x, y, w, h, r);
    return;
  }
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h - r);
  ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
  ctx.lineTo(x + r, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - r);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
  ctx.closePath();
}

/**
 * 시간표 PNG Blob 생성
 * @param {Object} args
 *   sessions: 현재 주 sessions
 *   weekStart: ISO 월요일
 *   firstStartTime, lastStartTime: 표시 범위
 *   durationMinutes: 기본 수업 길이
 */
export async function renderScheduleImage({ sessions, weekStart, firstStartTime, lastStartTime, durationMinutes }) {
  const firstSetting = timeToMinutes(firstStartTime || "10:00");
  const lastSetting = timeToMinutes(lastStartTime || "23:00");
  const duration = durationMinutes || 50;

  const starts = sessions.length ? sessions.map((s) => timeToMinutes(s.start)) : [firstSetting];
  const ends = sessions.length ? sessions.map((s) => timeToMinutes(s.end)) : [lastSetting + duration];
  const firstHour = Math.floor(Math.min(firstSetting, ...starts) / 60);
  const lastHour = Math.ceil(Math.max(lastSetting + duration, ...ends) / 60);
  const hourCount = lastHour - firstHour;

  // 레이아웃 (CSS px)
  const PAD = 24;
  const W = 1200;
  const TITLE_H = 80;
  const HEAD_H = 56;
  const TIME_COL_W = 96;
  const HOUR_H = 76;
  const DAY_COL_W = (W - PAD * 2 - TIME_COL_W) / 7;
  const GRID_H = HEAD_H + hourCount * HOUR_H;
  const FOOTER_H = 36;
  const H = PAD * 2 + TITLE_H + GRID_H + FOOTER_H;

  // 레티나용 2배 해상도
  const dpr = 2;
  const canvas = document.createElement("canvas");
  canvas.width = W * dpr;
  canvas.height = H * dpr;
  const ctx = canvas.getContext("2d");
  ctx.scale(dpr, dpr);
  ctx.textBaseline = "alphabetic";

  // 배경
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, W, H);

  // 제목
  ctx.fillStyle = "#17202a";
  ctx.font = "bold 26px ui-sans-serif, system-ui, -apple-system, 'Apple SD Gothic Neo', sans-serif";
  ctx.textAlign = "left";
  ctx.fillText(`${formatWeekRange(weekStart)}`, PAD, PAD + 28);

  ctx.fillStyle = "#65717f";
  ctx.font = "14px ui-sans-serif, system-ui, sans-serif";
  ctx.fillText("자동 시간표", PAD, PAD + 52);

  const gridTop = PAD + TITLE_H;
  const dateIsos = weekDays(weekStart);
  const today = todayISO();

  // 헤더 배경
  ctx.fillStyle = "#eef2f5";
  ctx.fillRect(PAD, gridTop, W - PAD * 2, HEAD_H);

  // 오늘 컬럼 강조 (헤더)
  for (let i = 0; i < 7; i++) {
    if (dateIsos[i] === today) {
      const x = PAD + TIME_COL_W + DAY_COL_W * i;
      ctx.fillStyle = "#e8f4f3";
      ctx.fillRect(x, gridTop, DAY_COL_W, HEAD_H);
    }
  }

  // 시간 헤더
  ctx.fillStyle = "#65717f";
  ctx.font = "bold 13px ui-sans-serif, system-ui, sans-serif";
  ctx.textAlign = "center";
  ctx.fillText("시간", PAD + TIME_COL_W / 2, gridTop + 32);

  // 요일 헤더
  for (let i = 0; i < 7; i++) {
    const x = PAD + TIME_COL_W + DAY_COL_W * i;
    const day = DAYS[i];
    const iso = dateIsos[i];
    const isToday = iso === today;
    ctx.fillStyle = isToday ? "#0b5553" : "#17202a";
    ctx.font = "bold 14px ui-sans-serif, system-ui, sans-serif";
    ctx.fillText(day, x + DAY_COL_W / 2, gridTop + 22);
    ctx.fillStyle = isToday ? "#0b5553" : "#65717f";
    ctx.font = "12px ui-sans-serif, system-ui, sans-serif";
    ctx.fillText(formatDateShort(iso), x + DAY_COL_W / 2, gridTop + 42);
  }

  // 본문 배경
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(PAD, gridTop + HEAD_H, W - PAD * 2, hourCount * HOUR_H);

  // 오늘 컬럼 강조 (본문, 옅게)
  for (let i = 0; i < 7; i++) {
    if (dateIsos[i] === today) {
      const x = PAD + TIME_COL_W + DAY_COL_W * i;
      ctx.fillStyle = "rgba(18, 109, 106, 0.04)";
      ctx.fillRect(x, gridTop + HEAD_H, DAY_COL_W, hourCount * HOUR_H);
    }
  }

  // 그리드 선
  ctx.strokeStyle = "#dbe1e8";
  ctx.lineWidth = 1;

  // 수평 (시간) 선 + 시간 라벨
  for (let i = 0; i <= hourCount; i++) {
    const y = gridTop + HEAD_H + i * HOUR_H;
    ctx.beginPath();
    ctx.moveTo(PAD, y);
    ctx.lineTo(W - PAD, y);
    ctx.stroke();
    if (i < hourCount) {
      ctx.fillStyle = "#65717f";
      ctx.font = "bold 12px ui-sans-serif, system-ui, sans-serif";
      ctx.textAlign = "right";
      ctx.fillText(hourLabel(firstHour + i), PAD + TIME_COL_W - 10, y + 20);
    }
  }

  // 수직 (요일) 선
  for (let i = 0; i <= 7; i++) {
    const x = PAD + TIME_COL_W + DAY_COL_W * i;
    ctx.beginPath();
    ctx.moveTo(x, gridTop);
    ctx.lineTo(x, gridTop + GRID_H);
    ctx.stroke();
  }
  // 시간/날짜 구분 선
  ctx.beginPath();
  ctx.moveTo(PAD + TIME_COL_W, gridTop);
  ctx.lineTo(PAD + TIME_COL_W, gridTop + GRID_H);
  ctx.stroke();
  // 헤더 하단 선
  ctx.beginPath();
  ctx.moveTo(PAD, gridTop + HEAD_H);
  ctx.lineTo(W - PAD, gridTop + HEAD_H);
  ctx.stroke();
  // 외곽
  ctx.strokeRect(PAD, gridTop, W - PAD * 2, GRID_H);

  // 이벤트
  ctx.textAlign = "center";
  for (const s of sessions) {
    const dayIdx = DAYS.indexOf(s.day);
    if (dayIdx < 0) continue;
    const startMin = timeToMinutes(s.start);
    const endMin = timeToMinutes(s.end);
    const top = gridTop + HEAD_H + ((startMin - firstHour * 60) / 60) * HOUR_H;
    const height = ((endMin - startMin) / 60) * HOUR_H;
    const x = PAD + TIME_COL_W + DAY_COL_W * dayIdx + 4;
    const eventW = DAY_COL_W - 8;

    const [bg, ink] = colorForName(s.name);

    // 상태에 따른 투명도
    let alpha = 1;
    if (s.status === "completed") alpha = 0.78;
    else if (s.status === "no-show") alpha = 0.6;
    else if (s.status === "cancelled") alpha = 0.4;
    ctx.globalAlpha = alpha;

    // 박스
    ctx.fillStyle = bg;
    drawRoundRect(ctx, x, top + 2, eventW, height - 4, 6);
    ctx.fill();

    // 잠금 아이콘 (좌상단)
    if (s.source === "fixed") {
      ctx.font = "11px ui-sans-serif, system-ui, sans-serif";
      ctx.fillStyle = ink;
      ctx.textAlign = "left";
      ctx.fillText("🔒", x + 4, top + 16);
    }

    // 이름
    ctx.fillStyle = ink;
    ctx.font = `bold 14px ui-sans-serif, system-ui, sans-serif`;
    ctx.textAlign = "center";
    ctx.fillText(s.name, x + eventW / 2, top + 26);

    // 시간
    ctx.font = "12px ui-sans-serif, system-ui, sans-serif";
    ctx.fillText(`${s.start}-${s.end}`, x + eventW / 2, top + 44);

    // 취소 라인
    if (s.status === "cancelled") {
      ctx.strokeStyle = ink;
      ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(x + 8, top + 22);
      ctx.lineTo(x + eventW - 8, top + 22);
      ctx.stroke();
    }

    // 상태 아이콘 (우상단)
    ctx.globalAlpha = 1;
    ctx.textAlign = "right";
    if (s.status === "completed") {
      ctx.font = "bold 15px ui-sans-serif, system-ui, sans-serif";
      ctx.fillStyle = "#14532d";
      ctx.fillText("✓", x + eventW - 6, top + 18);
    } else if (s.status === "no-show") {
      ctx.font = "bold 15px ui-sans-serif, system-ui, sans-serif";
      ctx.fillStyle = "#7a1818";
      ctx.fillText("✗", x + eventW - 6, top + 18);
    }
  }

  ctx.globalAlpha = 1;

  // 푸터
  ctx.fillStyle = "#65717f";
  ctx.font = "11px ui-sans-serif, system-ui, sans-serif";
  ctx.textAlign = "right";
  ctx.fillText(`생성일: ${new Date().toLocaleString("ko-KR")}`, W - PAD, H - PAD);

  return new Promise((resolve) => {
    canvas.toBlob((blob) => resolve(blob), "image/png", 0.95);
  });
}

export async function shareScheduleImage(args) {
  const blob = await renderScheduleImage(args);
  if (!blob) throw new Error("이미지 생성 실패");

  const filename = `schedule-${args.weekStart}.png`;
  const file = new File([blob], filename, { type: "image/png" });

  // 1) Web Share API (모바일에서 카톡/문자 시트)
  if (typeof navigator !== "undefined" && navigator.canShare && navigator.canShare({ files: [file] })) {
    try {
      await navigator.share({
        title: `${formatWeekRange(args.weekStart)} 시간표`,
        files: [file],
      });
      return { method: "share" };
    } catch (e) {
      // 사용자 취소도 throw — 다운로드로 폴백 안 함
      if (e?.name === "AbortError") return { method: "cancelled" };
      throw e;
    }
  }

  // 2) 다운로드 폴백
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1500);
  return { method: "download" };
}
