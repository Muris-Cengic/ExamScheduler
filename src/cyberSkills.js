import * as XLSX from "xlsx/xlsx.mjs";
import { parseExamTimeRange } from "./imports.js";

const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const text = (value) => String(value ?? "").trim();
const normalise = (value) => text(value).toLowerCase().replace(/[^a-z0-9]/g, "");
const DAY_MS = 86400000;
const isoDate = (year, month, day) => {
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) throw new Error("Invalid CyberSkills date.");
  return date.toISOString().slice(0, 10);
};
function readDate(value, date1904) {
  if (typeof value === "number") {
    const date = XLSX.SSF.parse_date_code(value, { date1904 });
    if (date) return isoDate(date.y, date.m, date.d);
  }
  const iso = text(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return isoDate(...iso.slice(1).map(Number));
  // The supplied weekly workbook uses US-formatted dates; ISO is recommended for CSV.
  const us = text(value).match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if (us) return isoDate(Number(us[3]) + (us[3].length === 2 ? 2000 : 0), Number(us[1]), Number(us[2]));
  throw new Error("Use an Excel date or YYYY-MM-DD for CyberSkills dates.");
}
export function cyberSkillsWeekStart(startDate) {
  const date = new Date(readDate(startDate) + "T00:00:00Z");
  date.setUTCDate(date.getUTCDate() - (date.getUTCDay() + 6) % 7);
  return date;
}

export function parseCyberSkillsWorkbook(workbook, startDate, maxWeeks = 10) {
  const first = cyberSkillsWeekStart(startDate).getTime();
  const sessions = new Map();
  const add = (instructor, date, range, label, sheet) => {
    const timestamp = Date.parse(date + "T00:00:00Z");
    if (timestamp < first || timestamp >= first + maxWeeks * 7 * DAY_MS) return;
    if (!DAYS[new Date(timestamp).getUTCDay() - 1]) throw new Error(`${sheet}: CyberSkills dates must be Monday-Friday.`);
    const name = text(instructor).replace(/^Trainer\s*:\s*/i, "").trim();
    if (!name) throw new Error(`${sheet}: CyberSkills instructor is missing.`);
    const entry = { instructor: name, date, start: range.startMinutes, end: range.endMinutes, label: text(label) || "CyberSkills", sheet };
    sessions.set(JSON.stringify([name.toLowerCase(), date, entry.start, entry.end]), entry);
  };
  workbook.SheetNames.forEach((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true, range: 0 });
    let flat = null;
    let grid = null;
    let range = null;
    let modules = {};
    rows.forEach((row, index) => {
      const headers = row.map(normalise);
      const dateCol = headers.indexOf("date"), timeCol = headers.indexOf("time");
      const staffCol = headers.findIndex((name) => ["instructor", "trainer", "staffname"].includes(name));
      if (dateCol >= 0 && timeCol >= 0 && staffCol >= 0) { flat = { dateCol, timeCol, staffCol, labelCol: headers.indexOf("session") }; grid = null; return; }
      const dayColumns = DAYS.map((day) => headers.indexOf(day.toLowerCase()));
      if (timeCol >= 0 && dayColumns.filter((column) => column >= 0).length >= 2) {
        const dates = dayColumns.map((column, dayIndex) => {
          if (column < 0) return null;
          const date = readDate(rows[index + 1]?.[column], workbook.Workbook?.WBProps?.date1904);
          if (new Date(date + "T00:00:00Z").getUTCDay() !== dayIndex + 1) throw new Error(`${sheetName}: ${date} does not match ${DAYS[dayIndex]}.`);
          return date;
        });
        grid = dates.some((date) => date && Date.parse(date + "T00:00:00Z") >= first && Date.parse(date + "T00:00:00Z") < first + maxWeeks * 7 * DAY_MS)
          ? { timeCol, dayColumns, dates } : null;
        flat = null; range = null; modules = {}; return;
      }
      if (flat) {
        if (!text(row[flat.dateCol]) && !text(row[flat.staffCol])) return;
        add(row[flat.staffCol], readDate(row[flat.dateCol], workbook.Workbook?.WBProps?.date1904),
          parseExamTimeRange(row[flat.timeCol]), row[flat.labelCol], sheetName);
        return;
      }
      if (!grid) return;
      if (headers.includes("staffname") || /^break$/i.test(text(row[grid.timeCol]))) { range = null; return; }
      const time = text(row[grid.timeCol]);
      if (time) { range = parseExamTimeRange(time); modules = {}; }
      if (!range) return;
      grid.dayColumns.forEach((column, dayIndex) => {
        if (column < 0) return;
        const value = text(row[column]);
        if (/^Module\s*:/i.test(value)) modules[column] = value.replace(/^Module\s*:\s*/i, "");
        if (/^(?:Trainer\s*:|(?:Dr|Mr|Ms|Mrs|Prof)\.?\s+)/i.test(value)) add(value, grid.dates[dayIndex], range, modules[column], sheetName);
      });
    });
  });
  if (!sessions.size) throw new Error("No dated CyberSkills duties found within ten weeks of the exam start date. Check the file and Exam Setup date.");
  return [...sessions.values()].sort((a, b) => a.date.localeCompare(b.date) || a.start - b.start || a.instructor.localeCompare(b.instructor));
}

const nameTokens = (name) => text(name).replace(/^(?:(?:dr|mr|ms|mrs|prof|professor)(?:\.\s*|\s+))+/i, "")
  .normalize("NFKD").replace(/[\u0300-\u036f]/g, "").toLowerCase().match(/[a-z0-9]+/g) || [];
export function resolveCyberSkillsLinks(schedule, people) {
  return Object.fromEntries([...new Set(schedule?.sessions.map((entry) => entry.instructor) || [])].map((name) => {
    const saved = schedule.links[name];
    if (Object.hasOwn(schedule.links, name)) return [name, saved === null || people.some((person) => person.id === saved) ? saved : ""];
    const tokens = nameTokens(name);
    const candidates = people.filter((person) => {
      const other = nameTokens(person.name);
      return tokens.length && tokens.every((token) => other.includes(token));
    });
    return [name, candidates.length === 1 ? candidates[0].id : ""];
  }));
}

export function readCyberSkillsSchedule(value) {
  if (value == null) return null;
  if (typeof value !== "object" || typeof value.filename !== "string" || !Array.isArray(value.sessions) || !value.sessions.length ||
    !value.links || typeof value.links !== "object" || Array.isArray(value.links) || Object.values(value.links).some((id) => id !== null && typeof id !== "string")) throw new Error("Invalid CyberSkills schedule in snapshot.");
  value.sessions.forEach((entry) => {
    if (!entry || typeof entry.instructor !== "string" || !entry.instructor.trim() || typeof entry.label !== "string" || typeof entry.sheet !== "string" ||
      typeof entry.date !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(entry.date) || readDate(entry.date) !== entry.date ||
      !DAYS[new Date(entry.date + "T00:00:00Z").getUTCDay() - 1] || !Number.isInteger(entry.start) || !Number.isInteger(entry.end) ||
      entry.start < 0 || entry.end > 1440 || entry.end <= entry.start) throw new Error("Invalid CyberSkills duty in snapshot.");
  });
  return value;
}

export function withCyberSkillsAvailability(catalog, schedule, startDate) {
  if (!schedule) return catalog;
  const links = resolveCyberSkillsLinks(schedule, catalog.invigilators);
  const first = cyberSkillsWeekStart(startDate).getTime();
  return { ...catalog,
    availabilityIssues: Object.entries(links).filter(([, id]) => id === "").map(([name]) => `Match CyberSkills trainer ${name} to an invigilator, or explicitly mark them as not in the pool.`),
    invigilators: catalog.invigilators.map((person) => ({ ...person, externalDuties: schedule.sessions
      .filter((entry) => links[entry.instructor] === person.id).map((entry) => ({ ...entry,
        week: Math.floor((Date.parse(entry.date + "T00:00:00Z") - first) / (7 * DAY_MS)) + 1,
        day: DAYS[new Date(entry.date + "T00:00:00Z").getUTCDay() - 1],
      })),
    })),
  };
}
