import * as XLSX from "xlsx/xlsx.mjs";
import StyledXLSX from "xlsx-js-style";
import { clock, invigilatorWorkloads, isTeachingTimeDuty, validateResourcePlan } from "./resources.js";
import { buildResourceWorkbookForWeek } from "./reports.js";

export const REPORT_VIEWS = [
  { id: "complete", label: "Complete report", audience: "For the exam team", description: "The existing report, with room assignments, invigilator pools and daily student sheets.", filename: "Exam_Schedule" },
  { id: "overview", label: "Exam overview", audience: "For sharing the timetable", description: "A course timetable with compact rooms, required invigilators and lab-time exams. No student names.", filename: "Exam_Overview" },
  { id: "staff", label: "Staff duties", audience: "For invigilators", description: "Named duties and a workload summary. Backups and replaced teaching hours stay separate.", filename: "Staff_Duties" },
  { id: "students", label: "Student room lists", audience: "For room check-in", description: "Every student sitting, with the course, time and confirmed room. Contains personal data.", filename: "Student_Room_Lists" },
  { id: "seating", label: "Course seating", audience: "For course sections", description: "One Excel file per exam, with a student-to-room sheet for each CRN.", filename: "Course_Seating" },
];
export const REPORT_DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

export function reportDate(startDate, week, day) {
  const date = new Date(startDate + "T00:00:00Z");
  date.setUTCDate(date.getUTCDate() + (week - 1) * 7 + REPORT_DAYS.indexOf(day));
  return date.toISOString().slice(0, 10);
}

export function formatReportDate(value, includeYear = false) {
  return new Intl.DateTimeFormat("en-GB", { day: "numeric", month: "short", ...(includeYear ? { year: "numeric" } : {}), timeZone: "UTC" }).format(new Date(value + "T00:00:00Z"));
}

export const reportTimeRange = (start, end) => clock(start) + "-" + clock(end);

export function formatRoomDisplayName(name) {
  return String(name ?? "").split(/\s*\/\s*/).map((part) => part.trim())
    .filter((part) => part && part.toUpperCase() !== "PAD").join("/");
}

export function buildAsdOverviewExams({ assignments, courseLookup, examDurations = {}, defaultDuration = 60 }) {
  const exams = [];
  Object.entries(assignments).forEach(([week, days]) => Object.entries(days).forEach(([day, slots]) => Object.entries(slots).forEach(([slotId, ids]) => {
    [...new Set(ids)].forEach((courseId) => {
      const course = courseLookup[courseId];
      const [hour, minute] = slotId.split(":").map(Number);
      const start = hour * 60 + minute;
      const duration = examDurations[courseId] ?? defaultDuration;
      if (!course || !REPORT_DAYS.includes(day) || !Number.isInteger(Number(week)) || Number(week) < 1 ||
        !/^\d{2}:\d{2}$/.test(slotId) || hour > 23 || minute > 59 || !Number.isFinite(duration) || duration <= 0 || start + duration > 1440) {
        throw new Error("Invalid ASD exam in the overview.");
      }
      // Reference exams deliberately carry no department resource or student totals.
      exams.push({ id: "asd/" + week + "/" + day + "/" + slotId + "/" + courseId, isAsd: true,
        week: Number(week), day, start, end: start + duration, code: course.code, title: course.title,
        crns: [...course.crns].sort(), studentCount: null, roomNames: [], primaryInvigilatorsNeeded: null, duringLab: null });
    });
  })));
  return exams;
}

export function buildExportModel({ sessions, catalog, plan, startDate, weeks, asdExams = [], includeAsd = false }) {
  if (!validateResourcePlan(sessions, catalog, plan).complete) throw new Error("Complete valid resource assignments before exporting.");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(startDate) || !Number.isFinite(Date.parse(startDate + "T00:00:00Z"))) throw new Error("Choose a valid exam start date.");
  const references = includeAsd ? asdExams : [];
  const available = [...new Set([...sessions, ...references].map((session) => session.week))].sort((a, b) => a - b);
  const selectedWeeks = [...new Set(weeks ?? available)].sort((a, b) => a - b);
  if (!selectedWeeks.length) throw new Error("Select at least one exam week.");
  if (selectedWeeks.some((week) => !available.includes(week))) throw new Error("Only weeks with scheduled exams can be exported.");
  const current = sessions.filter((session) => selectedWeeks.includes(session.week));
  const staff = new Map(catalog.invigilators.map((person) => [person.id, person]));
  const rooms = new Map(catalog.rooms.map((room) => [room.id, room]));
  const roomRows = [];
  const exams = [];
  const duties = [];
  const students = [];
  current.forEach((session) => {
    const date = reportDate(startDate, session.week, session.day);
    const timing = { week: session.week, day: session.day, date, start: session.start, end: session.end, sessionId: session.id };
    const courses = new Map();
    session.rooms.forEach((room) => {
      const allocation = plan.allocations[room.id];
      const roomName = rooms.get(allocation.roomId).name;
      const invigilators = allocation.invigilatorIds.map((id) => staff.get(id).name);
      const crns = [...new Set(room.students.map((student) => student.crn).filter(Boolean))].sort();
      roomRows.push({ ...timing, id: room.id, code: room.code, title: room.title, crns, roomName, invigilators,
        studentCount: room.students.length, maxStudents: room.maxStudents, distributionChoice: room.distributionChoice });
      const exam = courses.get(room.courseId) || { ...timing, id: session.id + "/" + room.courseId, code: room.code, title: room.title,
        crns: new Set(), roomNames: [], studentCount: 0, primaryInvigilatorsNeeded: 0,
        duringLab: session.replacedLabs.some((lab) => lab.code === room.code) };
      crns.forEach((crn) => exam.crns.add(crn));
      exam.roomNames.push(formatRoomDisplayName(roomName));
      exam.studentCount += room.students.length;
      exam.primaryInvigilatorsNeeded += room.requiredInvigilators;
      courses.set(room.courseId, exam);
      allocation.invigilatorIds.forEach((id) => {
        const teaching = isTeachingTimeDuty(staff.get(id), session, sessions);
        duties.push({ ...timing, id: room.id + "/" + id, invigilatorId: id, name: staff.get(id).name, role: "Exam", code: room.code,
          roomName, teaching, extraLoad: teaching ? 0 : 1 });
      });
      room.students.forEach((student) => students.push({ ...timing, examId: exam.id, id: student.id, name: student.name, crn: student.crn || "",
        code: room.code, title: room.title, roomName }));
    });
    courses.forEach((exam) => exams.push({ ...exam, crns: [...exam.crns].sort() }));
    (plan.backups[session.id] || []).forEach((id) => {
      const teaching = isTeachingTimeDuty(staff.get(id), session, sessions);
      duties.push({ ...timing, id: session.id + "/backup/" + id, invigilatorId: id, name: staff.get(id).name, role: "Backup",
        code: [...courses.values()].map((exam) => exam.code).join(", "), roomName: "Slot standby", teaching, extraLoad: teaching ? 0 : 1 });
    });
  });
  const overviewExams = [...exams, ...references.filter((exam) => selectedWeeks.includes(exam.week))
    .map((exam) => ({ ...exam, date: reportDate(startDate, exam.week, exam.day) }))]
    .sort((a, b) => a.week - b.week || REPORT_DAYS.indexOf(a.day) - REPORT_DAYS.indexOf(b.day) ||
      a.start - b.start || Number(Boolean(a.isAsd)) - Number(Boolean(b.isAsd)) || a.code.localeCompare(b.code));
  return { selectedWeeks, sessions: current, roomRows, exams, overviewExams, includeAsd, duties, students,
    workloads: invigilatorWorkloads(catalog, current, plan),
    summary: { exams: exams.length, roomSessions: roomRows.length, students: new Set(students.map((student) => student.id)).size, sittings: students.length } };
}

export function buildCourseSeating(model) {
  return model.exams.map((exam) => {
    const groups = new Map();
    model.students.filter((student) => student.examId === exam.id).forEach((student) => {
      const rows = groups.get(student.crn) || [];
      rows.push([String(student.id), formatRoomDisplayName(student.roomName)]);
      groups.set(student.crn, rows);
    });
    const crnSheets = [...groups].sort(([a], [b]) => a.localeCompare(b, "en", { numeric: true }))
      .map(([crn, rows]) => ({ crn, rows: rows.sort((a, b) => a[1].localeCompare(b[1], "en", { numeric: true }) ||
        a[0].localeCompare(b[0], "en", { numeric: true })) }));
    return { ...exam, crnSheets };
  });
}

export function courseSeatingInfo(exam) {
  return [["Course", exam.code], ["Course title", exam.title],
    ["Date", exam.date], ["Day", exam.day], ["Time", reportTimeRange(exam.start, exam.end)]];
}

function uniqueExportName(base, used, maxLength) {
  let name = base.slice(0, maxLength);
  for (let number = 2; used.has(name.toLowerCase()); number += 1) {
    const suffix = "_" + number;
    name = base.slice(0, maxLength - suffix.length) + suffix;
  }
  used.add(name.toLowerCase());
  return name;
}

function styleCourseSeatingSheet(sheet, info, studentRows) {
  const navy = "FF244761";
  const paleBlue = "FFF0F5F9";
  const white = "FFFFFFFF";
  const headerRow = info.length + 2;
  const fill = (rgb) => ({ patternType: "solid", fgColor: { rgb } });
  const rule = (rgb, style = "thin") => ({ style, color: { rgb } });
  const styleCell = (row, column, { font, alignment, ...style } = {}) => {
    const address = XLSX.utils.encode_cell({ r: row, c: column });
    sheet[address] ??= { t: "s", v: "" };
    sheet[address].s = {
      font: { name: "Arial", sz: 11, color: { rgb: "FF263747" }, ...font },
      fill: fill(white), ...style,
      alignment: { vertical: "center", horizontal: "left", wrapText: true, indent: 1, ...alignment },
    };
  };
  const lineCount = (value, width) => String(value).split(/\r\n|\r|\n/)
    .reduce((count, line) => count + Math.max(1, Math.ceil(line.length / width)), 0);
  sheet["!cols"] = [{ wch: 24 }, { wch: 44 }];
  sheet["!rows"] = [{ hpt: 32 }, ...info.map(([, value]) => ({ hpt: Math.max(24, lineCount(value, 38) * 18) })),
    { hpt: 10 }, { hpt: 26 }];
  [0, 1].forEach((column) => styleCell(0, column, {
    font: { sz: 16, bold: true, color: { rgb: navy } },
    border: { bottom: rule(navy) },
  }));
  info.forEach((_, index) => {
    styleCell(index + 1, 0, { fill: fill(paleBlue), font: { bold: true, color: { rgb: navy } } });
    styleCell(index + 1, 1, index === 0 ? { font: { bold: true, color: { rgb: navy } } } : {});
  });
  [0, 1].forEach((column) => styleCell(headerRow, column, {
    fill: fill(navy), font: { bold: true, color: { rgb: white } },
    alignment: { horizontal: "center", indent: 0 },
    ...(column === 0 ? { border: { right: rule(white) } } : {}),
  }));
  let roomGroup = 0;
  studentRows.forEach((values, index) => {
    const nextRoom = index > 0 && values[1] !== studentRows[index - 1][1];
    if (nextRoom) roomGroup += 1;
    const row = headerRow + 1 + index;
    sheet["!rows"][row] = { hpt: Math.max(22, Math.max(lineCount(values[0], 20), lineCount(values[1], 38)) * 18) };
    values.forEach((_, column) => styleCell(row, column, {
      numFmt: "@", fill: fill(roomGroup % 2 ? paleBlue : white),
      border: { ...(nextRoom ? { top: rule("FF9CB2C3") } : {}),
        ...(index === studentRows.length - 1 ? { bottom: rule("FF9CB2C3") } : {}) },
    }));
  });
}

function courseSeatingWorkbook(exam) {
  const workbook = XLSX.utils.book_new();
  const sheetNames = new Set();
  exam.crnSheets.forEach((crnSheet) => {
    const info = courseSeatingInfo(exam).map(([label, value]) => [label, label === "Date"
      ? { t: "d", v: new Date(value + "T00:00:00Z"), z: "dd mmm yyyy" } : value]);
    const headerRow = info.length + 2;
    const sheet = XLSX.utils.aoa_to_sheet([["Student seating"], ...info, [], ["StudentID", "Room"], ...crnSheet.rows]);
    sheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
    styleCourseSeatingSheet(sheet, info, crnSheet.rows);
    sheet["!autofilter"] = { ref: "A" + (headerRow + 1) + ":B" + (headerRow + 1 + crnSheet.rows.length) };
    const base = ("CRN " + (crnSheet.crn || "Unspecified")).replace(/[\\/?*[\]:'\p{Cc}]/gu, "_");
    XLSX.utils.book_append_sheet(workbook, sheet, uniqueExportName(base, sheetNames, 31));
  });
  workbook.Props = { Title: exam.code + " - " + exam.title, Subject: "Confirmed student room assignments by CRN" };
  return workbook;
}

export function reportTable(report, model) {
  if (report === "overview") return {
    columns: ["Week", "Date", "Day", "Start", "End", "Course", "Course title", "CRNs", "Students", "Rooms", "Assigned rooms", "Primary invigilators needed", "During lab time", "Schedule"],
    widths: [10, 18, 16, 12, 12, 18, 44, 22, 12, 12, 48, 30, 20, 24],
    rows: model.overviewExams.map((exam) => [exam.week, exam.date, exam.day, clock(exam.start), clock(exam.end), exam.code, exam.title,
      exam.crns.join(", "), exam.studentCount, exam.isAsd ? null : exam.roomNames.length, exam.isAsd ? null : exam.roomNames.join("; "),
      exam.primaryInvigilatorsNeeded, exam.isAsd ? null : exam.duringLab ? "Yes" : "No", exam.isAsd ? "ASD (reference)" : "Department"]),
  };
  if (report === "staff") return {
    columns: ["Week", "Date", "Day", "Start", "End", "Invigilator", "Duty", "Room", "Course", "Extra load", "During teaching hours"],
    widths: [10, 18, 16, 12, 12, 34, 14, 40, 26, 14, 24],
    rows: model.duties.map((duty) => [duty.week, duty.date, duty.day, clock(duty.start), clock(duty.end), duty.name, duty.role,
      duty.roomName, duty.code, duty.extraLoad, duty.teaching ? "Yes" : "No"]),
  };
  if (report === "students") return {
    columns: ["Week", "Date", "Day", "Start", "End", "Student ID", "Student name", "CRN", "Course", "Course title", "Assigned room"],
    widths: [10, 18, 16, 12, 12, 20, 38, 14, 18, 44, 40],
    rows: model.students.map((student) => [student.week, student.date, student.day, clock(student.start), clock(student.end),
      student.id, student.name, student.crn, student.code, student.title, student.roomName]),
  };
  throw new Error("Choose a supported focused report.");
}

function appendTable(workbook, name, columns, rows, widths, dateColumn) {
  const typedRows = rows.map((row) => row.map((value, index) => index === dateColumn
    ? { t: "d", v: new Date(value + "T00:00:00Z"), z: "dd mmm yyyy" } : value));
  const sheet = XLSX.utils.aoa_to_sheet([columns, ...typedRows]);
  sheet["!cols"] = widths.map((wch) => ({ wch }));
  sheet["!autofilter"] = { ref: sheet["!ref"] };
  XLSX.utils.book_append_sheet(workbook, sheet, name);
}

export function buildCombinedResourceWorkbook(options) {
  options = { ...options, includeAsd: false };
  const model = buildExportModel(options);
  const workbook = XLSX.utils.book_new();
  const rows = model.selectedWeeks.map((week) => {
    const exams = model.exams.filter((exam) => exam.week === week);
    const roomRows = model.roomRows.filter((room) => room.week === week);
    return [week, reportDate(options.startDate, week, "Monday"), exams.length, roomRows.length,
      roomRows.reduce((sum, room) => sum + room.studentCount, 0),
      { t: "s", v: "Open Week " + week, l: { Target: "#'Week " + week + " Invigilators'!A1" } }];
  });
  appendTable(workbook, "Schedule Index", ["Week", "Week commencing", "Exams", "Room sessions", "Student sittings", "Weekly report"], rows, [10, 22, 14, 18, 20, 28], 1);
  model.selectedWeeks.forEach((week) => {
    const weekly = buildResourceWorkbookForWeek({ ...options, week });
    // Week-qualified sheet names preserve every existing cross-sheet formula.
    weekly.SheetNames.forEach((name) => XLSX.utils.book_append_sheet(workbook, weekly.Sheets[name], name));
  });
  return workbook;
}

function focusedWorkbook(report, model) {
  const workbook = XLSX.utils.book_new();
  const table = reportTable(report, model);
  const title = REPORT_VIEWS.find((view) => view.id === report).label;
  appendTable(workbook, title, table.columns, table.rows, table.widths, 1);
  if (report === "staff") appendTable(workbook, "Workload Summary",
    ["Invigilator", "Exam load", "Backup load", "During teaching hours", "Total extra duties", "Included in pool"],
    model.workloads.map((person) => [person.name, person.exam, person.backup, person.teaching, person.exam + person.backup, person.enabled ? "Yes" : "No"]),
    [34, 18, 18, 26, 22, 20]);
  return workbook;
}

function csvValue(value) {
  let text = String(value ?? "");
  // Spreadsheet apps must not execute formula-like names, IDs or course titles from CSV.
  if (typeof value === "string" && /^[\s]*[=+\-@]|^[\t\r\n]/.test(text)) text = "'" + text;
  return '"' + text.replace(/"/g, '""') + '"';
}

export function buildExportFiles({ report = "complete", format = "xlsx", packaging = report === "seating" ? "course" : "combined", examIds, ...options }) {
  const view = REPORT_VIEWS.find((item) => item.id === report);
  if (!view || !["xlsx", "csv"].includes(format) || !(report === "seating" ? packaging === "course" : ["combined", "weekly"].includes(packaging))) throw new Error("Invalid export options.");
  if (report === "complete" && format !== "xlsx") throw new Error("The complete multi-sheet report is available as Excel.");
  if (report === "seating" && format !== "xlsx") throw new Error("Course seating uses Excel files with one sheet per CRN.");
  if (examIds !== undefined && report !== "seating") throw new Error("Exam selection is only available for course seating.");
  options = { ...options, includeAsd: report === "overview" && Boolean(options.includeAsd) };
  const model = buildExportModel(options);
  if (report === "seating") {
    const exams = buildCourseSeating(model);
    if (examIds !== undefined && (!Array.isArray(examIds) || !examIds.length || examIds.some((id) => !exams.some((exam) => exam.id === id)))) {
      throw new Error("Choose an exam from the included weeks.");
    }
    const filenames = new Set();
    return exams.filter((exam) => examIds === undefined || examIds.includes(exam.id)).map((exam) => {
      const base = (exam.code + " - " + exam.title).replace(/[<>:"/\\|?*\p{Cc}]/gu, "_").trim().replace(/[. ]+$/g, "");
      return { filename: uniqueExportName(base, filenames, 180) + ".xlsx", mimeType: XLSX_MIME,
        data: StyledXLSX.write(courseSeatingWorkbook(exam), { bookType: "xlsx", type: "array", compression: true }) };
    });
  }
  const groups = packaging === "weekly" ? model.selectedWeeks.map((week) => [week]) : [model.selectedWeeks];
  return groups.map((weeks) => {
    const scoped = buildExportModel({ ...options, weeks });
    const filename = (packaging === "weekly" ? "Week_" + weeks[0] + "_" : "") + view.filename + "." + format;
    if (format === "csv") {
      const table = reportTable(report, scoped);
      return { filename, mimeType: "text/csv;charset=utf-8", data: "\uFEFF" + [table.columns, ...table.rows].map((row) => row.map(csvValue).join(",")).join("\r\n") + "\r\n" };
    }
    const workbook = report === "complete"
      ? packaging === "weekly" ? buildResourceWorkbookForWeek({ ...options, week: weeks[0] }) : buildCombinedResourceWorkbook({ ...options, weeks })
      : focusedWorkbook(report, scoped);
    workbook.Props = { Title: view.label, Subject: "Confirmed department exam schedule" };
    return { filename, mimeType: XLSX_MIME, data: XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true }) };
  });
}
