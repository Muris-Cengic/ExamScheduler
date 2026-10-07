import * as XLSX from "xlsx/xlsx.mjs";

const WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const MONTHS = [
  "january", "february", "march", "april", "may", "june",
  "july", "august", "september", "october", "november", "december",
];
const normalise = (value) =>
  String(value ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
const textValue = (value) => String(value ?? "").trim();

function headerColumns(row) {
  return new Map(row.map((value, index) => [normalise(value), index]));
}

function cell(row, columns, ...names) {
  for (const name of names) {
    const index = columns.get(normalise(name));
    if (index !== undefined && textValue(row[index])) {
      return textValue(row[index]);
    }
  }
  return "";
}

function hasColumn(columns, ...names) {
  return names.some((name) => columns.has(normalise(name)));
}

function enrolmentRows(workbook) {
  const result = [];
  workbook.SheetNames.forEach((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1, defval: "", raw: false,
    });
    let columns = null;
    let grouped = false;
    let currentStudent = null;

    rows.forEach((row, index) => {
      const candidate = headerColumns(row);
      const hasCode = hasColumn(candidate, "Course Code", "SCBCRSE_SUBJ_CODE_SCBCRSE_CRSE") ||
        (hasColumn(candidate, "SSBSECT_SUBJ_CODE") && hasColumn(candidate, "SSBSECT_CRSE_NUMB"));
      if (hasCode && (hasColumn(candidate, "Student ID", "SPRIDEN_ID") ||
        (hasColumn(candidate, "CRN", "SSBSECT_CRN") && hasColumn(candidate, "Pidm")))) {
        columns = candidate;
        grouped = !hasColumn(candidate, "Student ID", "SPRIDEN_ID");
        currentStudent = null;
        return;
      }
      if (!columns) return;

      if (grouped) {
        const identity = row.map(textValue).find((value) => /^STUDENT_ID\s*:/i.test(value));
        if (identity) {
          const match = identity.match(/^STUDENT_ID\s*:\s*([^,]+),\s*Student Name\s*:\s*(.*?)(?:,\s*Program Code\s*:|$)/i);
          if (!match || !match[1].trim() || !match[2].trim()) {
            throw new Error(`Invalid student identity at ${sheetName}, row ${index + 1}.`);
          }
          currentStudent = { id: match[1].trim(), name: match[2].trim() };
          return;
        }
        if (row.some((value) => /^sum(?:\s|$)/i.test(textValue(value)))) {
          currentStudent = null;
          return;
        }
      }

      const crn = cell(row, columns, "CRN", "SSBSECT_CRN");
      const code = cell(row, columns, "Course Code", "SCBCRSE_SUBJ_CODE_SCBCRSE_CRSE") ||
        [cell(row, columns, "SSBSECT_SUBJ_CODE"), cell(row, columns, "SSBSECT_CRSE_NUMB")].filter(Boolean).join("-");
      if (!code || (grouped && !crn)) return;

      const id = grouped ? currentStudent?.id : cell(row, columns, "Student ID", "SPRIDEN_ID");
      if (!id) {
        if (grouped) {
          throw new Error(`Course without a student identity at ${sheetName}, row ${index + 1}.`);
        }
        return;
      }
      result.push({
        id,
        name: grouped ? currentStudent.name : cell(row, columns, "Student Name"),
        code,
        title: cell(row, columns, "Title", "Course Title", "SCBCRSE_TITLE") || code,
        crn,
        section: cell(row, columns, "Section", "SSBSECT_SEQ_NUMB"),
        instructor: cell(row, columns, "Instructor", "Instructor Name", "CF_INSTRUCTOR"),
      });
    });
  });
  if (!result.length) {
    throw new Error("No enrolments found. Select a registration report or a table with Student ID and Course Code columns.");
  }
  return result;
}

export function parseEnrolmentWorkbook(workbook, studentsPerRoom = 25) {
  const groups = new Map();
  const studentDirectory = {};
  const capacity = Math.max(1, Number(studentsPerRoom) || 1);

  enrolmentRows(workbook).forEach((row) => {
    const name = row.name || row.id;
    studentDirectory[row.id] = name;
    if (!groups.has(row.code)) {
      groups.set(row.code, {
        id: row.code, code: row.code, title: row.title,
        sections: new Set(), crns: new Set(), students: new Map(),
        instructors: new Set(), crnDetails: new Map(),
      });
    }
    const group = groups.get(row.code);
    if (row.section) group.sections.add(row.section);
    if (row.instructor) group.instructors.add(row.instructor);
    const crn = row.crn || `${row.code}${row.section ? `-${row.section}` : ""}`;
    group.crns.add(crn);
    const student = { id: row.id, name, crn };
    group.students.set(row.id, student);
    if (!group.crnDetails.has(crn)) {
      group.crnDetails.set(crn, { crn, instructor: row.instructor, students: new Map() });
    }
    const detail = group.crnDetails.get(crn);
    if (row.instructor) detail.instructor = row.instructor;
    detail.students.set(row.id, student);
  });

  const sortStudents = (values) => [...values].sort(
    (a, b) => a.name.localeCompare(b.name) || a.id.localeCompare(b.id),
  );
  const courses = [...groups.values()].map((group) => {
    const students = sortStudents(group.students.values());
    const instructors = [...group.instructors].sort();
    return {
      id: group.id, code: group.code, title: group.title,
      sections: [...group.sections].sort(), crns: [...group.crns].sort(),
      students, studentCount: students.length,
      roomsNeeded: Math.ceil(students.length / capacity),
      instructors, primaryInstructor: instructors[0] || "",
      crnDetails: [...group.crnDetails.values()].map((detail) => ({
        crn: detail.crn, instructor: detail.instructor,
        students: sortStudents(detail.students.values()),
      })).sort((a, b) => a.crn.localeCompare(b.crn)),
    };
  });
  return { courses, studentDirectory };
}

function dateISO(year, month, day) {
  const date = new Date(Date.UTC(year, month, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month || date.getUTCDate() !== day) {
    throw new Error("Invalid exam date.");
  }
  return date.toISOString().slice(0, 10);
}

function parseExamDate(value, year) {
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) throw new Error("Invalid Excel exam date.");
    return dateISO(parsed.y, parsed.m - 1, parsed.d);
  }
  if (value instanceof Date) {
    return dateISO(value.getFullYear(), value.getMonth(), value.getDate());
  }
  const text = textValue(value);
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) return dateISO(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  const match = text.match(/^(?:(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),?\s+)?(\d{1,2})\s+([a-z]+)(?:\s+(\d{4}))?$/i);
  if (!match) throw new Error(`Unrecognised exam date: "${text}".`);
  const month = MONTHS.findIndex((name) => name === match[3].toLowerCase() || name.slice(0, 3) === match[3].toLowerCase());
  if (month === -1) throw new Error(`Unrecognised month: "${match[3]}".`);
  const result = dateISO(Number(match[4] || year), month, Number(match[2]));
  if (match[1]) {
    const actual = new Intl.DateTimeFormat("en", { weekday: "long", timeZone: "UTC" }).format(new Date(result));
    if (actual.toLowerCase() !== match[1].toLowerCase()) {
      throw new Error(`"${text}" does not match the year ${result.slice(0, 4)}. Check the exam start date.`);
    }
  }
  return result;
}

export function parseExamTimeRange(value) {
  const match = textValue(value).match(/^(\d{1,2}):(\d{2})\s*(AM|PM)?\s*[-\u2013\u2014]\s*(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
  if (!match) throw new Error(`Unrecognised exam time: "${value}".`);
  const startHour = Number(match[1]);
  const endHour = Number(match[4]);
  const startMinute = Number(match[2]);
  const endMinute = Number(match[5]);
  const endSuffix = match[6]?.toUpperCase();
  let startSuffix = match[3]?.toUpperCase();
  // A shared PM suffix in "12:00 - 1:00 PM" applies to both endpoints.
  if (!startSuffix && endSuffix) {
    startSuffix = endSuffix;
    if (endSuffix === "PM" && startHour !== 12 && (startHour > endHour || endHour === 12)) {
      startSuffix = "AM";
    }
  }
  const minutes = (hour, minute, suffix) => {
    if (minute > 59 || hour > (suffix ? 12 : 23) || hour < (suffix ? 1 : 0)) {
      throw new Error(`Invalid exam time: "${value}".`);
    }
    return (suffix ? hour % 12 + (suffix === "PM" ? 12 : 0) : hour) * 60 + minute;
  };
  const startMinutes = minutes(startHour, startMinute, startSuffix);
  const endMinutes = minutes(endHour, endMinute, endSuffix);
  if (endMinutes <= startMinutes) throw new Error(`Exam must end after it starts: "${value}".`);
  return { startMinutes, endMinutes };
}

export function parseAsdWorkbook(workbook, courses, year) {
  if (!courses.length) throw new Error("Load enrolment data before importing an ASD schedule.");
  const lookup = new Map();
  courses.forEach((course) => {
    new Set([normalise(course.title), normalise(course.code)]).forEach((key) => {
      if (!lookup.has(key)) lookup.set(key, []);
      lookup.get(key).push(course);
    });
  });
  const exams = [];
  const unmatched = new Set();
  const seen = new Map();
  let asdRows = 0;
  workbook.SheetNames.forEach((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
    let columns = null;
    rows.forEach((row, index) => {
      const candidate = headerColumns(row);
      if (["Course", "Department", "Day", "Time"].every((name) => hasColumn(candidate, name))) {
        columns = candidate;
        return;
      }
      if (!columns || cell(row, columns, "Department").toUpperCase() !== "ASD") return;
      const title = cell(row, columns, "Course");
      if (!title) throw new Error(`Missing ASD course title at ${sheetName}, row ${index + 1}.`);
      asdRows += 1;
      const matches = lookup.get(normalise(title)) || [];
      if (!matches.length) {
        unmatched.add(title);
        return;
      }
      if (matches.length > 1) throw new Error(`ASD course "${title}" matches several course codes. Use a course code in the Course column.`);
      let date;
      let range;
      try {
        date = parseExamDate(row[columns.get("day")], year);
        range = parseExamTimeRange(cell(row, columns, "Time"));
      } catch (error) {
        throw new Error(`${sheetName}, row ${index + 1}: ${error.message}`);
      }
      const courseId = matches[0].id;
      const signature = `${date}|${range.startMinutes}|${range.endMinutes}`;
      if (seen.has(courseId)) {
        if (seen.get(courseId) !== signature) throw new Error(`ASD course "${title}" has different exam sessions.`);
        return;
      }
      seen.set(courseId, signature);
      exams.push({ courseId, date, ...range });
    });
  });
  if (!asdRows) throw new Error("No ASD rows found. Expected Course, Department, Day and Time columns.");
  if (!exams.length) throw new Error(`No ASD courses match the loaded enrolment. Unmatched: ${[...unmatched].join(", ")}.`);
  return { exams, unmatched: [...unmatched] };
}

export function planAsdImport(exams, settings, existingWeeks = [1], baseDate = null, maxWeeks = 10) {
  const dates = exams.map((exam) => exam.date).sort();
  const first = new Date(`${baseDate || dates[0]}T00:00:00Z`);
  first.setUTCDate(first.getUTCDate() - (first.getUTCDay() + 6) % 7);
  const startDate = first.toISOString().slice(0, 10);
  const interval = exams.every((exam) =>
    exam.startMinutes % settings.slotIntervalMinutes === 0 &&
    exam.endMinutes % settings.slotIntervalMinutes === 0,
  ) ? settings.slotIntervalMinutes : 30;
  if (exams.some((exam) => exam.startMinutes % interval || exam.endMinutes % interval)) {
    throw new Error("ASD exam times must align to 30-minute slots.");
  }
  const nextSettings = {
    ...settings, slotIntervalMinutes: interval,
    startHour: Math.min(settings.startHour, Math.floor(Math.min(...exams.map((exam) => exam.startMinutes)) / 60)),
    endHour: Math.max(settings.endHour, Math.ceil(Math.max(...exams.map((exam) => exam.endMinutes)) / 60)),
  };
  if (nextSettings.endHour > 23) throw new Error("ASD exams must finish by 11 PM.");
  const weekSet = new Set(existingWeeks);
  const placements = exams.map((exam) => {
    const date = new Date(`${exam.date}T00:00:00Z`);
    const day = WEEKDAYS[date.getUTCDay() - 1];
    const week = Math.floor((date - first) / (7 * 86400000)) + 1;
    if (!day) throw new Error(`ASD exam on ${exam.date} is outside the Monday-Friday timetable.`);
    if (week < 1 || week > maxWeeks) throw new Error("ASD dates are outside the timetable range. Check the exam start date.");
    weekSet.add(week);
    const hour = Math.floor(exam.startMinutes / 60);
    const minute = exam.startMinutes % 60;
    return { ...exam, week, day, slotId: `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}` };
  });
  const weeks = [...weekSet].sort((a, b) => a - b);
  return {
    settings: nextSettings, startDate, weeks, placements,
    examDurations: Object.fromEntries(exams.map((exam) => [exam.courseId, exam.endMinutes - exam.startMinutes])),
  };
}

export function continuingCourses(dayAssignments, timeSlots, slotIndex, slotsPerExam, examSlots = {}) {
  const result = new Set();
  for (let previousIndex = 0; previousIndex < slotIndex; previousIndex += 1) {
    (dayAssignments?.[timeSlots[previousIndex].id] || []).forEach((courseId) => {
      if (slotIndex - previousIndex < (examSlots[courseId] || slotsPerExam)) result.add(courseId);
    });
  }
  return [...result];
}
