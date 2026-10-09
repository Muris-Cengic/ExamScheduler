import * as XLSX from "xlsx/xlsx.mjs";
import { examRoomSizes } from "./examRooms.js";

const normalise = (value) => String(value ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
const text = (value) => String(value ?? "").trim();
const DAY_CODES = { M: "Monday", T: "Tuesday", W: "Wednesday", R: "Thursday", F: "Friday" };

export function instructorIdentity(value) {
  const raw = text(value);
  if (!raw) return null;
  const match = raw.match(/^(\d+)\s*:\s*(.+)$/);
  const name = match ? match[2].trim() : raw;
  return { id: match ? `staff:${match[1]}` : `staff:${normalise(name)}`, name };
}

export function roomIdentity(campus, building, room) {
  return `room:${[campus, building, room].map(normalise).join(":")}`;
}

export function labExamEndMinutes(lab) {
  // Listed :50 lab endings include the final ten minutes of that exam hour.
  return lab.endMinutes % 60 === 50 ? lab.endMinutes + 10 : lab.endMinutes;
}

const BANNER_COLUMNS = [
  "Campus", "Crn No", "Course Code", "Title", "Cr", "Maximum Load", "No Of Enrolled",
  "Session Id", "DAYS", "Time", "Primary Instructor", "Second Instructor", "Type", "Building", "Room",
];

export function defaultHasExam(course) {
  return !/graduation\s*project|capstone|on\s*job\s*training|internship|\b(?:GP|OCT)\b/i.test(course.title);
}

function meetingTime(value) {
  const match = text(value).match(/^(\d{2}):?(\d{2})\s*-\s*(\d{2}):?(\d{2})$/);
  if (!match) throw new Error(`Invalid meeting time "${value}". Expected 0800 - 0950 or 08:00 - 09:50.`);
  const [, h1, m1, h2, m2] = match.map(Number);
  const startMinutes = h1 * 60 + m1;
  const endMinutes = h2 * 60 + m2;
  if (h1 > 23 || h2 > 23 || m1 > 59 || m2 > 59 || endMinutes <= startMinutes) {
    throw new Error(`Invalid meeting time "${value}".`);
  }
  return { startMinutes, endMinutes };
}

export function parseDepartmentWorkbook(workbook) {
  const sheetName = workbook.SheetNames[0];
  if (!workbook.Sheets[sheetName]) throw new Error("The CRN workbook's first sheet is missing.");
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: false });
  let columns;
  const groups = new Map();
  rows.forEach((row, index) => {
    const candidate = row.map(normalise);
    if (candidate.includes("coursecode") && (candidate.includes("crnno") || candidate.includes("crn"))) {
      columns = candidate;
      return;
    }
    // The department-only sheet uses the same 15-column Banner layout without a header.
    if (!columns && row.length >= 15 && /^\d+$/.test(text(row[1])) && /^[A-Z]+[- ]?\d{4}$/i.test(text(row[2]))) {
      columns = BANNER_COLUMNS.map(normalise);
    }
    if (!columns) return;
    const cell = (...names) => {
      for (const name of names) {
        const value = text(row[columns.indexOf(normalise(name))]);
        if (value) return value;
      }
      return "";
    };
    const crn = cell("Crn No", "CRN");
    const code = cell("Course Code");
    if (!crn || !code) return;
    if (!groups.has(normalise(code))) {
      groups.set(normalise(code), { code, title: cell("Title", "Course Title") || code, crns: [], meetings: [] });
    }
    const course = groups.get(normalise(code));
    if (!course.crns.includes(crn)) course.crns.push(crn);
    const type = cell("Type", "Session Type").toUpperCase();
    const primary = instructorIdentity(cell("Primary Instructor", "Instructor"));
    const second = instructorIdentity(cell("Second Instructor"));
    const instructors = [...new Map([primary, second].filter(Boolean).map((person) => [person.id, person])).values()];
    const isLab = ["OL", "LAB", "LB", "L/B", "PRA", "PRACTICAL"].includes(type);
    const lead = isLab ? second || primary : primary || second;
    const daysText = cell("DAYS").toUpperCase().replace(/\s/g, "");
    const timeText = cell("Time");
    let range = {};
    if (timeText && !/^(-|TBA|TBD|N\/A)$/i.test(timeText)) {
      try { range = meetingTime(timeText); }
      catch (error) { throw new Error(`${sheetName}, row ${index + 1}: ${error.message}`); }
    }
    if (daysText && !/^[MTWRF]+$/.test(daysText)) {
      throw new Error(`${sheetName}, row ${index + 1}: unsupported meeting days "${daysText}".`);
    }
    const meeting = {
      crn, type, instructor: lead?.name || "", instructors,
      teachingInstructorIds: lead ? [lead.id] : [], labInstructorId: isLab ? lead?.id || "" : "",
      days: [...new Set([...daysText].map((day) => DAY_CODES[day]))],
      ...range, room: cell("Room"), building: cell("Building"), campus: cell("Campus"), isLab,
    };
    if (!course.meetings.some((entry) => JSON.stringify(entry) === JSON.stringify(meeting))) course.meetings.push(meeting);
  });
  if (!groups.size) throw new Error(`No courses found in ${sheetName}. Use a CRN list with Course Code and CRN columns.`);
  return { sheetName, courses: [...groups.values()].sort((a, b) => a.code.localeCompare(b.code)) };
}

export function scopeDepartmentCourses(enrolmentCourses, catalog, studentsPerRoom = 25, roomDistributionChoices = {}) {
  const lookup = new Map(enrolmentCourses.map((course) => [normalise(course.code), course]));
  const missingCrns = [];
  const courses = catalog.map((entry) => {
    const source = lookup.get(normalise(entry.code));
    const details = entry.crns.map((crn) => {
      const detail = source?.crnDetails.find((item) => item.crn === crn);
      if (!detail?.students.length) missingCrns.push(`${entry.code}: ${crn}`);
      const instructor = detail?.instructor || entry.meetings.find((meeting) => meeting.crn === crn && meeting.instructor)?.instructor || "";
      return { crn, instructor, students: detail?.students || [] };
    });
    const students = [...new Map(details.flatMap((detail) => detail.students).map((student) => [student.id, student])).values()];
    const instructors = [...new Set(details.map((detail) => detail.instructor).filter(Boolean))];
    return {
      ...source, id: source?.id || entry.code, code: entry.code, title: source?.title || entry.title,
      crns: entry.crns, sections: [], crnDetails: details, students, studentCount: students.length,
      roomsNeeded: examRoomSizes(students.length, studentsPerRoom, roomDistributionChoices[source?.id || entry.code]).length,
      instructors, primaryInstructor: instructors[0] || "", meetings: entry.meetings,
      labSessions: entry.meetings.filter((meeting) => meeting.isLab && meeting.days.length && Number.isFinite(meeting.startMinutes)),
    };
  });
  return { courses, missingCrns };
}

export function assignmentIds(assignments) {
  return new Set(Object.values(assignments).flatMap((week) => Object.values(week).flatMap((day) => Object.values(day).flat())));
}

export function retainAssignments(assignments, allowedIds) {
  return Object.fromEntries(Object.entries(assignments).map(([week, dayMap]) => [week,
    Object.fromEntries(Object.entries(dayMap).map(([day, slots]) => [day,
      Object.fromEntries(Object.entries(slots).map(([slot, ids]) => [slot, ids.filter((id) => allowedIds.has(id))])),
    ])),
  ]));
}

export function readDepartmentSelection(value) {
  if (!value || !Array.isArray(value.courses) || !value.courses.length || typeof value.sheetName !== "string") {
    throw new Error("The snapshot is missing its department CRN list. Upload enrolment and the CRN list to start a new timetable.");
  }
  for (const course of value.courses) {
    if (typeof course.code !== "string" || typeof course.title !== "string" || !Array.isArray(course.crns) ||
      !course.crns.length || course.crns.some((crn) => typeof crn !== "string") || !Array.isArray(course.meetings)) {
      throw new Error("Invalid department CRN list in snapshot.");
    }
    for (const meeting of course.meetings) {
      if (!course.crns.includes(meeting.crn) || !Array.isArray(meeting.days) ||
        meeting.days.some((day) => !Object.values(DAY_CODES).includes(day)) ||
        typeof meeting.isLab !== "boolean" || typeof meeting.instructor !== "string" ||
        typeof meeting.campus !== "string" || typeof meeting.building !== "string" || typeof meeting.room !== "string" ||
        typeof meeting.labInstructorId !== "string" || !Array.isArray(meeting.instructors) ||
        meeting.instructors.some((person) => typeof person.id !== "string" || typeof person.name !== "string") ||
        !Array.isArray(meeting.teachingInstructorIds) || meeting.teachingInstructorIds.some((id) => !meeting.instructors.some((person) => person.id === id)) ||
        (meeting.startMinutes !== undefined && (!Number.isFinite(meeting.startMinutes) || !Number.isFinite(meeting.endMinutes) ||
          meeting.startMinutes < 0 || meeting.endMinutes > 1440 || meeting.endMinutes <= meeting.startMinutes))) {
        throw new Error("Invalid meeting in snapshot CRN list.");
      }
    }
  }
  return value;
}
