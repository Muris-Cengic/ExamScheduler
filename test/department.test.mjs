import test from "node:test";
import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import * as XLSX from "xlsx/xlsx.mjs";
import { assignmentIds, defaultHasExam, parseDepartmentWorkbook, readDepartmentSelection, retainAssignments, scopeDepartmentCourses } from "../src/department.js";
import { autoSchedule } from "../src/autoSchedule.js";
import { parseAsdWorkbook, parseEnrolmentWorkbook, planAsdImport } from "../src/imports.js";

const HEADERS = ["Campus", "Crn No", "Course Code", "Title", "Cr", "Maximum Load", "No Of Enrolled", "Session Id", "DAYS", "Time", "Primary Instructor", "Second Instructor", "Type", "Building", "Room"];
const row = (crn, code, title, days, time, type = "OL") => ["AD", crn, code, title, 3, 25, 20, "01", days, time, "123: Instructor One", "", type, "A", "Lab 1"];
const book = (rows, name = "Input") => {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), name);
  return workbook;
};
const student = (id, crn = "101") => ({ id, name: id, crn });
const course = (id, students = [student("S1")], crns = ["101"], labSessions = []) => ({
  id, code: id, title: id, students, studentCount: students.length, crns, labSessions,
  crnDetails: crns.map((crn) => ({ crn, instructor: "", students: students.filter((s) => s.crn === crn) })),
});
const settings = { slotIntervalMinutes: 30, startHour: 8, endHour: 18, studentsPerRoom: 25, invigilatorCount: 15, examDurationMinutes: 60 };
const slots = (start = 8, end = 18) => Array.from({ length: (end - start) * 2 }, (_, i) => {
  const minute = start * 60 + i * 30;
  return { id: `${String(Math.floor(minute / 60)).padStart(2, "0")}:${String(minute % 60).padStart(2, "0")}` };
});
const draft = (courses, options = {}) => autoSchedule({ courses, courseLookup: Object.fromEntries(courses.map((c) => [c.id, c])),
  assignments: {}, weeks: [1], timeSlots: slots(), settings, ...options });

test("only the first CRN sheet is used, even when a later sheet is named ISET CRNs", () => {
  const workbook = book([
    row("101", "ICT-1000", "Test", "M W", "0800 - 0950"),
    row("101", "ICT-1000", "Test", "M W", "0800 - 0950"),
    row("102", "ICT-1000", "Test", "R", "1300 - 1450"),
    row("101", "ICT-1000", "Test", "T", "1200 - 1250", "T"),
  ], "Department");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([HEADERS, row("999", "OTHER-1000", "Other", "M", "0800 - 0950")]), "All CRNs");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([HEADERS, row("998", "IGNORE-1000", "Ignore", "M", "2500 - 2600")]), "ISET CRNs");
  const parsed = parseDepartmentWorkbook(workbook);
  assert.equal(parsed.sheetName, "Department");
  assert.equal(parsed.courses.length, 1);
  assert.deepEqual(parsed.courses[0].crns, ["101", "102"]);
  assert.equal(parsed.courses[0].meetings.length, 3);
  assert.equal(parsed.courses[0].meetings.filter((meeting) => meeting.isLab).length, 2);
  assert.deepEqual(parsed.courses[0].meetings[0].days, ["Monday", "Wednesday"]);
  assert.equal(parsed.courses[0].meetings[0].instructor, "Instructor One");
  assert.equal(parsed.courses[0].code, "ICT-1000");
  assert.deepEqual(readDepartmentSelection(JSON.parse(JSON.stringify(parsed))), parsed);
});

test("an empty or missing first CRN sheet cannot fall back to another sheet", () => {
  const workbook = book([], "First");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([HEADERS, row("101", "ICT-1000", "Test", "M", "0800 - 0950")]), "ISET CRNs");
  assert.throws(() => parseDepartmentWorkbook(workbook), /No courses found in First/);
  assert.throws(() => parseDepartmentWorkbook({ SheetNames: [], Sheets: {} }), /first sheet is missing/);
});

test("CRN list handles unscheduled GP rows; invalid meeting times/days fail with their location", () => {
  const parsed = parseDepartmentWorkbook(book([HEADERS, row("101", "ICT-1000", "Graduation Project I", "", "", "GP")]));
  assert.deepEqual(parsed.courses[0].meetings[0].days, []);
  assert.equal(defaultHasExam(parsed.courses[0]), false);
  assert.throws(() => parseDepartmentWorkbook(book([HEADERS, row("101", "ICT-1000", "Test", "M", "2500 - 2600")])), /row 2: Invalid/);
  assert.throws(() => parseDepartmentWorkbook(book([HEADERS, row("101", "ICT-1000", "Test", "S", "0800 - 0950")])), /unsupported meeting days/);
  assert.throws(() => readDepartmentSelection({}), /missing its department/);
});

test("scope includes only selected CRN rosters and enriches instructors from the list", () => {
  const source = course("ICT-1000", [student("S1", "101"), student("S2", "102")], ["101", "102"]);
  const catalog = parseDepartmentWorkbook(book([HEADERS, row("101", "ict-1000", "Test", "M", "0800 - 0950"), row("103", "ict-1000", "Test", "R", "1300 - 1450")])).courses;
  const scoped = scopeDepartmentCourses([source], catalog);
  assert.deepEqual(scoped.courses[0].students.map((s) => s.id), ["S1"]);
  assert.equal(scoped.courses[0].id, "ICT-1000");
  assert.deepEqual(scoped.courses[0].crns, ["101", "103"]);
  assert.deepEqual(scoped.missingCrns, ["ict-1000: 103"]);
  assert.equal(scoped.courses[0].crnDetails[0].instructor, "Instructor One");
  assert.equal(scoped.courses[0].labSessions.length, 2);
  assert.equal(source.studentCount, 2);
});

test("exam defaults exclude projects and training, not taught project-management courses", () => {
  for (const title of ["Graduation Project I", "Graduation Project II", "Capstone Project", "GP", "On Job Training (Internship)"]) assert.equal(defaultHasExam({ title }), false);
  assert.equal(defaultHasExam({ title: "InfoSec Project Management" }), true);
  const assignments = { 1: { Monday: { "12:00": ["EXAM", "NO_EXAM", "OTHER"] } } };
  assert.deepEqual([...assignmentIds(retainAssignments(assignments, new Set(["EXAM"])))], ["EXAM"]);
  assert.equal(assignments[1].Monday["12:00"].length, 3);
});

test("OCT courses default to no exam without matching OCT inside unrelated words", () => {
  for (const title of ["OCT", "OCT I", "Ethic.Hack. & Pen.Test.- OCT I", "Pen. Testing in-Depth - OCT I", "Digit. Foren.& Invest.- OCT II", "Secure Windows/Linux OS-OCT I", "Training - oct ii"]) {
    assert.equal(defaultHasExam({ title }), false, title);
  }
  for (const title of ["Doctoral Research", "October Seminar", "Secure Windows/Linux OS"]) {
    assert.equal(defaultHasExam({ title }), true, title);
  }
});

test("single CRN uses its lab weekday and fits fully inside the lab time", () => {
  const single = course("SINGLE", [student("S1")], ["101"], [{ days: ["Thursday"], startMinutes: 900, endMinutes: 1010 }]);
  const result = draft([single]);
  assert.equal(result.unplaced.length, 0);
  assert.equal(result.placed[0].day, "Thursday");
  assert.equal(result.placed[0].slotId, "15:00");
  assert.equal(result.placed[0].end, 960);
});

test("early morning labs use 09:00-10:00, not the first hour or a later part of a long lab", () => {
  for (const endMinutes of [590, 600, 710]) {
    const c = course("MORNING", [student("S1")], ["101"], [{ days: ["Thursday"], startMinutes: 480, endMinutes }]);
    const result = draft([c]);
    assert.equal(result.unplaced.length, 0);
    assert.equal(result.placed[0].day, "Thursday");
    assert.equal(result.placed[0].slotId, "09:00");
    assert.equal(result.placed[0].end, 600);
    assert.equal(c.labSessions[0].startMinutes, 480, "Listed lab metadata must not be changed");
    const blocked = draft([c], {
      asdAssignments: { 1: { Thursday: { "09:00": ["ASD"] } } },
      courseLookup: { MORNING: c, ASD: course("ASD") },
    });
    assert.equal(blocked.placed.length, 0, "Do not fall back to 08:00 or 10:00 when the second hour conflicts");
    assert.match(blocked.unplaced[0].reason, /student overlap/);
    const unrelatedAsd = draft([c], {
      asdAssignments: { 1: { Thursday: { "09:00": ["ASD"] } } },
      courseLookup: { MORNING: c, ASD: course("ASD", [student("OTHER")]) },
    });
    assert.equal(unrelatedAsd.placed[0].slotId, "09:00", "ASD overlap without shared students is still allowed");
  }
});

test("morning lab allowance never extends shorter labs, shortens exams or exceeds timetable hours", () => {
  const c = course("MORNING", [student("S1")], ["101"], [{ days: ["Monday"], startMinutes: 480, endMinutes: 580 }]);
  const result = draft([c]);
  assert.equal(result.placed.length, 0);
  assert.match(result.unplaced[0].reason, /09:00-10:00.*09:50.*10:00.*not shortened/);
  const shorter = draft([c], { settings: { ...settings, examDurationMinutes: 30 } });
  assert.equal(shorter.placed[0].slotId, "09:00");
  assert.equal(shorter.placed[0].end, 570);
  const fullLab = course("FULL", [student("S1")], ["101"], [{ days: ["Monday"], startMinutes: 480, endMinutes: 710 }]);
  assert.equal(draft([fullLab], { timeSlots: slots(8, 9) }).placed.length, 0);
  assert.equal(draft([fullLab], { settings: { ...settings, examDurationMinutes: 90 } }).placed.length, 0);
});

test("multiple CRNs use only common windows, even when labs are available", () => {
  const multi = course("MULTI", [student("S1")], ["101", "102"], [{ days: ["Thursday"], startMinutes: 900, endMinutes: 1010 }]);
  const result = draft([multi], { timeSlots: slots(12, 18) });
  assert.equal(result.placed[0].slotId, "12:00");
  assert.equal(result.placed[0].day, "Monday");
  const middayBlocked = { 1: Object.fromEntries(["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"].map((day) => [day, { "12:00": ["ASD"] }])) };
  const evening = draft([multi], { timeSlots: slots(12, 18), asdAssignments: middayBlocked, courseLookup: { MULTI: multi, ASD: course("ASD") } });
  assert.equal(evening.placed[0].slotId, "17:00");
});

test("earlier common windows outrank lighter days and weeks", () => {
  const c = course("MAIN", [student("S1")], ["101", "102"]);
  const existing = course("EXISTING", Array.from({ length: 16 }, (_, i) => student("OTHER" + i)));
  const result = draft([c], {
    assignments: { 1: { Friday: { "09:00": ["EXISTING"] } }, 2: { Friday: { "09:00": ["EXISTING"] } } },
    courseLookup: { MAIN: c, EXISTING: existing }, weeks: [1, 2], settings: { ...settings, invigilatorCount: 4 },
  });
  assert.equal(result.placed[0].slotId, "09:00", "An occupied early slot beats empty noon/evening slots in any week");
  assert.equal(result.placed[0].day, "Friday");
  const blocked = draft([c], {
    asdAssignments: { 1: { Friday: { "09:00": ["ASD"] } } }, courseLookup: { MAIN: c, ASD: course("ASD") },
  });
  assert.equal(blocked.placed[0].slotId, "10:30", "Friday's second morning slot beats an empty weekday noon slot");
});

test("feasible exams can share noon slots instead of moving to an empty evening slot", () => {
  const c = course("MAIN", [student("S1")], ["101", "102"]);
  const existing = course("EXISTING", [student("OTHER")]);
  const assignments = { 1: Object.fromEntries(["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    .map((day) => [day, { "12:00": ["EXISTING"] }])) };
  const options = { assignments, courseLookup: { MAIN: c, EXISTING: existing }, timeSlots: slots(12, 18) };
  const result = draft([c], { ...options, settings: { ...settings, invigilatorCount: 3 } });
  assert.equal(result.unplaced.length, 0);
  assert.equal(result.placed[0].slotId, "12:00");
  assert.equal(result.placed[0].day, "Monday");
  assert.deepEqual(result.assignments[1].Monday["12:00"], ["EXISTING", "MAIN"]);
  assert.deepEqual(assignments[1].Monday["12:00"], ["EXISTING"], "Existing input maps are not mutated");
  const limited = draft([c], { ...options, settings: { ...settings, invigilatorCount: 2 } });
  assert.equal(limited.placed[0].slotId, "17:00", "Concurrent room duties plus backup capacity still limit noon placement");
  const conflict = draft([c], { ...options, courseLookup: { MAIN: c, EXISTING: course("EXISTING") } });
  assert.equal(conflict.placed[0].slotId, "17:00", "Shared students still block the earlier slot");
});

test("a batch fills feasible noon slots in parallel before using evening slots", () => {
  const courses = Array.from({ length: 7 }, (_, i) => course("EXAM-" + i, [student("STUDENT-" + i)], ["101", "102"]));
  const options = { timeSlots: slots(12, 18) };
  const enough = draft(courses, { ...options, settings: { ...settings, invigilatorCount: 3 } });
  assert.equal(enough.unplaced.length, 0);
  assert.ok(enough.placed.every((exam) => exam.slotId === "12:00"));
  assert.ok(Object.values(enough.assignments[1]).some((day) => day["12:00"].length === 2), "Independent exams can share noon");
  assert.ok(Object.values(enough.assignments[1]).every((day) => day["12:00"].length <= 2), "Reserve the slot backup as well as each room duty");
  const limited = draft(courses, { ...options, settings: { ...settings, invigilatorCount: 2 } });
  assert.equal(limited.placed.filter((exam) => exam.slotId === "12:00").length, 5);
  assert.equal(limited.placed.filter((exam) => exam.slotId === "17:00").length, 2);
});

test("earlier-slot priority still prevents a third exam for a student on Friday", () => {
  const courses = ["A", "B", "C"].map((id) => course(id, [student("S1")], ["101", "102"]));
  const result = draft(courses);
  assert.equal(result.unplaced.length, 0);
  assert.equal(result.placed[0].slotId, "09:00");
  assert.equal(result.placed[1].slotId, "10:30");
  assert.equal(result.placed[2].slotId, "12:00");
  assert.equal(result.placed[2].day, "Monday");
  assert.equal(result.placed.filter((exam) => exam.day === "Friday").length, 2);
});

test("equally early slots retain day and week load balancing", () => {
  const c = course("MAIN", [student("S1")], ["101", "102"]);
  const existing = course("EXISTING", [student("OTHER")]);
  const options = { assignments: { 1: { Monday: { "12:00": ["EXISTING"] } } },
    courseLookup: { MAIN: c, EXISTING: existing }, timeSlots: slots(12, 18) };
  assert.equal(draft([c], options).placed[0].day, "Tuesday", "Choose a lighter day without changing the preferred start time");
  assert.equal(draft([c], { ...options, weeks: [1, 2] }).placed[0].week, 2, "Choose a lighter week when start times and day loads tie");
});

test("single-CRN labs also prefer earlier feasible starts without leaving their lab window", () => {
  const c = course("LAB", [student("S1")], ["101"], [{ days: ["Thursday"], startMinutes: 900, endMinutes: 1010 }]);
  const existing = course("EXISTING", [student("OTHER")]);
  const result = draft([c], {
    assignments: { 1: { Thursday: { "15:00": ["EXISTING"] } } },
    courseLookup: { LAB: c, EXISTING: existing }, settings: { ...settings, invigilatorCount: 3 },
  });
  assert.equal(result.placed[0].slotId, "15:00");
  assert.equal(result.placed[0].end, 960);
});

test("one CRN without a lab uses common windows; long exams stay unplaced", () => {
  assert.equal(draft([course("NO_LAB")]).placed[0].slotId, "09:00");
  const result = draft([course("LONG", [student("S1")], ["101", "102"])], { settings: { ...settings, examDurationMinutes: 120 } });
  assert.equal(result.placed.length, 0);
  assert.match(result.unplaced[0].reason, /No matching window/);
  assert.equal(draft([course("EARLY_ONLY")], { timeSlots: slots(8, 9) }).placed.length, 0);
});

test("Friday morning windows are available for multiple CRNs and labless courses only on Friday", () => {
  for (const crns of [["101", "102"], ["101"]]) {
    const c = course("MAIN", [student("S1")], crns);
    const morning = draft([c], { timeSlots: slots(8, 12) });
    assert.equal(morning.placed[0].day, "Friday");
    assert.equal(morning.placed[0].slotId, "09:00");
    assert.equal(morning.placed[0].end, 600);
    const later = draft([c], {
      timeSlots: slots(8, 12),
      asdAssignments: { 1: { Friday: { "09:00": ["ASD"] } } },
      courseLookup: { MAIN: c, ASD: course("ASD") },
    });
    assert.equal(later.placed[0].day, "Friday");
    assert.equal(later.placed[0].slotId, "10:30");
    assert.equal(later.placed[0].end, 690);
  }
});

test("Friday common windows do not replace a single-CRN course's lab rule", () => {
  const c = course("LAB", [student("S1")], ["101"], [{ days: ["Thursday"], startMinutes: 900, endMinutes: 1010 }]);
  assert.equal(draft([c], { timeSlots: slots(8, 12) }).placed.length, 0);
});

test("ASD overlap blocks labs; an adjacent exam at ASD's actual end is allowed", () => {
  const single = course("SINGLE", [student("S1")], ["101"], [{ days: ["Monday"], startMinutes: 600, endMinutes: 710 }]);
  const asd = course("ASD");
  const options = { asdAssignments: { 1: { Monday: { "10:00": ["ASD"] } } }, courseLookup: { SINGLE: single, ASD: asd } };
  assert.equal(draft([single], options).placed.length, 0);
  const adjacent = draft([single], { ...options, asdExamDurations: { ASD: 30 } });
  assert.equal(adjacent.placed[0].slotId, "10:30");
  assert.equal(adjacent.placed[0].end, 690);
});

test("more than two exams in a day is avoided across both schedules", () => {
  const c = course("MAIN", [student("S1")], ["101"], [{ days: ["Monday"], startMinutes: 900, endMinutes: 1010 }]);
  const result = draft([c], { asdAssignments: { 1: { Monday: { "08:00": ["ASD1"], "10:00": ["ASD2"] } } }, courseLookup: { MAIN: c, ASD1: course("ASD1"), ASD2: course("ASD2") } });
  assert.equal(result.placed.length, 0);
  assert.match(result.unplaced[0].reason, /two exams/);
});

test("staffing reserves two invigilators above 15 students plus slot backups; ASD excluded", () => {
  const c = course("MAIN", Array.from({ length: 16 }, (_, i) => student(`S${i}`)));
  c.labSessions.push({ days: ["Monday"], startMinutes: 720, endMinutes: 780 });
  const short = draft([c], { settings: { ...settings, invigilatorCount: 2 } });
  assert.equal(short.placed.length, 0);
  assert.match(short.unplaced[0].reason, /insufficient invigilators/);
  const asd = course("ASD", Array.from({ length: 100 }, (_, i) => student(`OTHER${i}`)));
  const enough = draft([c], { settings: { ...settings, invigilatorCount: 3 }, asdAssignments: { 1: { Monday: { "12:00": ["ASD"] } } }, courseLookup: { MAIN: c, ASD: asd } });
  assert.equal(enough.placed.length, 1);
});

test("auto-scheduling staffing uses balanced room sizes rather than capacity-filled rooms", () => {
  const large = course("LARGE", Array.from({ length: 55 }, (_, i) => student(`S${i}`)));
  const short = draft([large], { settings: { ...settings, invigilatorCount: 6 } });
  assert.equal(short.placed.length, 0, "19/18/18 needs six room invigilators plus one backup");
  assert.match(short.unplaced[0].reason, /insufficient invigilators/);
  assert.equal(draft([large], { settings: { ...settings, invigilatorCount: 7 } }).placed.length, 1);
  const small = course("SMALL", Array.from({ length: 26 }, (_, i) => student(`S${i}`)));
  assert.equal(draft([small], { settings: { ...settings, invigilatorCount: 3 } }).placed.length, 1, "13/13 needs only two room invigilators plus one backup");
});

test("automatic staffing respects only explicitly approved room consolidation", () => {
  const c = course("MAIN", Array.from({ length: 52 }, (_, i) => student(`S${i}`)));
  assert.equal(draft([c], { settings: { ...settings, invigilatorCount: 5 } }).placed.length, 0);
  assert.equal(draft([c], { settings: { ...settings, invigilatorCount: 5 }, roomDistributionChoices: { MAIN: "distribute" } }).placed.length, 1);
  assert.equal(draft([c], { settings: { ...settings, invigilatorCount: 5, studentsPerRoom: 20 }, roomDistributionChoices: { MAIN: "distribute" } }).placed.length, 0, "A lower configured limit cannot be overridden");
});

test("simultaneous courses have separate rooms and invigilators", () => {
  const c = course("NEW", [student("NEW_STUDENT")], ["101"], [{ days: ["Monday"], startMinutes: 720, endMinutes: 780 }]);
  const existing = course("EXISTING", [student("OTHER")]);
  const result = draft([c], { assignments: { 1: { Monday: { "12:00": ["EXISTING"] } } }, courseLookup: { NEW: c, EXISTING: existing }, settings: { ...settings, invigilatorCount: 3 } });
  assert.equal(result.placed.length, 1);
});

test("existing schedules are preserved, input maps are not mutated and reruns do not duplicate exams", () => {
  const c = course("MAIN");
  const original = { 1: { Monday: { "12:00": ["MAIN"] } } };
  const result = draft([c], { assignments: original });
  assert.deepEqual(result.assignments, original);
  assert.equal(result.placed.length, 0);
  const first = draft([c]);
  const second = draft([c], { assignments: first.assignments });
  assert.equal(second.placed.length, 0);
  assert.equal(assignmentIds(second.assignments).size, 1);
  assert.equal(original[1].Monday["12:00"].length, 1);
});

test("staffing checks actual concurrent slots rather than combining consecutive exams", () => {
  const c = course("MAIN", [student("S1")], ["101"], [{ days: ["Monday"], startMinutes: 930, endMinutes: 990 }]);
  const result = draft([c], {
    assignments: { 1: { Monday: { "15:00": ["EARLY"], "16:00": ["LATE"] } } },
    courseLookup: { MAIN: c, EARLY: course("EARLY", [student("S2")]), LATE: course("LATE", [student("S3")]) },
    settings: { ...settings, invigilatorCount: 3 },
  });
  assert.equal(result.placed[0].slotId, "15:30");
});

test("an additional week provides a valid lab slot when the first week is blocked", () => {
  const c = course("MAIN", [student("S1")], ["101"], [{ days: ["Monday"], startMinutes: 480, endMinutes: 600 }]);
  const result = draft([c], { weeks: [1, 2], asdAssignments: { 1: { Monday: { "09:00": ["ASD"] } } }, courseLookup: { MAIN: c, ASD: course("ASD") } });
  assert.equal(result.placed[0].week, 2);
  assert.equal(result.placed[0].slotId, "09:00");
  assert.equal(result.unplaced.length, 0);
});

const root = "data/input/26-27 S1/";
const files = ["Students Registration 26-27 S1.xlsx", "CRN List 26-27 S1.xlsx", "Other Departments Schedule/Midterm Schedule 26-27 S1.xlsx"];
test("supplied department list scopes enrolment and produces a valid automatic draft around ASD", { skip: files.some((file) => !existsSync(root + file)) }, () => {
  const read = (file) => XLSX.read(readFileSync(root + file), { type: "buffer" });
  const enrolment = parseEnrolmentWorkbook(read(files[0]));
  const catalog = parseDepartmentWorkbook(read(files[1]));
  const scope = scopeDepartmentCourses(enrolment.courses, catalog.courses);
  assert.equal(scope.courses.length, 37);
  assert.equal(scope.courses.flatMap((c) => c.crns).length, 78);
  assert.equal(scope.courses.reduce((sum, c) => sum + c.labSessions.length, 0), 46);
  assert.deepEqual(scope.missingCrns, []);
  const selected = scope.courses.filter(defaultHasExam);
  assert.equal(selected.length, 26);
  assert.equal(scope.courses.filter((course) => /\bOCT\b/i.test(course.title)).length, 4);
  assert.ok(selected.every((course) => !/\bOCT\b/i.test(course.title)));
  const asd = parseAsdWorkbook(read(files[2]), enrolment.courses, 2026);
  const plan = planAsdImport(asd.exams, settings);
  const asdAssignments = {};
  plan.placements.forEach(({ week, day, slotId, courseId }) => {
    asdAssignments[week] ??= {};
    asdAssignments[week][day] ??= {};
    asdAssignments[week][day][slotId] ??= [];
    asdAssignments[week][day][slotId].push(courseId);
  });
  const courseLookup = Object.fromEntries([...enrolment.courses, ...scope.courses].map((c) => [c.id, c]));
  const result = draft(selected, { courseLookup, asdAssignments, asdExamDurations: plan.examDurations, weeks: plan.weeks });
  assert.equal(result.placed.length + result.unplaced.length, selected.length);
  assert.equal(result.placed.length, 26);
  assert.deepEqual(result.unplaced, []);
  for (const placed of result.placed) {
    const c = courseLookup[placed.courseId];
    if (c.crns.length === 1 && c.labSessions.length) {
      assert.ok(c.labSessions.some((lab) => lab.days.includes(placed.day) && placed.start >= lab.startMinutes &&
        placed.end <= (lab.startMinutes < 540 && lab.endMinutes === 590 ? 600 : lab.endMinutes)));
      if (c.labSessions.every((lab) => lab.startMinutes < 540)) {
        assert.equal(placed.slotId, "09:00");
        assert.equal(placed.end, 600);
      }
    } else assert.ok([720, 1020].includes(placed.start) || (placed.day === "Friday" && [540, 630].includes(placed.start)));
    assert.ok(defaultHasExam(c));
    for (const other of [...result.placed, ...plan.placements.map((p) => ({ ...p, start: p.startMinutes, end: p.endMinutes }))]) {
      if (other.courseId === placed.courseId || other.week !== placed.week || other.day !== placed.day || placed.start >= other.end || other.start >= placed.end) continue;
      const otherIds = new Set(courseLookup[other.courseId].students.map((s) => s.id));
      assert.ok(c.students.every((s) => !otherIds.has(s.id)), `Overlap ${c.code} and ${other.courseId}`);
    }
  }
});
