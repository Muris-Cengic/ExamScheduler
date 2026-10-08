import test from "node:test";
import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import * as XLSX from "xlsx/xlsx.mjs";
import { assignmentIds, defaultHasExam, parseDepartmentWorkbook, readDepartmentSelection, retainAssignments, roomIdentity, scopeDepartmentCourses } from "../src/department.js";
import { buildExamSessions, parseResourceCatalog, standbyAvailability, validateResourcePlan } from "../src/resources.js";
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
  id, code: id, title: id, students, studentCount: students.length, crns,
  labSessions: labSessions.map((lab) => ({ campus: "Campus", building: "Building", room: id + " Lab", labInstructorId: "I0", ...lab })),
  crnDetails: crns.map((crn) => ({ crn, instructor: "", students: students.filter((s) => s.crn === crn) })),
});
const settings = { slotIntervalMinutes: 30, startHour: 8, endHour: 18, studentsPerRoom: 25, examDurationMinutes: 60 };
const slots = (start = 8, end = 18) => Array.from({ length: (end - start) * 2 }, (_, i) => {
  const minute = start * 60 + i * 30;
  return { id: `${String(Math.floor(minute / 60)).padStart(2, "0")}:${String(minute % 60).padStart(2, "0")}` };
});
function resourceFixture(courses, staffCount = 15, roomCount = 15) {
  const rooms = new Map(Array.from({ length: roomCount }, (_, i) => ["R" + i, { id: "R" + i, name: "Room " + i, enabled: true, busy: [] }]));
  courses.forEach((course) => (course.labSessions || []).forEach((lab) => {
    const id = roomIdentity(lab.campus, lab.building, lab.room);
    rooms.set(id, { id, name: lab.room, enabled: true, busy: [] });
  }));
  return { rooms: [...rooms.values()], invigilators: Array.from({ length: staffCount }, (_, i) => ({ id: "I" + i, name: "Staff " + i, enabled: true, busy: [] })) };
}
const draft = (courses, options = {}) => {
  const courseLookup = options.courseLookup || Object.fromEntries(courses.map((c) => [c.id, c]));
  return autoSchedule({ courses, courseLookup, assignments: {}, weeks: [1], timeSlots: slots(), settings,
    catalog: options.catalog || resourceFixture(Object.values(courseLookup), options.staffCount), ...options });
};

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

test("daytime scheduling prioritizes earlier weeks and weekdays before using evenings", () => {
  const c = course("MAIN", [student("S1")], ["101", "102"]);
  const existing = course("EXISTING", Array.from({ length: 16 }, (_, i) => student("OTHER" + i)));
  const result = draft([c], {
    assignments: { 1: { Friday: { "09:00": ["EXISTING"] } }, 2: { Friday: { "09:00": ["EXISTING"] } } },
    courseLookup: { MAIN: c, EXISTING: existing }, weeks: [2, 1], staffCount: 4,
  });
  assert.equal(result.placed[0].week, 1);
  assert.equal(result.placed[0].slotId, "12:00", "Monday noon comes before Friday morning");
  assert.equal(result.placed[0].day, "Monday");
  const blocked = draft([c], {
    asdAssignments: { 1: { Monday: { "12:00": ["ASD"] } } }, courseLookup: { MAIN: c, ASD: course("ASD") },
  });
  assert.equal(blocked.placed[0].slotId, "12:00", "Tuesday noon comes before Monday evening");
  assert.equal(blocked.placed[0].day, "Tuesday");
});

test("noon in a later configured week comes before an evening in the first week", () => {
  const c = course("MAIN", [student("S1")], ["101", "102"]);
  const blocked = { 1: Object.fromEntries(["Monday", "Tuesday", "Wednesday", "Thursday"]
    .map((day) => [day, { "12:00": ["ASD"] }])) };
  blocked[1].Friday = { "09:00": ["ASD"], "10:30": ["ASD"] };
  const result = draft([c], { weeks: [2, 1], asdAssignments: blocked,
    courseLookup: { MAIN: c, ASD: course("ASD") } });
  assert.deepEqual(result.placed.map((exam) => [exam.week, exam.day, exam.slotId]), [[2, "Monday", "12:00"]]);
  assert.deepEqual(result.unplaced, []);
});

test("Friday morning is tried before earlier weekday evenings when all noon sessions conflict", () => {
  const c = course("MAIN", [student("S1")], ["101", "102"]);
  const blocked = { 1: Object.fromEntries(["Monday", "Tuesday", "Wednesday", "Thursday"]
    .map((day) => [day, { "12:00": ["ASD"] }])) };
  const result = draft([c], { asdAssignments: blocked, courseLookup: { MAIN: c, ASD: course("ASD") } });
  assert.deepEqual(result.placed.map((exam) => [exam.day, exam.slotId]), [["Friday", "09:00"]]);
});

test("feasible exams can share noon slots instead of moving to an empty evening slot", () => {
  const c = course("MAIN", [student("S1")], ["101", "102"]);
  const existing = course("EXISTING", [student("OTHER")]);
  const assignments = { 1: Object.fromEntries(["Monday", "Tuesday", "Wednesday", "Thursday"]
    .map((day) => [day, { "12:00": ["EXISTING"] }])) };
  const options = { assignments, courseLookup: { MAIN: c, EXISTING: existing }, timeSlots: slots(12, 18) };
  const result = draft([c], { ...options, staffCount: 4 });
  assert.equal(result.unplaced.length, 0);
  assert.equal(result.placed[0].slotId, "12:00");
  assert.equal(result.placed[0].day, "Monday");
  assert.deepEqual(result.assignments[1].Monday["12:00"], ["EXISTING", "MAIN"]);
  assert.deepEqual(assignments[1].Monday["12:00"], ["EXISTING"], "Existing input maps are not mutated");
  const limited = draft([c], { ...options, staffCount: 2 });
  assert.equal(limited.placed[0].slotId, "17:00", "Concurrent room duties plus backup capacity still limit noon placement");
  const conflict = draft([c], { ...options, courseLookup: { MAIN: c, EXISTING: course("EXISTING") } });
  assert.equal(conflict.placed[0].slotId, "17:00", "Shared students still block the earlier slot");
});

test("a batch fills parallel noon slots across all days before moving to evenings", () => {
  const courses = Array.from({ length: 7 }, (_, i) => course("EXAM-" + i, [student("STUDENT-" + i)], ["101", "102"]));
  const options = { timeSlots: slots(12, 18) };
  const enough = draft(courses, { ...options, staffCount: 4 });
  assert.equal(enough.unplaced.length, 0);
  assert.deepEqual(enough.placed.map((exam) => [exam.day, exam.slotId]),
    [["Monday", "12:00"], ["Monday", "12:00"], ["Tuesday", "12:00"], ["Tuesday", "12:00"],
      ["Wednesday", "12:00"], ["Wednesday", "12:00"], ["Thursday", "12:00"]]);
  assert.deepEqual(enough.warnings, []);
  const limited = draft(courses, { ...options, staffCount: 2 });
  assert.equal(limited.unplaced.length, 0);
  assert.equal(limited.placed.filter((exam) => exam.day === "Monday").length, 2);
  assert.equal(limited.placed.filter((exam) => exam.day === "Tuesday").length, 2);
  assert.deepEqual(limited.placed.map((exam) => [exam.day, exam.slotId]),
    [["Monday", "12:00"], ["Tuesday", "12:00"], ["Wednesday", "12:00"], ["Thursday", "12:00"],
      ["Monday", "17:00"], ["Tuesday", "17:00"], ["Wednesday", "17:00"]]);
  assert.equal(limited.warnings.length, 7, "One-backup fallback is explicit when two standby people cannot fit anywhere");
});

test("daytime placement puts a student's conflicting exams in noon slots on successive days", () => {
  const courses = ["A", "B", "C"].map((id) => course(id, [student("S1")], ["101", "102"]));
  const result = draft(courses);
  assert.equal(result.unplaced.length, 0);
  assert.deepEqual(result.placed.map((exam) => [exam.day, exam.slotId]),
    [["Monday", "12:00"], ["Tuesday", "12:00"], ["Wednesday", "12:00"]]);
});

test("daytime capacity goes to exams needing more invigilators, accounting for room consolidation", () => {
  const standard = course("STANDARD", Array.from({ length: 51 }, (_, i) => student("S" + i)), ["101", "102"]);
  const consolidated = course("CONSOLIDATED", Array.from({ length: 52 }, (_, i) => student("C" + i)), ["101", "102"]);
  const courses = [consolidated, standard];
  const catalog = resourceFixture(courses, 7, 3);
  catalog.rooms.forEach((room) => { room.busy = [{ code: "CLASS", crn: "900", days: ["Tuesday", "Wednesday", "Thursday", "Friday"], start: 0, end: 1440, isLab: false }]; });
  const choices = { CONSOLIDATED: "distribute" };
  const result = draft(courses, { catalog, roomDistributionChoices: choices });
  assert.deepEqual(result.placed.map((exam) => [exam.courseId, exam.day, exam.slotId]),
    [["STANDARD", "Monday", "12:00"], ["CONSOLIDATED", "Monday", "17:00"]]);
  const sessions = buildExamSessions(result.assignments, Object.fromEntries(courses.map((c) => [c.id, c])), 60, 25, choices);
  assert.deepEqual(sessions.map((session) => session.rooms.reduce((sum, room) => sum + room.requiredInvigilators, 0)), [5, 4],
    "The larger consolidated exam needs less staffing and belongs in evening, not the 51-student exam");
  assert.equal(validateResourcePlan(sessions, catalog, result.resourcePlan).complete, true);
  assert.deepEqual(result.warnings, []);
});

test("evening fallback places exams needing the fewest invigilators first, not the largest exam first", () => {
  const courses = [["LARGE", 51], ["MEDIUM", 16], ["SMALL", 10]].map(([id, count]) =>
    course(id, Array.from({ length: count }, (_, i) => student(id + i)), ["101", "102"]));
  const catalog = resourceFixture(courses, 7, 3);
  catalog.rooms.forEach((room) => { room.busy = [
    { code: "NOON", crn: "900", days: ["Monday", "Tuesday", "Wednesday", "Thursday"], start: 720, end: 780, isLab: false },
    { code: "FRIDAY", crn: "901", days: ["Friday"], start: 540, end: 690, isLab: false },
  ]; });
  const result = draft(courses, { catalog });
  assert.deepEqual(result.placed.map((exam) => [exam.courseId, exam.day, exam.slotId]),
    [["SMALL", "Monday", "17:00"], ["MEDIUM", "Monday", "17:00"], ["LARGE", "Tuesday", "17:00"]]);
  const sessions = buildExamSessions(result.assignments, Object.fromEntries(courses.map((c) => [c.id, c])), 60);
  assert.equal(validateResourcePlan(sessions, catalog, result.resourcePlan).complete, true);
  assert.deepEqual(result.warnings, []);
});

test("fixed evening labs keep their required window while flexible courses still try daytime first", () => {
  const lab = course("LAB", [student("LAB-STUDENT")], ["101"], [{ days: ["Monday"], startMinutes: 1020, endMinutes: 1080 }]);
  const flexible = course("FLEXIBLE", [student("OTHER")], ["101", "102"]);
  const result = draft([flexible, lab]);
  assert.deepEqual(result.placed.map((exam) => [exam.courseId, exam.day, exam.slotId]),
    [["LAB", "Monday", "17:00"], ["FLEXIBLE", "Monday", "12:00"]]);
  const sessions = buildExamSessions(result.assignments, { LAB: lab, FLEXIBLE: flexible }, 60);
  const labRoom = sessions.find((session) => session.slotId === "17:00").rooms[0];
  assert.equal(result.resourcePlan.allocations[labRoom.id].roomId, labRoom.fixedRoomId);
  assert.equal(result.resourcePlan.allocations[labRoom.id].invigilatorIds[0], labRoom.fixedInvigilatorId);
});

test("earliest dates outrank lighter days and weeks when resource and student constraints allow", () => {
  const c = course("MAIN", [student("S1")], ["101", "102"]);
  const existing = course("EXISTING", [student("OTHER")]);
  const options = { assignments: { 1: { Monday: { "12:00": ["EXISTING"] } } },
    courseLookup: { MAIN: c, EXISTING: existing }, timeSlots: slots(12, 18) };
  assert.equal(draft([c], options).placed[0].day, "Monday", "An occupied earlier slot is retained when capacity allows");
  assert.equal(draft([c], { ...options, weeks: [1, 2] }).placed[0].week, 1, "Do not use a later week just to balance the timetable");
});

test("single-CRN labs also prefer earlier feasible starts without leaving their lab window", () => {
  const c = course("LAB", [student("S1")], ["101"], [{ days: ["Thursday"], startMinutes: 900, endMinutes: 1010 }]);
  const existing = course("EXISTING", [student("OTHER")]);
  const result = draft([c], {
    assignments: { 1: { Thursday: { "15:00": ["EXISTING"] } } },
    courseLookup: { LAB: c, EXISTING: existing }, staffCount: 3,
  });
  assert.equal(result.placed[0].slotId, "15:00");
  assert.equal(result.placed[0].end, 960);
});

test("one CRN without a lab uses common windows; long exams stay unplaced", () => {
  assert.equal(draft([course("NO_LAB")]).placed[0].slotId, "12:00");
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

test("Friday never falls back to noon or evening after its two morning sessions are blocked", () => {
  for (const crns of [["101", "102"], ["101"]]) {
    const c = course("MAIN", [student("S1")], crns);
    const asd = course("ASD", [student("S1")]);
    const blocked = { 1: Object.fromEntries(["Monday", "Tuesday", "Wednesday", "Thursday"]
      .map((day) => [day, { "12:00": ["ASD"], "17:00": ["ASD"] }])) };
    const first = draft([c], { asdAssignments: blocked, courseLookup: { MAIN: c, ASD: asd } });
    assert.equal(first.placed[0].day, "Friday");
    assert.equal(first.placed[0].slotId, "09:00");
    blocked[1].Friday = { "09:00": ["ASD"] };
    const second = draft([c], { asdAssignments: blocked, courseLookup: { MAIN: c, ASD: asd } });
    assert.equal(second.placed[0].slotId, "10:30");
    blocked[1].Friday["10:30"] = ["ASD"];
    const none = draft([c], { asdAssignments: blocked, courseLookup: { MAIN: c, ASD: asd } });
    assert.equal(none.placed.length, 0, "Free Friday noon/evening resources cannot bypass the allowed windows");
    assert.equal(none.unplaced.length, 1);
  }
});

test("Friday lab exams must fit both an allowed session and their own listed lab hours", () => {
  const morning = course("MORNING", [student("S1")], ["101"], [{ days: ["Friday"], startMinutes: 480, endMinutes: 590 }]);
  const lateMorning = course("LATE", [student("S2")], ["101"], [{ days: ["Friday"], startMinutes: 600, endMinutes: 710 }]);
  assert.equal(draft([morning]).placed[0].slotId, "09:00");
  assert.equal(draft([lateMorning]).placed[0].slotId, "10:30");
  for (const [startMinutes, endMinutes] of [[720, 830], [900, 1010], [600, 660]]) {
    const c = course("LAB", [student("S1")], ["101"], [{ days: ["Friday"], startMinutes, endMinutes }]);
    const result = draft([c]);
    assert.equal(result.placed.length, 0);
    assert.match(result.unplaced[0].reason, /Friday exams must use 09:00-10:00 or 10:30-11:30/);
  }
  assert.equal(draft([lateMorning], { settings: { ...settings, examDurationMinutes: 120 } }).placed.length, 0);
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
  c.labSessions.push({ days: ["Monday"], startMinutes: 720, endMinutes: 780, campus: "Campus", building: "Building", room: "Lab", labInstructorId: "I0" });
  const short = draft([c], { staffCount: 2 });
  assert.equal(short.placed.length, 0);
  assert.match(short.unplaced[0].reason, /Slot backup not assigned|Invigilators not assigned/);
  const asd = course("ASD", Array.from({ length: 100 }, (_, i) => student(`OTHER${i}`)));
  const enough = draft([c], { staffCount: 3, asdAssignments: { 1: { Monday: { "12:00": ["ASD"] } } }, courseLookup: { MAIN: c, ASD: asd } });
  assert.equal(enough.placed.length, 1);
});

test("auto-scheduling staffing uses the last-room saving while keeping slot backups separate", () => {
  const large = course("LARGE", Array.from({ length: 55 }, (_, i) => student(`S${i}`)));
  const short = draft([large], { staffCount: 5 });
  assert.equal(short.placed.length, 0, "20/20/15 needs five room invigilators plus one backup");
  assert.match(short.unplaced[0].reason, /Slot backup not assigned|Invigilators not assigned/);
  assert.equal(draft([large], { staffCount: 6 }).placed.length, 1);
  const small = course("SMALL", Array.from({ length: 26 }, (_, i) => student(`S${i}`)));
  assert.equal(draft([small], { staffCount: 3 }).placed.length, 1, "13/13 needs only two room invigilators plus one backup");
});

test("automatic staffing respects only explicitly approved room consolidation", () => {
  const c = course("MAIN", Array.from({ length: 52 }, (_, i) => student(`S${i}`)));
  assert.equal(draft([c], { staffCount: 5 }).placed.length, 0);
  assert.equal(draft([c], { staffCount: 5, roomDistributionChoices: { MAIN: "distribute" } }).placed.length, 1);
  assert.equal(draft([c], { staffCount: 5, settings: { ...settings, studentsPerRoom: 20 }, roomDistributionChoices: { MAIN: "distribute" } }).placed.length, 0, "A lower configured limit cannot be overridden");
});

test("simultaneous courses have separate rooms and invigilators", () => {
  const c = course("NEW", [student("NEW_STUDENT")], ["101"], [{ days: ["Monday"], startMinutes: 720, endMinutes: 780 }]);
  const existing = course("EXISTING", [student("OTHER")]);
  const result = draft([c], { assignments: { 1: { Monday: { "12:00": ["EXISTING"] } } }, courseLookup: { NEW: c, EXISTING: existing }, staffCount: 3 });
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
    staffCount: 4,
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

test("resource-aware placement checks rooms and staff for the entire exam rather than pool totals", () => {
  const c = course("MAIN", Array.from({ length: 16 }, (_, i) => student("S" + i)), ["101", "102"]);
  const catalog = resourceFixture([c], 4, 1);
  catalog.rooms[0].busy = [{ code: "ROOM-CLASS", crn: "900", days: ["Monday"], start: 750, end: 780, isLab: false }];
  catalog.invigilators[0].busy = [{ code: "STAFF-CLASS", crn: "901", days: ["Monday"], start: 750, end: 780, isLab: false }];
  const before = JSON.stringify(catalog);
  const result = draft([c], { catalog, settings: { ...settings, invigilatorCount: 999 } });
  assert.deepEqual(result.placed.map((exam) => [exam.week, exam.day, exam.slotId]), [[1, "Tuesday", "12:00"]]);
  const sessions = buildExamSessions(result.assignments, { MAIN: c }, 60);
  assert.equal(validateResourcePlan(sessions, catalog, result.resourcePlan).complete, true);
  assert.deepEqual(standbyAvailability(sessions[0], sessions, catalog, result.resourcePlan), { count: 2, assigned: 1, unassigned: 1 });
  assert.equal(result.resourcePlan.backups[sessions[0].id].length, 1, "Two-person capacity cannot add a second report backup");
  assert.deepEqual(result.warnings, []);
  assert.equal(JSON.stringify(catalog), before);
});

test("the preferred two-person standby capacity can defer a slot; a one-backup fallback is explicit", () => {
  const c = course("MAIN", [student("S1")], ["101", "102"]);
  const catalog = resourceFixture([c], 3, 1);
  catalog.invigilators[2].busy = [{ code: "CLASS", crn: "901", days: ["Monday"], start: 720, end: 780, isLab: false }];
  const later = draft([c], { catalog });
  assert.equal(later.placed[0].slotId, "12:00");
  assert.equal(later.placed[0].day, "Tuesday");
  assert.deepEqual(later.warnings, []);
  catalog.invigilators[2].enabled = false;
  const fallback = draft([c], { catalog, weeks: [2, 1] });
  assert.equal(fallback.placed[0].week, 1);
  assert.equal(fallback.placed[0].day, "Monday");
  assert.equal(fallback.placed[0].slotId, "12:00");
  assert.equal(fallback.warnings.length, 1);
  assert.match(fallback.warnings[0].message, /Week 1 \/ Monday \/ 12:00-13:00: only 1 available standby/);
  const sessions = buildExamSessions(fallback.assignments, { MAIN: c }, 60);
  assert.equal(validateResourcePlan(sessions, catalog, fallback.resourcePlan).complete, true);
});

test("valid daytime backup coverage outranks gaining a second standby person in the evening", () => {
  const c = course("MAIN", [student("S1")], ["101", "102"]);
  const catalog = resourceFixture([c], 3, 1);
  catalog.invigilators[2].busy = [
    { code: "NOON", crn: "900", days: ["Monday", "Tuesday", "Wednesday", "Thursday"], start: 720, end: 780, isLab: false },
    { code: "FRIDAY", crn: "901", days: ["Friday"], start: 540, end: 690, isLab: false },
  ];
  const result = draft([c], { catalog });
  assert.deepEqual(result.placed.map((exam) => [exam.day, exam.slotId]), [["Monday", "12:00"]]);
  assert.equal(result.warnings.length, 1);
  const sessions = buildExamSessions(result.assignments, { MAIN: c }, 60);
  assert.equal(validateResourcePlan(sessions, catalog, result.resourcePlan).complete, true);
  assert.deepEqual(standbyAvailability(sessions[0], sessions, catalog, result.resourcePlan), { count: 1, assigned: 1, unassigned: 0 });
});

test("excluded resources and unspecified class times affect placement before generation", () => {
  const c = course("MAIN", [student("S1")], ["101", "102"]);
  const catalog = resourceFixture([c], 4, 2);
  catalog.rooms[1].enabled = false;
  catalog.invigilators[3].enabled = false;
  const unknown = { code: "UNKNOWN", crn: "901", days: ["Monday"], start: 0, end: 1440, isLab: false, unknownTime: true };
  catalog.rooms[0].busy = [unknown];
  catalog.invigilators[0].busy = [unknown];
  const result = draft([c], { catalog });
  assert.equal(result.placed[0].day, "Tuesday");
  assert.equal(result.placed[0].slotId, "12:00");
  assert.ok(Object.values(result.resourcePlan.allocations).every((allocation) =>
    allocation.roomId !== "R1" && !allocation.invigilatorIds.includes("I3")));
  assert.ok(Object.values(result.resourcePlan.backups).every((ids) => !ids.includes("I3")));
  catalog.rooms[0].enabled = false;
  const noRooms = draft([c], { catalog });
  assert.equal(noRooms.placed.length, 0);
  assert.match(noRooms.unplaced[0].reason, /Room not assigned/);
  assert.equal(assignmentIds(noRooms.assignments).size, 0);
});

test("single-CRN exams retain the lab room and teacher and skip unrelated teaching conflicts", () => {
  const lab = { days: ["Monday", "Wednesday"], crn: "101", startMinutes: 480, endMinutes: 590 };
  const c = course("LAB", Array.from({ length: 16 }, (_, i) => student("S" + i)), ["101"], [lab]);
  const catalog = resourceFixture([c], 4, 0);
  const ownLab = { code: "LAB", crn: "101", days: ["Monday", "Wednesday"], start: 480, end: 590, isLab: true };
  catalog.rooms[0].busy = [ownLab];
  catalog.invigilators[0].busy = [ownLab,
    { code: "OTHER-LECTURE", crn: "902", days: ["Monday"], start: 570, end: 600, isLab: false }];
  const result = draft([c], { catalog });
  assert.equal(result.placed[0].day, "Wednesday");
  assert.equal(result.placed[0].slotId, "09:00");
  const sessions = buildExamSessions(result.assignments, { LAB: c }, 60);
  const allocation = result.resourcePlan.allocations[sessions[0].rooms[0].id];
  assert.equal(allocation.roomId, catalog.rooms[0].id);
  assert.equal(allocation.invigilatorIds[0], "I0");
  assert.equal(allocation.invigilatorIds.length, 2);
  assert.equal(validateResourcePlan(sessions, catalog, result.resourcePlan).complete, true);
  catalog.invigilators[0].enabled = false;
  const excluded = draft([c], { catalog });
  assert.equal(excluded.placed.length, 0, "A different available teacher cannot replace the required lab teacher");
  assert.match(excluded.unplaced[0].reason, /Lab instructor unavailable/);
});

test("overlapping start times reserve distinct people and retain two standby people for each slot", () => {
  const c = course("MAIN", [student("S1")], ["101"], [{ days: ["Monday"], startMinutes: 930, endMinutes: 990 }]);
  const lookup = { MAIN: c, EARLY: course("EARLY", [student("S2")]), LATE: course("LATE", [student("S3")]) };
  const assignments = { 1: { Monday: { "15:00": ["EARLY"], "16:00": ["LATE"] } } };
  const catalog = resourceFixture(Object.values(lookup), 5, 2);
  const result = draft([c], { assignments, courseLookup: lookup, catalog });
  assert.equal(result.placed[0].slotId, "15:30");
  const sessions = buildExamSessions(result.assignments, lookup, 60);
  assert.equal(validateResourcePlan(sessions, catalog, result.resourcePlan).complete, true);
  assert.ok(sessions.every((session) => standbyAvailability(session, sessions, catalog, result.resourcePlan).count >= 2));
  const tooFew = draft([c], { assignments, courseLookup: lookup, catalog: resourceFixture(Object.values(lookup), 3, 2) });
  assert.equal(tooFew.placed.length, 0, "Existing backups at other overlapping starts cannot be reused");
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
  const result = draft(selected, { courseLookup, asdAssignments, asdExamDurations: plan.examDurations, weeks: plan.weeks, catalog: parseResourceCatalog(catalog) });
  assert.equal(result.placed.length + result.unplaced.length, selected.length);
  assert.equal(result.placed.length, 26);
  assert.deepEqual(result.unplaced, []);
  const sessions = buildExamSessions(result.assignments, courseLookup, 60);
  const resources = parseResourceCatalog(catalog);
  assert.equal(validateResourcePlan(sessions, resources, result.resourcePlan).complete, true);
  assert.ok(sessions.every((session) => standbyAvailability(session, sessions, resources, result.resourcePlan).count >= 2));
  assert.deepEqual(result.warnings, []);
  for (const placed of result.placed) {
    const c = courseLookup[placed.courseId];
    if (c.crns.length === 1 && c.labSessions.length) {
      assert.ok(c.labSessions.some((lab) => lab.days.includes(placed.day) && placed.start >= lab.startMinutes &&
        placed.end <= (lab.startMinutes < 540 && lab.endMinutes === 590 ? 600 : lab.endMinutes)));
      if (c.labSessions.every((lab) => lab.startMinutes < 540)) {
        assert.equal(placed.slotId, "09:00");
        assert.equal(placed.end, 600);
      }
    } else assert.ok(placed.day === "Friday" ? [540, 630].includes(placed.start) : [720, 1020].includes(placed.start));
    assert.ok(defaultHasExam(c));
    for (const other of [...result.placed, ...plan.placements.map((p) => ({ ...p, start: p.startMinutes, end: p.endMinutes }))]) {
      if (other.courseId === placed.courseId || other.week !== placed.week || other.day !== placed.day || placed.start >= other.end || other.start >= placed.end) continue;
      const otherIds = new Set(courseLookup[other.courseId].students.map((s) => s.id));
      assert.ok(c.students.every((s) => !otherIds.has(s.id)), `Overlap ${c.code} and ${other.courseId}`);
    }
  }
});
