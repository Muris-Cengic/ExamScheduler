import test from "node:test";
import assert from "node:assert/strict";
import * as XLSX from "xlsx/xlsx.mjs";
import { parseDepartmentWorkbook, roomIdentity } from "../src/department.js";
import { assignResources, backupTarget, buildExamSessions, emptyResourcePlan, invigilatorWorkloads, isTeachingTimeDuty, parseResourceCatalog, readResourceCatalog, readResourcePlan, reconcileRoomDistribution, resourceBusyReason, resourceFingerprint, validateResourcePlan } from "../src/resources.js";
import { buildResourceWorkbookForWeek } from "../src/reports.js";

const course = (id, count, options = {}) => ({
  id, code: id, title: id, crns: ["101"], labSessions: [], primaryInstructor: "Lecturer",
  students: Array.from({ length: count }, (_, index) => ({ id: `${id}-S${index}`, name: `${id} Student ${String(index).padStart(3, "0")}`, crn: "101" })), ...options,
});
const catalog = (rooms = 3, staff = 8) => ({
  rooms: Array.from({ length: rooms }, (_, i) => ({ id: `R${i}`, name: `Building / ${i}`, enabled: true, busy: [] })),
  invigilators: Array.from({ length: staff }, (_, i) => ({ id: `I${i}`, name: `Staff ${i}`, enabled: true, busy: [] })),
});
const sessionsFor = (c, time = "12:00", duration = 60) => buildExamSessions({ 1: { Monday: { [time]: [c.id] } } }, { [c.id]: c }, duration);
const classTime = (code, start, end, options = {}) => ({ code, crn: "101", days: ["Monday"], start, end, isLab: false, ...options });
const headers = ["Campus", "Crn No", "Course Code", "Title", "Cr", "Maximum Load", "No Of Enrolled", "Session Id", "DAYS", "Time", "Primary Instructor", "Second Instructor", "Type", "Building", "Room"];
const meetingRow = (code, time, type = "LEC", room = "1", crn = "101") => ["Campus", crn, code, code, 3, 25, 20, "01", "M", time, "1: Lecturer", "2: Lab Teacher", type, "Building", room];

test("resource pools and teaching availability use only the first CRN sheet", () => {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    meetingRow("ICT-1000", "0800 - 0950", "OL"), meetingRow("ICT-1000", "1200 - 1250"),
  ]), "ISET CRNs");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([headers,
    meetingRow("ICT-1000", "0800 - 0950", "OL"), meetingRow("ICT-1000", "1200 - 1250"),
    meetingRow("OTHER-1000", "1500 - 1650", "LEC", "9", "999"),
    meetingRow("CSTP-1011", "0800 - 1150", "OL", "1", "888"),
    meetingRow("INVALID-1000", "2500 - 2600", "LEC", "9", "777"),
  ]), "All CRNs");
  const selection = parseDepartmentWorkbook(workbook);
  assert.equal(selection.courses[0].meetings[0].labInstructorId, "staff:2");
  assert.equal(selection.courses[0].meetings[0].instructor, "Lab Teacher");
  const resources = parseResourceCatalog(selection);
  assert.equal(resources.rooms.length, 1);
  assert.equal(resources.invigilators.length, 2);
  const lecturer = resources.invigilators.find((person) => person.id === "staff:1");
  const labTeacher = resources.invigilators.find((person) => person.id === "staff:2");
  assert.deepEqual(lecturer.busy.map((entry) => entry.start), [720]);
  assert.deepEqual(labTeacher.busy.map((entry) => entry.start), [480]);
  assert.deepEqual(resources.rooms[0].busy.map((entry) => entry.start).sort((a, b) => a - b), [480, 720]);
  assert.ok([...resources.rooms, ...resources.invigilators].every((resource) => resource.busy.every((entry) => entry.code === "ICT-1000")));
  const labExam = sessionsFor(course("ICT-1000", 10, { labSessions: selection.courses[0].meetings.filter((meeting) => meeting.isLab) }), "08:00");
  assert.equal(resourceBusyReason(labTeacher, labExam[0], labExam), "", "CSTP on a later sheet must not block the replaced lab instructor");
  assert.equal(resourceBusyReason(resources.rooms[0], labExam[0], labExam), "", "A later sheet must not block the original lab room");
  assert.equal(resourceBusyReason(lecturer, { week: 1, day: "Monday", start: 900, end: 960 }, []), "");
  assert.ok(!resources.invigilators.some((person) => "type" in person));
  assert.deepEqual(readResourceCatalog(JSON.parse(JSON.stringify(resources))), resources);
  const exam = sessionsFor(course("EXAM", 16), "17:00");
  const plan = assignResources(exam, resources);
  assert.deepEqual(new Set(plan.allocations[exam[0].rooms[0].id].invigilatorIds), new Set([lecturer.id, labTeacher.id]), "Either course instructor can invigilate outside their classes");
});

test("rooms are distinguished by campus/building and unspecified class times block the named days", () => {
  assert.notEqual(roomIdentity("Campus", "Building1", "1"), roomIdentity("Campus", "Building2", "1"));
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([headers, meetingRow("ICT-1000", "-")]), "ISET CRNs");
  const resources = parseResourceCatalog(parseDepartmentWorkbook(workbook));
  assert.match(resourceBusyReason(resources.rooms[0], { week: 1, day: "Monday", start: 720, end: 780 }, []), /time unspecified/);
  assert.equal(resourceBusyReason(resources.rooms[0], { week: 2, day: "Tuesday", start: 720, end: 780 }, []), "");
});

test("exam rosters are balanced under the room limit and staffing follows each room's actual count", () => {
  for (const [count, roomSizes, invigilators] of [[15, [15], [1]], [16, [16], [2]], [25, [25], [2]], [26, [13, 13], [1, 1]], [31, [16, 15], [2, 1]], [55, [19, 18, 18], [2, 2, 2]], [60, [20, 20, 20], [2, 2, 2]]]) {
    const sessions = sessionsFor(course("EXAM", count));
    assert.deepEqual(sessions[0].rooms.map((room) => room.students.length), roomSizes);
    assert.deepEqual(sessions[0].rooms.map((room) => room.requiredInvigilators), invigilators);
    assert.equal(new Set(sessions[0].rooms.flatMap((room) => room.students.map((student) => student.id))).size, count);
    const resources = catalog();
    assert.equal(validateResourcePlan(sessions, resources, assignResources(sessions, resources)).complete, true);
  }
  const c = course("LIMIT", 30);
  const limited = buildExamSessions({ 1: { Monday: { "12:00": [c.id] } } }, { LIMIT: c }, 60, 500);
  assert.deepEqual(limited[0].rooms.map((room) => room.students.length), [15, 15]);
  const lower = buildExamSessions({ 1: { Monday: { "12:00": [c.id] } } }, { LIMIT: c }, 60, 12);
  assert.deepEqual(lower[0].rooms.map((room) => room.students.length), [10, 10, 10]);
});

test("balanced room rosters keep courses separate and retain each enrolled student exactly once", () => {
  const a = course("A", 55);
  const b = course("B", 26);
  a.students.push({ ...a.students[0] });
  const sessions = buildExamSessions({ 1: { Monday: { "12:00": ["A", "B"] } } }, { A: a, B: b }, 60);
  for (const [c, sizes] of [[a, [19, 18, 18]], [b, [13, 13]]]) {
    const rooms = sessions[0].rooms.filter((room) => room.courseId === c.id);
    assert.deepEqual(rooms.map((room) => room.students.length), sizes);
    const roster = rooms.flatMap((room) => room.students);
    assert.deepEqual(new Set(roster.map((student) => student.id)), new Set(c.students.map((student) => student.id)));
    assert.equal(new Set(roster.map((student) => student.id)).size, roster.length);
    assert.ok(roster.every((student) => student.id.startsWith(`${c.id}-`)));
  }
});

test("single-CRN lab exams keep their room and second instructor; extra rooms/staff cover overflow", () => {
  const labRoom = roomIdentity("Campus", "Building", "Lab");
  const lab = { days: ["Monday"], crn: "101", startMinutes: 480, endMinutes: 590, room: "Lab", building: "Building", campus: "Campus", labInstructorId: "I2" };
  const c = course("LAB", 26, { labSessions: [lab] });
  const sessions = sessionsFor(c, "08:00");
  const resources = catalog(2, 5);
  resources.rooms[0].id = labRoom;
  const busyLab = classTime("LAB", 480, 590, { isLab: true });
  resources.rooms[0].busy = [busyLab];
  resources.invigilators[2].busy = [busyLab];
  const plan = assignResources(sessions, resources);
  assert.equal(plan.allocations[sessions[0].rooms[0].id].roomId, labRoom);
  assert.equal(plan.allocations[sessions[0].rooms[0].id].invigilatorIds[0], "I2");
  assert.deepEqual(sessions[0].rooms.map((room) => room.students.length), [13, 13]);
  assert.equal(plan.allocations[sessions[0].rooms[0].id].invigilatorIds.length, 1);
  assert.notEqual(plan.allocations[sessions[0].rooms[1].id].roomId, labRoom);
  const labTeacherLoad = invigilatorWorkloads(resources, sessions, plan).find((person) => person.id === "I2");
  assert.equal(labTeacherLoad.teaching, 1);
  assert.equal(labTeacherLoad.exam, 0);
  assert.deepEqual(labTeacherLoad.bySlot, {});
  assert.equal(invigilatorWorkloads(resources, sessions, plan).reduce((sum, person) => sum + person.exam, 0), 1, "Only the overflow room's invigilator carries extra load");
  assert.equal(validateResourcePlan(sessions, resources, plan).complete, true);
  assert.equal(resourceBusyReason(resources.rooms[0], { ...sessions[0], week: 2 }, sessions), "Class LAB (CRN 101) on Monday, 08:00-09:50.");
  const outsideLab = sessionsFor(c, "12:00");
  assert.match(outsideLab[0].issues[0], /move this single-CRN/);
  const larger = sessionsFor(course("LAB", 32, { labSessions: [lab] }), "08:00");
  const largerPlan = assignResources(larger, resources);
  assert.deepEqual(larger[0].rooms.map((room) => room.students.length), [16, 16]);
  assert.ok(larger[0].rooms.every((room) => largerPlan.allocations[room.id].invigilatorIds.filter(Boolean).length === 2));
  assert.equal(largerPlan.allocations[larger[0].rooms[0].id].invigilatorIds[0], "I2");
  assert.equal(validateResourcePlan(larger, resources, largerPlan).complete, true);
});

test("09:00-10:00 exams retain the 09:50 lab's room and instructor without extra teaching-time load", () => {
  const labRoom = roomIdentity("Campus", "Building", "Lab");
  const lab = { days: ["Monday"], crn: "101", startMinutes: 480, endMinutes: 590, room: "Lab", building: "Building", campus: "Campus", labInstructorId: "I0" };
  const c = course("LAB", 16, { labSessions: [lab] });
  const sessions = sessionsFor(c, "09:00");
  assert.deepEqual(sessions[0].issues, []);
  assert.deepEqual(sessions[0].replacedLabs, [{ code: "LAB", crn: "101", start: 480, end: 590 }], "Keep the original lab identity for availability matching");
  const resources = catalog(2, 4);
  resources.rooms[0].id = labRoom;
  resources.rooms[0].busy = [classTime("LAB", 480, 590, { isLab: true })];
  resources.invigilators[0].busy = [classTime("LAB", 480, 590, { isLab: true })];
  const plan = assignResources(sessions, resources);
  const allocation = plan.allocations[sessions[0].rooms[0].id];
  assert.equal(allocation.roomId, labRoom);
  assert.equal(allocation.invigilatorIds[0], "I0");
  assert.equal(allocation.invigilatorIds.length, 2);
  assert.equal(validateResourcePlan(sessions, resources, plan).complete, true);
  const load = invigilatorWorkloads(resources, sessions, plan).find((person) => person.id === "I0");
  assert.equal(load.teaching, 1);
  assert.equal(load.exam, 0);
  assert.deepEqual(load.bySlot, {});

  // The allowance never overrides another class in the final ten minutes.
  resources.invigilators[0].busy.push(classTime("OTHER", 590, 600));
  assert.match(resourceBusyReason(resources.invigilators[0], sessions[0], sessions), /Class OTHER.*09:50-10:00/);
  const blockedTeacher = assignResources(sessions, resources);
  assert.equal(blockedTeacher.allocations[sessions[0].rooms[0].id].invigilatorIds[0], "");
  assert.equal(validateResourcePlan(sessions, resources, blockedTeacher).complete, false);
  resources.invigilators[0].busy.pop();
  resources.rooms[0].busy.push(classTime("OTHER", 590, 600));
  assert.match(resourceBusyReason(resources.rooms[0], sessions[0], sessions), /Class OTHER.*09:50-10:00/);
  const blockedRoom = assignResources(sessions, resources);
  assert.equal(blockedRoom.allocations[sessions[0].rooms[0].id].roomId, "");
  assert.equal(validateResourcePlan(sessions, resources, blockedRoom).complete, false);

  assert.match(sessionsFor({ ...c, labSessions: [{ ...lab, endMinutes: 580 }] }, "09:00")[0].issues[0], /move this single-CRN/, "A shorter lab cannot use the allowance");
  assert.match(sessionsFor(c, "09:00", 90)[0].issues[0], /move this single-CRN/, "The allowance cannot extend an exam beyond 10:00");
});

test("replacing a lab never cancels an unrelated class or the same course's lecture", () => {
  const c = course("LAB", 10, { labSessions: [{ days: ["Monday"], crn: "101", startMinutes: 480, endMinutes: 590, room: "Lab", building: "Building", campus: "Campus", labInstructorId: "I0" }] });
  const sessions = sessionsFor(c, "08:00");
  const resources = catalog();
  resources.invigilators[0].busy = [classTime("LAB", 480, 590, { isLab: true }), classTime("OTHER", 480, 590)];
  const plan = assignResources(sessions, resources);
  assert.equal(plan.allocations[sessions[0].rooms[0].id].invigilatorIds[0], "");
  assert.ok(validateResourcePlan(sessions, resources, plan).issues.some((issue) => /Class OTHER/.test(issue.message)));
  resources.invigilators[0].busy = [classTime("LAB", 480, 590)];
  assert.match(resourceBusyReason(resources.invigilators[0], sessions[0], sessions), /Class LAB/);
});

test("lab conflicts name the instructor, class and CRN with human-readable exam context and a resolution", () => {
  const labRoom = roomIdentity("Campus", "Building", "Lab");
  const lab = { days: ["Wednesday"], crn: "101", startMinutes: 480, endMinutes: 590, room: "Lab", building: "Building", campus: "Campus", labInstructorId: "I2" };
  const c = course("ISET-4001", 16, { labSessions: [lab] });
  const sessions = buildExamSessions({ 1: { Wednesday: { "08:00": [c.id] } } }, { [c.id]: c }, 60);
  const resources = catalog(1, 3);
  resources.rooms[0].id = labRoom;
  resources.invigilators[2].name = "Lab Teacher";
  resources.invigilators[2].busy = [
    classTime("ISET-4001", 480, 590, { days: ["Wednesday"], isLab: true }),
    classTime("CSTP-1011", 480, 710, { days: ["Wednesday"], crn: "999" }),
  ];
  const validation = validateResourcePlan(sessions, resources, assignResources(sessions, resources));
  assert.equal(validation.complete, false);
  assert.equal(validation.issues.length, 1, "The fixed-instructor conflict must not also produce a generic missing-invigilator warning");
  const issue = validation.issues[0];
  assert.equal(issue.title, "Lab instructor unavailable");
  assert.equal(issue.context, "Week 1 / Wednesday / 08:00-09:00");
  assert.equal(issue.exam, "ISET-4001 / Exam room 1 of 1 / 16 students");
  assert.match(issue.message, /Lab Teacher.*CSTP-1011.*CRN 999.*Wednesday, 08:00-11:50/);
  assert.match(issue.action, /CRN list.*another listed lab time.*Only this exam's lab is replaced/);
  assert.ok(!issue.message.includes("first invigilator"));
  // Assigning the unavailable instructor manually must still report the same named conflict.
  const plan = assignResources(sessions, resources);
  plan.allocations[sessions[0].rooms[0].id].invigilatorIds[0] = "I2";
  assert.ok(validateResourcePlan(sessions, resources, plan).issues.some((entry) => entry.title === issue.title && entry.message === issue.message));
});

test("fixed resources excluded from the pool get inclusion instructions, not advice to change teaching times", () => {
  const labRoom = roomIdentity("Campus", "Building", "Lab");
  const lab = { days: ["Monday"], crn: "101", startMinutes: 480, endMinutes: 590, room: "Lab", building: "Building", campus: "Campus", labInstructorId: "I0" };
  const sessions = sessionsFor(course("LAB", 10, { labSessions: [lab] }), "08:00");
  const resources = catalog(1, 3);
  resources.rooms[0].id = labRoom;
  resources.rooms[0].enabled = false;
  resources.invigilators[0].enabled = false;
  const issues = validateResourcePlan(sessions, resources, assignResources(sessions, resources)).issues;
  for (const [title, name] of [["Original lab room unavailable", resources.rooms[0].name], ["Lab instructor unavailable", resources.invigilators[0].name]]) {
    const issue = issues.find((entry) => entry.title === title);
    assert.ok(issue);
    assert.ok(issue.message.includes(name));
    assert.ok(issue.action.includes(`Include ${name} in Review Resource Pool`));
    assert.ok(!issue.action.includes("another listed lab time"));
  }
  assert.ok(!issues.some((entry) => entry.title === "Room not assigned" || entry.title === "Invigilators not assigned"));
});

test("double-booking issues name the other exam or backup duty and do not confuse it with a regular class", () => {
  const courses = { A: course("A", 10), B: course("B", 10) };
  const sessions = buildExamSessions({ 1: { Monday: { "08:00": ["A"], "08:30": ["B"] } } }, courses, 60);
  const resources = catalog(2, 4);
  const plan = assignResources(sessions, resources);
  plan.allocations[sessions[1].rooms[0].id].invigilatorIds[0] = plan.backups[sessions[0].id][0];
  const issues = validateResourcePlan(sessions, resources, plan).issues;
  const examIssue = issues.find((issue) => issue.roomId === sessions[1].rooms[0].id);
  assert.match(examIssue.message, /Already assigned to slot backup duty on Monday, 08:00-09:00/);
  assert.match(examIssue.action, /Select an available invigilator/);
  const backupIssue = issues.find((issue) => issue.title === "Backup invigilator unavailable");
  assert.match(backupIssue.message, /Already assigned to the B exam/);
  assert.match(backupIssue.action, /Slot backups/);
});

test("missing extra staff and slot backups explain the student threshold and slot-level coverage", () => {
  const sessions = sessionsFor(course("EXAM", 16));
  const resources = catalog(1, 1);
  const issues = validateResourcePlan(sessions, resources, assignResources(sessions, resources)).issues;
  const staffIssue = issues.find((issue) => issue.title === "Invigilators not assigned");
  assert.match(staffIssue.message, /1 invigilator\(s\) still needed.*16 students require 2.*more than 15/);
  const backupIssue = issues.find((issue) => issue.title === "Slot backup not assigned");
  assert.match(backupIssue.message, /1 exam room.*0 backup.*at least 1 and at most 1.*not each room/);
  assert.match(backupIssue.action, /cannot also cover an exam room/);
  const unknown = catalog(1, 3);
  unknown.invigilators[0].busy = [classTime("UNKNOWN", 0, 1440, { unknownTime: true })];
  const plan = assignResources(sessions, unknown);
  plan.allocations[sessions[0].rooms[0].id].invigilatorIds[0] = "I0";
  const unknownIssue = validateResourcePlan(sessions, unknown, plan).issues.find((issue) => /time unspecified/.test(issue.message));
  assert.match(unknownIssue.action, /missing class time.*CRN list/);
});

test("availability covers the entire exam and repeats each week; adjacent bookings are allowed", () => {
  const c = course("EXAM", 10);
  const sessions = sessionsFor(c, "08:30");
  const resources = catalog(2, 4);
  resources.rooms[0].busy = [classTime("OTHER", 540, 600)];
  resources.invigilators[0].busy = [classTime("OTHER", 540, 600)];
  const plan = assignResources(sessions, resources);
  assert.notEqual(plan.allocations[sessions[0].rooms[0].id].roomId, "R0");
  assert.ok(!plan.allocations[sessions[0].rooms[0].id].invigilatorIds.includes("I0"));
  assert.ok(!plan.backups[sessions[0].id].includes("I0"));
  assert.equal(validateResourcePlan(sessions, resources, plan).complete, true);
  assert.match(resourceBusyReason(resources.rooms[0], { ...sessions[0], week: 2 }, sessions), /Class OTHER/);
  assert.equal(resourceBusyReason(resources.rooms[0], { ...sessions[0], start: 600, end: 660 }, sessions), "");
});

test("no room or staff double booking is allowed for different starts with overlapping exams", () => {
  const a = course("A", 10);
  const b = course("B", 10);
  const sessions = buildExamSessions({ 1: { Monday: { "08:00": ["A"], "08:30": ["B"] } } }, { A: a, B: b }, 60);
  const resources = catalog(2, 4);
  const plan = assignResources(sessions, resources);
  assert.equal(validateResourcePlan(sessions, resources, plan).complete, true);
  const first = plan.allocations[sessions[0].rooms[0].id];
  const second = plan.allocations[sessions[1].rooms[0].id];
  assert.notEqual(first.roomId, second.roomId);
  assert.equal(new Set([...first.invigilatorIds, ...second.invigilatorIds, ...Object.values(plan.backups).flat()]).size, 4);
  second.roomId = first.roomId;
  second.invigilatorIds[0] = first.invigilatorIds[0];
  assert.ok(validateResourcePlan(sessions, resources, plan).issues.some((issue) => /Already assigned/.test(issue.message)));
});

test("backups are slot-level, separate from room duties, and may be fewer than the maximum", () => {
  const sessions = sessionsFor(course("EXAM", 150));
  const resources = catalog(6, 14);
  const plan = assignResources(sessions, resources);
  assert.equal(backupTarget(sessions[0].rooms.length), 2);
  assert.equal(plan.backups[sessions[0].id].length, 2);
  assert.equal(validateResourcePlan(sessions, resources, plan).complete, true);
  plan.backups[sessions[0].id].pop();
  assert.equal(validateResourcePlan(sessions, resources, plan).complete, true);
  plan.backups[sessions[0].id][0] = plan.allocations[sessions[0].rooms[0].id].invigilatorIds[0];
  assert.ok(validateResourcePlan(sessions, resources, plan).issues.some((issue) => /Already assigned/.test(issue.message)));
});

test("exam and backup workloads are balanced independently over the whole timetable", () => {
  const c = course("EXAM", 10);
  const assignments = {};
  for (let week = 1; week <= 3; week += 1) {
    assignments[week] = Object.fromEntries(["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"].map((day) => [day, { "09:00": ["EXAM"], "12:00": ["EXAM"] }]));
  }
  const sessions = buildExamSessions(assignments, { EXAM: c }, 60);
  const resources = catalog(1, 4);
  const plan = assignResources(sessions, resources);
  assert.equal(validateResourcePlan(sessions, resources, plan).complete, true);
  const loads = invigilatorWorkloads(resources, sessions, plan);
  for (const role of ["exam", "backup"]) assert.ok(Math.max(...loads.map((load) => load[role])) - Math.min(...loads.map((load) => load[role])) <= 1);
  for (const slot of ["09:00", "12:00"]) assert.ok(Math.max(...loads.map((load) => load.bySlot[slot] || 0)) - Math.min(...loads.map((load) => load.bySlot[slot] || 0)) <= 1);
  for (const slot of ["09:00", "12:00"]) assert.ok(Math.max(...loads.map((load) => load.backupBySlot[slot] || 0)) - Math.min(...loads.map((load) => load.backupBySlot[slot] || 0)) <= 1);
});

test("replaced lab duties do not disadvantage the instructor when balancing extra invigilations", () => {
  const labRoom = roomIdentity("Campus", "Building", "Lab");
  const lab = { days: ["Monday"], crn: "101", startMinutes: 480, endMinutes: 590, room: "Lab", building: "Building", campus: "Campus", labInstructorId: "I0" };
  const courses = { LAB: course("LAB", 10, { labSessions: [lab] }), EXTRA: course("EXTRA", 10) };
  const sessions = buildExamSessions({ 1: { Monday: { "08:00": ["LAB"], "12:00": ["EXTRA"] } } }, courses, 60);
  const resources = catalog(1, 3);
  resources.rooms[0].id = labRoom;
  resources.rooms[0].busy = [classTime("LAB", 480, 590, { isLab: true })];
  resources.invigilators[0].busy = [classTime("LAB", 480, 590, { isLab: true })];
  const plan = assignResources(sessions, resources);
  assert.equal(validateResourcePlan(sessions, resources, plan).complete, true);
  assert.equal(plan.allocations[sessions[1].rooms[0].id].invigilatorIds[0], "I0", "The earlier lab duty must not increase the balancing counter");
  const load = invigilatorWorkloads(resources, sessions, plan).find((person) => person.id === "I0");
  assert.equal(load.teaching, 1);
  assert.equal(load.exam, 1);
  assert.deepEqual(load.bySlot, { "12:00": 1 });
});

test("teaching-hour exemption needs the same week's replaced class and the entire exam within its hours", () => {
  const lab = { days: ["Monday"], crn: "101", startMinutes: 480, endMinutes: 590, room: "Lab", building: "Building", campus: "Campus", labInstructorId: "I0" };
  const sessions = sessionsFor(course("LAB", 10, { labSessions: [lab] }), "08:00");
  const teacher = catalog().invigilators[0];
  teacher.busy = [classTime("LAB", 480, 590, { isLab: true })];
  assert.equal(isTeachingTimeDuty(teacher, sessions[0], sessions), true);
  assert.equal(isTeachingTimeDuty(teacher, { ...sessions[0], start: 540, end: 590 }, sessions), true);
  assert.equal(isTeachingTimeDuty(teacher, { ...sessions[0], start: 540, end: 600 }, sessions), true, "The morning lab's 09:50 ending includes its exam allowance through 10:00");
  assert.equal(isTeachingTimeDuty(teacher, { ...sessions[0], start: 540, end: 610 }, sessions), false);
  assert.equal(isTeachingTimeDuty(teacher, { ...sessions[0], week: 2 }, sessions), false);
  assert.equal(isTeachingTimeDuty(teacher, { ...sessions[0], day: "Tuesday" }, sessions), false);
  teacher.busy = [classTime("LAB", 480, 590)];
  assert.equal(isTeachingTimeDuty(teacher, sessions[0], sessions), false, "An uncancelled lecture is not a teaching-hour exam duty");
  teacher.busy = [classTime("OTHER", 480, 590, { isLab: true })];
  assert.equal(isTeachingTimeDuty(teacher, sessions[0], sessions), false);
  teacher.busy = [classTime("LAB", 0, 1440, { isLab: true, unknownTime: true })];
  assert.equal(isTeachingTimeDuty(teacher, sessions[0], sessions), false);
});

test("teaching-hour backup duties do not add backup load but still prevent double booking", () => {
  const labRoom = roomIdentity("Campus", "Building", "Lab");
  const lab = { days: ["Monday"], crn: "101", startMinutes: 480, endMinutes: 600, room: "Lab", building: "Building", campus: "Campus", labInstructorId: "I0" };
  const courses = { LAB: course("LAB", 10, { labSessions: [lab] }), EXTRA: course("EXTRA", 10) };
  const sessions = buildExamSessions({ 1: { Monday: { "08:00": ["LAB"], "09:00": ["EXTRA"] } } }, courses, 60);
  const resources = catalog(1, 3);
  resources.rooms[0].id = labRoom;
  resources.invigilators[0].busy = [classTime("LAB", 480, 600, { isLab: true })];
  const plan = assignResources(sessions, resources);
  plan.allocations[sessions[1].rooms[0].id].invigilatorIds = ["I1"];
  plan.backups[sessions[1].id] = ["I0"];
  assert.equal(validateResourcePlan(sessions, resources, plan).complete, true);
  const load = invigilatorWorkloads(resources, sessions, plan).find((person) => person.id === "I0");
  assert.equal(load.teaching, 2);
  assert.equal(load.exam, 0);
  assert.equal(load.backup, 0);
  assert.deepEqual(load.backupBySlot, {});
  plan.backups[sessions[0].id] = ["I0"];
  assert.ok(validateResourcePlan(sessions, resources, plan).issues.some((issue) => /Already assigned/.test(issue.message)));
});

test("incomplete, disabled or stale allocations block export and snapshots validate their resource shape", () => {
  const sessions = sessionsFor(course("EXAM", 25));
  const resources = catalog(1, 3);
  assert.equal(validateResourcePlan(sessions, resources, emptyResourcePlan(sessions, resources)).complete, false);
  const plan = assignResources(sessions, resources);
  assert.deepEqual(readResourcePlan(JSON.parse(JSON.stringify(plan))), plan);
  resources.invigilators[0].enabled = false;
  assert.ok(validateResourcePlan(sessions, resources, plan).issues.some((issue) => /outdated/.test(issue.message)));
  assert.ok(validateResourcePlan(sessions, resources, plan).issues.some((issue) => /Not in the active/.test(issue.message)));
  assert.throws(() => buildResourceWorkbookForWeek({ week: 1, sessions, catalog: resources, plan }), /Complete valid/);
  assert.throws(() => readResourcePlan({ fingerprint: "x", allocations: { x: { roomId: "R", invigilatorIds: "not an array" } }, backups: {} }), /Invalid resource/);
  assert.throws(() => readResourceCatalog({ rooms: [{}], invigilators: [] }), /Invalid resource/);
});

test("reports use reviewed resources and per-room student references, without an invigilator type", () => {
  const sessions = sessionsFor(course("EXAM", 60));
  const resources = catalog();
  const plan = assignResources(sessions, resources);
  const workbook = buildResourceWorkbookForWeek({ week: 1, sessions, catalog: resources, plan, startDate: "2026-10-19", templateHeaders: {
    invigilatorHeader: ["CRN", "Course Code", "Course Title", "No of Students", "Date", "Time", "Instructor Name", "Invigilator room", "Invigilator1", "Invigilator2", "Backup Invigilator"],
    studentHeader: ["CRN", "Code", "Title", "Student ID", "Student Name", "Class room", "Present/ Absent"],
  } });
  const inv = workbook.Sheets["Week 1 Invigilators"];
  assert.deepEqual([inv.D2.v, inv.D3.v, inv.D4.v], [20, 20, 20]);
  assert.ok(inv.I2.v && inv.J2.v && inv.I3.v && inv.J3.v && inv.I4.v);
  assert.ok(inv.J4.v);
  assert.equal(XLSX.utils.sheet_to_json(inv, { header: 1 }).slice(1).filter((row) => row[10]).length, 1);
  const daily = workbook.Sheets["Week 1 Monday"];
  assert.equal(daily.F2.f, "'Week 1 Invigilators'!H2");
  assert.equal(daily.F21.f, "'Week 1 Invigilators'!H2");
  assert.equal(daily.F22.f, "'Week 1 Invigilators'!H3");
  assert.equal(daily.F41.f, "'Week 1 Invigilators'!H3");
  assert.equal(daily.F42.f, "'Week 1 Invigilators'!H4");
  assert.ok(!XLSX.utils.sheet_to_json(workbook.Sheets["Week 1 Invigilator Pool"], { header: 1 })[0].includes("Type"));
  const saved = XLSX.read(XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }), { type: "buffer" });
  assert.equal(saved.Sheets["Week 1 Monday"].F22.f, "'Week 1 Invigilators'!H3");
});

test("reports exclude teaching-hour duties from cached loads and recalculation formulas", () => {
  const labRoom = roomIdentity("Campus", "Building", "Lab");
  const lab = { days: ["Monday"], crn: "101", startMinutes: 480, endMinutes: 600, room: "Lab", building: "Building", campus: "Campus", labInstructorId: "I0" };
  const courses = { LAB: course("LAB", 16, { labSessions: [lab] }), EXTRA: course("EXTRA", 10), LATE: course("LATE", 10) };
  const sessions = buildExamSessions({ 1: { Monday: { "08:00": ["LAB"], "09:00": ["EXTRA"], "12:00": ["LATE"] } } }, courses, 60);
  const resources = catalog(1, 3);
  resources.rooms[0].id = labRoom;
  resources.invigilators[0].busy = [classTime("LAB", 480, 600, { isLab: true })];
  const plan = assignResources(sessions, resources);
  plan.allocations[sessions[1].rooms[0].id].invigilatorIds = ["I1"];
  plan.backups[sessions[1].id] = ["I0"];
  plan.allocations[sessions[2].rooms[0].id].invigilatorIds = ["I0"];
  plan.backups[sessions[2].id] = ["I1"];
  const workbook = buildResourceWorkbookForWeek({ week: 1, sessions, catalog: resources, plan, startDate: "2026-10-19", templateHeaders: {
    invigilatorHeader: ["CRN", "Course Code", "Course Title", "No of Students", "Date", "Time", "Instructor Name", "Invigilator room", "Primary 1", "Primary 2", "Backup"],
    studentHeader: ["CRN", "Code", "Title", "Student ID", "Student Name", "Class room", "Present/ Absent"],
  } });
  const inv = workbook.Sheets["Week 1 Invigilators"];
  const pool = workbook.Sheets["Week 1 Invigilator Pool"];
  assert.equal(inv.I1.v, "Invigilator 1");
  assert.equal(inv.J1.v, "Invigilator 2");
  assert.equal(inv.L2.v, 0, "The lab instructor has no extra load");
  assert.equal(inv.M2.v, 1, "The additional instructor carries extra load");
  assert.equal(inv.N3.v, 0, "A backup within replaced teaching hours is also exempt");
  assert.equal(inv.L4.v, 1, "The same instructor's later exam adds load");
  assert.deepEqual([pool.B2.v, pool.C2.v, pool.D2.v, pool.E2.v], [1, 0, 2, 1]);
  assert.match(pool.B2.f, /SUMIF.*L:L.*SUMIF.*M:M/);
  assert.match(pool.C2.f, /SUMIF.*N:N/);
  assert.ok(inv["!cols"].slice(11, 14).every((column) => column.hidden));
  assert.ok(!inv["!cols"][14].hidden);
  const rows = XLSX.utils.sheet_to_json(inv, { header: 1 }).slice(1);
  for (let index = 0; index < resources.invigilators.length; index += 1) {
    const row = index + 2;
    const name = pool[`A${row}`].v;
    const examLoad = rows.reduce((sum, entry) => sum + (entry[8] === name ? entry[11] : 0) + (entry[9] === name ? entry[12] : 0), 0);
    const backupLoad = rows.reduce((sum, entry) => sum + (entry[10] === name ? entry[13] : 0), 0);
    assert.equal(pool[`B${row}`].v, examLoad);
    assert.equal(pool[`C${row}`].v, backupLoad);
  }
  const saved = XLSX.read(XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }), { type: "buffer", cellStyles: true });
  assert.equal(saved.Sheets["Week 1 Invigilators"].L2.v, 0);
  assert.equal(saved.Sheets["Week 1 Invigilator Pool"].D2.v, 2);
  assert.equal(saved.Sheets["Week 1 Invigilator Pool"].B2.f, pool.B2.f);
  assert.ok(saved.Sheets["Week 1 Invigilators"]["!cols"].slice(11, 14).every((column) => column.hidden));
});

test("approving consolidation preserves the remaining physical rooms, staff, backups and unrelated exams", () => {
  const lookup = { A: course("A", 26), B: course("B", 20) };
  const assignments = { 1: { Monday: { "12:00": ["A", "B"] } } };
  const original = buildExamSessions(assignments, lookup, 60);
  const resources = catalog(3, 5);
  const originalPlan = assignResources(original, resources);
  const aRooms = original[0].rooms.filter((room) => room.courseId === "A");
  const bRoom = original[0].rooms.find((room) => room.courseId === "B");
  const next = buildExamSessions(assignments, lookup, 60, 25, { A: "merge" });
  const plan = reconcileRoomDistribution(original, next, resources, originalPlan, "A");
  const a = next[0].rooms.find((room) => room.courseId === "A");
  assert.equal(a.students.length, 26);
  assert.equal(a.requiredInvigilators, 2);
  assert.equal(a.maxStudents, 27);
  assert.equal(plan.allocations[a.id].roomId, originalPlan.allocations[aRooms[0].id].roomId);
  assert.deepEqual(new Set(plan.allocations[a.id].invigilatorIds), new Set(aRooms.flatMap((room) => originalPlan.allocations[room.id].invigilatorIds)));
  assert.deepEqual(plan.allocations[bRoom.id], originalPlan.allocations[bRoom.id]);
  assert.deepEqual(plan.backups, originalPlan.backups);
  assert.equal(validateResourcePlan(next, resources, plan).complete, true);
  assert.equal(Object.keys(plan.allocations).length, 2);
  assert.deepEqual(original[0].rooms.map((room) => room.students.length), [13, 13, 20], "The original model is not mutated");
  const restoredPlan = reconcileRoomDistribution(next, original, resources, plan, "A");
  assert.deepEqual(restoredPlan.allocations[bRoom.id], originalPlan.allocations[bRoom.id]);
  assert.equal(restoredPlan.allocations[aRooms[1].id].roomId, "", "Restoring the extra room needs a fresh physical-room assignment");
  assert.equal(validateResourcePlan(original, resources, restoredPlan).complete, false);
  resources.invigilators[0].enabled = false;
  assert.deepEqual(reconcileRoomDistribution(original, next, resources, originalPlan, "A").allocations, {}, "Stale resource plans are not reused");
});

test("consolidated lab exams retain their original instructor and room while requiring two invigilators", () => {
  const labRoom = roomIdentity("Campus", "Building", "Lab");
  const lab = { days: ["Monday"], crn: "101", startMinutes: 480, endMinutes: 590, room: "Lab", building: "Building", campus: "Campus", labInstructorId: "I2" };
  const c = course("LAB", 27, { labSessions: [lab] });
  const assignments = { 1: { Monday: { "08:00": ["LAB"] } } };
  const resources = catalog(2, 4);
  resources.rooms[0].id = labRoom;
  resources.invigilators[2].busy = [classTime("LAB", 480, 590, { isLab: true })];
  const original = buildExamSessions(assignments, { LAB: c }, 60);
  const next = buildExamSessions(assignments, { LAB: c }, 60, 25, { LAB: "merge" });
  const plan = reconcileRoomDistribution(original, next, resources, assignResources(original, resources), "LAB");
  const room = next[0].rooms[0];
  assert.equal(room.students.length, 27);
  assert.equal(plan.allocations[room.id].roomId, labRoom);
  assert.equal(plan.allocations[room.id].invigilatorIds[0], "I2");
  assert.equal(plan.allocations[room.id].invigilatorIds.filter(Boolean).length, 2);
  assert.equal(invigilatorWorkloads(resources, next, plan).find((person) => person.id === "I2").exam, 0);
  assert.equal(validateResourcePlan(next, resources, plan).complete, true);
});

test("capacity exceptions are recorded in reports and unapproved or excessive room counts block export", () => {
  const c = course("EXAM", 52);
  const assignments = { 1: { Monday: { "12:00": ["EXAM"] } } };
  const sessions = buildExamSessions(assignments, { EXAM: c }, 60, 25, { EXAM: "distribute" });
  assert.deepEqual(sessions[0].rooms.map((room) => room.students.length), [26, 26]);
  const resources = catalog(2, 5);
  const plan = assignResources(sessions, resources);
  const workbook = buildResourceWorkbookForWeek({ week: 1, sessions, catalog: resources, plan, startDate: "2026-10-19", templateHeaders: {
    invigilatorHeader: ["CRN", "Code", "Title", "Students", "Date", "Time", "Instructor", "Room"],
    studentHeader: ["CRN", "Code", "Title", "Student ID", "Student Name", "Room", "Present"],
  } });
  const inv = workbook.Sheets["Week 1 Invigilators"];
  assert.deepEqual([inv.D2.v, inv.D3.v], [26, 26]);
  assert.ok(inv.I2.v && inv.J2.v && inv.I3.v && inv.J3.v);
  assert.match(inv.O2.v, /Approved up to 27.*distribute/);
  assert.equal(workbook.Sheets["Week 1 Room Pool"].C2.v, 25);
  assert.equal(workbook.Sheets["Week 1 Room Pool"].D2.v, 27);
  assert.equal(workbook.Sheets["Week 1 Monday"].F28.f, "'Week 1 Invigilators'!H3");
  sessions[0].rooms[0].distributionChoice = "standard";
  plan.fingerprint = resourceFingerprint(sessions, resources);
  assert.ok(validateResourcePlan(sessions, resources, plan).issues.some((issue) => issue.title === "Room capacity exceeded"));
  assert.throws(() => buildResourceWorkbookForWeek({ week: 1, sessions, catalog: resources, plan }), /Complete valid/);
  sessions[0].rooms[0].distributionChoice = "distribute";
  sessions[0].rooms[0].students.push(...c.students.slice(0, 2));
  plan.fingerprint = resourceFingerprint(sessions, resources);
  assert.ok(validateResourcePlan(sessions, resources, plan).issues.some((issue) => /approved maximum is 27/.test(issue.message)));
});
