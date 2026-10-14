import test from "node:test";
import assert from "node:assert/strict";
import * as XLSX from "xlsx/xlsx.mjs";
import { parseDepartmentWorkbook } from "../src/department.js";
import { balanceInvigilatorPlan, buildExamSessions, emptyResourcePlan, invigilatorWorkloads, isTeachingTimeDuty, parseResourceCatalog, readResourceCatalog, scopeInvigilatorAvailability, resourceBusyReason, resourceChoiceReason, standbyAvailability, validateResourcePlan, assignResources } from "../src/resources.js";
import { buildExportFiles, buildExportModel } from "../src/exportReports.js";

const course = (id = "EXAM") => ({ id, code: id, title: id, crns: ["101", "102"], labSessions: [],
  students: Array.from({ length: 10 }, (_, i) => ({ id: "S" + i, name: "Student " + i, crn: "101" })) });
const catalog = () => ({ rooms: [{ id: "R", name: "Room", enabled: true, busy: [] }],
  invigilators: Array.from({ length: 3 }, (_, i) => ({ id: "I" + i, name: "Staff " + i, enabled: true, busy: [] })) });
function planWith(sessions, resources, primary) {
  const plan = emptyResourcePlan(sessions, resources);
  sessions.forEach((session, index) => {
    const id = primary(index, session);
    plan.allocations[session.rooms[0].id] = { roomId: "R", invigilatorIds: [id] };
    plan.backups[session.id] = [id === "I0" ? "I1" : "I0"];
  });
  assert.equal(validateResourcePlan(sessions, resources, plan).complete, true);
  return plan;
}

function teachingHoursFixture() {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
    ["CRN", "Course Code", "Title", "DAYS", "Time", "Type", "Primary Instructor", "Second Instructor", "Campus", "Building", "Room"],
    ["101", "TUTORIAL", "Tutorial", "M", "1200 - 1250", "T", "1: Teacher", "2: Other", "PAD", "Building", "13"],
    ["102", "LAB", "Lab", "T", "1200 - 1250", "OL", "2: Other", "1: Teacher", "PAD", "Building", "13"],
    ["103", "LECTURE", "Lecture", "W", "1200 - 1250", "LEC", "1: Teacher", "2: Other", "PAD", "Building", "13"],
  ]), "First");
  const selection = parseDepartmentWorkbook(book);
  selection.courses = selection.courses.map((entry) => ({ ...entry, id: entry.code }));
  const raw = parseResourceCatalog(selection);
  raw.rooms.push({ id: "R", name: "Exam Room", enabled: true, busy: [] });
  raw.invigilators.push({ id: "staff:3", name: "Spare", enabled: true, busy: [], teachingSchedule: [] });
  const choices = { TUTORIAL: true, LAB: false, LECTURE: false };
  const resources = scopeInvigilatorAvailability(raw, selection.courses, choices);
  const sessions = buildExamSessions({ 1: Object.fromEntries(["Monday", "Tuesday", "Wednesday"]
    .map((day) => [day, { "12:00": ["EXAM"] }])) }, { EXAM: course() }, 60);
  const plan = emptyResourcePlan(sessions, resources);
  sessions.forEach((session) => {
    plan.allocations[session.rooms[0].id] = { roomId: "R", invigilatorIds: ["staff:1"] };
    plan.backups[session.id] = ["staff:2"];
  });
  return { selection, raw, choices, resources, sessions, plan };
}

test("ignored tutorials and unchecked labs retain zero-load credit, but lectures and non-teaching instructors do not", () => {
  const { raw, resources, sessions, plan } = teachingHoursFixture();
  const teacher = resources.invigilators.find((person) => person.id === "staff:1");
  const other = resources.invigilators.find((person) => person.id === "staff:2");
  assert.deepEqual(teacher.busy, [], "Ignoring classes for availability is independent of workload credit");
  assert.equal(teacher.teachingSchedule.length, 3);
  assert.deepEqual(sessions.map((session) => isTeachingTimeDuty(teacher, session)), [true, true, false]);
  assert.ok(sessions.every((session) => !isTeachingTimeDuty(other, session)), "Being listed as the other course instructor is not a teaching duty");
  assert.ok(sessions.every((session) => resourceBusyReason(teacher, session, sessions) === ""));
  assert.match(resourceBusyReason(resources.rooms[0], sessions[0], sessions), /Class TUTORIAL/, "Tutorial room bookings still apply");
  assert.match(resourceBusyReason(raw.invigilators.find((person) => person.id === "staff:1"), sessions[1], sessions), /Class LAB/);
  assert.equal(validateResourcePlan(sessions, resources, plan).complete, true);
  const load = invigilatorWorkloads(resources, sessions, plan).find((person) => person.id === teacher.id);
  assert.deepEqual([load.exam, load.teaching, load.bySlot], [1, 2, { "12:00": 1 }]);
  assert.match(resourceChoiceReason(teacher, "invigilator", sessions[0], sessions, plan, sessions[0].id + "/backup/0"), /Already assigned/,
    "A zero-load exam still reserves the invigilator for its full duration");
});

test("auto-assignment and final balancing both use retained lab/tutorial hours for workload credit", () => {
  const { resources, sessions, plan } = teachingHoursFixture();
  const automatic = assignResources(sessions, resources);
  assert.equal(validateResourcePlan(sessions, resources, automatic).complete, true);
  for (const session of sessions.slice(0, 2)) {
    assert.deepEqual(automatic.allocations[session.rooms[0].id].invigilatorIds, ["staff:1"]);
    plan.allocations[session.rooms[0].id].invigilatorIds = ["staff:2"];
    plan.backups[session.id] = ["staff:1"];
  }
  const balanced = balanceInvigilatorPlan(sessions, resources, plan);
  assert.equal(validateResourcePlan(sessions, resources, balanced).complete, true);
  assert.ok(sessions.every((session) => balanced.allocations[session.rooms[0].id].roomId === "R"));
  assert.deepEqual(invigilatorWorkloads(resources, sessions, balanced).map((person) => person.exam).sort(), [0, 0, 1]);
  assert.equal(invigilatorWorkloads(resources, sessions, balanced).reduce((sum, person) => sum + person.teaching, 0), 2);
});

test("lab/tutorial credit survives saved resources and CRN scoping without changing the raw catalog", () => {
  const { selection, raw, choices } = teachingHoursFixture();
  const saved = JSON.parse(JSON.stringify(raw));
  assert.deepEqual(readResourceCatalog(saved), raw);
  const before = JSON.stringify(saved);
  const active = scopeInvigilatorAvailability(saved, selection.courses, choices);
  assert.equal(JSON.stringify(saved), before);
  assert.equal(active.invigilators.find((person) => person.id === "staff:1").teachingSchedule.length, 3);
  assert.equal(scopeInvigilatorAvailability(active, selection.courses, choices).invigilators
    .find((person) => person.id === "staff:1").teachingSchedule.length, 3, "Repeated scoping must not duplicate original teaching hours");
  delete saved.invigilators.find((person) => person.id === "staff:1").teachingSchedule;
  assert.equal(scopeInvigilatorAvailability(saved, selection.courses, choices).invigilators
    .find((person) => person.id === "staff:1").teachingSchedule.length, 3, "CRN meeting data supplies actual teaching assignments");
  saved.invigilators[0].teachingSchedule = "not an array";
  assert.throws(() => readResourceCatalog(saved), /Invalid resource/);
  saved.invigilators[0].teachingSchedule = [{ code: "TUTORIAL", crn: "101", isLab: false, days: ["Monday"], start: 720, end: 710 }];
  assert.throws(() => readResourceCatalog(saved), /Invalid resource availability/);
});

test("retained lab/tutorial hours give identical credit in resource workloads and exported reports", () => {
  const { resources, sessions, plan } = teachingHoursFixture();
  const options = { sessions, catalog: resources, plan, startDate: "2026-10-19", templateHeaders: {
    invigilatorHeader: ["CRN", "Code", "Title", "Students", "Date", "Time", "Instructor", "Room"],
    studentHeader: ["CRN", "Code", "Title", "ID", "Name", "Room", "Notes"],
  } };
  const model = buildExportModel(options);
  assert.deepEqual(model.duties.map((duty) => duty.extraLoad), [0, 0, 1]);
  const load = model.workloads.find((person) => person.id === "staff:1");
  assert.deepEqual([load.exam, load.teaching], [1, 2]);
  const complete = XLSX.read(buildExportFiles(options)[0].data, { type: "array" });
  assert.deepEqual([2, 3, 4].map((row) => complete.Sheets["Invigilator Duties"]["L" + row].v), [0, 0, 1]);
  const pool = XLSX.utils.sheet_to_json(complete.Sheets["Invigilator Pool"], { header: 1 });
  assert.deepEqual(pool.find((row) => row[0] === "Teacher").slice(1, 3), [1, 2]);
  const staff = XLSX.read(buildExportFiles({ ...options, report: "staff" })[0].data, { type: "array" });
  const duties = XLSX.utils.sheet_to_json(staff.Sheets["Staff duties"], { header: 1 }).slice(1);
  assert.deepEqual(duties.map((row) => row[9]), [0, 0, 1]);
});

test("tutorials free staff and their day span, but retain room bookings and lecture commitments", () => {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
    ["CRN", "Course Code", "Title", "DAYS", "Time", "Type", "Primary Instructor", "Campus", "Building", "Room"],
    ["101", "EXAM", "Exam", "M", "0800 - 0850", "T", "1: Teacher", "PAD", "Building", "13"],
    ["102", "EXAM", "Exam", "T", "1200 - 1250", "LEC", "1: Teacher", "PAD", "Building", "13"],
  ]), "First");
  const selection = parseDepartmentWorkbook(book);
  const resources = parseResourceCatalog(selection);
  assert.equal(resources.rooms[0].busy.length, 2);
  assert.deepEqual(resources.invigilators[0].busy.map((entry) => entry.type), ["LEC"]);
  const tutorial = { ...resources.rooms[0].busy[0] };
  delete tutorial.type;
  resources.invigilators[0].busy.push(tutorial);
  const before = JSON.stringify(resources);
  const active = scopeInvigilatorAvailability(resources, selection.courses, { EXAM: true });
  assert.equal(active.invigilators[0].busy.length, 1, "The CRN list also identifies tutorial commitments lacking a type on the resource entry");
  assert.equal(active.rooms, resources.rooms);
  assert.equal(JSON.stringify(resources), before);
  const sessions = buildExamSessions({ 1: { Monday: { "17:00": ["EXAM"] } } }, { EXAM: course() }, 60);
  active.invigilators.push({ id: "staff:2", name: "Z Spare", enabled: true, busy: [] });
  const plan = assignResources(sessions, active);
  assert.equal(plan.allocations[sessions[0].rooms[0].id].invigilatorIds[0], "staff:1", "Ignored 08:00 tutorials must not make a 17:00 exam a ten-hour day");
  assert.equal(resourceBusyReason(active.invigilators[0], { ...sessions[0], day: "Tuesday", start: 720, end: 780 }, sessions).startsWith("Class EXAM"), true);
  assert.equal(resourceBusyReason(active.rooms[0], { ...sessions[0], start: 480, end: 530 }, sessions).startsWith("Class EXAM"), true);
});

test("the final pass balances primary duties and exchanges backup names without changing rooms or the source plan", () => {
  const sessions = buildExamSessions({ 1: Object.fromEntries(["Monday", "Tuesday", "Wednesday", "Thursday"].map((day) => [day, { "12:00": ["EXAM"] }])) }, { EXAM: course() }, 60);
  const resources = catalog();
  const original = planWith(sessions, resources, () => "I0");
  const before = JSON.stringify(original);
  const plan = balanceInvigilatorPlan(sessions, resources, original);
  assert.equal(JSON.stringify(original), before);
  assert.equal(validateResourcePlan(sessions, resources, plan).complete, true);
  const loads = invigilatorWorkloads(resources, sessions, plan);
  assert.deepEqual(loads.map((load) => load.exam).sort(), [1, 1, 2]);
  assert.ok(Object.values(plan.allocations).every((allocation) => allocation.roomId === "R"));
  assert.ok(sessions.every((session) => standbyAvailability(session, sessions, resources, plan).count === 2));
  assert.ok(Object.values(plan.backups).some((ids) => ids[0] !== "I1"), "A backup name can be exchanged so it does not lock a lighter-loaded person out of room duty");
  assert.ok(loads.every((load) => !("backup" in load) && !("backupBySlot" in load)));
});

test("the final pass swaps equally loaded people to improve time-slot balance", () => {
  const sessions = buildExamSessions({ 1: { Monday: { "09:00": ["EXAM"], "12:00": ["EXAM"], "15:00": ["EXAM"] },
    Tuesday: { "09:00": ["EXAM"], "12:00": ["EXAM"], "15:00": ["EXAM"] } } }, { EXAM: course() }, 60);
  const resources = catalog();
  const original = planWith(sessions, resources, (index) => "I" + index % 3);
  const plan = balanceInvigilatorPlan(sessions, resources, original);
  assert.equal(validateResourcePlan(sessions, resources, plan).complete, true);
  const loads = invigilatorWorkloads(resources, sessions, plan);
  assert.ok(loads.every((load) => load.exam === 2));
  assert.ok(loads.every((load) => Object.values(load.bySlot).every((count) => count <= 1)));
});

test("reassigning staff retains valid reviewed rooms and replaces only unavailable room selections", () => {
  const sessions = buildExamSessions({ 1: { Monday: { "12:00": ["EXAM"] } } }, { EXAM: course() }, 60);
  const resources = catalog();
  resources.rooms.push({ id: "OTHER", name: "A first choice", enabled: true, busy: [] });
  const previousPlan = planWith(sessions, resources, () => "I0");
  const plan = assignResources(sessions, resources, { previousPlan });
  assert.equal(plan.allocations[sessions[0].rooms[0].id].roomId, "R");
  resources.rooms[0].enabled = false;
  const revised = assignResources(sessions, resources, { previousPlan });
  assert.equal(revised.allocations[sessions[0].rooms[0].id].roomId, "OTHER");
  assert.equal(validateResourcePlan(sessions, resources, revised).complete, true);
});

test("better workload balance cannot introduce a longer-than-eight-hour exam day", () => {
  const sessions = buildExamSessions({ 1: { Monday: { "17:00": ["EXAM"] }, Tuesday: { "17:00": ["EXAM"] }, Wednesday: { "17:00": ["EXAM"] } } }, { EXAM: course() }, 60);
  const resources = catalog();
  resources.invigilators.forEach((person, index) => {
    person.busy = [{ code: "LECTURE", crn: "101", isLab: false, days: ["Monday", "Tuesday", "Wednesday"], start: index === 0 ? 600 : 480, end: index === 0 ? 660 : 540 }];
  });
  const original = planWith(sessions, resources, () => "I0");
  const plan = balanceInvigilatorPlan(sessions, resources, original);
  assert.deepEqual(plan.allocations, original.allocations);
  assert.equal(validateResourcePlan(sessions, resources, plan).complete, true);
});

test("backup names remain in the complete workbook, never in staff rows, counters or formulas", () => {
  const sessions = buildExamSessions({ 1: { Monday: { "12:00": ["EXAM"] } } }, { EXAM: course() }, 60);
  const resources = catalog(), plan = planWith(sessions, resources, () => "I0");
  const options = { sessions, catalog: resources, plan, startDate: "2026-10-19", templateHeaders: {
    invigilatorHeader: ["CRN", "Code", "Title", "Students", "Date", "Time", "Instructor", "Room"],
    studentHeader: ["CRN", "Code", "Title", "ID", "Name", "Room", "Notes"],
  } };
  const model = buildExportModel(options);
  assert.equal(model.duties.length, 1);
  const changed = structuredClone(plan);
  changed.backups[sessions[0].id] = ["I2"];
  assert.deepEqual(buildExportModel({ ...options, plan: changed }).workloads, model.workloads);
  const complete = XLSX.read(buildExportFiles(options)[0].data, { type: "array" });
  assert.equal(complete.Sheets["Invigilator Duties"].K2.v, "Staff 1");
  assert.equal(complete.Sheets["Invigilator Duties"].N1.v, "Room allocation decision");
  const pool = XLSX.utils.sheet_to_json(complete.Sheets["Invigilator Pool"], { header: 1 });
  assert.deepEqual(pool[0], ["Invigilator", "Invigilation load", "During teaching hours", "Included in pool"]);
  assert.ok(!Object.values(complete.Sheets["Invigilator Pool"]).some((cell) => cell?.f?.includes("!K:K")));
  const staff = XLSX.read(buildExportFiles({ ...options, report: "staff" })[0].data, { type: "array" });
  assert.ok(!JSON.stringify(staff.Sheets).includes("Backup load"));
  assert.equal(XLSX.utils.sheet_to_json(staff.Sheets["Staff duties"]).length, 1);
});
