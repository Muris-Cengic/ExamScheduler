import test from "node:test";
import assert from "node:assert/strict";
import * as XLSX from "xlsx/xlsx.mjs";
import { roomIdentity } from "../src/department.js";
import { assignResources, buildExamSessions, emptyResourcePlan, scopeInvigilatorAvailability, validateResourcePlan } from "../src/resources.js";
import { buildExportFiles, buildExportModel, buildStaffDistribution, buildStaffSessionRosters } from "../src/exportReports.js";

function fixture() {
  const course = (id, extra = {}) => ({ id, code: id, title: id, crns: ["101", "102"], labSessions: [],
    students: Array.from({ length: 10 }, (_, i) => ({ id: id + i, name: "Student " + i, crn: "101" })), ...extra });
  const lab = { days: ["Monday"], crn: "101", startMinutes: 480, endMinutes: 710, campus: "PAD", building: "Building", room: "13", labInstructorId: "A" };
  const courses = { LAB: course("LAB", { crns: ["101"], labSessions: [lab] }), EXAM: course("EXAM"), OTHER: course("OTHER") };
  const sessions = buildExamSessions({ 1: { Monday: { "09:00": ["LAB"], "12:00": ["EXAM"], "17:00": ["OTHER"] }, Tuesday: { "12:00": ["EXAM"] } },
    2: { Monday: { "12:00": ["EXAM"] } } }, courses, 60);
  const commitment = (code, day, start, end, extra = {}) => ({ code, crn: "101", days: [day], start, end, isLab: false, ...extra });
  const raw = { rooms: [{ id: roomIdentity(lab.campus, lab.building, lab.room), name: "Lab", enabled: true, busy: [] }, { id: "R", name: "Exam Room", enabled: true, busy: [] }],
    invigilators: [
      { id: "A", name: "Alice", enabled: true, busy: [commitment("LAB", "Monday", 480, 710, { isLab: true }),
        commitment("LECTURE", "Monday", 900, 950), commitment("TUTORIAL", "Monday", 420, 470, { type: "T" }), commitment("IGNORE", "Tuesday", 480, 530)] },
      { id: "B", name: "Bob", enabled: true, busy: [commitment("LECTURE", "Monday", 510, 560), commitment("LECTURE", "Tuesday", 480, 530)] },
      { id: "C", name: "Chris", enabled: true, busy: [commitment("UNKNOWN", "Wednesday", 0, 1440, { unknownTime: true })] },
      { id: "D", name: "Disabled", enabled: false, busy: [] },
      { id: "E", name: "Erin", enabled: true, busy: [] },
    ] };
  const catalog = scopeInvigilatorAvailability(raw, [course("IGNORE")], { IGNORE: false });
  const plan = emptyResourcePlan(sessions, catalog);
  sessions.forEach((session, index) => {
    plan.allocations[session.rooms[0].id] = { roomId: session.rooms[0].fixedRoomId || "R", invigilatorIds: [index === 2 || index === 3 ? "B" : "A"] };
    plan.backups[session.id] = ["E"];
  });
  assert.equal(validateResourcePlan(sessions, catalog, plan).complete, true);
  return { sessions, catalog, plan, startDate: "2026-10-19" };
}

test("session rosters are chronological, preserve zero-load duties and exclude backup-only or unassigned staff", () => {
  const options = fixture();
  const model = buildExportModel(options);
  const before = JSON.stringify(model);
  const rosters = buildStaffSessionRosters({ ...model, sessions: [...model.sessions].reverse(), duties: [...model.duties].reverse() });
  assert.equal(JSON.stringify(model), before);
  assert.deepEqual(rosters.map((session) => session.id), options.sessions.map((session) => session.id));
  assert.equal(rosters[0].duties[0].extraLoad, 0);
  assert.equal(rosters[0].duties[0].name, "Alice");
  assert.ok(rosters.every((session) => session.roomCount === 1 && session.invigilatorCount === 1));
  assert.deepEqual(rosters.flatMap((session) => session.duties.map((duty) => duty.id)).sort(), model.duties.map((duty) => duty.id).sort());
  assert.ok(rosters.every((session) => session.duties.every((duty) => !["C", "D", "E"].includes(duty.invigilatorId))));
  assert.deepEqual(rosters.filter((session) => session.day === "Monday" && session.start === 720).map((session) => session.date), ["2026-10-19", "2026-10-26"]);
  const week2 = buildStaffSessionRosters(buildExportModel({ ...options, weeks: [2] }));
  assert.equal(week2.length, 1);
  assert.equal(week2[0].week, 2);
});

test("one session roster combines concurrent courses and rooms without duplicating invigilators", () => {
  const course = (id, count) => ({ id, code: id, title: id, crns: ["101", "102"], labSessions: [],
    students: Array.from({ length: count }, (_, i) => ({ id: id + i, name: "Student " + i, crn: "101" })) });
  const sessions = buildExamSessions({ 1: { Monday: { "12:00": ["A", "B"] } } }, { A: course("A", 26), B: course("B", 20) }, 60);
  const catalog = { rooms: Array.from({ length: 3 }, (_, i) => ({ id: "R" + i, name: "Room " + i, enabled: true, busy: [] })),
    invigilators: Array.from({ length: 6 }, (_, i) => ({ id: "I" + i, name: "Staff " + (6 - i), enabled: true, busy: [] })) };
  const plan = assignResources(sessions, catalog);
  const model = buildExportModel({ sessions, catalog, plan, startDate: "2026-10-19" });
  const [roster] = buildStaffSessionRosters(model);
  assert.equal(roster.roomCount, 3);
  assert.equal(roster.invigilatorCount, 4);
  assert.deepEqual(roster.courses, ["A", "B"]);
  assert.equal(roster.duties.length, 4);
  assert.equal(new Set(roster.duties.map((duty) => duty.invigilatorId)).size, 4);
  assert.deepEqual(roster.duties.map((duty) => duty.name), roster.duties.map((duty) => duty.name).sort());
  assert.ok(roster.duties.every((duty) => duty.start === 720 && duty.end === 780));
  assert.ok(roster.duties.every((duty) => !plan.backups[sessions[0].id].includes(duty.invigilatorId)));
});

test("session distribution counts actual exam duties, including zero-load duties and idle enabled staff, never backups", () => {
  const options = fixture();
  const distribution = buildStaffDistribution(buildExportModel(options));
  assert.deepEqual(distribution.rows.map((person) => person.name), ["Alice", "Bob", "Chris", "Erin"]);
  assert.deepEqual(distribution.rows.map((person) => [person.exam, person.teaching, person.total]), [[2, 1, 3], [2, 0, 2], [0, 0, 0], [0, 0, 0]]);
  const alice = distribution.rows[0];
  assert.deepEqual(alice.weeks.map((week) => [week.exam, week.teaching]), [[1, 1], [1, 0]]);
  assert.equal(alice.weeks[0].days[0].slots[0].duty.extraLoad, 0);
  assert.equal(alice.weeks[0].days[0].slots[1].duty.extraLoad, 1);
  assert.deepEqual(alice.weeks[0].days[1].slots.map((slot) => slot.scheduled), [false, true, false]);
  assert.ok(alice.weeks[0].days[1].slots.every((slot) => !slot.duty));
  options.sessions.forEach((session) => { options.plan.backups[session.id] = ["C"]; });
  assert.deepEqual(buildStaffDistribution(buildExportModel(options)), distribution, "Changing standby names must not affect duties or daily spans");
});

test("daily start/end uses remaining teaching and exams, excludes ignored classes, and cancels labs only in their replaced week", () => {
  const distribution = buildStaffDistribution(buildExportModel(fixture()));
  const alice = distribution.rows[0];
  assert.deepEqual(alice.weeks[0].days[0].span, { start: 540, end: 950, minutes: 410, overrun: 0, unknownTime: false });
  assert.deepEqual(alice.weeks[1].days[0].span, { start: 480, end: 950, minutes: 470, overrun: 0, unknownTime: false });
  assert.equal(alice.weeks[0].days[1].span, null, "Unchecked courses do not extend a day with no exam duty");
  assert.deepEqual(distribution.rows[1].weeks[0].days[0].span, { start: 510, end: 1080, minutes: 570, overrun: 90, unknownTime: false });
  assert.deepEqual(distribution.rows[1].weeks[0].days[1].span, { start: 480, end: 780, minutes: 300, overrun: 0, unknownTime: false });
  assert.deepEqual(distribution.rows[2].weeks[0].days[2].span, { start: null, end: null, minutes: null, overrun: 0, unknownTime: true });
  assert.ok(distribution.rows[3].weeks.every((week) => week.days.every((day) => day.span === null)), "Backup-only staff have no daily duty span");
});

test("distribution periods and differing session end times follow only included weeks", () => {
  const options = fixture();
  const distribution = buildStaffDistribution(buildExportModel({ ...options, weeks: [2] }));
  assert.deepEqual(distribution.weeks.map((week) => week.week), [2]);
  assert.deepEqual(distribution.weeks[0].slots, [{ key: "720/780", start: 720, end: 780 }]);
  assert.deepEqual(distribution.rows.map((person) => person.total), [1, 0, 0, 0]);
  assert.equal(distribution.weeks[0].days[0].date, "2026-10-26");
  assert.equal(distribution.rows[0].weeks[0].days[0].span.start, 480);
  const model = buildExportModel(options);
  model.sessions.push({ ...model.sessions[1], id: "extra", end: 810 });
  assert.deepEqual(buildStaffDistribution(model).weeks[0].slots.map((slot) => slot.key), ["540/600", "720/780", "720/810", "1020/1080"]);
});

test("staff Excel adds distribution and daily duty hours with the same totals, scope and unknown-time warnings as the view", () => {
  const options = fixture();
  const workbook = XLSX.read(buildExportFiles({ ...options, report: "staff" })[0].data, { type: "array" });
  assert.deepEqual(workbook.SheetNames, ["Staff duties", "Workload Summary", "Session Distribution", "Daily Duty Hours"]);
  const stats = XLSX.utils.sheet_to_json(workbook.Sheets["Session Distribution"], { header: 1 });
  assert.deepEqual(stats[1].slice(0, 4), ["Alice", 2, 1, 3]);
  assert.equal(stats[1][4], "Lab/tutorial (0 load)");
  assert.equal(stats[1][5], "Exam (+1 load)");
  assert.equal(stats.length, 5, "Only enabled invigilators appear, including idle staff");
  const hours = XLSX.utils.sheet_to_json(workbook.Sheets["Daily Duty Hours"], { header: 1 });
  const aliceMonday = hours.find((row) => row[0] === 1 && row[2] === "Monday" && row[3] === "Alice");
  assert.deepEqual(aliceMonday.slice(4, 11), ["09:00", "15:50", 6.83, 2, 1, 1, "No"]);
  const unknown = hours.find((row) => row[3] === "Chris" && row[2] === "Wednesday");
  assert.deepEqual(unknown.slice(4, 7), ["", "", ""]);
  assert.match(unknown[10], /span may be incomplete/);
  assert.ok(!hours.some((row) => row[3] === "Erin" || row[3] === "Disabled"));
  const weekly = buildExportFiles({ ...options, report: "staff", packaging: "weekly", weeks: [2] });
  const week2 = XLSX.read(weekly[0].data, { type: "array" });
  const week2Stats = XLSX.utils.sheet_to_json(week2.Sheets["Session Distribution"], { header: 1 });
  assert.deepEqual(week2Stats[1].slice(0, 4), ["Alice", 1, 0, 1]);
  assert.ok(week2Stats[0].slice(4).every((header) => header.startsWith("Week 2")));
  assert.ok(XLSX.utils.sheet_to_json(week2.Sheets["Daily Duty Hours"]).every((row) => row.Week === 2));
  assert.ok(!JSON.stringify(workbook.Sheets).includes("Backup load"));
});
