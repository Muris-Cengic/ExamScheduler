import test from "node:test";
import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import * as XLSX from "xlsx/xlsx.mjs";
import { parseCyberSkillsWorkbook, readCyberSkillsSchedule, resolveCyberSkillsLinks, withCyberSkillsAvailability } from "../src/cyberSkills.js";
import { parseDepartmentWorkbook } from "../src/department.js";
import { assignResources, buildExamSessions, emptyResourcePlan, invigilatorDaySpan, invigilatorWorkloads, isTeachingTimeDuty, parseResourceCatalog, resourceBusyReason, resourceChoiceReason, resourceChoiceSummary, resourceFingerprint, scopeInvigilatorAvailability, validateResourcePlan } from "../src/resources.js";
import { buildExportFiles, buildExportModel, buildStaffDistribution } from "../src/exportReports.js";
import { autoSchedule } from "../src/autoSchedule.js";

const book = (rows, name = "Sheet") => { const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), name); return workbook; };
const flat = (date = "2026-10-19", time = "08:00 - 09:00", name = "Dr. Alex Smith") => book([
  ["Date", "Time", "Instructor", "Session"], [date, time, name, "Computer Networks"],
]);
const catalog = () => ({ rooms: [{ id: "R", name: "Room", enabled: true, busy: [] }], invigilators: [
  { id: "A", name: "Alex James Smith", enabled: true, busy: [] },
  ...Array.from({ length: 3 }, (_, i) => ({ id: "B" + i, name: "Other " + i, enabled: true, busy: [] })),
] });
const schedule = (date = "2026-10-19", time = "08:00 - 09:00", name = "Dr. Alex Smith") => ({ filename: "CyberSkills.xlsx", sessions: parseCyberSkillsWorkbook(flat(date, time, name), "2026-10-19"), links: {} });
const course = { id: "EXAM", code: "EXAM", title: "Exam", crns: ["101", "102"], labSessions: [], students: [{ id: "S1", name: "Student", crn: "101" }] };
const examSessions = (time = "08:30") => buildExamSessions({ 1: { Monday: { [time]: ["EXAM"] } } }, { EXAM: course }, 60);

test("weekly CyberSkills grids use calendar dates, trainer rows and full times, ignoring summaries, breaks and old sheets", () => {
  const workbook = book([
    ["CyberSkills Week 5"], ["Time", "Monday", "Tuesday"], ["", 46314, 46315],
    ["08:00 - 09:00", "Module: Computer Networks", "Module: AI"],
    ["", "Topics: examples", "Topics: examples"], ["", "Trainer: Dr. Alex Smith", "Ms. Beth Jones"],
    ["Break", "", ""], ["", "Trainer: ignored break row", ""],
    ["13:00 - 14:00", "Module: Networks", "Module: AI"], ["", "Trainer: Dr. Alex Smith", "Trainer: Ms. Beth Jones"],
    ["Staff Name", "Total Hours", "Role"], ["", "Dr. Footer Name", "Trainer"],
  ], "Week 5");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["Staff Name", "Week 5"], ["Dr. Alex Smith", 9]]), "Overall");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["Date", "Time", "Instructor"], ["2025-10-20", "08:00 - 09:00", "Old Trainer"]]), "Old");
  const duties = parseCyberSkillsWorkbook(workbook, "2026-10-21");
  assert.equal(duties.length, 4);
  assert.deepEqual([...new Set(duties.map((duty) => duty.date))], ["2026-10-19", "2026-10-20"]);
  assert.deepEqual(duties.filter((duty) => duty.date === "2026-10-19").map((duty) => [duty.start, duty.end]), [[480, 540], [780, 840]]);
  assert.equal(duties[0].label, "Computer Networks");
  const active = withCyberSkillsAvailability(catalog(), { filename: "Schedule", sessions: duties, links: { "Ms. Beth Jones": null } }, "2026-10-21");
  assert.equal(active.invigilators[0].externalDuties[0].week, 1, "Training Week 5 is exam Week 1 by date");
});

test("flat dates/times, deduplication and malformed input fail explicitly instead of inventing commitments", () => {
  const workbook = flat();
  XLSX.utils.book_append_sheet(workbook, workbook.Sheets.Sheet, "Duplicate");
  assert.equal(parseCyberSkillsWorkbook(workbook, "2026-10-19").length, 1);
  assert.equal(parseCyberSkillsWorkbook(flat("10/19/26", "08:00 - 08:50"), "2026-10-19")[0].end, 530, "CyberSkills times are exact, with no teaching-hour rounding");
  for (const input of [flat("2026-02-31"), flat("2026-10-24"), flat("2026-10-19", "09:00 - 08:00"), flat("2026-10-19", "25:00 - 26:00"), flat("2026-10-19", "08:00 - 09:00", "")]) {
    assert.throws(() => parseCyberSkillsWorkbook(input, "2026-10-19"));
  }
  assert.throws(() => parseCyberSkillsWorkbook(flat("2025-10-20"), "2026-10-19"), /No dated CyberSkills/);
  assert.throws(() => parseCyberSkillsWorkbook(book([["Time", "Monday", "Tuesday"], ["", "2026-10-20", "2026-10-20"]]), "2026-10-19"), /does not match Monday/);
  assert.throws(() => parseCyberSkillsWorkbook(book([["Time", "Monday", "Tuesday"], ["", 46314, 46315],
    ["08:00 - 09:00", "Trainer: Dr. Alex Smith", ""], ["10:0011:00", "Trainer: Dr. Alex Smith", ""]]), "2026-10-19"), /Unrecognised exam time/,
  "A malformed new time block cannot silently inherit the previous block's hours");
});

test("name matching removes titles and accepts omitted middle names only when unique; explicit mappings and ignored people persist", () => {
  const raw = schedule();
  const people = catalog().invigilators;
  assert.equal(resolveCyberSkillsLinks(raw, people)["Dr. Alex Smith"], "A");
  people.push({ id: "AMBIGUOUS", name: "Alex Smith" });
  assert.equal(resolveCyberSkillsLinks(raw, people)["Dr. Alex Smith"], "");
  raw.links["Dr. Alex Smith"] = "A";
  assert.equal(resolveCyberSkillsLinks(raw, people)["Dr. Alex Smith"], "A");
  raw.links["Dr. Alex Smith"] = null;
  assert.equal(resolveCyberSkillsLinks(raw, people)["Dr. Alex Smith"], null);
  raw.links["Dr. Alex Smith"] = "";
  assert.equal(resolveCyberSkillsLinks(raw, [people[0]])["Dr. Alex Smith"], "", "Clearing an auto-match must stay unresolved until reviewed");
  raw.links["Dr. Alex Smith"] = "MISSING";
  assert.equal(resolveCyberSkillsLinks(raw, [people[0]])["Dr. Alex Smith"], "", "A missing manually matched ID cannot silently relink");
  assert.equal(resolveCyberSkillsLinks(schedule("2026-10-19", "08:00 - 09:00", "Professor Drake Smith"), [{ id: "D", name: "Drake Smith" }])["Professor Drake Smith"], "D");
});

test("CyberSkills overlaps block exam and backup choices on actual dates, allowing adjacent times and other weeks", () => {
  const raw = catalog(), before = JSON.stringify(raw);
  const active = withCyberSkillsAvailability(raw, schedule(), "2026-10-19");
  assert.equal(JSON.stringify(raw), before);
  assert.equal(active.rooms, raw.rooms);
  const sessions = examSessions(), person = active.invigilators[0], plan = emptyResourcePlan(sessions, active);
  assert.match(resourceChoiceReason(person, "invigilator", sessions[0], sessions, plan, "room/invigilator/0"), /CyberSkills.*2026-10-19.*08:00-09:00/);
  assert.match(resourceChoiceReason(person, "invigilator", sessions[0], sessions, plan, sessions[0].id + "/backup/0"), /CyberSkills/);
  assert.equal(resourceChoiceSummary(person, "invigilator", sessions[0], sessions, plan, "room/invigilator/0"), "CyberSkills 08:00-09:00");
  assert.equal(resourceBusyReason(person, { ...sessions[0], week: 2 }, sessions), "");
  assert.equal(resourceBusyReason(person, { ...sessions[0], start: 540, end: 600 }, sessions), "");
  const assigned = assignResources(sessions, active);
  assert.equal(validateResourcePlan(sessions, active, assigned).complete, true);
  assert.ok(!assigned.allocations[sessions[0].rooms[0].id].invigilatorIds.includes("A"));
  assert.ok(!assigned.backups[sessions[0].id].includes("A"));
  assigned.allocations[sessions[0].rooms[0].id].invigilatorIds = ["A"];
  assert.ok(validateResourcePlan(sessions, active, assigned).issues.some((issue) => issue.message.includes("CyberSkills")));
});

test("CyberSkills is unaffected by course exclusions or lab replacement and cannot cancel a required lab instructor", () => {
  const raw = catalog();
  raw.invigilators[0].busy = [{ code: "LAB", crn: "101", days: ["Monday"], start: 480, end: 600, isLab: true }];
  const active = withCyberSkillsAvailability(scopeInvigilatorAvailability(raw, [{ id: "LAB", code: "LAB", title: "Lab", meetings: [] }], { LAB: false }), schedule(), "2026-10-19");
  const sessions = examSessions("08:00");
  sessions[0].replacedLabs = [{ code: "LAB", crn: "101", start: 480, end: 600 }];
  sessions[0].rooms[0].fixedInvigilatorId = "A";
  const plan = assignResources(sessions, active);
  const validation = validateResourcePlan(sessions, active, plan);
  assert.equal(validation.complete, false);
  const issue = validation.issues.find((issue) => issue.title === "Lab instructor unavailable");
  assert.match(issue.message, /CyberSkills/);
  assert.match(issue.action, /step 3/);
});

test("CyberSkills extends only matching dates' duty spans and influences staff selection without affecting workload", () => {
  const raw = catalog();
  const active = withCyberSkillsAvailability(raw, schedule(), "2026-10-19");
  const sessions = examSessions("17:00");
  const span = invigilatorDaySpan(active.invigilators[0], sessions[0], sessions, [sessions[0]]);
  assert.deepEqual([span.start, span.end, span.minutes, span.overrun], [480, 1080, 600, 120]);
  assert.equal(invigilatorDaySpan(active.invigilators[0], { ...sessions[0], week: 2 }, sessions, [{ ...sessions[0], week: 2 }]).minutes, 60);
  assert.equal(isTeachingTimeDuty(active.invigilators[0], { ...sessions[0], start: 480, end: 540 }), false, "External time never grants lab/tutorial credit");
  const plan = assignResources(sessions, active);
  assert.notEqual(plan.allocations[sessions[0].rooms[0].id].invigilatorIds[0], "A", "Prefer a shorter day over adding a late exam to CyberSkills morning duty");
  assert.deepEqual(invigilatorWorkloads(active, sessions, plan).map((person) => [person.exam, person.teaching]), invigilatorWorkloads(raw, sessions, plan).map((person) => [person.exam, person.teaching]));
  const model = buildExportModel({ sessions, catalog: active, plan, startDate: "2026-10-19" });
  const alex = buildStaffDistribution(model).rows.find((person) => person.id === "A");
  assert.equal(alex.total, 0);
  assert.deepEqual([alex.weeks[0].days[0].span.start, alex.weeks[0].days[0].span.end], [480, 540]);
  assert.equal(model.duties.length, 1);
  const file = buildExportFiles({ sessions, catalog: active, plan, startDate: "2026-10-19", report: "staff" })[0];
  const workbook = XLSX.read(file.data, { type: "array" });
  assert.equal(XLSX.utils.sheet_to_json(workbook.Sheets["Staff duties"]).length, 1);
  const hours = XLSX.utils.sheet_to_json(workbook.Sheets["Daily Duty Hours"]).find((row) => row.Invigilator === "Alex James Smith");
  assert.deepEqual([hours.Start, hours.End, hours["Extra exam load"]], ["08:00", "09:00", 0]);
});

test("unresolved trainers block resource validation and exports until matched or explicitly ignored", () => {
  const raw = schedule("2026-10-19", "08:00 - 09:00", "Unknown Trainer");
  const sessions = examSessions("12:00");
  const active = withCyberSkillsAvailability(catalog(), raw, "2026-10-19");
  const plan = assignResources(sessions, active);
  assert.ok(validateResourcePlan(sessions, active, plan).issues.some((issue) => issue.title === "CyberSkills trainer needs matching"));
  assert.throws(() => buildExportModel({ sessions, catalog: active, plan, startDate: "2026-10-19" }), /Complete valid resource/);
  raw.links["Unknown Trainer"] = null;
  const ignored = withCyberSkillsAvailability(catalog(), raw, "2026-10-19");
  assert.equal(validateResourcePlan(sessions, ignored, assignResources(sessions, ignored)).complete, true);
});

test("automatic scheduling checks CyberSkills staff availability before accepting an exam slot", () => {
  const workbook = book([["Date", "Time", "Instructor", "Session"], ...["2026-10-19", "2026-10-20", "2026-10-21", "2026-10-22"]
    .flatMap((date) => [0, 1, 2].map((i) => [date, "12:00 - 13:00", "Other " + i, "CyberSkills"]))]);
  const active = withCyberSkillsAvailability(catalog(), { filename: "Noon training", sessions: parseCyberSkillsWorkbook(workbook, "2026-10-19"), links: {} }, "2026-10-19");
  const result = autoSchedule({ courses: [course], courseLookup: { EXAM: course }, assignments: {}, weeks: [1],
    timeSlots: [{ id: "12:00" }, { id: "12:30" }, { id: "17:00" }, { id: "17:30" }],
    settings: { slotIntervalMinutes: 30, startHour: 8, endHour: 18, studentsPerRoom: 25, examDurationMinutes: 60 }, catalog: active });
  assert.equal(result.placed.length, 1);
  assert.deepEqual([result.placed[0].day, result.placed[0].slotId], ["Monday", "17:00"], "No noon slot has enough staff and backup coverage");
  const sessions = buildExamSessions(result.assignments, { EXAM: course }, 60);
  assert.equal(validateResourcePlan(sessions, active, result.resourcePlan).complete, true);
});

test("saved CyberSkills data validates and round-trips deterministically; date or schedule changes invalidate old plans", () => {
  const raw = schedule(), people = catalog();
  const restored = readCyberSkillsSchedule(JSON.parse(JSON.stringify(raw)));
  assert.deepEqual(restored, raw);
  assert.equal(readCyberSkillsSchedule(null), null);
  assert.deepEqual(withCyberSkillsAvailability(people, null, "2026-10-19"), people);
  const active = withCyberSkillsAvailability(people, raw, "2026-10-19"), sessions = examSessions("12:00");
  assert.equal(resourceFingerprint(sessions, active), resourceFingerprint(sessions, withCyberSkillsAvailability(people, restored, "2026-10-19")));
  assert.notEqual(resourceFingerprint(sessions, active), resourceFingerprint(sessions, withCyberSkillsAvailability(people, raw, "2026-10-26")));
  assert.notEqual(resourceFingerprint(sessions, active), resourceFingerprint(sessions, people));
  for (const value of [{ ...raw, links: [] }, { ...raw, links: { wrong: 12 } }, { ...raw, sessions: [] }, { ...raw, sessions: [{ ...raw.sessions[0], end: 400 }] }, { ...raw, sessions: [{ ...raw.sessions[0], date: "2026-02-31" }] }]) assert.throws(() => readCyberSkillsSchedule(value));
});

const cyberPath = "data/input/26-27 S1/Cyber skills schedule/CyberSkills - Batch 3 Week 5-6.xlsx";
const crnPath = "data/input/26-27 S1/2. CRN List 26-27 S1.xlsx";
test("supplied CyberSkills workbook imports only the current 64 duties and matches all six trainers", { skip: !existsSync(cyberPath) || !existsSync(crnPath) }, () => {
  const sessions = parseCyberSkillsWorkbook(XLSX.read(readFileSync(cyberPath), { type: "buffer" }), "2026-10-19");
  const people = parseResourceCatalog(parseDepartmentWorkbook(XLSX.read(readFileSync(crnPath), { type: "buffer" }))).invigilators;
  const links = resolveCyberSkillsLinks({ sessions, links: {} }, people);
  assert.equal(sessions.length, 64);
  assert.deepEqual([...new Set(sessions.map((entry) => entry.sheet))], ["Week 5", "Week 6"]);
  assert.equal(Object.keys(links).length, 6);
  assert.ok(Object.values(links).every(Boolean));
  assert.equal(people.find((person) => person.id === links["Mr. Atif Chowhan"]).name, "Muhammad Atif Chowhan");
});
