import test from "node:test";
import assert from "node:assert/strict";
import * as XLSX from "xlsx/xlsx.mjs";
import { roomIdentity } from "../src/department.js";
import { buildExamSessions, emptyResourcePlan, scopeInvigilatorAvailability, validateResourcePlan } from "../src/resources.js";
import { buildExportFiles, buildExportModel, buildStaffCalendarFiles, buildStaffSchedules, reportFileName, staffScheduleSlots, staffSlotEntries } from "../src/exportReports.js";
import { buildInvigilatorCalendar } from "../src/invigilatorCalendar.js";

function fixture() {
  const lab = { days: ["Monday"], crn: "101", startMinutes: 480, endMinutes: 710, campus: "PAD", building: "P-B-4F", room: "13", labInstructorId: "A" };
  const course = (id, extra = {}) => ({ id, code: id, title: id + " title", crns: ["101", "102"], labSessions: [],
    students: [{ id: "PRIVATE-STUDENT-ID", name: "PRIVATE-STUDENT-NAME", crn: "101" }], ...extra });
  const sessions = buildExamSessions({ 1: { Monday: { "09:00": ["LAB"], "12:00": ["EXAM"] } }, 2: { Tuesday: { "12:00": ["EXAM"] } } },
    { LAB: course("LAB", { crns: ["101"], labSessions: [lab] }), EXAM: course("EXAM") }, 60);
  const entry = (code, day, start, end, extra = {}) => ({ code, crn: "101", days: [day], start, end, isLab: false, ...extra });
  const raw = { rooms: [{ id: roomIdentity(lab.campus, lab.building, lab.room), name: "PAD / P-B-4F / 13", enabled: true, busy: [] }], invigilators: [
    { id: "A", name: "Alice", enabled: true, busy: [entry("LAB", "Monday", 480, 710, { isLab: true, roomName: "PAD / P-B-4F / 13" }),
      entry("LECTURE", "Wednesday", 480, 530), entry("TUTORIAL", "Monday", 720, 770, { type: "T" }), entry("UNCHECKED", "Tuesday", 480, 590)],
      externalDuties: [{ week: 1, day: "Thursday", date: "2026-10-22", start: 600, end: 720, label: "Training topic" }] },
    { id: "B", name: "Bob", enabled: true, busy: [] },
    { id: "C", name: "Idle", enabled: true, busy: [entry("UNKNOWN", "Friday", 0, 1440, { unknownTime: true })] },
    { id: "D", name: "Excluded", enabled: false, busy: [] },
  ] };
  const catalog = scopeInvigilatorAvailability(raw, [course("UNCHECKED")], { UNCHECKED: false });
  const plan = emptyResourcePlan(sessions, catalog);
  sessions.forEach((session) => {
    plan.allocations[session.rooms[0].id] = { roomId: raw.rooms[0].id, invigilatorIds: ["A"] };
    plan.backups[session.id] = ["B"];
  });
  assert.equal(validateResourcePlan(sessions, catalog, plan).complete, true);
  return { sessions, catalog, plan, startDate: "2026-10-19" };
}
const allEntries = (people) => people.flatMap((person) => person.weeks.flatMap((week) => week.days.flatMap((day) => day.entries)));
const unfold = (data) => data.replace(/\r\n[ \t]/g, "");
const read = (file) => XLSX.read(file.data, { type: "array", cellStyles: true });
const workbookText = (book) => book.SheetNames.map((name) => XLSX.utils.sheet_to_json(book.Sheets[name], { header: 1 }).flat().join(" ")).join(" ");

test("staff schedules default to exam duties for every enabled invigilator without mutating the report model", () => {
  const model = buildExportModel(fixture());
  const before = JSON.stringify(model);
  const people = buildStaffSchedules(model);
  assert.deepEqual(people.map((person) => person.name), ["Alice", "Bob", "Idle"]);
  assert.equal(allEntries(people).length, model.duties.length);
  assert.ok(allEntries(people).every((entry) => entry.kind === "exam"));
  assert.equal(people[0].weeks[0].days[0].start, 540);
  assert.equal(people[0].weeks[0].days[0].end, 780);
  assert.equal(people[0].weeks[0].days[0].entries[0].extraLoad, 0);
  assert.equal(people[0].weeks[0].days[0].entries[0].roomName, "P-B-4F/13");
  assert.equal(allEntries(people.slice(1)).length, 0);
  assert.equal(JSON.stringify(model), before);
});

test("optional teaching and CyberSkills layers recur correctly, omit replaced labs and keep unknown times explicit", () => {
  const model = buildExportModel(fixture());
  const people = buildStaffSchedules(model, { includeTeaching: true, includeCyberSkills: true });
  const alice = people[0];
  assert.ok(!alice.weeks[0].days[0].entries.some((entry) => entry.kind === "teaching" && entry.code === "LAB"));
  assert.equal(alice.weeks[1].days[0].entries.find((entry) => entry.code === "LAB").roomName, "P-B-4F/13");
  assert.equal(alice.weeks[0].days[0].entries.find((entry) => entry.code === "TUTORIAL").label, "Tutorial");
  assert.ok(alice.weeks.every((week) => week.days[1].entries.some((entry) => entry.code === "UNCHECKED")));
  const cyber = allEntries(people).filter((entry) => entry.kind === "cyber");
  assert.equal(cyber.length, 1);
  assert.equal(cyber[0].date, "2026-10-22");
  assert.equal(cyber[0].title, "Training topic");
  assert.equal(cyber[0].extraLoad, null);
  assert.equal(people[2].weeks[0].days[4].start, null);
  assert.equal(people[2].weeks[0].days[4].unknownTime, true);
  const onlyWeek2 = buildStaffSchedules(buildExportModel({ ...fixture(), weeks: [2] }), { includeTeaching: true, includeCyberSkills: true });
  assert.deepEqual(onlyWeek2[0].weeks.map((week) => week.week), [2]);
  assert.equal(allEntries(onlyWeek2).filter((entry) => entry.kind === "cyber").length, 0);
  assert.ok(allEntries(onlyWeek2).some((entry) => entry.code === "LAB"));
  assert.equal(allEntries(buildStaffSchedules(model, { includeCyberSkills: true })).filter((entry) => entry.kind === "teaching").length, 0);
});

test("staff timetable slots preserve overlapping ranges and put unspecified times last", () => {
  const entries = [{ id: "unknown", unknownTime: true, start: 0, end: 1440 }, { id: "exam", start: 720, end: 780 },
    { id: "tutorial", start: 720, end: 770 }, { id: "lab", start: 480, end: 710 }];
  const week = { days: [{ entries }, { entries: [{ ...entries[1], id: "other-day" }] }] };
  const slots = staffScheduleSlots(week);
  assert.deepEqual(slots.map((slot) => slot.key), ["480/710", "720/770", "720/780", "unknown"]);
  assert.deepEqual(slots.flatMap((slot) => staffSlotEntries(week.days[0], slot)).map((entry) => entry.id), ["lab", "tutorial", "exam", "unknown"]);
  assert.deepEqual(staffSlotEntries(week.days[1], slots[2]).map((entry) => entry.id), ["other-day"]);
});

test("calendar exports contain only assigned exams, retain zero-load duties and use UAE times independently of the host", () => {
  const model = buildExportModel(fixture());
  const files = buildStaffCalendarFiles(model);
  assert.equal(files.length, 1, "Backup-only, idle and disabled staff must not get a calendar");
  assert.match(files[0].filename, /Fall 26-27 Staff Duties - Alice\.ics$/);
  assert.equal(files[0].mimeType, "text/calendar;charset=utf-8");
  const data = unfold(files[0].data);
  assert.equal((data.match(/BEGIN:VEVENT/g) || []).length, 3);
  assert.match(data, /DTSTART:20261019T050000Z\r\nDTEND:20261019T060000Z/);
  assert.match(data, /DTSTART:20261027T080000Z/);
  assert.match(data, /zero extra load/);
  assert.ok(!/TUTORIAL|LECTURE|UNCHECKED|Training topic|Backup|PRIVATE-STUDENT/.test(data));
  assert.ok(!/ATTENDEE|ORGANIZER|METHOD:REQUEST/.test(data), "Downloads are calendar imports, not emailed meeting invitations");
  const later = buildInvigilatorCalendar(model.workloads[0], model.duties, new Date("2027-01-01T00:00:00Z"));
  const uids = (value) => unfold(value).match(/^UID:.*$/gm);
  assert.deepEqual(uids(later), uids(files[0].data), "Repeat exports must keep stable event IDs");
  assert.equal((unfold(buildInvigilatorCalendar(model.workloads[0], [...model.duties, model.duties[0]])).match(/BEGIN:VEVENT/g) || []).length, 3);
  const lateDuty = { ...model.duties[0], start: 1380, end: 1440 };
  assert.match(unfold(buildInvigilatorCalendar(model.workloads[0], [lateDuty])), /DTEND:20261019T200000Z/);
  assert.match(unfold(buildInvigilatorCalendar(model.workloads[0], [{ ...lateDuty, start: 60, end: 120 }])), /DTSTART:20261018T210000Z/);
});

test("iCalendar text is escaped and folded by UTF-8 bytes without splitting Unicode characters or allowing extra properties", () => {
  const model = buildExportModel(fixture());
  const person = { ...model.workloads[0], name: "Instructor, A; B\\C\r\nBEGIN:VEVENT" };
  const title = "\u4e2d\u6587 \u00e9 ".repeat(45);
  const data = buildInvigilatorCalendar(person, [{ ...model.duties[0], title }]);
  assert.equal((data.match(/^BEGIN:VEVENT$/gm) || []).length, 1);
  assert.ok(data.endsWith("END:VCALENDAR\r\n"));
  assert.ok(data.split("\r\n").every((line) => Buffer.byteLength(line, "utf8") <= 75));
  assert.ok(!data.includes("\ufffd"));
  const plain = unfold(data);
  assert.ok(plain.includes("Instructor\\, A\\; B\\\\C\\nBEGIN:VEVENT"));
  assert.ok(plain.includes(title));
});

test("staff exports add one full-period calendar per invigilator, not one duplicate for each weekly workbook", () => {
  const options = fixture();
  for (const packaging of ["combined", "weekly"]) {
    const files = buildExportFiles({ ...options, report: "staff", packaging, includeCalendars: true });
    assert.equal(files.filter((file) => file.filename.endsWith(".xlsx")).length, packaging === "weekly" ? 2 : 1);
    const calendars = files.filter((file) => file.filename.endsWith(".ics"));
    assert.equal(calendars.length, 1);
    assert.equal((unfold(calendars[0].data).match(/BEGIN:VEVENT/g) || []).length, 3);
  }
  const week2 = buildExportFiles({ ...options, report: "staff", weeks: [2], includeCalendars: true });
  assert.equal((unfold(week2[1].data).match(/BEGIN:VEVENT/g) || []).length, 1);
  assert.match(week2[1].filename, /Week 2 - Alice\.ics$/);
  assert.ok(!unfold(week2[1].data).includes("20261019"));
  assert.equal(reportFileName({ report: "staff", startDate: options.startDate, format: "zip", includeCalendars: true }),
    "Midterm Exam Schedule Fall 26-27 Staff Duties With Calendars.zip");
});

test("staff schedule Excel exports styled weekly boards for all enabled staff with selected schedule layers only", () => {
  const options = { ...fixture(), report: "staff", staffView: "schedules" };
  const [file] = buildExportFiles(options);
  assert.equal(file.filename, "Midterm Exam Schedule Fall 26-27 Staff Schedules.xlsx");
  const book = read(file);
  assert.equal(book.SheetNames.length, 6);
  assert.ok(book.SheetNames.every((name) => !name.includes("Excluded")));
  assert.ok(!/TUTORIAL|LECTURE|UNCHECKED|Training topic|PRIVATE-STUDENT/.test(workbookText(book)));
  assert.equal(book.Sheets[book.SheetNames[0]].A1.v, "Alice / Week 1");
  assert.equal(book.Sheets[book.SheetNames[0]]["!merges"].length, 2);
  assert.ok(book.Sheets[book.SheetNames[0]].A1.s.fgColor || book.Sheets[book.SheetNames[0]].A1.s);
  const layered = buildExportFiles({ ...options, includeTeaching: true, includeCyberSkills: true, includeCalendars: true });
  const text = workbookText(read(layered[0]));
  assert.match(text, /TUTORIAL/);
  assert.match(text, /UNCHECKED/);
  assert.match(text, /Training topic/);
  assert.match(text, /Time not listed/);
  assert.ok(!text.includes("PRIVATE-STUDENT"));
  assert.ok(!unfold(layered[1].data).includes("Training topic"));
});

test("staff schedule CSV and individual downloads respect the requested staff, weeks and layers but never preview filters", () => {
  const options = { ...fixture(), report: "staff", staffView: "schedules", includeTeaching: true, includeCyberSkills: true };
  const files = buildExportFiles({ ...options, staffIds: ["A"], includeCalendars: true });
  assert.match(files[0].filename, /Staff Schedules - Alice\.xlsx$/);
  assert.equal(read(files[0]).SheetNames.length, 2);
  assert.equal(files.length, 2);
  const [csv] = buildExportFiles({ ...options, format: "csv", weeks: [2], staffIds: ["A"] });
  assert.match(csv.data, /"Invigilator","Week","Date","Day","Start","End","Activity"/);
  assert.match(csv.data, /TUTORIAL/);
  assert.ok(!/Training topic|2026-10-19|"Bob"|"Idle"|PRIVATE-STUDENT/.test(csv.data));
  const examOnly = buildExportFiles({ ...options, staffView: "duties", format: "csv" })[0];
  assert.ok(!/TUTORIAL|UNCHECKED|Training topic/.test(examOnly.data));
  assert.throws(() => buildExportFiles({ ...options, staffIds: [] }), /included invigilator/);
  assert.throws(() => buildExportFiles({ ...options, staffIds: ["D"] }), /included invigilator/);
  assert.throws(() => buildExportFiles({ ...options, staffView: "duties", staffIds: ["A"] }), /only available in Staff schedules/);
  assert.throws(() => buildExportFiles({ ...options, report: "overview", includeCalendars: true }), /only available for Staff duties/);
});

test("calendar filenames and staff sheet names remain safe and unique with duplicate, long and formula-like names", () => {
  const options = fixture();
  const long = "=A/name:[with]*?characters " + "long ".repeat(20);
  options.catalog.invigilators[0].name = long;
  options.catalog.invigilators[1].name = long.toUpperCase();
  options.plan.fingerprint = emptyResourcePlan(options.sessions, options.catalog).fingerprint;
  const session = options.sessions[2];
  options.plan.allocations[session.rooms[0].id].invigilatorIds = ["B"];
  options.plan.backups[session.id] = ["A"];
  const files = buildExportFiles({ ...options, report: "staff", staffView: "schedules", includeCalendars: true });
  const book = read(files[0]);
  assert.equal(new Set(book.SheetNames.map((name) => name.toLowerCase())).size, book.SheetNames.length);
  assert.ok(book.SheetNames.every((name) => name.length <= 31 && !/[\\/?*[\]:]/.test(name)));
  const calendars = files.slice(1);
  assert.equal(calendars.length, 2);
  assert.equal(new Set(calendars.map((file) => file.filename.toLowerCase())).size, 2);
  assert.ok(calendars.every((file) => file.filename.length <= 184 && !/[<>:"/\\|?*]/.test(file.filename)));
  assert.equal(book.Sheets[book.SheetNames[0]].A1.t, "s");
  assert.equal(book.Sheets[book.SheetNames[0]].A1.f, undefined);
});
