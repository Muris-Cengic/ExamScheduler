import test from "node:test";
import assert from "node:assert/strict";
import * as XLSX from "xlsx/xlsx.mjs";
import { assignResources, buildExamSessions, emptyResourcePlan } from "../src/resources.js";
import { roomIdentity } from "../src/department.js";
import { buildAsdOverviewExams, buildCombinedResourceWorkbook, buildExportFiles, buildExportModel, formatRoomDisplayName, reportDate, reportTable } from "../src/exportReports.js";

const course = (id, count, options = {}) => ({
  id, code: id, title: id + " course", crns: ["101"], labSessions: [],
  students: Array.from({ length: count }, (_, i) => ({ id: id + "-S" + i, name: "Student " + i, crn: "101" })), ...options,
});
function fixture() {
  const labRoom = roomIdentity("Campus", "Building", "Lab");
  const lab = { days: ["Monday"], crn: "101", startMinutes: 480, endMinutes: 590, room: "Lab", building: "Building", campus: "Campus", labInstructorId: "I0" };
  const courses = { A: course("A", 16, { labSessions: [lab] }), B: course("B", 26) };
  const sessions = buildExamSessions({ 1: { Monday: { "09:00": ["A"] } }, 2: { Tuesday: { "12:00": ["B"] } }, 3: {} }, courses, 60);
  const catalog = {
    rooms: Array.from({ length: 3 }, (_, i) => ({ id: i === 0 ? labRoom : "R" + i, name: "Building / Room " + i, enabled: true, busy: [] })),
    invigilators: Array.from({ length: 6 }, (_, i) => ({ id: "I" + i, name: "Invigilator " + i, enabled: true, busy: [] })),
  };
  const busy = { code: "A", crn: "101", isLab: true, days: ["Monday"], start: 480, end: 590 };
  catalog.rooms[0].busy = [busy];
  catalog.invigilators[0].busy = [busy];
  return { sessions, catalog, plan: assignResources(sessions, catalog), startDate: "2026-10-19",
    templateHeaders: { invigilatorHeader: ["CRN", "Code", "Title", "Students", "Date", "Time", "Instructor", "Room"],
      studentHeader: ["CRN", "Code", "Title", "ID", "Name", "Room", "Notes"] } };
}
const read = (file) => XLSX.read(file.data, { type: "array", cellDates: true });
const rows = (workbook, name) => XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1 });

test("invalid Friday placements block every main report without restricting ASD overview references", () => {
  const options = fixture();
  for (const time of ["08:00", "09:30", "10:00", "12:00", "17:00"]) {
    const sessions = buildExamSessions({ 1: { Friday: { [time]: ["MAIN"] } } }, { MAIN: course("MAIN", 5) }, 60);
    const plan = assignResources(sessions, options.catalog);
    assert.ok(sessions[0].issues.includes("Friday exams must use 09:00-10:00 or 10:30-11:30."));
    for (const report of ["complete", "overview", "staff", "students"]) {
      assert.throws(() => buildExportFiles({ ...options, sessions, plan, report }), /Complete valid resource/);
    }
  }
  for (const time of ["09:00", "10:30"]) {
    const sessions = buildExamSessions({ 1: { Friday: { [time]: ["MAIN"] } } }, { MAIN: course("MAIN", 5) }, 60);
    const plan = assignResources(sessions, options.catalog);
    assert.equal(buildExportFiles({ ...options, sessions, plan, report: "overview" }).length, 1);
  }
  const asdExams = buildAsdOverviewExams({ assignments: { 1: { Friday: { "12:00": ["ASD"] } } },
    courseLookup: { ASD: course("ASD", 5) } });
  const overview = buildExportFiles({ ...options, asdExams, includeAsd: true, report: "overview", format: "csv" })[0];
  assert.match(overview.data, /"Friday","12:00","13:00","ASD"/);
});

function asdFixture() {
  return buildAsdOverviewExams({
    assignments: { 1: { Monday: { "09:00": ["ASD-1", "ASD-1"], "08:00": ["ASD-2"] } },
      3: { Wednesday: { "10:30": ["ASD-1"] } } },
    courseLookup: { "ASD-1": course("ASD-1", 99), "ASD-2": course("ASD-2", 5) },
    examDurations: { "ASD-1": 90 }, defaultDuration: 75,
  });
}

test("ASD overview references retain imported durations without copying students or resources", () => {
  const references = asdFixture();
  assert.equal(references.length, 3, "Duplicate IDs in a slot are not duplicated in the overview");
  assert.deepEqual(references.map((exam) => [exam.week, exam.start, exam.end]), [[1, 540, 630], [1, 480, 555], [3, 630, 720]]);
  assert.ok(references.every((exam) => exam.isAsd && exam.studentCount === null &&
    exam.primaryInvigilatorsNeeded === null && exam.duringLab === null && !exam.roomNames.length));
  assert.ok(!JSON.stringify(references).includes("Student"));
  assert.equal(new Set(references.map((exam) => exam.id)).size, 3);
  const input = { assignments: { 1: { Monday: { "09:00": ["ASD-1"] } } }, courseLookup: { "ASD-1": course("ASD-1", 1) } };
  const before = JSON.stringify(input);
  assert.equal(buildAsdOverviewExams(input)[0].end, 600);
  assert.equal(JSON.stringify(input), before);
  assert.throws(() => buildAsdOverviewExams({ ...input, courseLookup: {} }), /Invalid ASD/);
  assert.throws(() => buildAsdOverviewExams({ ...input, examDurations: { "ASD-1": -1 } }), /Invalid ASD/);
  assert.throws(() => buildAsdOverviewExams({ ...input, assignments: { 1: { Monday: { "25:00": ["ASD-1"] } } } }), /Invalid ASD/);
});

test("ASD opt-in extends only the overview model, including ASD-only weeks and chronological reference entries", () => {
  const options = { ...fixture(), asdExams: asdFixture() };
  const before = JSON.stringify(options);
  const normal = buildExportModel(options);
  const included = buildExportModel({ ...options, includeAsd: true });
  assert.deepEqual(normal.selectedWeeks, [1, 2]);
  assert.deepEqual(normal.overviewExams, normal.exams);
  assert.deepEqual(included.selectedWeeks, [1, 2, 3]);
  assert.deepEqual(included.overviewExams.map((exam) => exam.code), ["ASD-2", "A", "ASD-1", "B", "ASD-1"],
    "Earlier times come first, with department exams before ASD at an identical start time");
  for (const key of ["summary", "exams", "roomRows", "students", "duties", "workloads", "sessions"]) {
    assert.deepEqual(included[key], normal[key], key + " stays department-only");
  }
  const week3 = buildExportModel({ ...options, includeAsd: true, weeks: [3] });
  assert.deepEqual(week3.summary, { exams: 0, roomSessions: 0, students: 0, sittings: 0 });
  assert.equal(week3.overviewExams[0].date, "2026-11-04");
  assert.equal(week3.overviewExams[0].end, 720);
  assert.ok(week3.workloads.every((person) => person.exam + person.backup + person.teaching === 0));
  assert.throws(() => buildExportModel({ ...options, weeks: [3] }), /Only weeks with scheduled/);
  assert.throws(() => buildExportModel({ ...options, includeAsd: true, weeks: [4] }), /Only weeks with scheduled/);
  assert.throws(() => buildExportModel({ ...options, includeAsd: true, weeks: [3], plan: emptyResourcePlan(options.sessions, options.catalog) }), /Complete valid resource/);
  assert.equal(JSON.stringify(options), before, "Reference display cannot modify resource assignments");
});

test("overview Excel and CSV export ASD references only when opted in, in combined and weekly files", () => {
  const options = { ...fixture(), asdExams: asdFixture(), report: "overview" };
  const excluded = rows(read(buildExportFiles(options)[0]), "Exam overview");
  assert.equal(excluded.length, 3);
  assert.ok(excluded.slice(1).every((row) => row.at(-1) === "Department"));
  for (const format of ["xlsx", "csv"]) {
    for (const packaging of ["combined", "weekly"]) {
      const files = buildExportFiles({ ...options, includeAsd: true, format, packaging });
      assert.equal(files.length, packaging === "weekly" ? 3 : 1);
      if (packaging === "weekly") assert.equal(files[2].filename, "Week_3_Exam_Overview." + format);
      const exported = files.flatMap((file) => {
        const table = format === "xlsx" ? rows(read(file), "Exam overview")
          : XLSX.utils.sheet_to_json(XLSX.read(file.data, { type: "string", raw: true }).Sheets.Sheet1, { header: 1 });
        assert.equal(table[0].at(-1), "Schedule");
        return table.slice(1);
      });
      assert.equal(exported.length, 5);
      const references = exported.filter((row) => row[13] === "ASD (reference)");
      assert.equal(references.length, 3);
      assert.ok(references.every((row) => row.slice(8, 13).every((value) => value === undefined || value === "")),
        "ASD resource and student totals are blank, not misleading zeroes");
      assert.ok(!JSON.stringify(exported).includes("Student"));
      assert.ok(!JSON.stringify(exported).includes("Invigilator"));
      assert.ok(references.some((row) => Number(row[0]) === 3 && row[3] === "10:30" && row[4] === "12:00"));
    }
  }
  const week3 = read(buildExportFiles({ ...options, includeAsd: true, weeks: [3] })[0]);
  assert.equal(rows(week3, "Exam overview").length, 2);
});

test("ASD inclusion cannot affect complete, staff or student reports, or their weekly exports", () => {
  const options = { ...fixture(), asdExams: asdFixture(), includeAsd: true };
  for (const report of ["complete", "staff", "students"]) {
    for (const packaging of ["combined", "weekly"]) {
      const baseline = buildExportFiles({ ...options, report, packaging, includeAsd: false });
      const files = buildExportFiles({ ...options, report, packaging });
      assert.deepEqual(files.map((file) => file.filename), baseline.map((file) => file.filename));
      for (const [index, file] of files.entries()) {
        const workbook = read(file);
        const original = read(baseline[index]);
        assert.deepEqual(workbook.SheetNames, original.SheetNames);
        for (const name of workbook.SheetNames) assert.deepEqual(rows(workbook, name), rows(original, name));
        assert.ok(!JSON.stringify(workbook.Sheets).includes("ASD-"));
      }
    }
    assert.throws(() => buildExportFiles({ ...options, report, weeks: [3] }), /Only weeks with scheduled/);
    if (report !== "complete") assert.equal(buildExportFiles({ ...options, report, format: "csv" })[0].data,
      buildExportFiles({ ...options, report, format: "csv", includeAsd: false })[0].data);
  }
  const direct = buildCombinedResourceWorkbook(options);
  assert.deepEqual(rows(direct, "Schedule Index").slice(1).map((row) => row[0]), [1, 2]);
  assert.ok(!direct.SheetNames.some((name) => name.startsWith("Week 3")));
});

test("export model uses confirmed resources, excludes empty weeks and counts sittings separately", () => {
  const options = fixture();
  const model = buildExportModel(options);
  assert.deepEqual(model.selectedWeeks, [1, 2]);
  assert.deepEqual(model.summary, { exams: 2, roomSessions: 3, students: 42, sittings: 42 });
  assert.deepEqual(model.exams.map((exam) => exam.studentCount), [16, 26]);
  assert.deepEqual(model.exams.map((exam) => exam.primaryInvigilatorsNeeded), [2, 2]);
  assert.deepEqual(model.exams.map((exam) => exam.duringLab), [true, false]);
  assert.deepEqual(model.roomRows.map((room) => room.studentCount), [16, 13, 13]);
  assert.deepEqual(model.exams.map((exam) => exam.date), ["2026-10-19", "2026-10-27"]);
  const teaching = model.duties.find((duty) => duty.invigilatorId === "I0" && duty.week === 1);
  assert.equal(teaching.extraLoad, 0);
  assert.equal(teaching.teaching, true);
  assert.ok(model.duties.filter((duty) => duty.role === "Backup").every((duty) => duty.roomName === "Slot standby"));
  assert.equal(model.duties.filter((duty) => duty.role === "Backup").length, 2, "Backups appear once per slot, not once per room");
  const week2 = buildExportModel({ ...options, weeks: [2, 2] });
  assert.deepEqual(week2.selectedWeeks, [2]);
  assert.equal(week2.students.length, 26);
  assert.ok(week2.duties.every((duty) => duty.week === 2));
  assert.equal(reportDate("2026-12-28", 2, "Friday"), "2027-01-08", "Calendar dates cross year boundaries without host-timezone shifts");
});

test("combined complete report preserves every weekly formula and adds a navigable schedule index", () => {
  const options = fixture();
  const workbook = buildCombinedResourceWorkbook(options);
  assert.equal(workbook.SheetNames[0], "Schedule Index");
  assert.deepEqual(rows(workbook, "Schedule Index").slice(1).map((row) => row[0]), [1, 2]);
  assert.equal(workbook.Sheets["Schedule Index"].F2.l.Target, "#'Week 1 Invigilators'!A1");
  for (const week of [1, 2]) {
    assert.ok(workbook.Sheets["Week " + week + " Invigilators"]);
    const day = week === 1 ? "Monday" : "Tuesday";
    const daily = workbook.Sheets["Week " + week + " " + day];
    assert.equal(daily.F2.f, "'Week " + week + " Invigilators'!H2");
    const staff = workbook.Sheets["Week " + week + " Invigilator Pool"];
    assert.match(staff.B2.f, new RegExp("Week " + week + " Invigilators"));
  }
  // Round-trip the file, checking that every referenced sheet actually exists.
  const restored = read(buildExportFiles(options)[0]);
  for (const sheet of Object.values(restored.Sheets)) {
    for (const cell of Object.values(sheet)) {
      if (cell?.f) for (const match of cell.f.matchAll(/'([^']+)'!/g)) assert.ok(restored.Sheets[match[1]], "Missing referenced sheet: " + match[1]);
    }
  }
  assert.equal(restored.Sheets["Week 1 Invigilator Pool"].B2.v, 0);
  assert.equal(restored.Sheets["Week 1 Invigilator Pool"].D2.v, 1);
});

test("complete Excel packaging supports one workbook, weekly files and non-contiguous selected weeks", () => {
  const options = fixture();
  const combined = buildExportFiles(options);
  assert.equal(combined.length, 1);
  assert.equal(combined[0].filename, "Exam_Schedule.xlsx");
  const weekly = buildExportFiles({ ...options, packaging: "weekly" });
  assert.deepEqual(weekly.map((file) => file.filename), ["Week_1_Exam_Schedule.xlsx", "Week_2_Exam_Schedule.xlsx"]);
  for (const [index, file] of weekly.entries()) {
    assert.ok(read(file).SheetNames.every((name) => name.startsWith("Week " + (index + 1) + " ")));
  }
  const week2 = read(buildExportFiles({ ...options, weeks: [2] })[0]);
  assert.ok(!week2.SheetNames.some((name) => name.startsWith("Week 1 ")));
  assert.equal(week2.Sheets["Schedule Index"].A2.v, 2);
  assert.ok(!week2.SheetNames.some((name) => name.startsWith("Week 3 ")));
});

test("focused workbooks contain the right audience's data and staff has a separate workload summary", () => {
  const options = fixture();
  const overview = read(buildExportFiles({ ...options, report: "overview" })[0]);
  assert.deepEqual(overview.SheetNames, ["Exam overview"]);
  const schedule = rows(overview, "Exam overview");
  assert.equal(schedule.length, 3);
  assert.ok(!schedule[0].some((header) => /student name|student id|invigilator name/i.test(header)));
  assert.deepEqual(schedule[0].slice(-3), ["Primary invigilators needed", "During lab time", "Schedule"]);
  assert.deepEqual(schedule.slice(1).map((row) => row.slice(-3)), [[2, "Yes", "Department"], [2, "No", "Department"]]);
  assert.equal(overview.Sheets["Exam overview"].L2.t, "n", "The requirement remains a numeric Excel value");
  assert.ok(!JSON.stringify(schedule.slice(1)).includes("Invigilator 0"), "Counts do not expose staff names");
  assert.equal(schedule[1][1].toISOString().slice(0, 10), "2026-10-19");
  const staff = read(buildExportFiles({ ...options, report: "staff", weeks: [1] })[0]);
  assert.deepEqual(staff.SheetNames, ["Staff duties", "Workload Summary"]);
  const load = rows(staff, "Workload Summary").find((row) => row[0] === "Invigilator 0");
  assert.deepEqual(load.slice(1, 5), [0, 0, 1, 0]);
  assert.equal(rows(staff, "Staff duties").filter((row) => row[6] === "Backup").length, 1);
  const students = read(buildExportFiles({ ...options, report: "students", weeks: [2] })[0]);
  const roster = rows(students, "Student room lists");
  assert.equal(roster.length, 27);
  assert.equal(new Set(roster.slice(1).map((row) => row[5])).size, 26);
  assert.ok(roster.slice(1).every((row) => row[0] === 2 && row[8] === "B" && row[10].startsWith("Building / Room")));
  assert.ok(!roster[0].includes("Invigilator"));
});

test("overview room labels remove standalone PAD tokens and compact separators without changing room identities", () => {
  for (const [input, expected] of [
    ["PAD / P-B-4F / 13", "P-B-4F/13"],
    [" pad / P-B-4F / 13 ", "P-B-4F/13"],
    ["P-B-4F/13", "P-B-4F/13"],
    ["PAD / P-W-4F / CISCO LAB", "P-W-4F/CISCO LAB"],
    ["Campus / PADDOCK / 13", "Campus/PADDOCK/13"],
    ["Room 13", "Room 13"], [null, ""],
  ]) assert.equal(formatRoomDisplayName(input), expected);
  const options = fixture();
  options.catalog.rooms[0].name = "PAD / P-B-4F / 13";
  options.plan = assignResources(options.sessions, options.catalog);
  const before = JSON.stringify(options);
  const model = buildExportModel(options);
  assert.deepEqual(model.exams[0].roomNames, ["P-B-4F/13"]);
  assert.equal(model.roomRows[0].roomName, "PAD / P-B-4F / 13");
  assert.ok(model.students.filter((student) => student.week === 1).every((student) => student.roomName === "PAD / P-B-4F / 13"));
  assert.ok(model.duties.filter((duty) => duty.week === 1 && duty.role === "Exam").every((duty) => duty.roomName === "PAD / P-B-4F / 13"));
  const overview = read(buildExportFiles({ ...options, report: "overview" })[0]);
  assert.equal(overview.Sheets["Exam overview"].K2.v, "P-B-4F/13");
  const csv = buildExportFiles({ ...options, report: "overview", format: "csv", weeks: [1] })[0].data;
  assert.match(csv, /"Primary invigilators needed","During lab time"/);
  assert.match(csv, /"P-B-4F\/13","2","Yes"/);
  assert.ok(!csv.includes("PAD"));
  const complete = read(buildExportFiles(options)[0]);
  assert.equal(complete.Sheets["Week 1 Invigilators"].H2.v, "PAD / P-B-4F / 13");
  assert.equal(complete.Sheets["Week 1 Monday"].F2.f, "'Week 1 Invigilators'!H2");
  assert.equal(complete.Sheets["Week 1 Monday"].F2.v, "PAD / P-B-4F / 13");
  assert.equal(JSON.stringify(options), before, "Display formatting cannot mutate saved resources or assignments");
});

test("overview primary requirements follow staffing-aware rooms and approved consolidation, not extra load or backup duties", () => {
  const catalog = {
    rooms: Array.from({ length: 3 }, (_, i) => ({ id: "R" + i, name: "Room " + i, enabled: true, busy: [] })),
    invigilators: Array.from({ length: 8 }, (_, i) => ({ id: "I" + i, name: "Staff " + i, enabled: true, busy: [] })),
  };
  for (const [count, choice, expectedSizes, needed] of [
    [15, "standard", [15], 1], [16, "standard", [16], 2],
    [26, "standard", [13, 13], 2], [31, "standard", [16, 15], 3],
    [32, "standard", [17, 15], 3], [55, "standard", [20, 20, 15], 5],
    [52, "standard", [19, 18, 15], 5], [66, "standard", [22, 22, 22], 6],
    [52, "distribute", [26, 26], 4],
  ]) {
    const sessions = buildExamSessions({ 1: { Monday: { "12:00": ["A"] } } }, { A: course("A", count) }, 60, 25, { A: choice });
    const plan = assignResources(sessions, catalog);
    const model = buildExportModel({ sessions, catalog, plan, startDate: "2026-10-19" });
    assert.deepEqual(sessions[0].rooms.map((room) => room.students.length), expectedSizes);
    assert.equal(model.exams[0].primaryInvigilatorsNeeded, needed, count + " students / " + choice);
    assert.equal(model.duties.filter((duty) => duty.role === "Exam").length, needed);
    assert.equal(model.duties.filter((duty) => duty.role === "Backup").length, 1, "A backup is not part of the exam's primary requirement");
  }
  const labModel = buildExportModel(fixture());
  assert.equal(labModel.exams[0].primaryInvigilatorsNeeded, 2, "A zero-load lab instructor still covers a required room duty");
  assert.equal(labModel.duties.filter((duty) => duty.week === 1 && duty.role === "Exam").reduce((sum, duty) => sum + duty.extraLoad, 0), 1);
});

test("lab indicators identify only the replaced course, including all its overflow rooms in a shared slot", () => {
  const options = fixture();
  const lab = { days: ["Monday"], crn: "101", startMinutes: 480, endMinutes: 590,
    room: "Lab", building: "Building", campus: "Campus", labInstructorId: "I0" };
  const sessions = buildExamSessions({ 1: { Monday: { "09:00": ["A", "B"] } } },
    { A: course("A", 31, { labSessions: [lab] }), B: course("B", 16) }, 60);
  const plan = assignResources(sessions, options.catalog);
  const model = buildExportModel({ ...options, sessions, plan });
  assert.deepEqual(model.exams.map((exam) => [exam.code, exam.roomNames.length, exam.primaryInvigilatorsNeeded, exam.duringLab]),
    [["A", 2, 3, true], ["B", 1, 2, false]]);
  assert.equal(model.exams[0].end, 600, "The approved 09:50 lab ending still qualifies for a 09:00-10:00 exam");
  assert.equal(reportTable("overview", model).rows[1].at(-2), "No", "A simultaneous non-lab exam cannot inherit the lab marker");
});

test("CSV exports include selected weeks, preserve Unicode and quotes, and neutralise spreadsheet formula injection", () => {
  const options = fixture();
  const student = options.sessions[1].rooms[0].students[0];
  student.name = '=HYPERLINK("https://example.invalid", "click")';
  student.id = "+123";
  options.sessions[1].rooms[1].students[0].name = "Muri\u0161, \"Student\"";
  options.plan = assignResources(options.sessions, options.catalog);
  const files = buildExportFiles({ ...options, report: "students", format: "csv", packaging: "weekly", weeks: [2] });
  assert.equal(files.length, 1);
  assert.equal(files[0].filename, "Week_2_Student_Room_Lists.csv");
  assert.equal(files[0].data[0], "\uFEFF");
  assert.ok(files[0].data.includes("\"'=HYPERLINK(\"\"https://example.invalid\"\", \"\"click\"\")\""));
  assert.ok(files[0].data.includes("\"'+123\""));
  assert.ok(files[0].data.includes('Muri\u0161, ""Student""'));
  const overview = buildExportFiles({ ...options, report: "overview", format: "csv" })[0];
  assert.equal(overview.data.split("\r\n").filter(Boolean).length, 3);
  assert.ok(!overview.data.includes("HYPERLINK"));
  assert.equal(reportTable("staff", buildExportModel(options)).rows.filter((row) => row[6] === "Backup").length, 2);
});

test("all export paths reject incomplete resources and invalid or empty selections", () => {
  const options = fixture();
  for (const report of ["complete", "overview", "staff", "students"]) {
    assert.throws(() => buildExportFiles({ ...options, report, plan: emptyResourcePlan(options.sessions, options.catalog) }), /Complete valid resource/);
  }
  assert.throws(() => buildExportFiles({ ...options, weeks: [] }), /Select at least one/);
  assert.throws(() => buildExportFiles({ ...options, weeks: [3] }), /Only weeks with scheduled/);
  assert.throws(() => buildExportFiles({ ...options, report: "complete", format: "csv" }), /multi-sheet/);
  assert.throws(() => buildExportFiles({ ...options, report: "unknown" }), /Invalid export/);
  assert.throws(() => buildExportFiles({ ...options, packaging: "unknown" }), /Invalid export/);
  assert.throws(() => buildExportModel({ ...options, startDate: "invalid" }), /valid exam start date/);
});
