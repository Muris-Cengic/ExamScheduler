import test from "node:test";
import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import * as XLSX from "xlsx/xlsx.mjs";
import {
  continuingCourses, parseAsdWorkbook, parseEnrolmentWorkbook,
  parseExamTimeRange, planAsdImport,
} from "../src/imports.js";

function workbook(rows, name = "Input") {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), name);
  return book;
}

const enrolment = () => workbook([
  ["", "Crn", "Course Code", "Title", "Pidm"],
  ["", "STUDENT_ID : A001, Student Name : Student One, Program Code : AB-ISET", "", "", ""],
  ["", 101, "MATH-1010", "Calculus I", 999],
  ["", 101, "MATH-1010", "Calculus I", 999],
  ["", 102, "ICT-1011", "Intr. to Prog.&Problem Solving", 999],
  ["Sum", "", "", "", ""],
  ["", "STUDENT_ID : A002, Student Name : Student Two, Program Code : AB-ISET", "", "", ""],
  ["", 103, "MATH-1010", "Calculus I", 888],
]);

const schedule = (rows) => workbook([
  ["Timeslot", "Course", "Department", "No. of Students", "Day", "Time"],
  ...rows,
], "Schedule");

test("grouped reports associate registrations with student blocks and retain CRNs", () => {
  const parsed = parseEnrolmentWorkbook(enrolment(), 1);
  assert.equal(parsed.courses.length, 2);
  assert.deepEqual(parsed.studentDirectory, { A001: "Student One", A002: "Student Two" });
  const calculus = parsed.courses.find((course) => course.id === "MATH-1010");
  assert.equal(calculus.studentCount, 2);
  assert.equal(calculus.roomsNeeded, 2);
  assert.deepEqual(calculus.crns, ["101", "103"]);
  assert.equal(calculus.crnDetails[0].students.length, 1);
  assert.equal(calculus.primaryInstructor, "");
  assert.deepEqual(calculus.students.map((student) => student.id), ["A001", "A002"]);
});

test("flat CSV-style tables and Banner headers still retain instructor and section data", () => {
  const rows = [
    ["Student ID", "Student Name", "Course Code", "Course Title", "CRN", "Section", "Instructor Name"],
    ["A001", "Student One", "ICT-1011", "Programming", "101", "01", "Instructor One"],
    ["A002", "Student Two", "ICT-1011", "Programming", "102", "02", "Instructor Two"],
  ];
  const flat = parseEnrolmentWorkbook(workbook(rows)).courses[0];
  assert.deepEqual(flat.sections, ["01", "02"]);
  assert.deepEqual(flat.instructors, ["Instructor One", "Instructor Two"]);
  const csv = XLSX.utils.sheet_to_csv(XLSX.utils.aoa_to_sheet(rows));
  assert.equal(parseEnrolmentWorkbook(XLSX.read(csv, { type: "string" })).courses[0].studentCount, 2);
  const banner = workbook([
    ["SPRIDEN_ID", "STUDENT_NAME", "SSBSECT_SUBJ_CODE", "SSBSECT_CRSE_NUMB", "SCBCRSE_TITLE", "SSBSECT_CRN", "CF_INSTRUCTOR"],
    ["A003", "Student Three", "ICT", "1011", "Programming", "103", "Instructor Three"],
  ]);
  assert.equal(parseEnrolmentWorkbook(banner).courses[0].id, "ICT-1011");
});

test("identity resets at sheet boundaries and subtotals cannot invent registrations", () => {
  const book = enrolment();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
    ["", "Crn", "Course Code", "Title", "Pidm"],
    ["", 111, "MATH-1010", "Calculus I", 123],
  ]), "Other");
  assert.throws(() => parseEnrolmentWorkbook(book), /without a student identity/);
  assert.throws(() => parseEnrolmentWorkbook(workbook([["A", "B"], [1, 2]])), /No enrolments found/);
});

test("ASD imports match titles, ignore other departments and flag unmatched courses", () => {
  const { courses } = parseEnrolmentWorkbook(enrolment());
  const parsed = parseAsdWorkbook(schedule([
    [1, "calculUs I", "ASD", 999, "Monday, 19 October", "12:00 - 1:00 PM"],
    [2, "Intr to Prog & Problem Solving", "ASD", 999, "Tuesday, 20 October", "10:30 - 11:30 AM"],
    [3, "Chemistry I", "ASD", 999, "Wednesday, 21 October", "12:00 - 1:00 PM"],
    [4, "Calculus I", "EMET", 999, "Friday, 23 October", "5:00 - 6:00 PM"],
  ]), courses, 2026);
  assert.equal(parsed.exams.length, 2);
  assert.equal(parsed.exams[0].date, "2026-10-19");
  assert.equal(parsed.exams[0].startMinutes, 720);
  assert.equal(parsed.exams[0].endMinutes, 780);
  assert.deepEqual(parsed.unmatched, ["Chemistry I"]);
  assert.equal(courses[0].studentCount, 2, "schedule headcount must not replace the loaded roster");
});

test("time ranges resolve noon, afternoon, midnight and meridiem transitions", () => {
  assert.deepEqual(parseExamTimeRange("12:00 - 1:00 PM"), { startMinutes: 720, endMinutes: 780 });
  assert.deepEqual(parseExamTimeRange("09:00 - 10:00 AM"), { startMinutes: 540, endMinutes: 600 });
  assert.deepEqual(parseExamTimeRange("5:00 - 6:00 PM"), { startMinutes: 1020, endMinutes: 1080 });
  assert.deepEqual(parseExamTimeRange("11:00 - 12:00 PM"), { startMinutes: 660, endMinutes: 720 });
  assert.deepEqual(parseExamTimeRange("11:30 AM - 1:00 PM"), { startMinutes: 690, endMinutes: 780 });
  assert.deepEqual(parseExamTimeRange("12:00 AM - 1:00 AM"), { startMinutes: 0, endMinutes: 60 });
  assert.throws(() => parseExamTimeRange("12:60 - 1:00 PM"), /Invalid exam time/);
  assert.throws(() => parseExamTimeRange("2:00 PM - 1:00 PM"), /end after/);
});

test("ambiguous or conflicting course matches are rejected; identical rows are deduplicated", () => {
  const courses = [{ id: "A", code: "A", title: "Calculus I" }];
  const row = [1, "Calculus I", "ASD", 1, "Monday, 19 October", "12:00 - 1:00 PM"];
  assert.equal(parseAsdWorkbook(schedule([row, row]), courses, 2026).exams.length, 1);
  assert.throws(() => parseAsdWorkbook(schedule([row]), [
    ...courses, { id: "B", code: "B", title: "Calculus I" },
  ], 2026), /several course codes/);
  assert.throws(() => parseAsdWorkbook(schedule([
    row, [2, "Calculus I", "ASD", 1, "Tuesday, 20 October", "12:00 - 1:00 PM"],
  ]), courses, 2026), /different exam sessions/);
});

test("missing-year dates validate the selected year; explicit and Excel dates are accepted", () => {
  const courses = [{ id: "A", code: "A", title: "Calculus I" }];
  const make = (day) => schedule([[1, "Calculus I", "ASD", 1, day, "12:00 - 1:00 PM"]]);
  assert.throws(() => parseAsdWorkbook(make("Monday, 19 October"), courses, 2025), /Check the exam start date/);
  assert.equal(parseAsdWorkbook(make("Monday, 19 October 2026"), courses, 2025).exams[0].date, "2026-10-19");
  assert.equal(parseAsdWorkbook(make(46314), courses, 2026).exams[0].date, "2026-10-19");
  assert.throws(() => parseAsdWorkbook(make("2026-02-31"), courses, 2026), /Invalid exam date/);
});

const settings = {
  slotIntervalMinutes: 60, startHour: 8, endHour: 17,
  examDurationMinutes: 120, studentsPerRoom: 25,
  invigilatorCount: 15,
};
test("calendar plans extend hours/weeks and retain imported durations independently of main duration", () => {
  const plan = planAsdImport([
    { courseId: "A", date: "2026-10-19", startMinutes: 630, endMinutes: 690 },
    { courseId: "B", date: "2026-10-30", startMinutes: 1020, endMinutes: 1080 },
  ], settings);
  assert.equal(plan.startDate, "2026-10-19");
  assert.deepEqual(plan.weeks, [1, 2]);
  assert.equal(plan.settings.slotIntervalMinutes, 30);
  assert.equal(plan.settings.endHour, 18);
  assert.equal(plan.settings.examDurationMinutes, 120);
  assert.deepEqual(plan.placements.map(({ week, day, slotId }) => [week, day, slotId]), [
    [1, "Monday", "10:30"], [2, "Friday", "17:00"],
  ]);
  assert.deepEqual(plan.examDurations, { A: 60, B: 60 });
});

test("existing main calendar stays anchored and unsupported dates/times fail explicitly", () => {
  const exam = { courseId: "A", date: "2026-10-26", startMinutes: 720, endMinutes: 780 };
  assert.equal(planAsdImport([exam], settings, [1, 3], "2026-10-19").placements[0].week, 2);
  assert.deepEqual(planAsdImport([exam], settings, [1, 3], "2026-10-19").weeks, [1, 2, 3]);
  assert.throws(() => planAsdImport([{ ...exam, date: "2026-10-18" }], settings), /Monday-Friday/);
  assert.throws(() => planAsdImport([{ ...exam, startMinutes: 725 }], settings), /30-minute/);
  assert.throws(() => planAsdImport([exam], settings, [1], "2026-11-02"), /outside the timetable range/);
});

test("imported ASD exams stop continuing at their actual end, even with two-hour main exams", () => {
  const slots = ["12:00", "12:30", "13:00", "13:30", "14:00"].map((id) => ({ id }));
  const day = { "12:00": ["ASD", "MAIN"] };
  assert.deepEqual(continuingCourses(day, slots, 1, 4, { ASD: 2 }), ["ASD", "MAIN"]);
  assert.deepEqual(continuingCourses(day, slots, 2, 4, { ASD: 2 }), ["MAIN"]);
  assert.deepEqual(continuingCourses(day, slots, 4, 4, { ASD: 2 }), []);
});

const enrolmentPath = "data/input/26-27 S1/Students Registration 26-27 S1.xlsx";
const asdPath = "data/input/26-27 S1/Other Departments Schedule/Midterm Schedule 26-27 S1.xlsx";
test("supplied 26-27 S1 workbooks import together", {
  skip: !existsSync(enrolmentPath) || !existsSync(asdPath),
}, () => {
  const parsed = parseEnrolmentWorkbook(XLSX.read(readFileSync(enrolmentPath), { type: "buffer" }));
  assert.equal(parsed.courses.length, 59);
  assert.equal(Object.keys(parsed.studentDirectory).length, 536);
  assert.equal(new Set(parsed.courses.flatMap((course) => course.crns)).size, 195);
  const asd = parseAsdWorkbook(XLSX.read(readFileSync(asdPath), { type: "buffer" }), parsed.courses, 2026);
  assert.equal(asd.exams.length, 13);
  assert.deepEqual(asd.unmatched, ["Calculus II", "Applied Mathematics", "Chemistry I", "Fluid Flow & Heat Transfer"]);
  const plan = planAsdImport(asd.exams, settings);
  assert.equal(plan.startDate, "2026-10-19");
  assert.deepEqual(plan.weeks, [1, 2]);
  assert.equal(plan.settings.slotIntervalMinutes, 30);
  assert.equal(plan.placements.find((exam) => exam.courseId === "MATH-2012").slotId, "10:30");
  assert.ok(Object.values(plan.examDurations).every((minutes) => minutes === 60));
});
