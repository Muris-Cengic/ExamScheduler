import test from "node:test";
import assert from "node:assert/strict";
import * as XLSX from "xlsx/xlsx.mjs";
import { importExample } from "../src/importExamples.js";
import { parseAsdWorkbook, parseEnrolmentWorkbook } from "../src/imports.js";
import { parseDepartmentWorkbook, scopeDepartmentCourses } from "../src/department.js";
import { parseResourceCatalog } from "../src/resources.js";

function workbook(example, format) {
  const sheet = XLSX.utils.aoa_to_sheet([example.columns, ...example.rows]);
  if (format === "csv") return XLSX.read(XLSX.utils.sheet_to_csv(sheet), { type: "string" });
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, "Example");
  return XLSX.read(XLSX.write(book, { type: "array", bookType: "xlsx" }), { type: "array" });
}

test("displayed enrollment and CRN examples are accepted and populate courses, labs and resources", () => {
  for (const format of ["xlsx", "csv"]) {
    const enrolment = parseEnrolmentWorkbook(workbook(importExample("enrolment"), format));
    const department = parseDepartmentWorkbook(workbook(importExample("crn"), format));
    const scope = scopeDepartmentCourses(enrolment.courses, department.courses);
    assert.equal(enrolment.courses[0].students[0].id, "S001");
    assert.equal(scope.courses[0].code, "ICT-1201");
    assert.deepEqual(scope.missingCrns, []);
    const lab = scope.courses[0].labSessions[0];
    assert.deepEqual(lab.days, ["Monday"]);
    assert.equal(lab.startMinutes, 780);
    assert.equal(lab.endMinutes, 890);
    assert.equal(lab.labInstructorId, "staff:20");
    const catalog = parseResourceCatalog(department);
    assert.equal(catalog.rooms.length, 1);
    assert.deepEqual(catalog.invigilators.map((person) => person.name).sort(), ["Alex", "Sam"]);
  }
});

test("displayed ASD example uses the selected timetable date and matches the enrollment example", () => {
  for (const format of ["xlsx", "csv"]) {
    const enrolment = parseEnrolmentWorkbook(workbook(importExample("enrolment"), format));
    const example = importExample("asd", "2027-01-11");
    const parsed = parseAsdWorkbook(workbook(example, format), enrolment.courses, 2027);
    assert.deepEqual(parsed.unmatched, []);
    assert.deepEqual(parsed.exams, [{ courseId: "ICT-1201", date: "2027-01-11", startMinutes: 720, endMinutes: 780 }]);
  }
  assert.throws(() => importExample("unknown"), /Unknown import example/);
});
