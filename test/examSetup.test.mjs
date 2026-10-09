import test from "node:test";
import assert from "node:assert/strict";
import { inferAcademicTerm } from "../src/examSetup.js";

test("semester and academic year follow an August-to-July academic calendar", () => {
  for (const year of [2026, 2027]) {
    for (let month = 1; month <= 12; month += 1) {
      const date = year + "-" + String(month).padStart(2, "0") + "-15";
      const start = month >= 8 ? year : year - 1;
      assert.deepEqual(inferAcademicTerm(date), {
        semester: month >= 8 ? "Fall" : month >= 6 ? "Summer" : "Spring",
        academicYear: start + "-" + (start + 1),
      }, date);
    }
  }
});

test("term boundaries use the entered date, with Fall taking precedence throughout August", () => {
  for (const [date, semester, academicYear] of [
    ["2026-05-31", "Spring", "2025-2026"], ["2026-06-01", "Summer", "2025-2026"],
    ["2026-07-31", "Summer", "2025-2026"], ["2026-08-01", "Fall", "2026-2027"],
    ["2026-08-31", "Fall", "2026-2027"], ["2026-09-01", "Fall", "2026-2027"],
    ["2026-12-31", "Fall", "2026-2027"], ["2027-01-01", "Spring", "2026-2027"],
    ["2028-02-29", "Spring", "2027-2028"],
  ]) assert.deepEqual(inferAcademicTerm(date), { semester, academicYear }, date);
});

test("invalid dates cannot produce misleading academic terms", () => {
  for (const value of [undefined, null, 2026, "", "not-a-date", "2026-2-1", "2026-02-29", "2026-04-31", "2026-13-01", "2026-00-10", "2026-01-00", "2026-01-01T00:00:00Z"]) {
    assert.equal(inferAcademicTerm(value), null, String(value));
  }
});
