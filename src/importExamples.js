export function importExample(kind, startDate = "2026-10-19") {
  const spreadsheets = { formats: "Excel (.xlsx, .xls) or CSV", accept: ".xlsx,.xls,.csv" };
  if (kind === "enrolment") return {
    ...spreadsheets, title: "Student Enrollment", action: "Upload Enrollment File",
    columns: ["Student ID", "Student Name", "Course Code", "Course Title", "CRN"],
    rows: [["S001", "Student One", "ICT-1201", "Computer Networks", "101"]],
    note: "Student registration reports and Banner exports are also supported.",
  };
  if (kind === "crn") return {
    ...spreadsheets, title: "CRN Info", action: "Upload CRN List",
    columns: ["CRN", "Course Code", "Title", "DAYS", "Time", "Type", "Primary Instructor", "Second Instructor", "Campus", "Building", "Room"],
    rows: [["101", "ICT-1201", "Computer Networks", "M", "1300 - 1450", "OL", "10: Alex", "20: Sam", "PAD", "P-B-4F", "13"]],
    note: "First sheet only. OL = lab; the second instructor teaches it. Days: M/T/W/R/F.",
  };
  if (kind === "asd") return {
    title: "ASD Schedule", action: "Load ASD Excel / JSON", formats: "Excel (.xlsx, .xls), CSV or saved ASD JSON",
    columns: ["Course", "Department", "Day", "Time"],
    rows: [["ICT-1201", "ASD", startDate, "12:00 - 1:00 PM"]],
    note: "Optional. Course code or title must match enrollment. Day must include the exam date; only ASD rows are used.",
  };
  throw new Error("Unknown import example.");
}
