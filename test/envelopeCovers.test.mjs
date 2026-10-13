import test from "node:test";
import assert from "node:assert/strict";
import { inflateSync } from "node:zlib";
import { jsPDF } from "jspdf";
import * as XLSX from "xlsx/xlsx.mjs";
import { assignResources, buildExamSessions, emptyResourcePlan } from "../src/resources.js";
import { buildEnvelopeCourses, buildEnvelopeCoverFiles, buildExportFiles, buildExportModel, envelopeCoverInfo, envelopeSessionInfo, reportFileName } from "../src/exportReports.js";
import { envelopeCoverLayout, envelopeCoverPdf } from "../src/envelopePdf.js";

function fixture() {
  const course = (code, total) => ({ id: code, code, title: code + " title", crns: ["00101", "00202"], labSessions: [],
    students: Array.from({ length: total }, (_, i) => ({ id: "PRIVATE-ID-" + i, name: "PRIVATE-NAME-" + i, crn: i < 28 ? "00101" : "00202" })) });
  const sessions = buildExamSessions({ 1: { Monday: { "12:00": ["A"] }, Tuesday: { "12:00": ["B"] } },
    2: { Wednesday: { "12:00": ["A"] } } }, { A: course("A", 57), B: course("B", 7) }, 60);
  const catalog = {
    rooms: Array.from({ length: 5 }, (_, i) => ({ id: "R" + i, name: "PAD / P-B-4F / " + (13 + i), enabled: true, busy: [] })),
    invigilators: Array.from({ length: 10 }, (_, i) => ({ id: "I" + i, name: "PRIVATE-STAFF-" + i, enabled: true, busy: [] })),
  };
  return { sessions, catalog, plan: assignResources(sessions, catalog), startDate: "2026-10-19",
    asdExams: [{ id: "ASD", code: "ASD", title: "ASD excluded", isAsd: true, week: 3, day: "Monday", start: 720, end: 780 }], includeAsd: true };
}

function pdfContents(data) {
  const bytes = Buffer.from(data);
  const source = bytes.toString("latin1");
  return [...source.matchAll(/stream\r?\n/g)].map((match) => {
    const start = match.index + match[0].length;
    const end = source.indexOf("endstream", start);
    try { return inflateSync(bytes.subarray(start, end)).toString("latin1"); } catch { return ""; }
  }).join("\n");
}

test("envelope pages use reviewed room rosters, counts and per-room CRNs without personal data or ASD", () => {
  const options = fixture();
  const before = JSON.stringify(options);
  const model = buildExportModel(options);
  const courses = buildEnvelopeCourses(model);
  assert.deepEqual(courses.map((course) => course.code), ["A", "B"]);
  assert.equal(courses[0].pages.length, 6, "Three rooms in each of two weeks belong to one course PDF");
  assert.deepEqual([...new Set(courses[0].pages.map((page) => page.week))], [1, 2]);
  for (const page of courses.flatMap((course) => course.pages)) {
    const room = model.roomRows.find((row) => row.id === page.id);
    assert.equal(page.studentCount, room.studentCount);
    assert.deepEqual(page.crns, room.crns);
    assert.equal(page.roomName, room.roomName.replace("PAD / ", "").replaceAll(" / ", "/"));
    assert.equal("students" in page, false);
    assert.equal("invigilators" in page, false);
  }
  assert.deepEqual(envelopeCoverInfo(courses[1].pages[0]), [
    ["Course Title", "B title"], ["CRNs", "00101"], ["# of students", "7"],
    ["# of exam copies", "", "Include one extra copy."], ["Instructor Name", ""],
  ]);
  assert.deepEqual(envelopeSessionInfo(courses[1].pages[0]), [
    ["Date", "20/10/2026"], ["Time", "12:00-13:00"], ["Location", courses[1].pages[0].roomName.replace("/", "\n")],
  ]);
  assert.ok(!JSON.stringify(courses).includes("PRIVATE"));
  assert.ok(!JSON.stringify(courses).includes("ASD"));
  assert.equal(JSON.stringify(options), before);
});

test("cover locations separate the building and room into two lines without changing room identities", () => {
  for (const [roomName, location] of [
    ["P-B-4F/13", "P-B-4F\n13"], ["PAD / P-B-4F / 13", "P-B-4F\n13"],
    ["Campus / Building / Lab", "Campus Building\nLab"], ["Room 1", "Room 1"], ["", ""],
  ]) {
    const page = { date: "2026-10-19", start: 720, end: 780, roomName };
    assert.equal(envelopeSessionInfo(page)[2][1], location);
    assert.equal(page.roomName, roomName);
  }
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const page = buildEnvelopeCourses(buildExportModel(fixture()))[0].pages[0];
  const location = envelopeCoverLayout(doc, page).session[2];
  assert.deepEqual(location.lines, page.roomName.split("/"));
  assert.ok(!pdfContents(envelopeCoverPdf([page])).includes("P-B-4F/"));
});

test("envelope PDF files group rooms by course, respect selected weeks and leave Instructor Name blank", async () => {
  const options = fixture();
  const files = await buildEnvelopeCoverFiles(options);
  assert.deepEqual(files.map((file) => file.filename), ["A - A title.pdf", "B - B title.pdf"]);
  assert.ok(files.every((file) => file.mimeType === "application/pdf"));
  for (const [index, file] of files.entries()) {
    const bytes = Buffer.from(file.data).toString("latin1");
    assert.ok(bytes.startsWith("%PDF-"));
    assert.equal([...bytes.matchAll(/\/Type \/Page\b/g)].length, index === 0 ? 6 : 1);
    const dimensions = bytes.match(/MediaBox \[0 0 ([\d.]+) ([\d.]+)\]/);
    assert.ok(dimensions && Math.abs(Number(dimensions[1]) - 595.28) < 0.01 && Math.abs(Number(dimensions[2]) - 841.89) < 0.01);
    const content = pdfContents(file.data);
    for (const label of ["Course Title", "CRNs", "# of students", "# of exam copies", "Instructor Name", "Date", "Time", "Location"]) assert.ok(content.includes(label), label);
    assert.ok(content.includes("Include one extra copy."));
    for (const unwanted of ["Department", "PRIVATE-ID", "PRIVATE-NAME", "PRIVATE-STAFF", "ASD excluded", "PAD / "]) assert.ok(!content.includes(unwanted), unwanted);
  }
  const weekly = await buildEnvelopeCoverFiles({ ...options, weeks: [2], courseCodes: ["A"] });
  assert.equal(weekly.length, 1);
  assert.equal([...Buffer.from(weekly[0].data).toString("latin1").matchAll(/\/Type \/Page\b/g)].length, 3);
  assert.ok(!pdfContents(weekly[0].data).includes("19/10/2026"));
  assert.ok(pdfContents(weekly[0].data).includes("28/10/2026"));
  assert.equal(reportFileName({ report: "envelopes", startDate: options.startDate, format: "zip" }),
    "Midterm Exam Schedule Fall 26-27 Envelope Covers.zip");
});

test("envelope files have safe unique course names after replacing Windows path characters", async () => {
  const options = fixture();
  options.sessions.forEach((session) => session.rooms.forEach((room) => { room.code = session.week === 1 ? "A/B" : "A\\B"; room.title = "Same <Title>?"; }));
  options.plan = assignResources(options.sessions, options.catalog);
  const files = await buildEnvelopeCoverFiles(options);
  assert.deepEqual(files.map((file) => file.filename), ["A_B - Same _Title__.pdf", "A_B - Same _Title___2.pdf"]);
});

test("envelope packing workbooks contain dated session sheets, course totals and individual room counts", () => {
  const options = fixture();
  const files = buildExportFiles({ ...options, report: "envelopes", examStartDate: "2026-08-01" });
  assert.equal(files[0].filename, "Midterm Exam Schedule Fall 26-27 Envelope Covers.xlsx");
  const workbook = XLSX.read(files[0].data, { type: "array", cellStyles: true });
  assert.deepEqual(workbook.SheetNames, ["2026-10-19 1200-1300", "2026-10-20 1200-1300", "2026-10-28 1200-1300"]);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  assert.deepEqual(rows.slice(0, 5), [["Date", "19/10/2026"], ["Time", "12:00-13:00"], [], ["Course / room", "Students"], ["A - A title", 57]]);
  assert.equal(rows.slice(5).reduce((total, row) => total + row[1], 0), 57);
  assert.equal(rows.length, 8);
  assert.ok(sheet.A1.s.fgColor.rgb.endsWith("D8E7E2"));
  assert.ok(!JSON.stringify(rows).includes("PRIVATE"));
  const weekly = buildExportFiles({ ...options, report: "envelopes", packaging: "weekly" });
  assert.equal(weekly.length, 2);
  assert.ok(weekly[1].filename.endsWith("Week 2.xlsx"));
  assert.deepEqual(XLSX.read(weekly[1].data, { type: "array" }).SheetNames, ["2026-10-28 1200-1300"]);
});

test("envelope export rejects invalid resources, unsupported formats and selections outside included weeks", async () => {
  const options = fixture();
  await assert.rejects(buildEnvelopeCoverFiles({ ...options, weeks: [3] }), /Only weeks with scheduled exams/);
  for (const courseCodes of [[], ["ASD"], ["B"], "A"]) {
    await assert.rejects(buildEnvelopeCoverFiles({ ...options, weeks: [2], courseCodes }), /Choose a course/);
  }
  await assert.rejects(buildEnvelopeCoverFiles({ ...options, format: "csv" }), /course PDF/);
  await assert.rejects(buildEnvelopeCoverFiles({ ...options, packaging: "weekly" }), /course PDF/);
  const plan = emptyResourcePlan(options.sessions, options.catalog);
  await assert.rejects(buildEnvelopeCoverFiles({ ...options, plan }), /Complete valid resource/);
  assert.throws(() => buildExportFiles({ ...options, report: "envelopes", plan }), /Complete valid resource/);
  assert.throws(() => buildExportFiles({ ...options, report: "envelopes", format: "csv" }), /packing list uses Excel/);
  assert.throws(() => envelopeCoverPdf([]), /No envelope covers/);
});

test("A4 envelope layout wraps long values and shrinks only when necessary without clipping", () => {
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const page = buildEnvelopeCourses(buildExportModel(fixture()))[0].pages[0];
  const layouts = [page, { ...page, title: "Advanced Information Security and Computer Network Architecture ".repeat(5),
    crns: Array.from({ length: 45 }, (_, i) => "10" + i), roomName: "P-B-4F/13 " + "Long room name ".repeat(5) }]
    .map((item) => envelopeCoverLayout(doc, item));
  assert.equal(layouts[0].scale, 1);
  assert.ok(layouts[1].fields[0].lines.length > 1);
  for (const layout of layouts) {
    const counts = layout.fieldRows[2];
    assert.equal(counts.fields.length, 2);
    assert.deepEqual(counts.fields.map((field) => field.label), ["# of students", "# of exam copies"]);
    assert.deepEqual(counts.fields[1].lines, [" "]);
    assert.deepEqual(counts.fields[1].noteLines, ["Include one extra copy."]);
    assert.equal(counts.fieldWidth * 2 + 16, layout.width);
    const contentHeight = layout.titleHeight + layout.fieldRows.reduce((sum, row) => sum + row.height + 22 * layout.scale, 0) + layout.sessionHeight;
    assert.ok(contentHeight <= doc.internal.pageSize.getHeight() - 88);
  }
  assert.throws(() => envelopeCoverLayout(doc, { ...page, title: "unreasonably long ".repeat(2000) }), /too long to fit/);
});
