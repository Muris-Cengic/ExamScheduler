import test from "node:test";
import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import { createContext, runInContext } from "node:vm";
import { transformWithEsbuild } from "vite";
import * as XLSX from "xlsx/xlsx.mjs";
import JSZip from "jszip";
import * as imports from "../src/imports.js";
import * as department from "../src/department.js";
import { autoSchedule } from "../src/autoSchedule.js";
import * as resources from "../src/resources.js";
import * as examRooms from "../src/examRooms.js";
import { buildResourceWorkbookForWeek } from "../src/reports.js";
import * as exportReports from "../src/exportReports.js";

// Exercise the real App event handlers/render tree without a browser or new test dependencies.
const appSource = readFileSync(new URL("../src/App.jsx", import.meta.url), "utf8")
  .replace(/^import .*;\r?$/gm, "")
  .replace(/import\.meta\.url/g, JSON.stringify(new URL("../src/App.jsx", import.meta.url).href))
  .replace("export default App;", "globalThis.App = App;");
const { code } = await transformWithEsbuild(appSource, "App.jsx", { loader: "jsx", jsxFactory: "h", jsxFragment: "Fragment" });
const resourceSource = readFileSync(new URL("../src/ResourceAssignment.jsx", import.meta.url), "utf8")
  .replace(/^import .*;\r?$/gm, "")
  .replace("export default function ResourceAssignment", "function ResourceAssignment") + "\nglobalThis.ResourceAssignment = ResourceAssignment;";
const { code: resourceComponentCode } = await transformWithEsbuild(resourceSource, "ResourceAssignment.jsx", { loader: "jsx", jsxFactory: "h", jsxFragment: "Fragment" });
const exportSource = readFileSync(new URL("../src/ExportStudio.jsx", import.meta.url), "utf8")
  .replace(/^import .*;\r?$/gm, "")
  .replace("export default function ExportStudio", "function ExportStudio") + "\nglobalThis.ExportStudio = ExportStudio;";
const { code: exportComponentCode } = await transformWithEsbuild(exportSource, "ExportStudio.jsx", { loader: "jsx", jsxFactory: "h", jsxFragment: "Fragment" });

function harness() {
  const hooks = [];
  const downloads = [];
  const downloadNames = [];
  let printCount = 0;
  let index = 0;
  let dirty = false;
  let effects = [];
  class TestURL extends URL {
    static createObjectURL(blob) { downloads.push(blob); return "blob:test"; }
    static revokeObjectURL() {}
  }
  const context = createContext({
    ...imports, ...department, ...resources, ...examRooms, ...exportReports, buildResourceWorkbookForWeek, autoSchedule, XLSX, JSZip, Blob, URL: TestURL, console,
    CourseSelection: function CourseSelection() {}, Fragment: "fragment",
    ResourceAssignment: function ResourceAssignment() {},
    h: (type, props, ...children) => ({ type, props: props || {}, children: typeof type === "function" && type.name !== "CourseSelection" ? [type(props)] : children.flat(Infinity) }),
    useState(initial) {
      const current = index++;
      hooks[current] ??= { value: typeof initial === "function" ? initial() : initial };
      return [hooks[current].value, (value) => {
        const previous = hooks[current].value;
        hooks[current].value = typeof value === "function" ? value(previous) : value;
        dirty ||= !Object.is(previous, hooks[current].value);
      }];
    },
    useRef(initial) {
      const current = index++;
      hooks[current] ??= { current: initial };
      return hooks[current];
    },
    useMemo(factory, dependencies) {
      const current = index++;
      const previous = hooks[current];
      if (!previous || dependencies.some((value, i) => !Object.is(value, previous.dependencies[i]))) {
        hooks[current] = { dependencies, value: factory() };
      }
      return hooks[current].value;
    },
    useEffect(effect, dependencies) {
      const current = index++;
      const previous = hooks[current];
      if (!previous || dependencies.some((value, i) => !Object.is(value, previous.dependencies[i]))) {
        effects.push(effect);
        hooks[current] = { dependencies };
      }
    },
    window: { addEventListener() {}, removeEventListener() {}, print() { printCount += 1; } },
    document: { body: { appendChild() {}, removeChild() {} }, createElement: () => ({ click() { downloadNames.push(this.download); } }) },
    fetch: async () => ({ ok: true, arrayBuffer: async () => buffer("data/ReportReference/Report Template.xlsx") }),
  });
  runInContext(resourceComponentCode, context);
  runInContext(exportComponentCode, context);
  runInContext(code, context);
  return {
    downloads,
    downloadNames,
    prints: () => printCount,
    render() {
      let tree;
      let attempts = 0;
      do {
        dirty = false;
        index = 0;
        effects = [];
        tree = context.App();
        effects.forEach((effect) => effect());
        assert.ok(++attempts < 10, "Render must settle");
      } while (dirty);
      return tree;
    },
  };
}

const buffer = (path) => {
  const bytes = readFileSync(path);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
};
const fileEvent = (path) => ({ target: { files: [{ name: path.split("/").pop(), arrayBuffer: async () => buffer(path) }], value: path } });
const jsonEvent = (value) => ({ target: { files: [{ name: "snapshot.json", text: async () => JSON.stringify(value) }], value: "snapshot.json" } });
const text = (node) => typeof node === "string" || typeof node === "number" ? String(node) : node?.children?.map(text).join("") || "";
const nodes = (node) => node && typeof node === "object" ? [node, ...(node.children || []).flatMap(nodes)] : [];
const find = (tree, predicate) => {
  const result = nodes(tree).find(predicate);
  assert.ok(result, "Expected control exists");
  return result;
};
const button = (tree, label) => find(tree, (node) => node.type === "button" && text(node) === label);
const review = (tree) => find(tree, (node) => node.type?.name === "CourseSelection");
const root = "data/input/26-27 S1/";
const enrolmentPath = root + "Students Registration 26-27 S1.xlsx";
const crnPath = root + "CRN List 26-27 S1.xlsx";
const asdPath = root + "Other Departments Schedule/Midterm Schedule 26-27 S1.xlsx";

test("wizard imports, exam toggles, draft, saved-state round trip and report exclusion work together", {
  skip: [enrolmentPath, crnPath, asdPath].some((path) => !existsSync(path)),
}, async () => {
  const app = harness();
  let tree = app.render();
  const enrolmentLabel = find(tree, (node) => node.type === "label" && text(node) === "Upload Enrollment File");
  await find(enrolmentLabel, (node) => node.type === "input").props.onChange(fileEvent(enrolmentPath));
  tree = app.render();
  assert.equal(button(tree, "Continue To ASD").props.disabled, true);
  await review(tree).props.onUpload(fileEvent(crnPath));
  tree = app.render();
  assert.equal(review(tree).props.courses.length, 37);
  assert.equal(review(tree).props.selection.sheetName, "ISET CRNs");
  assert.ok(!("onSheetChange" in review(tree).props));
  assert.ok(!("sheetNames" in review(tree).props));
  const octCourses = review(tree).props.courses.filter((course) => /\bOCT\b/i.test(course.title));
  assert.equal(octCourses.length, 4);
  octCourses.forEach((course) => assert.equal(review(tree).props.examChoices[course.id], false));
  const octId = octCourses[0].id;
  review(tree).props.onExamChange(octId, true);
  tree = app.render();
  assert.equal(review(tree).props.examChoices[octId], true, "OCT defaults can be overridden");
  review(tree).props.onExamChange(octId, false);
  tree = app.render();
  const firstId = review(tree).props.courses.find((course) => department.defaultHasExam(course)).id;
  review(tree).props.onExamChange(firstId, false);
  button(tree, "Continue To ASD").props.onClick();
  tree = app.render();
  const asdInput = find(tree, (node) => node.type === "input" && node.props.accept === ".json,.xlsx,.xls,.csv");
  await asdInput.props.onChange(fileEvent(asdPath));
  tree = app.render();
  button(tree, "Auto-Schedule Remaining Exams").props.onClick();
  tree = app.render();
  assert.match(text(tree), /25 exams added. 0 remain unplaced/);
  const conflicts = find(tree, (node) => node.props.className === "conflicts conflicts--full");
  assert.match(text(conflicts), /No conflicts detected/);
  button(tree, "Save Timetable").props.onClick();
  const snapshot = JSON.parse(await app.downloads.at(-1).text());
  assert.equal(snapshot.version, 5);
  assert.deepEqual(snapshot.roomDistributionChoices, {});
  assert.equal(snapshot.resourceCatalog.rooms.length, 11);
  assert.equal(snapshot.resourceCatalog.invigilators.length, 15);
  const firstSheetCrns = new Set(snapshot.departmentSelection.courses.flatMap((course) => course.crns));
  assert.ok([...snapshot.resourceCatalog.rooms, ...snapshot.resourceCatalog.invigilators].every((resource) => resource.busy.every((entry) => firstSheetCrns.has(entry.crn))), "Saved availability comes only from the first sheet");
  assert.equal(snapshot.departmentSelection.courses.length, 37);
  assert.equal(snapshot.examChoices[firstId], false);
  assert.equal(department.assignmentIds(snapshot.assignments).size, 25);
  assert.ok(snapshot.assignments[1].Wednesday["09:00"]?.includes("ISET-4001") || snapshot.assignments[2].Wednesday["09:00"]?.includes("ISET-4001"));
  assert.ok(snapshot.assignments[1].Thursday["09:00"]?.includes("NCS-3202") || snapshot.assignments[2].Thursday["09:00"]?.includes("NCS-3202"));
  octCourses.forEach((course) => assert.equal(department.assignmentIds(snapshot.assignments).has(course.id), false));
  assert.equal(department.assignmentIds(snapshot.asdAssignments).size, 13);
  assert.equal(snapshot.settings.endHour, 18);
  assert.equal(department.assignmentIds(snapshot.assignments).has(firstId), false);

  const reloaded = harness();
  let loadedTree = reloaded.render();
  await find(loadedTree, (node) => node.type === "input" && node.props.accept === "application/json").props.onChange(jsonEvent(snapshot));
  loadedTree = reloaded.render();
  button(loadedTree, "Review Department Exams").props.onClick();
  loadedTree = reloaded.render();
  assert.equal(review(loadedTree).props.examChoices[firstId], false);
  assert.equal(review(loadedTree).props.courses.reduce((sum, course) => sum + course.labSessions.length, 0), 46);
  const assignedId = [...department.assignmentIds(snapshot.assignments)].find((id) => id !== "ISET-4001");
  review(loadedTree).props.onExamChange(assignedId, false);
  button(loadedTree, "Continue To ASD").props.onClick();
  loadedTree = reloaded.render();
  button(loadedTree, "Continue To Main").props.onClick();
  loadedTree = reloaded.render();
  button(loadedTree, "Save Timetable").props.onClick();
  const afterToggle = JSON.parse(await reloaded.downloads.at(-1).text());
  assert.equal(department.assignmentIds(afterToggle.assignments).size, 24);
  assert.equal(department.assignmentIds(afterToggle.assignments).has(assignedId), false);
  button(loadedTree, "Proceed To Resources").props.onClick();
  loadedTree = reloaded.render();
  assert.equal(button(loadedTree, "Proceed To Export").props.disabled, true);
  find(loadedTree, (node) => node.type?.name === "ResourceAssignment").props.onAssign();
  loadedTree = reloaded.render();
  const fullResources = find(loadedTree, (node) => node.type?.name === "ResourceAssignment").props;
  assert.ok(!text(loadedTree).includes("Specialist"));
  assert.ok(!text(loadedTree).includes("Primary"));
  assert.match(text(loadedTree), /During Teaching Hours/);
  const pinned = fullResources.sessions.flatMap((session) => session.rooms).find((room) => room.fixedInvigilatorId);
  const leadSelect = find(loadedTree, (node) => node.type === "select" && node.props["aria-label"] === `Invigilator 1 for ${pinned.id}`);
  assert.ok(leadSelect.children.filter((node) => node.type === "option" && node.props.value && node.props.value !== pinned.fixedInvigilatorId).every((node) => node.props.disabled));
  assert.ok(!fullResources.catalog.invigilators.some((person) => person.busy.some((entry) => /CSTP-1011/.test(entry.code))), "Second-sheet CSTP commitments must not enter the resource pool");
  const isetSession = fullResources.sessions.find((session) => session.rooms.some((room) => room.code === "ISET-4001" && room.fixedInvigilatorId));
  assert.ok(isetSession);
  assert.equal(isetSession.slotId, "09:00");
  assert.equal(isetSession.end, 600);
  assert.deepEqual(isetSession.issues, []);
  const isetRoom = isetSession.rooms.find((room) => room.code === "ISET-4001");
  const isetInstructor = fullResources.catalog.invigilators.find((person) => person.id === isetRoom.fixedInvigilatorId);
  assert.equal(resources.resourceBusyReason(isetInstructor, isetSession, fullResources.sessions), "", "Second-sheet CSTP commitments no longer block the ISET-4001 lab instructor");
  assert.equal(resources.isTeachingTimeDuty(isetInstructor, isetSession, fullResources.sessions), true, "The 09:50-to-10:00 allowance stays within the replaced lab duty");
  assert.ok(!fullResources.validation.issues.some((issue) => /CSTP-1011/.test(issue.message)));
  // Excluding the fixed lab instructor gives a deterministic issue to exercise the issue cards.
  fullResources.onPoolChange("invigilators", pinned.fixedInvigilatorId, { enabled: false });
  loadedTree = reloaded.render();
  const issuePanel = find(loadedTree, (node) => node.props.className === "resource-issues");
  assert.match(text(issuePanel), /How to resolve:/);
  assert.match(text(issuePanel), /Lab instructor unavailable/);
  assert.ok(!text(issuePanel).includes("1/Wednesday/08:00"), "Internal session IDs must not be shown as issue descriptions");
  const issueLinks = nodes(issuePanel).filter((node) => node.type === "a");
  assert.ok(issueLinks.length > 0);
  issueLinks.forEach((link) => {
    const target = decodeURIComponent(link.props.href.slice(1));
    assert.ok(nodes(loadedTree).some((node) => node.props.id === target), "Each issue links to a real assignment row or time slot");
  });
  assert.equal(button(loadedTree, "Proceed To Export").props.disabled, true);
  find(loadedTree, (node) => node.type?.name === "ResourceAssignment").props.onPoolChange("invigilators", pinned.fixedInvigilatorId, { enabled: true });
  loadedTree = reloaded.render();

  // A small evening fixture isolates the successful resource/export/save round trip.
  const scoped = department.scopeDepartmentCourses(afterToggle.courses, afterToggle.departmentSelection.courses).courses;
  const eveningCourses = scoped.filter((course) => course.crns.length > 1 && afterToggle.examChoices[course.id]).slice(0, 2);
  const eveningSnapshot = { ...afterToggle, resourcePlan: null, assignments: { 1: {
    Monday: { "17:00": [eveningCourses[0].id] }, Tuesday: { "17:00": [eveningCourses[1].id] },
  } } };
  await find(loadedTree, (node) => node.type === "input" && node.props.accept === "application/json").props.onChange(jsonEvent(eveningSnapshot));
  loadedTree = reloaded.render();
  button(loadedTree, "Proceed To Resources").props.onClick();
  loadedTree = reloaded.render();
  find(loadedTree, (node) => node.type?.name === "ResourceAssignment").props.onAssign();
  loadedTree = reloaded.render();
  const assignedResources = find(loadedTree, (node) => node.type?.name === "ResourceAssignment").props;
  assert.equal(assignedResources.validation.complete, true, JSON.stringify(assignedResources.validation.issues));
  assignedResources.sessions.forEach((session) => {
    const courseIds = new Set(session.rooms.map((room) => room.courseId));
    courseIds.forEach((courseId) => {
      const sizes = session.rooms.filter((room) => room.courseId === courseId).map((room) => room.students.length);
      assert.ok(Math.max(...sizes) - Math.min(...sizes) <= 1);
    });
  });
  button(loadedTree, "Save Timetable").props.onClick();
  const savedResources = JSON.parse(await reloaded.downloads.at(-1).text());
  assert.ok(savedResources.resourcePlan.fingerprint);
  await find(loadedTree, (node) => node.type === "input" && node.props.accept === "application/json").props.onChange(jsonEvent(savedResources));
  loadedTree = reloaded.render();
  button(loadedTree, "Proceed To Resources").props.onClick();
  loadedTree = reloaded.render();
  assert.equal(button(loadedTree, "Proceed To Export").props.disabled, false);
  const ready = find(loadedTree, (node) => node.type?.name === "ResourceAssignment").props;
  const firstAllocation = ready.plan.allocations[ready.sessions[0].rooms[0].id];
  ready.onPoolChange("invigilators", firstAllocation.invigilatorIds[0], { enabled: false });
  loadedTree = reloaded.render();
  assert.equal(button(loadedTree, "Proceed To Export").props.disabled, true);
  find(loadedTree, (node) => node.type?.name === "ResourceAssignment").props.onPoolChange("invigilators", firstAllocation.invigilatorIds[0], { enabled: true });
  loadedTree = reloaded.render();
  assert.equal(button(loadedTree, "Proceed To Export").props.disabled, false);
  button(loadedTree, "Proceed To Export").props.onClick();
  loadedTree = reloaded.render();
  await button(loadedTree, "Export Timetable").props.onClick();
  const exportedBlob = reloaded.downloads.at(-1);
  let exportedBytes;
  if (exportedBlob.type.includes("spreadsheet")) {
    exportedBytes = [new Uint8Array(await exportedBlob.arrayBuffer())];
  } else {
    const zip = await JSZip.loadAsync(await exportedBlob.arrayBuffer());
    exportedBytes = await Promise.all(Object.values(zip.files).filter((file) => file.name.endsWith(".xlsx")).map((file) => file.async("uint8array")));
  }
  const exportedCourses = new Set();
  const asdCodes = new Set(snapshot.courses.filter((course) => department.assignmentIds(snapshot.asdAssignments).has(course.id)).map((course) => course.code));
  for (const bytes of exportedBytes) {
    const workbook = XLSX.read(bytes, { type: "array" });
    const invigilators = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames.find((name) => /Invigilators$/.test(name))], { header: 1 });
    invigilators.slice(1).forEach((row) => {
      assert.ok(!asdCodes.has(row[1]), "ASD excluded from invigilator reports");
      assert.ok(Number(row[3]) <= 25);
      assert.ok(row[8]);
      if (Number(row[3]) > 15) assert.ok(row[9]);
      assert.ok(!/^Room \d+$/.test(row[7]), "Real room names exported");
      assert.ok(!/^(Specialist )?Invigilator \d+$/.test(row[8]), "Real staff names exported");
    });
    const staffHeader = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames.find((name) => /Invigilator Pool$/.test(name))], { header: 1 })[0];
    assert.ok(!staffHeader.includes("Type"));
    assert.ok(!staffHeader.some((header) => /primary/i.test(header)));
    assert.ok(staffHeader.includes("Invigilation load"));
    assert.ok(staffHeader.includes("During teaching hours"));
    for (const name of workbook.SheetNames.filter((name) => /Monday|Tuesday|Wednesday|Thursday|Friday/.test(name))) {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1 });
      rows.slice(1).forEach((row) => {
        assert.ok(!asdCodes.has(row[1]), "ASD excluded from daily reports");
        assert.ok(snapshot.departmentSelection.courses.find((course) => course.code === row[1])?.crns.includes(row[0]), "Only selected CRNs are exported");
        exportedCourses.add(row[1]);
      });
    }
  }
  assert.equal(exportedCourses.size, 2);
});

test("room consolidation choices require a click, preserve other exams, survive save/load and appear in export", async () => {
  const rows = [["Student ID", "Student Name", "Course Code", "Course Title", "CRN"],
    ...Array.from({ length: 52 }, (_, index) => [`S${index}`, `Student ${index}`, "EXAM-1000", "Consolidation Exam", "101"]),
    ...Array.from({ length: 10 }, (_, index) => [`B${index}`, `Other Student ${index}`, "OTHER-1000", "Unrelated Exam", "201"]),
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "Enrolment");
  const enrolment = imports.parseEnrolmentWorkbook(workbook);
  const exam = enrolment.courses.find((course) => course.code === "EXAM-1000");
  const other = enrolment.courses.find((course) => course.code === "OTHER-1000");
  const snapshot = { type: "main", version: 5, startDate: "2026-10-19", weeks: [1], selectedWeek: 1,
    settings: { slotIntervalMinutes: 30, startHour: 8, endHour: 18, studentsPerRoom: 25, examDurationMinutes: 60 },
    assignments: { 1: { Monday: { "17:00": [exam.id, other.id] } } }, asdAssignments: {}, asdExamDurations: {}, hasAsdStep: false,
    courses: enrolment.courses, studentDirectory: enrolment.studentDirectory,
    departmentSelection: { sheetName: "First", courses: enrolment.courses.map((course) => ({ code: course.code, title: course.title, crns: course.crns, meetings: [] })) },
    examChoices: { [exam.id]: true, [other.id]: true }, roomDistributionChoices: {}, resourcePlan: null,
    resourceCatalog: {
      rooms: Array.from({ length: 4 }, (_, index) => ({ id: `R${index}`, name: `Building / ${index}`, enabled: true, busy: [] })),
      invigilators: Array.from({ length: 10 }, (_, index) => ({ id: `I${index}`, name: `Staff ${index}`, enabled: true, busy: [] })),
    },
  };
  const app = harness();
  let tree = app.render();
  await find(tree, (node) => node.type === "input" && node.props.accept === "application/json").props.onChange(jsonEvent(snapshot));
  tree = app.render();
  button(tree, "Proceed To Resources").props.onClick();
  tree = app.render();
  button(tree, "Auto-Assign Resources").props.onClick();
  tree = app.render();
  const panel = () => find(tree, (node) => node.type?.name === "ResourceAssignment").props;
  const examRooms = () => panel().sessions[0].rooms.filter((room) => room.courseId === exam.id);
  const decisionButton = (label) => find(tree, (node) => node.type === "button" && node.props["aria-label"] === `${label} for ${exam.code}`);
  assert.deepEqual(examRooms().map((room) => room.students.length), [18, 17, 17]);
  assert.equal(decisionButton("Keep the extra room (normal limit)").props["aria-pressed"], true);
  const otherRoom = panel().sessions[0].rooms.find((room) => room.courseId === other.id);
  const otherAllocation = panel().plan.allocations[otherRoom.id];
  decisionButton("Merge overflow into one room").props.onClick();
  tree = app.render();
  assert.deepEqual(examRooms().map((room) => room.students.length), [27, 25]);
  assert.equal(panel().validation.complete, true);
  assert.deepEqual(panel().plan.allocations[otherRoom.id], otherAllocation);
  assert.match(text(tree), /Exception approved for this exam/);
  decisionButton("Distribute overflow evenly").props.onClick();
  tree = app.render();
  assert.deepEqual(examRooms().map((room) => room.students.length), [26, 26]);
  assert.ok(examRooms().every((room) => room.requiredInvigilators === 2));
  button(tree, "Save Timetable").props.onClick();
  const saved = JSON.parse(await app.downloads.at(-1).text());
  assert.equal(saved.roomDistributionChoices[exam.id], "distribute");
  const reloaded = harness();
  tree = reloaded.render();
  await find(tree, (node) => node.type === "input" && node.props.accept === "application/json").props.onChange(jsonEvent(saved));
  tree = reloaded.render();
  button(tree, "Proceed To Resources").props.onClick();
  tree = reloaded.render();
  assert.equal(decisionButton("Distribute overflow evenly").props["aria-pressed"], true);
  assert.equal(panel().validation.complete, true);
  assert.equal(panel().sessions[0].rooms.find((room) => room.courseId === other.id).maxStudents, 25);
  button(tree, "Proceed To Export").props.onClick();
  tree = reloaded.render();
  await button(tree, "Export Timetable").props.onClick();
  const exported = XLSX.read(await reloaded.downloads.at(-1).arrayBuffer(), { type: "array" });
  const reportRows = XLSX.utils.sheet_to_json(exported.Sheets["Week 1 Invigilators"], { header: 1 });
  const examRows = reportRows.slice(1).filter((row) => row[1] === exam.code);
  assert.deepEqual(examRows.map((row) => row[3]), [26, 26]);
  assert.ok(examRows.every((row) => /Approved up to 27/.test(row[14])));
  button(tree, "Back To Resources").props.onClick();
  tree = reloaded.render();
  decisionButton("Keep the extra room (normal limit)").props.onClick();
  tree = reloaded.render();
  assert.deepEqual(examRooms().map((room) => room.students.length), [18, 17, 17]);
  assert.equal(button(tree, "Proceed To Export").props.disabled, true, "The restored room must be assigned before export");
});

test("export workspace previews audiences, combines or splits weeks, exports CSV and prints only selected report data", async () => {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
    ["Student ID", "Student Name", "Course Code", "Course Title", "CRN"],
    ...Array.from({ length: 16 }, (_, i) => ["A" + i, "Alice " + i, "MAIN-1000", "First Exam", "101"]),
    ...Array.from({ length: 10 }, (_, i) => ["B" + i, "Bob " + i, "MAIN-2000", "Second Exam", "201"]),
    ["ASD1", "ASD Student", "ASD-1000", "Excluded ASD Exam", "301"],
  ]), "Enrolment");
  const enrolment = imports.parseEnrolmentWorkbook(book);
  const main = enrolment.courses.filter((course) => course.code.startsWith("MAIN"));
  const asd = enrolment.courses.find((course) => course.code.startsWith("ASD"));
  const labRoomId = department.roomIdentity("PAD", "P-B-4F", "13");
  const labMeeting = { crn: "101", days: ["Monday"], isLab: true, startMinutes: 480, endMinutes: 590,
    room: "13", building: "P-B-4F", campus: "PAD", instructor: "Staff 0",
    labInstructorId: "I0", instructors: [{ id: "I0", name: "Staff 0" }], teachingInstructorIds: ["I0"] };
  const labBusy = { code: main[0].code, crn: "101", isLab: true, days: ["Monday"], start: 480, end: 590 };
  const snapshot = { type: "main", version: 5, startDate: "2026-10-20", weeks: [1, 2, 3], selectedWeek: 1,
    settings: { slotIntervalMinutes: 30, startHour: 8, endHour: 18, studentsPerRoom: 25, examDurationMinutes: 60 },
    assignments: { 1: { Monday: { "09:00": [main[0].id] } }, 2: { Tuesday: { "12:00": [main[1].id] } }, 3: {} },
    asdAssignments: { 1: { Monday: { "09:00": [asd.id] } } }, asdExamDurations: { [asd.id]: 60 }, hasAsdStep: true,
    courses: enrolment.courses, studentDirectory: enrolment.studentDirectory,
    departmentSelection: { sheetName: "First", courses: main.map((course, index) => ({
      code: course.code, title: course.title, crns: course.crns, meetings: index === 0 ? [labMeeting] : [],
    })) },
    examChoices: Object.fromEntries(main.map((course) => [course.id, true])), roomDistributionChoices: {}, resourcePlan: null,
    resourceCatalog: {
      rooms: Array.from({ length: 3 }, (_, i) => ({ id: i === 0 ? labRoomId : "R" + i,
        name: "PAD / P-B-4F / " + (13 + i), enabled: true, busy: i === 0 ? [labBusy] : [] })),
      invigilators: Array.from({ length: 6 }, (_, i) => ({ id: "I" + i, name: "Staff " + i, enabled: true, busy: i === 0 ? [labBusy] : [] })),
    },
  };
  const app = harness();
  let tree = app.render();
  await find(tree, (node) => node.type === "input" && node.props.accept === "application/json").props.onChange(jsonEvent(snapshot));
  tree = app.render();
  button(tree, "Proceed To Resources").props.onClick();
  tree = app.render();
  button(tree, "Auto-Assign Resources").props.onClick();
  tree = app.render();
  button(tree, "Proceed To Export").props.onClick();
  tree = app.render();
  const panel = () => find(tree, (node) => node.props.className === "export-studio");
  const preview = () => find(panel(), (node) => node.props.className === "export-preview");
  const reportButton = (label) => find(panel(), (node) => node.type === "button" && node.props["aria-label"] === label);
  const weekInput = (week) => find(panel(), (node) => node.type === "input" && node.props["aria-label"] === "Include Week " + week);
  const radio = (label) => find(find(panel(), (node) => node.type === "label" && text(node).includes(label)), (node) => node.type === "input" && node.props.type === "radio");
  const printView = () => find(panel(), (node) => node.props.className === "export-print");
  assert.equal(weekInput(1).props.checked, true);
  assert.equal(weekInput(2).props.checked, true);
  assert.ok(!nodes(panel()).some((node) => node.props["aria-label"] === "Include Week 3"), "Empty weeks are omitted");
  assert.ok(!text(panel()).includes("ASD-1000"));
  assert.match(text(preview()), /Week 1 Invigilators/);
  assert.match(text(preview()), /19 Oct 2026/, "Preview and downloads use the same Monday-aligned date");
  find(preview(), (node) => node.type === "input" && node.props.type === "search").props.onChange({ target: { value: "Staff 0" } });
  tree = app.render();
  assert.ok(!text(preview()).includes("No matching records"), "The room ledger can be searched by an assigned invigilator");
  await button(tree, "Export Timetable").props.onClick();
  const combined = XLSX.read(await app.downloads.at(-1).arrayBuffer(), { type: "array" });
  assert.ok(combined.Sheets["Schedule Index"]);
  assert.ok(combined.Sheets["Week 1 Invigilators"]);
  assert.ok(combined.Sheets["Week 2 Invigilators"]);
  assert.match(combined.Sheets["Week 1 Invigilators"].E2.v, /19 Oct 2026/);
  assert.equal(combined.Sheets["Week 2 Tuesday"].F2.f, "'Week 2 Invigilators'!H2");

  radio("Separate weekly files").props.onChange();
  tree = app.render();
  await button(tree, "Export Timetable").props.onClick();
  const zip = await JSZip.loadAsync(await app.downloads.at(-1).arrayBuffer());
  assert.deepEqual(Object.keys(zip.files).sort(), ["Week_1_Exam_Schedule.xlsx", "Week_2_Exam_Schedule.xlsx"]);
  assert.equal(app.downloadNames.at(-1), "Exam_Schedule_Weekly_Files.zip");
  weekInput(1).props.onChange();
  tree = app.render();
  assert.match(text(preview()), /Live previewWeek 2/);
  assert.ok(!text(preview()).includes("MAIN-1000"));
  await button(tree, "Export Timetable").props.onClick();
  assert.equal(app.downloadNames.at(-1), "Week_2_Exam_Schedule.xlsx");
  const single = XLSX.read(await app.downloads.at(-1).arrayBuffer(), { type: "array" });
  assert.ok(single.SheetNames.every((name) => name.startsWith("Week 2 ")));

  reportButton("Exam overview").props.onClick();
  tree = app.render();
  assert.ok(nodes(preview()).some((node) => node.props.className === "export-board"));
  assert.ok(!text(preview()).includes("Bob"));
  assert.ok(!text(printView()).includes("Bob"));
  assert.match(text(find(preview(), (node) => node.props.className === "export-board__exam")), /1 primary invigilator needed/);
  assert.ok(!nodes(preview()).some((node) => node.props.className === "export-lab"), "Non-lab exams have no lab badge");
  assert.ok(!text(preview()).includes("PAD"));
  weekInput(1).props.onChange();
  tree = app.render();
  button(preview(), "Week 1").props.onClick();
  tree = app.render();
  const labCard = find(preview(), (node) => node.props.className === "export-board__exam");
  assert.match(text(labCard), /2 primary invigilators needed/);
  assert.match(text(labCard), /P-B-4F\/13/);
  assert.equal(text(find(labCard, (node) => node.props.className === "export-lab")), "During lab time");
  const printedOverviews = nodes(printView()).filter((node) => node.type?.name === "PreviewTable" && node.props.label === "Exam overview");
  assert.deepEqual(JSON.parse(JSON.stringify(printedOverviews.map((table) => table.props.rows[0].slice(-2)))), [[2, "Yes"], [1, "No"]]);
  assert.ok(printedOverviews.every((table) => table.props.columns.includes("Primary invigilators needed") && table.props.columns.includes("During lab time")));
  assert.ok(!text(printView()).includes("PAD"));
  button(tree, "Chronological list").props.onClick();
  tree = app.render();
  const labList = find(preview(), (node) => node.type?.name === "PreviewTable" && node.props.label === "Exam overview preview");
  assert.deepEqual(JSON.parse(JSON.stringify(labList.props.rows[0].slice(-3))), ["P-B-4F/13", 2, "Yes"]);
  weekInput(1).props.onChange();
  tree = app.render();
  assert.match(text(preview()), /MAIN-2000/);
  radio("CSV (.csv)").props.onChange();
  tree = app.render();
  await button(tree, "Download CSV").props.onClick();
  assert.equal(app.downloadNames.at(-1), "Week_2_Exam_Overview.csv");
  const overviewCsv = await app.downloads.at(-1).text();
  assert.ok(!overviewCsv.includes("Bob"));
  assert.ok(!overviewCsv.includes("PAD"));
  assert.match(overviewCsv, /"Primary invigilators needed","During lab time"/);
  assert.match(overviewCsv, /"P-B-4F\/\d+","1","No"/);

  reportButton("Staff duties").props.onClick();
  tree = app.render();
  assert.match(text(preview()), /Workload balance/);
  assert.match(text(preview()), /Backup/);
  assert.ok(!text(preview()).includes("Bob"));
  radio("Excel (.xlsx)").props.onChange();
  tree = app.render();
  await button(tree, "Download Excel").props.onClick();
  const staff = XLSX.read(await app.downloads.at(-1).arrayBuffer(), { type: "array" });
  assert.ok(staff.Sheets["Workload Summary"]);

  weekInput(1).props.onChange();
  tree = app.render();
  button(preview(), "Overall").props.onClick();
  tree = app.render();
  assert.equal(button(preview(), "Overall").props["aria-pressed"], true);
  assert.equal(button(preview(), "Week 1").props["aria-pressed"], false);
  assert.equal(button(preview(), "Week 2").props["aria-pressed"], false);
  assert.match(text(preview()), /Workload balance \/ Overall/);
  assert.match(text(preview()), /Totals across included weeks: Week 1, Week 2/);
  assert.match(text(preview()), /Week \/ date \/ time/);
  assert.match(text(preview()), /MAIN-1000/);
  assert.match(text(preview()), /MAIN-2000/);
  assert.ok(!text(preview()).includes("ASD-1000"));
  const exportProps = find(tree, (node) => node.type?.name === "ExportStudio").props;
  const plainRows = (value) => JSON.parse(JSON.stringify(value));
  const expectedOverall = plainRows(resources.invigilatorWorkloads(exportProps.catalog, exportProps.sessions, exportProps.plan)
    .map((person) => [person.name, person.exam, person.backup, person.teaching, person.exam + person.backup]));
  const workloadRows = () => plainRows(find(preview(), (node) => node.type?.name === "PreviewTable" && node.props.label === "Workload balance").props.rows);
  assert.deepEqual(workloadRows(), expectedOverall, "Overall matches the full Resource Review Pool when all weeks are included");
  assert.ok(workloadRows().some((row) => row.slice(1).every((load) => load === 0)), "Overall includes idle invigilators for a complete pool comparison");
  const staffProps = () => find(tree, (node) => node.type?.name === "ExportStudio").props;
  const expectedWeek2 = plainRows(exportReports.buildExportModel({ ...exportProps, weeks: [2] }).workloads
    .map((person) => [person.name, person.exam, person.backup, person.teaching, person.exam + person.backup]));
  weekInput(1).props.onChange();
  tree = app.render();
  assert.equal(button(preview(), "Overall").props["aria-pressed"], true);
  assert.deepEqual(workloadRows(), expectedWeek2, "Overall follows the included-week selection");
  assert.ok(!text(preview()).includes("MAIN-1000"));
  await button(tree, "Download Excel").props.onClick();
  const filteredStaff = XLSX.read(await app.downloads.at(-1).arrayBuffer(), { type: "array" });
  assert.deepEqual(XLSX.utils.sheet_to_json(filteredStaff.Sheets["Workload Summary"], { header: 1 }).slice(1).map((row) => row.slice(0, 5)),
    expectedWeek2, "The Overall preview matches the exported workload summary");
  weekInput(2).props.onChange();
  tree = app.render();
  assert.equal(button(preview(), "Overall").props.disabled, true);
  assert.match(text(preview()), /Choose a week to preview/);
  weekInput(2).props.onChange();
  tree = app.render();
  button(preview(), "Week 2").props.onClick();
  tree = app.render();
  assert.equal(button(preview(), "Overall").props["aria-pressed"], false);
  assert.match(text(preview()), /Workload balance \/ Week 2/);
  const weeklyExpected = plainRows(exportReports.buildExportModel({ ...staffProps(), weeks: [2] }).workloads
    .filter((person) => person.exam + person.backup + person.teaching > 0)
    .map((person) => [person.name, person.exam, person.backup, person.teaching, person.exam + person.backup]));
  assert.deepEqual(workloadRows(), weeklyExpected, "Weekly tabs retain weekly-only totals");

  reportButton("Student room lists").props.onClick();
  tree = app.render();
  assert.match(text(preview()), /Bob 0/);
  const search = find(preview(), (node) => node.type === "input" && node.props.type === "search");
  search.props.onChange({ target: { value: "NO-MATCH" } });
  tree = app.render();
  assert.match(text(preview()), /No matching records/);
  assert.match(text(printView()), /Bob 0/, "Preview search must not filter print output");
  button(tree, "Print / Save PDF").props.onClick();
  assert.equal(app.prints(), 1);
  await button(tree, "Download Excel").props.onClick();
  const students = XLSX.read(await app.downloads.at(-1).arrayBuffer(), { type: "array" });
  assert.equal(XLSX.utils.sheet_to_json(students.Sheets["Student room lists"]).length, 10, "Preview search must not filter exports");
  assert.ok(!text(printView()).includes("Alice"));

  weekInput(2).props.onChange();
  tree = app.render();
  assert.equal(button(tree, "Download Excel").props.disabled, true);
  assert.equal(button(tree, "Print / Save PDF").props.disabled, true);
  assert.match(text(preview()), /Choose a week to preview/);
  weekInput(1).props.onChange();
  tree = app.render();
  assert.match(text(preview()), /Alice/);
  radio("CSV (.csv)").props.onChange();
  tree = app.render();
  reportButton("Complete report").props.onClick();
  tree = app.render();
  assert.equal(radio("Excel (.xlsx)").props.checked, true);
  assert.equal(radio("CSV (.csv)").props.disabled, true);
});
