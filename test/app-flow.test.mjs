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
import * as examWindows from "../src/examWindows.js";
import { buildResourceWorkbookForWeek } from "../src/reports.js";
import * as exportReports from "../src/exportReports.js";
import * as examSetup from "../src/examSetup.js";

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
const poolSource = readFileSync(new URL("../src/ResourcePool.jsx", import.meta.url), "utf8")
  .replace(/^import .*;\r?$/gm, "")
  .replace("export default function ResourcePool", "function ResourcePool") + "\nglobalThis.ResourcePool = ResourcePool;";
const { code: poolComponentCode } = await transformWithEsbuild(poolSource, "ResourcePool.jsx", { loader: "jsx", jsxFactory: "h", jsxFragment: "Fragment" });
const exportSource = readFileSync(new URL("../src/ExportStudio.jsx", import.meta.url), "utf8")
  .replace(/^import .*;\r?$/gm, "")
  .replace("export default function ExportStudio", "function ExportStudio") + "\nglobalThis.ExportStudio = ExportStudio;";
const { code: exportComponentCode } = await transformWithEsbuild(exportSource, "ExportStudio.jsx", { loader: "jsx", jsxFactory: "h", jsxFragment: "Fragment" });

function harness(component = "App", props) {
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
    ...imports, ...department, ...resources, ...examRooms, ...examWindows, ...exportReports, ...examSetup, buildResourceWorkbookForWeek, autoSchedule, XLSX, JSZip, Blob, URL: TestURL, console,
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
  runInContext(poolComponentCode, context);
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
        tree = context[component](props);
        effects.forEach((effect) => effect());
        assert.ok(++attempts < 10, "Render must settle");
      } while (dirty);
      return tree;
    },
  };
}

test("invigilator and backup selectors show available people first and update after assignments", () => {
  const makeCourse = (id, count) => ({ id, code: id, title: id, crns: ["101"], labSessions: [],
    students: Array.from({ length: count }, (_, i) => ({ id: id + i, name: "Student " + i, crn: "101" })) });
  const sessions = resources.buildExamSessions({ 1: { Monday: { "09:00": ["A"], "09:30": ["B"] } } },
    { A: makeCourse("A", 16), B: makeCourse("B", 10) }, 60);
  const [current, other] = sessions;
  const meeting = (code, start, end, extra = {}) => ({ code, crn: "101", days: ["Monday"], isLab: false, start, end, ...extra });
  const person = (id, name, busy = [], enabled = true) => ({ id, name, busy, enabled });
  const catalog = { rooms: [], invigilators: [
    person("CLASS", "Ada Class", [meeting("CLASS-1000", 480, 590)]),
    person("EXCLUDED", "Ben Excluded", [], false),
    person("EXAM", "Cal Exam"),
    person("BACKUP", "Dee Backup"),
    person("UNKNOWN", "Eli Unknown", [meeting("UNKNOWN-1000", 0, 1440, { unknownTime: true })]),
    person("OTHERBACKUP", "Finn Backup"),
    person("FREEZ", "Zoe Available", [meeting("TUESDAY-1000", 540, 600, { days: ["Tuesday"] })]),
    person("FREEA", "Alex Available", [meeting("EARLY-1000", 480, 540)]),
    person("CURRENT", "Sam Selected"),
    person("LATE", "Late Class", [meeting("LATE-1000", 585, 660)]),
  ] };
  const plan = resources.emptyResourcePlan(sessions, catalog);
  plan.allocations[current.rooms[0].id] = { roomId: "", invigilatorIds: ["CURRENT", ""] };
  plan.allocations[other.rooms[0].id] = { roomId: "", invigilatorIds: ["EXAM"] };
  plan.backups[current.id] = ["BACKUP"];
  plan.backups[other.id] = ["OTHERBACKUP"];
  const props = { sessions, catalog, plan, validation: { complete: false, issues: [] },
    onAssign() {}, onPoolChange() {}, onDistributionChange() {},
    onAllocationChange(id, allocation) { props.plan = { ...props.plan, allocations: { ...props.plan.allocations, [id]: allocation } }; },
    onBackupChange(id, ids) { props.plan = { ...props.plan, backups: { ...props.plan.backups, [id]: ids } }; },
  };
  const app = harness("ResourceAssignment", props);
  let tree = app.render();
  const select = (label) => find(tree, (node) => node.type === "select" && node.props["aria-label"] === label);
  const groups = (control) => nodes(control).filter((node) => node.type === "optgroup");
  const choices = (control) => nodes(control).filter((node) => node.type === "option" && node.props.value);
  const availableIds = (control) => choices(control).filter((node) => !node.props.disabled).map((node) => node.props.value);
  const option = (control, id) => find(control, (node) => node.type === "option" && node.props.value === id);
  const mainLabel = "Invigilator 1 for " + current.rooms[0].id;
  const extraLabel = "Invigilator 2 for " + current.rooms[0].id;
  const backupLabel = "Backup 1 for " + current.id;
  const main = select(mainLabel);
  assert.deepEqual(availableIds(main), ["FREEA", "CURRENT", "FREEZ"], "Available choices are alphabetic, not interleaved with disabled staff");
  assert.deepEqual(groups(main).map((group) => group.props.label), ["Available (3)", "Unavailable (7)"]);
  assert.equal(groups(main)[1].props.disabled, true);
  assert.ok(choices(groups(main)[1]).every((node) => node.props.disabled));
  assert.equal(text(option(main, "CLASS")), "Ada Class - Class CLASS-1000 08:00-09:50");
  assert.equal(text(option(main, "LATE")), "Late Class - Class LATE-1000 09:45-11:00", "Availability covers the whole exam, not just its start");
  assert.equal(text(option(main, "EXCLUDED")), "Ben Excluded - Excluded from pool");
  assert.equal(text(option(main, "EXAM")), "Cal Exam - Exam B 09:30-10:30");
  assert.equal(text(option(main, "BACKUP")), "Dee Backup - Backup duty 09:00-10:00");
  assert.equal(text(option(main, "OTHERBACKUP")), "Finn Backup - Backup duty 09:30-10:30");
  assert.equal(text(option(main, "UNKNOWN")), "Eli Unknown - Class UNKNOWN-1000 (time unknown; day blocked)");
  assert.equal(option(main, "CURRENT").props.disabled, undefined);
  assert.equal(option(select(extraLabel), "CURRENT").props.disabled, true, "A person cannot cover two places in the same slot");
  assert.deepEqual(availableIds(select(backupLabel)), ["FREEA", "BACKUP", "FREEZ"]);
  assert.equal(option(select(backupLabel), "CURRENT").props.disabled, true);

  select(extraLabel).props.onChange({ target: { value: "FREEA" } });
  tree = app.render();
  assert.equal(select(extraLabel).props.value, "FREEA");
  assert.ok(availableIds(select(extraLabel)).includes("FREEA"));
  assert.equal(option(select(mainLabel), "FREEA").props.disabled, true);
  assert.equal(option(select(backupLabel), "FREEA").props.disabled, true);
  assert.equal(select(mainLabel).props.value, "CURRENT", "Regrouping must not change current assignments");
  select(backupLabel).props.onChange({ target: { value: "FREEZ" } });
  tree = app.render();
  assert.equal(select(backupLabel).props.value, "FREEZ");
  assert.ok(availableIds(select(backupLabel)).includes("FREEZ"));
  assert.equal(option(select(mainLabel), "FREEZ").props.disabled, true);
  assert.ok(availableIds(select(mainLabel)).includes("BACKUP"), "The former backup becomes available immediately");

  catalog.invigilators.forEach((staff) => { staff.enabled = false; });
  tree = app.render();
  assert.deepEqual(availableIds(select(mainLabel)), []);
  assert.equal(groups(select(mainLabel)).length, 1);
  assert.match(text(find(select(mainLabel), (node) => node.type === "option" && node.props.value === "")), /none available/);
  assert.equal(select(mainLabel).props.value, "CURRENT", "An invalid selection stays visible until explicitly corrected");
  assert.equal(option(select(mainLabel), "CURRENT").props.disabled, true);
});

test("grouped lab selectors retain the required instructor and teaching-hour labels without restricting extra staff", () => {
  const labRoom = department.roomIdentity("PAD", "P-B-4F", "13");
  const lab = { days: ["Monday"], crn: "101", startMinutes: 480, endMinutes: 590,
    room: "13", building: "P-B-4F", campus: "PAD", labInstructorId: "I0" };
  const exam = { id: "LAB", code: "LAB", title: "Lab Exam", crns: ["101"], labSessions: [lab],
    students: Array.from({ length: 31 }, (_, i) => ({ id: "S" + i, name: "Student " + i, crn: "101" })) };
  const sessions = resources.buildExamSessions({ 1: { Monday: { "09:00": ["LAB"] } } }, { LAB: exam }, 60);
  const catalog = {
    rooms: [{ id: labRoom, name: "PAD / P-B-4F / 13", enabled: true, busy: [] }, { id: "R1", name: "Overflow", enabled: true, busy: [] }],
    invigilators: Array.from({ length: 6 }, (_, i) => ({ id: "I" + i, name: "Staff " + i, enabled: i !== 3, busy: [] })),
  };
  catalog.invigilators[0].busy = [{ code: "LAB", crn: "101", isLab: true, days: ["Monday"], start: 480, end: 590 }];
  catalog.invigilators[2].busy = [{ code: "OTHER", crn: "201", isLab: false, days: ["Monday"], start: 570, end: 630 }];
  const plan = resources.emptyResourcePlan(sessions, catalog);
  plan.allocations[sessions[0].rooms[0].id] = { roomId: labRoom, invigilatorIds: ["I0", "I1"] };
  plan.allocations[sessions[0].rooms[1].id] = { roomId: "R1", invigilatorIds: ["I5"] };
  plan.backups[sessions[0].id] = ["I4"];
  const app = harness("ResourceAssignment", { sessions, catalog, plan, validation: { complete: true, issues: [] },
    onAssign() {}, onPoolChange() {}, onAllocationChange() {}, onBackupChange() {}, onDistributionChange() {} });
  let tree = app.render();
  const control = (index) => find(tree, (node) => node.type === "select" && node.props["aria-label"] === "Invigilator " + index + " for " + sessions[0].rooms[0].id);
  const choices = (select) => nodes(select).filter((node) => node.type === "option" && node.props.value);
  assert.deepEqual(choices(control(1)).filter((node) => !node.props.disabled).map((node) => node.props.value), ["I0"]);
  assert.match(text(choices(control(1)).find((node) => node.props.value === "I0")), /teaching hours, no extra load/);
  assert.ok(choices(control(1)).filter((node) => node.props.value !== "I0").every((node) => node.props.disabled && /Lab instructor required/.test(text(node))));
  assert.ok(!text(control(2)).includes("Lab instructor required"), "Only the lab lead selector is pinned");
  assert.deepEqual(choices(control(2)).filter((node) => !node.props.disabled).map((node) => node.props.value), ["I1"]);
  catalog.invigilators[0].busy.push({ code: "CLASH", crn: "301", isLab: false, days: ["Monday"], start: 570, end: 630 });
  tree = app.render();
  assert.equal(choices(control(1)).filter((node) => !node.props.disabled).length, 0);
  assert.match(text(choices(control(1)).find((node) => node.props.value === "I0")), /Class CLASH 09:30-10:30/);
  assert.equal(control(1).props.value, "I0", "A lab instructor conflict cannot silently replace the required instructor");
});

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
const button = (tree, label) => find(tree, (node) => node.type === "button" && (text(node) === label || node.props["aria-label"] === label));
const step = (tree, id) => find(tree, (node) => node.type === "button" && node.props["data-step"] === id);
const exportReady = (tree) => find(tree, (node) => node.type?.name === "ResourceAssignment").props.validation.complete;
const review = (tree) => find(tree, (node) => node.type?.name === "CourseSelection");
const root = "data/input/26-27 S1/";
const enrolmentPath = root + "Students Registration 26-27 S1.xlsx";
const crnPath = root + "CRN List 26-27 S1.xlsx";
const asdPath = root + "Other Departments Schedule/Midterm Schedule 26-27 S1.xlsx";

test("the start screen shows only create and load choices before opening the workspace", () => {
  const app = harness();
  let tree = app.render();
  assert.equal(text(find(tree, (node) => node.type === "h1")), "Midterm Exam Scheduling Helper");
  assert.deepEqual(nodes(tree).filter((node) => node.type === "button").map((node) => node.props["aria-label"]),
    ["Create Schedule", "Load Schedule"]);
  assert.ok(!nodes(tree).some((node) => node.type === "nav" || node.props.role === "dialog" || node.props.className === "step-actions"));
  const inputs = nodes(tree).filter((node) => node.type === "input");
  assert.equal(inputs.length, 1);
  assert.equal(inputs[0].props.accept, "application/json");
  assert.equal(inputs[0].props.hidden, true);
  assert.ok(!nodes(tree).some((node) => node.type?.name === "CourseSelection" || node.type?.name === "ResourceAssignment" || node.props.className === "timetable"));
  button(tree, "Create Schedule").props.onClick();
  tree = app.render();
  assert.equal(step(tree, "setup").props["aria-current"], "step");
  assert.equal(nodes(tree).filter((node) => node.props["data-step"]).length, 8);
  assert.equal(text(find(tree, (node) => node.props.className === "app__position")), "Step 0 / 7");
  assert.ok(!nodes(tree).some((node) => node.props["aria-label"] === "Step prerequisites"));
  assert.equal(nodes(tree).filter((node) => node.type === "input" && node.props.type === "date").length, 1);
  assert.ok(!nodes(tree).some((node) => node.type === "label" && text(node) === "Upload Enrollment File"));
  button(tree, "Continue to Student Enrollment").props.onClick();
  tree = app.render();
  assert.equal(step(tree, "load").props["aria-current"], "step");
  find(tree, (node) => node.type === "label" && text(node) === "Upload Enrollment File");
  assert.ok(!nodes(tree).some((node) => node.type === "input" && node.props.type === "date"));
  assert.ok(!nodes(tree).some((node) => node.props.className === "start-screen"));
});

test("Exam Setup derives metadata from the entered date and restores it after saving and loading", async () => {
  const workbookEvent = (rows) => {
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "First");
    return { target: { files: [{ name: "input.xlsx", arrayBuffer: async () => XLSX.write(book, { type: "array", bookType: "xlsx" }) }], value: "input.xlsx" } };
  };
  const app = harness();
  let tree = app.render();
  button(tree, "Create Schedule").props.onClick();
  tree = app.render();
  const dateInput = () => find(tree, (node) => node.props.id === "start-date-input");
  const summary = () => text(find(tree, (node) => node.props.className === "exam-setup__summary"));
  for (const [date, semester, academicYear] of [
    ["2026-08-01", "Fall", "2026-2027"], ["2026-06-01", "Summer", "2025-2026"],
    ["2027-01-12", "Spring", "2026-2027"],
  ]) {
    dateInput().props.onChange({ target: { value: date } });
    tree = app.render();
    assert.equal(dateInput().props.value, date, "The date is not silently shifted into the preceding month or year");
    assert.equal(summary(), "Semester" + semester + "Academic year" + academicYear);
    assert.equal(nodes(tree).filter((node) => node.type === "input" && !node.props.hidden).length, 1, "Inferred metadata is read-only");
  }
  assert.match(text(find(tree, (node) => node.props.id === "exam-week-start")), /Monday, 2027-01-11/);
  button(tree, "Continue to Student Enrollment").props.onClick();
  tree = app.render();
  await find(find(tree, (node) => node.type === "label" && text(node) === "Upload Enrollment File"), (node) => node.type === "input").props.onChange(workbookEvent([
    ["Student ID", "Student Name", "Course Code", "Course Title", "CRN"], ["S1", "Student One", "EXAM-1000", "Exam Course", "101"],
  ]));
  tree = app.render();
  assert.equal(step(tree, "courses").props["aria-current"], "step");
  assert.ok(!nodes(tree).some((node) => node.props.type === "date"));
  await review(tree).props.onUpload(workbookEvent([
    ["Campus", "Crn No", "Course Code", "Title", "Cr", "Maximum Load", "No Of Enrolled", "Session Id", "DAYS", "Time", "Primary Instructor", "Second Instructor", "Type", "Building", "Room"],
    ["PAD", "101", "EXAM-1000", "Exam Course", 3, 25, 1, "01", "M", "0800 - 0950", "10: Alice", "20: Bob", "T", "Building", "A"],
  ]));
  tree = app.render();
  step(tree, "setup").props.onClick();
  tree = app.render();
  assert.equal(dateInput().props.value, "2027-01-12");
  assert.ok(!nodes(tree).some((node) => node.props.className === "overview" || node.type?.name === "CourseSelection"));
  step(tree, "main").props.onClick();
  tree = app.render();
  button(tree, "Save Timetable").props.onClick();
  const saved = JSON.parse(await app.downloads.at(-1).text());
  assert.equal(saved.startDate, "2027-01-12");
  assert.equal("semester" in saved, false, "Derived values must not duplicate the authoritative date");
  assert.equal("academicYear" in saved, false);
  const loaded = harness();
  tree = loaded.render();
  await find(tree, (node) => node.type === "input" && node.props.accept === "application/json").props.onChange(jsonEvent(saved));
  tree = loaded.render();
  assert.equal(step(tree, "main").props["aria-current"], "step");
  step(tree, "setup").props.onClick();
  tree = loaded.render();
  assert.equal(dateInput().props.value, "2027-01-12");
  assert.equal(summary(), "SemesterSpringAcademic year2026-2027");
});

test("canceled and invalid loads stay on the start screen and loading disables both choices", async () => {
  const app = harness();
  let tree = app.render();
  const input = () => find(tree, (node) => node.type === "input" && node.props.accept === "application/json");
  let pickerRequests = 0;
  input().props.ref.current = { click() { pickerRequests += 1; } };
  button(tree, "Load Schedule").props.onClick();
  tree = app.render();
  assert.equal(pickerRequests, 1);
  await input().props.onChange({ target: { files: [], value: "" } });
  tree = app.render();
  assert.ok(!nodes(tree).some((node) => node.type === "nav" || node.props.role === "alert"));
  assert.equal(button(tree, "Create Schedule").props.disabled, false);

  let resolveFile;
  const event = { target: { files: [{ name: "invalid.json", text: () => new Promise((resolve) => { resolveFile = resolve; }) }], value: "invalid.json" } };
  const pending = input().props.onChange(event);
  tree = app.render();
  assert.equal(text(find(tree, (node) => node.props.role === "status")), "Loading schedule...");
  assert.equal(button(tree, "Create Schedule").props.disabled, true);
  assert.equal(button(tree, "Load Schedule").props.disabled, true);
  assert.equal(input().props.disabled, true);
  button(tree, "Create Schedule").props.onClick();
  button(tree, "Load Schedule").props.onClick();
  tree = app.render();
  assert.ok(!nodes(tree).some((node) => node.type === "nav"));
  assert.equal(pickerRequests, 1, "Ignore new actions while a schedule is loading");
  resolveFile(JSON.stringify({ type: "main", version: -1 }));
  await pending;
  tree = app.render();
  assert.match(text(find(tree, (node) => node.props.role === "alert")), /version is not supported/);
  assert.ok(!nodes(tree).some((node) => node.type === "nav" || node.props.role === "status"));
  assert.equal(button(tree, "Create Schedule").props.disabled, false);
  assert.equal(button(tree, "Load Schedule").props.disabled, false);
  assert.equal(event.target.value, "", "An invalid file can be selected again after correction");
  button(tree, "Create Schedule").props.onClick();
  tree = app.render();
  assert.equal(step(tree, "setup").props["aria-current"], "step");
  assert.ok(!nodes(tree).some((node) => node.props.role === "alert"));
});

test("every step is directly accessible, with actions separate and prerequisites enforced", () => {
  const app = harness();
  let tree = app.render();
  button(tree, "Create Schedule").props.onClick();
  tree = app.render();
  const header = () => find(tree, (node) => node.props.className === "app__header");
  const navigation = () => find(header(), (node) => node.type === "nav" && node.props["aria-label"] === "Schedule steps");
  const actions = () => find(tree, (node) => node.props.className === "step-actions");
  assert.equal(nodes(navigation()).filter((node) => node.type === "button").length, 8);
  assert.equal(step(tree, "setup").props["aria-label"], "Step 0: Exam Setup");
  assert.equal(step(tree, "load").props["aria-label"], "Step 1: Student Enrollment");
  assert.equal(step(tree, "courses").props["aria-label"], "Step 2: CRN Info");
  assert.equal(step(tree, "asd").props["aria-label"], "Step 3: ASD Schedule (optional)");
  assert.ok(!nodes(header()).some((node) => node.type === "input"));
  assert.ok(!text(header()).includes("Upload student"));
  const scrollRequests = [];
  navigation().props.ref.current = {
    querySelector(selector) {
      assert.equal(selector, '[aria-current="step"]');
      return { scrollIntoView(options) { scrollRequests.push(options); } };
    },
  };
  for (const id of ["export", "asd", "resources", "main", "courses", "pool", "load", "setup"]) {
    step(tree, id).props.onClick();
    tree = app.render();
    assert.equal(step(tree, id).props["aria-current"], "step");
    assert.equal(nodes(navigation()).filter((node) => node.props["aria-current"] === "step").length, 1);
    assert.ok(nodes(navigation()).filter((node) => node.type === "button").every((node) => node.props.type === "button" && !node.props.disabled));
    assert.ok(!nodes(actions()).some((node) => node.props["data-step"]));
    assert.equal(nodes(actions()).filter((node) => node.type === "button" && text(node) === "Load Timetable").length, id === "load" ? 1 : 0,
      "Load Timetable is available only in Student Enrollment");
    if (!["setup", "load"].includes(id)) assert.match(text(find(tree, (node) => node.props["aria-label"] === "Step prerequisites")), /Load enrolment in Student Enrollment/);
    else assert.ok(!nodes(tree).some((node) => node.props["aria-label"] === "Step prerequisites"));
    assert.equal(nodes(tree).filter((node) => node.type === "input" && node.props.type === "date").length, id === "setup" ? 1 : 0);
    if (id === "asd") assert.equal(button(tree, "Create ASD Timetable").props.disabled, true);
    if (id === "main") assert.equal(button(tree, "Auto-Schedule Remaining Exams").props.disabled, true);
  }
  assert.equal(scrollRequests.length, 8, "Programmatic navigation keeps the active step visible on narrow screens");
  assert.ok(scrollRequests.every((options) => options.block === "nearest" && options.inline === "nearest"));
  button(tree, "Settings").props.onClick();
  tree = app.render();
  assert.ok(nodes(tree).some((node) => node.props.role === "dialog"));
  assert.ok(!nodes(header()).some((node) => node.props.role === "dialog"));
});

test("resource selection precedes generation and stays editable after the resource-aware draft", async () => {
  const workbook = (rows) => {
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "First");
    const data = XLSX.write(book, { type: "array", bookType: "xlsx" });
    return { target: { files: [{ name: "input.xlsx", arrayBuffer: async () => data }], value: "input.xlsx" } };
  };
  const app = harness();
  let tree = app.render();
  button(tree, "Create Schedule").props.onClick();
  tree = app.render();
  button(tree, "Continue to Student Enrollment").props.onClick();
  tree = app.render();
  const upload = find(find(tree, (node) => node.type === "label" && text(node) === "Upload Enrollment File"),
    (node) => node.type === "input");
  await upload.props.onChange(workbook([
    ["Student ID", "Student Name", "Course Code", "Course Title", "CRN"],
    ["S1", "Student One", "MAIN-1000", "Main Exam", "101"],
    ["S2", "Student Two", "MAIN-1000", "Main Exam", "102"],
  ]));
  tree = app.render();
  step(tree, "export").props.onClick();
  tree = app.render();
  assert.match(text(find(tree, (node) => node.props["aria-label"] === "Step prerequisites")), /Load your CRN list in CRN Info/);
  step(tree, "main").props.onClick();
  tree = app.render();
  assert.equal(button(tree, "Auto-Schedule Remaining Exams").props.disabled, true);
  assert.ok(!nodes(tree).some((node) => node.props.className === "timetable"));
  step(tree, "courses").props.onClick();
  tree = app.render();
  await review(tree).props.onUpload(workbook([
    ["Campus", "Crn No", "Course Code", "Title", "Cr", "Maximum Load", "No Of Enrolled", "Session Id", "DAYS", "Time", "Primary Instructor", "Second Instructor", "Type", "Building", "Room"],
    ["PAD", "101", "MAIN-1000", "Main Exam", 3, 25, 1, "01", "M", "0800 - 0950", "10: Alice", "20: Bob", "T", "P-B-4F", "13"],
    ["PAD", "102", "MAIN-1000", "Main Exam", 3, 25, 1, "02", "T", "0800 - 0950", "30: Carol", "", "T", "P-B-4F", "14"],
  ]));
  tree = app.render();
  step(tree, "asd").props.onClick();
  tree = app.render();
  button(tree, "Skip ASD Step").props.onClick();
  tree = app.render();
  const selection = () => find(tree, (node) => node.props["aria-label"] === "Resource selection");
  const pool = () => find(tree, (node) => node.type?.name === "ResourcePool");
  const include = (name) => find(selection(), (node) => node.type === "input" && node.props["aria-label"] === "Include " + name);
  assert.equal(pool().props.selectionOnly, true);
  assert.ok(!text(selection()).includes("Invigilation Load"), "Do not show assignment/load controls before generation");
  assert.ok(!nodes(tree).some((node) => node.type?.name === "ResourceAssignment"));
  assert.ok(!nodes(tree).some((node) => node.type === "button" && text(node) === "Auto-Schedule Remaining Exams"));
  assert.match(text(selection()), /3 invigilators selected/);
  assert.match(text(selection()), /Target: two available standby invigilators per slot/);
  assert.match(text(selection()), /08:00-09:50/);
  const schedulingDisabled = () => {
    step(tree, "main").props.onClick();
    tree = app.render();
    const disabled = button(tree, "Auto-Schedule Remaining Exams").props.disabled;
    step(tree, "pool").props.onClick();
    tree = app.render();
    return disabled;
  };
  const roomNames = pool().props.catalog.rooms.map((room) => room.name);
  include(roomNames[0]).props.onChange({ target: { checked: false } });
  tree = app.render();
  include(roomNames[1]).props.onChange({ target: { checked: false } });
  tree = app.render();
  assert.equal(schedulingDisabled(), true);
  include(roomNames[1]).props.onChange({ target: { checked: true } });
  include("Carol").props.onChange({ target: { checked: false } });
  include("Bob").props.onChange({ target: { checked: false } });
  tree = app.render();
  assert.equal(schedulingDisabled(), true, "One person cannot cover both an exam and backup");
  include("Carol").props.onChange({ target: { checked: true } });
  include("Bob").props.onChange({ target: { checked: true } });
  tree = app.render();
  assert.equal(schedulingDisabled(), false);
  step(tree, "main").props.onClick();
  tree = app.render();
  button(tree, "Auto-Schedule Remaining Exams").props.onClick();
  tree = app.render();
  assert.match(text(tree), /1 exams added. 0 remain unplaced/);
  for (const id of ["asd", "courses", "pool", "main"]) {
    step(tree, id).props.onClick();
    tree = app.render();
  }
  assert.ok(nodes(tree).some((node) => node.props.className === "auto-schedule-result"), "Jumping between steps keeps the draft result");
  step(tree, "resources").props.onClick();
  tree = app.render();
  const assignment = () => find(tree, (node) => node.type?.name === "ResourceAssignment").props;
  assert.equal(assignment().validation.complete, true, "The generated draft includes its checked allocation");
  assert.equal(!exportReady(tree), false);
  assert.match(text(tree), /Available standby capacity: 2 \(target 2\)/);
  assert.equal(assignment().sessions[0].day, "Monday");
  assert.equal(assignment().sessions[0].slotId, "12:00");
  assert.equal(assignment().plan.backups[assignment().sessions[0].id].length, 1);
  const excludedRoom = assignment().catalog.rooms.find((room) => room.name === roomNames[0]);
  assert.equal(excludedRoom.enabled, false);
  assert.ok(Object.values(assignment().plan.allocations).every((allocation) => allocation.roomId !== excludedRoom.id));
  const model = exportReports.buildExportModel({ ...assignment(), startDate: "2026-10-19" });
  assert.equal(model.duties.length, 2, "Unassigned standby people cannot become report duties");
  assert.equal(model.workloads.filter((person) => person.exam + person.backup + person.teaching === 0).length, 1);
  const primaryIds = Object.values(assignment().plan.allocations).flatMap((allocation) => allocation.invigilatorIds);
  const backupIds = Object.values(assignment().plan.backups).flat();
  const spare = assignment().catalog.invigilators.find((person) => !primaryIds.includes(person.id) && !backupIds.includes(person.id));
  assignment().onPoolChange("invigilators", spare.id, { enabled: false });
  tree = app.render();
  assert.equal(!exportReady(tree), true, "Pool edits invalidate the generated plan until review");
  assert.match(text(tree), /Confirm valid room and staff assignments to assess standby capacity/);
  const outdatedPlan = JSON.stringify(assignment().plan);
  step(tree, "export").props.onClick();
  tree = app.render();
  assert.equal(button(tree, "Export Timetable").props.disabled, true, "Jumping to export cannot bypass resource validation");
  assert.equal(button(tree, "Print / Save PDF").props.disabled, true);
  step(tree, "resources").props.onClick();
  tree = app.render();
  assert.equal(JSON.stringify(assignment().plan), outdatedPlan, "Navigation preserves allocations, including outdated ones that need review");
  button(tree, "Auto-Assign Resources").props.onClick();
  tree = app.render();
  assert.equal(assignment().validation.complete, true, "The existing one-backup report rule is still valid");
  assert.match(text(tree), /Available standby capacity: 1 \(target 2\)/);
  assert.ok(nodes(tree).some((node) => node.props.className === "resource-standby resource-standby--short"));
  assignment().onPoolChange("invigilators", spare.id, { enabled: true });
  tree = app.render();
  button(tree, "Auto-Assign Resources").props.onClick();
  tree = app.render();
  button(tree, "Save Timetable").props.onClick();
  const saved = JSON.parse(await app.downloads.at(-1).text());
  assert.equal(saved.resourceCatalog.rooms.find((room) => room.name === roomNames[0]).enabled, false);
  assert.ok(saved.resourcePlan.fingerprint);
  const mainId = assignment().sessions[0].rooms[0].courseId;
  for (const id of ["load", "asd", "export", "courses", "pool", "resources", "main"]) {
    step(tree, id).props.onClick();
    tree = app.render();
  }
  button(tree, "Save Timetable").props.onClick();
  const afterNavigation = JSON.parse(await app.downloads.at(-1).text());
  const withoutTimestamp = ({ savedAt: _savedAt, ...snapshot }) => snapshot;
  assert.deepEqual(withoutTimestamp(afterNavigation), withoutTimestamp(saved), "Step navigation must not change schedule data or resource selections");
  step(tree, "main").props.onClick();
  tree = app.render();
  const fridayCell = (time) => {
    const [hour, minute] = time.split(":").map(Number);
    const index = (hour * 60 + minute - saved.settings.startHour * 60) / saved.settings.slotIntervalMinutes;
    return find(tree, (node) => node.type === "td" && node.props["data-day"] === "Friday" && node.props["data-slot-index"] === index);
  };
  const drag = { preventDefault() {}, dataTransfer: { getData: () => mainId } };
  for (const time of ["08:00", "09:30", "10:00", "11:00", "12:00", "17:00"]) {
    assert.equal(fridayCell(time).props["aria-disabled"], true);
    assert.match(fridayCell(time).props.title, /Friday exams must use 09:00-10:00 or 10:30-11:30/);
    fridayCell(time).props.onDrop(drag);
    tree = app.render();
    button(tree, "Save Timetable").props.onClick();
    assert.deepEqual(JSON.parse(await app.downloads.at(-1).text()).assignments, saved.assignments,
      "An invalid Friday drop must not remove or move an existing exam");
  }
  let accepted = false;
  fridayCell("12:00").props.onDragOver({ preventDefault() { accepted = true; } });
  assert.equal(accepted, false);
  for (const time of ["09:00", "10:30"]) {
    assert.equal(fridayCell(time).props["aria-disabled"], undefined);
    fridayCell(time).props.onDrop(drag);
    tree = app.render();
    button(tree, "Save Timetable").props.onClick();
    const moved = JSON.parse(await app.downloads.at(-1).text());
    assert.deepEqual(moved.assignments[1].Friday[time], [mainId]);
    assert.equal(department.assignmentIds(moved.assignments).size, 1);
    assert.deepEqual(moved.assignments[1].Monday["12:00"], []);
  }
  step(tree, "pool").props.onClick();
  tree = app.render();
  assert.equal(include(roomNames[0]).props.checked, false);
  assert.equal(include(roomNames[1]).props.checked, true);
  assert.ok(!nodes(selection()).some((node) => node.type === "select"), "Pre-scheduling pool selection never exposes exam assignments");
  step(tree, "asd").props.onClick();
  tree = app.render();
  button(tree, "Create ASD Timetable").props.onClick();
  tree = app.render();
  assert.equal(fridayCell("12:00").props["aria-disabled"], undefined, "ASD is not restricted by the main timetable rule");
  fridayCell("12:00").props.onDrop(drag);
  tree = app.render();
  button(tree, "Save ASD Timetable").props.onClick();
  const asd = JSON.parse(await app.downloads.at(-1).text());
  assert.deepEqual(asd.assignments[1].Friday["12:00"], [mainId]);
  step(tree, "main").props.onClick();
  tree = app.render();
  button(tree, "Save Timetable").props.onClick();
  const afterAsd = JSON.parse(await app.downloads.at(-1).text());
  assert.equal(department.assignmentIds(afterAsd.assignments).has(mainId), false,
    "Placing an ASD exam removes its duplicate main placement without requiring a Continue button");
  assert.equal(department.assignmentIds(afterAsd.asdAssignments).has(mainId), true);
});

test("wizard imports, exam toggles, draft, saved-state round trip and report exclusion work together", {
  skip: [enrolmentPath, crnPath, asdPath].some((path) => !existsSync(path)),
}, async () => {
  const app = harness();
  let tree = app.render();
  button(tree, "Create Schedule").props.onClick();
  tree = app.render();
  button(tree, "Continue to Student Enrollment").props.onClick();
  tree = app.render();
  const enrolmentLabel = find(tree, (node) => node.type === "label" && text(node) === "Upload Enrollment File");
  await find(enrolmentLabel, (node) => node.type === "input").props.onChange(fileEvent(enrolmentPath));
  tree = app.render();
  assert.equal(step(tree, "asd").props.disabled, undefined);
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
  step(tree, "asd").props.onClick();
  tree = app.render();
  const asdInput = find(tree, (node) => node.type === "input" && node.props.accept === ".json,.xlsx,.xls,.csv");
  await asdInput.props.onChange(fileEvent(asdPath));
  tree = app.render();
  assert.match(text(tree), /Select Resources/);
  assert.ok(!nodes(tree).some((node) => node.props.className === "timetable"));
  step(tree, "main").props.onClick();
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
  button(loadedTree, "Load Schedule").props.onClick();
  await find(loadedTree, (node) => node.type === "input" && node.props.accept === "application/json").props.onChange(jsonEvent(snapshot));
  loadedTree = reloaded.render();
  assert.equal(step(loadedTree, "main").props["aria-current"], "step", "Loading from the start screen opens Timetable directly");
  assert.ok(!nodes(loadedTree).some((node) => node.props.className === "start-screen"));
  step(loadedTree, "courses").props.onClick();
  loadedTree = reloaded.render();
  assert.equal(review(loadedTree).props.examChoices[firstId], false);
  assert.equal(review(loadedTree).props.courses.reduce((sum, course) => sum + course.labSessions.length, 0), 46);
  const assignedId = [...department.assignmentIds(snapshot.assignments)].find((id) => id !== "ISET-4001");
  review(loadedTree).props.onExamChange(assignedId, false);
  step(loadedTree, "asd").props.onClick();
  loadedTree = reloaded.render();
  step(loadedTree, "pool").props.onClick();
  loadedTree = reloaded.render();
  step(loadedTree, "main").props.onClick();
  loadedTree = reloaded.render();
  button(loadedTree, "Save Timetable").props.onClick();
  const afterToggle = JSON.parse(await reloaded.downloads.at(-1).text());
  assert.equal(department.assignmentIds(afterToggle.assignments).size, 24);
  assert.equal(department.assignmentIds(afterToggle.assignments).has(assignedId), false);
  step(loadedTree, "resources").props.onClick();
  loadedTree = reloaded.render();
  assert.equal(!exportReady(loadedTree), true);
  find(loadedTree, (node) => node.type?.name === "ResourceAssignment").props.onAssign();
  loadedTree = reloaded.render();
  const fullResources = find(loadedTree, (node) => node.type?.name === "ResourceAssignment").props;
  assert.ok(!text(loadedTree).includes("Specialist"));
  assert.ok(!text(loadedTree).includes("Primary"));
  assert.match(text(loadedTree), /During Teaching Hours/);
  const pinned = fullResources.sessions.flatMap((session) => session.rooms).find((room) => room.fixedInvigilatorId);
  const leadSelect = find(loadedTree, (node) => node.type === "select" && node.props["aria-label"] === `Invigilator 1 for ${pinned.id}`);
  const alternativeLeads = nodes(leadSelect).filter((node) => node.type === "option" && node.props.value && node.props.value !== pinned.fixedInvigilatorId);
  assert.ok(alternativeLeads.length > 0);
  assert.ok(alternativeLeads.every((node) => node.props.disabled));
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
  assert.equal(!exportReady(loadedTree), true);
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
  step(loadedTree, "resources").props.onClick();
  loadedTree = reloaded.render();
  find(loadedTree, (node) => node.type?.name === "ResourceAssignment").props.onAssign();
  loadedTree = reloaded.render();
  const assignedResources = find(loadedTree, (node) => node.type?.name === "ResourceAssignment").props;
  assert.equal(assignedResources.validation.complete, true, JSON.stringify(assignedResources.validation.issues));
  assignedResources.sessions.forEach((session) => {
    const courseIds = new Set(session.rooms.map((room) => room.courseId));
    courseIds.forEach((courseId) => {
      const sizes = session.rooms.filter((room) => room.courseId === courseId).map((room) => room.students.length);
      const total = sizes.reduce((sum, size) => sum + size, 0);
      assert.deepEqual(sizes, examRooms.examRoomSizes(total, afterToggle.settings.studentsPerRoom));
    });
  });
  button(loadedTree, "Save Timetable").props.onClick();
  const savedResources = JSON.parse(await reloaded.downloads.at(-1).text());
  assert.ok(savedResources.resourcePlan.fingerprint);
  await find(loadedTree, (node) => node.type === "input" && node.props.accept === "application/json").props.onChange(jsonEvent(savedResources));
  loadedTree = reloaded.render();
  step(loadedTree, "resources").props.onClick();
  loadedTree = reloaded.render();
  assert.equal(!exportReady(loadedTree), false);
  const ready = find(loadedTree, (node) => node.type?.name === "ResourceAssignment").props;
  const firstAllocation = ready.plan.allocations[ready.sessions[0].rooms[0].id];
  ready.onPoolChange("invigilators", firstAllocation.invigilatorIds[0], { enabled: false });
  loadedTree = reloaded.render();
  assert.equal(!exportReady(loadedTree), true);
  find(loadedTree, (node) => node.type?.name === "ResourceAssignment").props.onPoolChange("invigilators", firstAllocation.invigilatorIds[0], { enabled: true });
  loadedTree = reloaded.render();
  assert.equal(!exportReady(loadedTree), false);
  step(loadedTree, "export").props.onClick();
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

test("CRN Info choices control invigilator availability across scheduling, review, export and save/load", async () => {
  const workbookEvent = (rows) => {
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "First");
    return { target: { files: [{ name: "input.xlsx", arrayBuffer: async () => XLSX.write(book, { type: "array", bookType: "xlsx" }) }], value: "input.xlsx" } };
  };
  const app = harness();
  let tree = app.render();
  button(tree, "Create Schedule").props.onClick();
  tree = app.render();
  button(tree, "Continue to Student Enrollment").props.onClick();
  tree = app.render();
  await find(find(tree, (node) => node.type === "label" && text(node) === "Upload Enrollment File"), (node) => node.type === "input").props.onChange(workbookEvent([
    ["Student ID", "Student Name", "Course Code", "Course Title", "CRN"], ["S1", "Student One", "EXAM-1000", "Selected Exam", "101"],
  ]));
  tree = app.render();
  await review(tree).props.onUpload(workbookEvent([
    ["Campus", "Crn No", "Course Code", "Title", "Cr", "Maximum Load", "No Of Enrolled", "Session Id", "DAYS", "Time", "Primary Instructor", "Second Instructor", "Type", "Building", "Room"],
    ["PAD", "101", "EXAM-1000", "Selected Exam", 3, 25, 1, "01", "M", "0800 - 0950", "10: Alice", "", "T", "Building", "A"],
    ["PAD", "201", "NOEXAM-1000", "Other Course", 3, 25, 1, "01", "M", "1200 - 1250", "10: Alice", "", "T", "Building", "B"],
    ["PAD", "301", "CAP-1000", "Capstone Project", 3, 25, 1, "01", "M", "1200 - 1250", "20: Bob", "", "GP", "Building", "C"],
    ["PAD", "401", "OCT-1000", "OCT", 3, 25, 1, "01", "M", "0800 - 0950", "30: Carol", "", "F", "Building", "D"],
  ]));
  tree = app.render();
  const other = review(tree).props.courses.find((course) => course.code === "NOEXAM-1000");
  const pool = () => find(tree, (node) => node.type?.name === "ResourcePool").props;
  const panel = () => find(tree, (node) => node.type?.name === "ResourceAssignment").props;
  step(tree, "pool").props.onClick();
  tree = app.render();
  assert.deepEqual(Array.from(pool().catalog.invigilators.find((person) => person.name === "Alice").busy, (entry) => entry.code), ["EXAM-1000", "NOEXAM-1000"]);
  assert.equal(pool().catalog.invigilators.find((person) => person.name === "Bob").busy.length, 0, "Default no-exam courses never block staff");
  pool().onPoolChange("invigilators", "staff:30", { enabled: false });
  tree = app.render();
  step(tree, "courses").props.onClick();
  tree = app.render();
  review(tree).props.onExamChange(other.id, false);
  tree = app.render();
  step(tree, "pool").props.onClick();
  tree = app.render();
  assert.deepEqual(Array.from(pool().catalog.invigilators.find((person) => person.name === "Alice").busy, (entry) => entry.code), ["EXAM-1000"]);
  assert.equal(pool().catalog.invigilators.find((person) => person.name === "Carol").enabled, false, "Changing exam choices preserves pool exclusions");
  assert.ok(pool().catalog.rooms.some((room) => room.busy.some((entry) => entry.code === "NOEXAM-1000")), "Teaching rooms remain blocked");
  step(tree, "main").props.onClick();
  tree = app.render();
  button(tree, "Auto-Schedule Remaining Exams").props.onClick();
  tree = app.render();
  step(tree, "resources").props.onClick();
  tree = app.render();
  assert.equal(panel().sessions[0].slotId, "12:00");
  assert.equal(panel().validation.complete, true);
  assert.deepEqual(Array.from(panel().plan.allocations[panel().sessions[0].rooms[0].id].invigilatorIds), ["staff:10"]);
  assert.deepEqual(Array.from(panel().plan.backups[panel().sessions[0].id]), ["staff:20"]);
  button(tree, "Save Timetable").props.onClick();
  const saved = JSON.parse(await app.downloads.at(-1).text());
  assert.ok(saved.resourceCatalog.invigilators.find((person) => person.name === "Alice").busy.some((entry) => entry.code === "NOEXAM-1000"));
  assert.equal(saved.examChoices[other.id], false);
  step(tree, "export").props.onClick();
  tree = app.render();
  assert.equal(find(tree, (node) => node.type?.name === "ExportStudio").props.ready, true);
  await button(tree, "Export Timetable").props.onClick();
  assert.equal(XLSX.read(await app.downloads.at(-1).arrayBuffer(), { type: "array" }).Sheets["Week 1 Invigilators"].G2.v, "Alice");
  const loaded = harness();
  tree = loaded.render();
  await find(tree, (node) => node.type === "input" && node.props.accept === "application/json").props.onChange(jsonEvent(saved));
  tree = loaded.render();
  step(tree, "resources").props.onClick();
  tree = loaded.render();
  assert.equal(panel().validation.complete, true, "The filtered catalog and assignment fingerprint survive save/load");
  step(tree, "courses").props.onClick();
  tree = loaded.render();
  review(tree).props.onExamChange(other.id, true);
  tree = loaded.render();
  step(tree, "resources").props.onClick();
  tree = loaded.render();
  assert.equal(panel().validation.complete, false);
  assert.ok(panel().validation.issues.some((issue) => /Alice.*NOEXAM-1000/.test(issue.message)), "Rechecking restores commitments, even when that course has no enrolment");
  assert.ok(panel().validation.issues.some((issue) => issue.title === "Resources need reassignment"));
});

test("manually moving a single-CRN exam to the final lab hour stays valid through resource review and save/load", async () => {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ["Student ID", "Student Name", "Course Code", "Course Title", "CRN"],
    ...Array.from({ length: 16 }, (_, index) => ["S" + index, "Student " + index, "LAB-1000", "Afternoon Lab", "101"]),
  ]), "Enrolment");
  const enrolment = imports.parseEnrolmentWorkbook(workbook);
  const exam = enrolment.courses[0];
  const labRoomId = department.roomIdentity("Campus", "Building", "Lab");
  const labMeeting = { crn: "101", days: ["Monday"], isLab: true, startMinutes: 780, endMinutes: 890,
    room: "Lab", building: "Building", campus: "Campus", instructor: "Lab Teacher",
    labInstructorId: "I0", instructors: [{ id: "I0", name: "Lab Teacher" }], teachingInstructorIds: ["I0"] };
  const labBusy = { code: exam.code, crn: "101", isLab: true, days: ["Monday"], start: 780, end: 890 };
  const snapshot = { type: "main", version: 5, startDate: "2026-10-19", weeks: [1], selectedWeek: 1,
    settings: { slotIntervalMinutes: 30, startHour: 8, endHour: 18, studentsPerRoom: 25, examDurationMinutes: 60 },
    assignments: { 1: { Monday: { "13:00": [exam.id] } } }, asdAssignments: {}, asdExamDurations: {}, hasAsdStep: false,
    courses: enrolment.courses, studentDirectory: enrolment.studentDirectory,
    departmentSelection: { sheetName: "First", courses: [{ code: exam.code, title: exam.title, crns: ["101"], meetings: [labMeeting] }] },
    examChoices: { [exam.id]: true }, roomDistributionChoices: {}, resourcePlan: null,
    resourceCatalog: {
      rooms: [{ id: labRoomId, name: "Campus / Building / Lab", enabled: true, busy: [labBusy] }],
      invigilators: Array.from({ length: 4 }, (_, index) => ({ id: "I" + index, name: index === 0 ? "Lab Teacher" : "Staff " + index,
        enabled: true, busy: index === 0 ? [labBusy] : [] })),
    },
  };
  const app = harness();
  let tree = app.render();
  await find(tree, (node) => node.type === "input" && node.props.accept === "application/json").props.onChange(jsonEvent(snapshot));
  tree = app.render();
  step(tree, "main").props.onClick();
  tree = app.render();
  const target = find(tree, (node) => node.type === "td" && node.props["data-day"] === "Monday" && node.props["data-slot-index"] === 12);
  target.props.onDrop({ preventDefault() {}, dataTransfer: { getData: () => exam.id } });
  tree = app.render();
  step(tree, "resources").props.onClick();
  tree = app.render();
  let panel = find(tree, (node) => node.type?.name === "ResourceAssignment").props;
  assert.equal(panel.sessions.length, 1);
  assert.equal(panel.sessions[0].start, 840);
  assert.equal(panel.sessions[0].end, 900);
  assert.deepEqual(Array.from(panel.sessions[0].issues), []);
  assert.ok(!panel.validation.issues.some((issue) => issue.title === "Exam scheduling issue"));
  button(tree, "Auto-Assign Resources").props.onClick();
  tree = app.render();
  panel = find(tree, (node) => node.type?.name === "ResourceAssignment").props;
  assert.equal(panel.validation.complete, true);
  const allocation = panel.plan.allocations[panel.sessions[0].rooms[0].id];
  assert.equal(allocation.roomId, labRoomId);
  assert.equal(allocation.invigilatorIds[0], "I0");
  assert.equal(resources.invigilatorWorkloads(panel.catalog, panel.sessions, panel.plan).find((person) => person.id === "I0").exam, 0);
  button(tree, "Save Timetable").props.onClick();
  const saved = JSON.parse(await app.downloads.at(-1).text());
  assert.deepEqual(saved.assignments[1].Monday["13:00"], []);
  assert.deepEqual(saved.assignments[1].Monday["14:00"], [exam.id]);
  assert.equal(saved.departmentSelection.courses[0].meetings[0].endMinutes, 890, "The original lab data is not rewritten");
  const loaded = harness();
  tree = loaded.render();
  await find(tree, (node) => node.type === "input" && node.props.accept === "application/json").props.onChange(jsonEvent(saved));
  tree = loaded.render();
  step(tree, "resources").props.onClick();
  tree = loaded.render();
  assert.equal(exportReady(tree), true);
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
  step(tree, "resources").props.onClick();
  tree = app.render();
  button(tree, "Auto-Assign Resources").props.onClick();
  tree = app.render();
  const panel = () => find(tree, (node) => node.type?.name === "ResourceAssignment").props;
  const examRooms = () => panel().sessions[0].rooms.filter((room) => room.courseId === exam.id);
  const decisionButton = (label) => find(tree, (node) => node.type === "button" && node.props["aria-label"] === `${label} for ${exam.code}`);
  assert.deepEqual(examRooms().map((room) => room.students.length), [19, 18, 15]);
  assert.deepEqual(examRooms().map((room) => room.requiredInvigilators), [2, 2, 1]);
  assert.equal(decisionButton("Keep the extra room (normal limit)").props["aria-pressed"], true);
  const otherRoom = panel().sessions[0].rooms.find((room) => room.courseId === other.id);
  const otherAllocation = panel().plan.allocations[otherRoom.id];
  decisionButton("Merge overflow into one room").props.onClick();
  tree = app.render();
  assert.deepEqual(examRooms().map((room) => room.students.length), [27, 25]);
  assert.equal(panel().validation.complete, true);
  assert.deepEqual(panel().plan.allocations[otherRoom.id], otherAllocation);
  assert.match(text(tree), /Exception approved for this exam/);
  decisionButton("Distribute across remaining rooms").props.onClick();
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
  step(tree, "resources").props.onClick();
  tree = reloaded.render();
  assert.equal(decisionButton("Distribute across remaining rooms").props["aria-pressed"], true);
  assert.equal(panel().validation.complete, true);
  assert.equal(panel().sessions[0].rooms.find((room) => room.courseId === other.id).maxStudents, 25);
  step(tree, "export").props.onClick();
  tree = reloaded.render();
  await button(tree, "Export Timetable").props.onClick();
  const exported = XLSX.read(await reloaded.downloads.at(-1).arrayBuffer(), { type: "array" });
  const reportRows = XLSX.utils.sheet_to_json(exported.Sheets["Week 1 Invigilators"], { header: 1 });
  const examRows = reportRows.slice(1).filter((row) => row[1] === exam.code);
  assert.deepEqual(examRows.map((row) => row[3]), [26, 26]);
  assert.ok(examRows.every((row) => /Approved up to 27/.test(row[14])));
  step(tree, "resources").props.onClick();
  tree = reloaded.render();
  decisionButton("Keep the extra room (normal limit)").props.onClick();
  tree = reloaded.render();
  assert.deepEqual(examRooms().map((room) => room.students.length), [19, 18, 15]);
  assert.equal(!exportReady(tree), true, "The restored room must be assigned before export");
});

test("export workspace previews audiences, combines or splits weeks, exports CSV and prints only selected report data", async () => {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
    ["Student ID", "Student Name", "Course Code", "Course Title", "CRN"],
    ...Array.from({ length: 16 }, (_, i) => ["A" + i, "Alice " + i, "MAIN-1000", "First Exam", "101"]),
    ...Array.from({ length: 10 }, (_, i) => ["B" + i, "Bob " + i, "MAIN-2000", "Second Exam", "201"]),
    ["ASD1", "ASD Student", "ASD-1000", "Excluded ASD Exam", "301"],
    ["ASD2", "ASD Private Student", "ASD-2000", "ASD-only Week Exam", "302"],
  ]), "Enrolment");
  const enrolment = imports.parseEnrolmentWorkbook(book);
  const main = enrolment.courses.filter((course) => course.code.startsWith("MAIN"));
  const asd = enrolment.courses.find((course) => course.code.startsWith("ASD"));
  const asdOnly = enrolment.courses.find((course) => course.code === "ASD-2000");
  const labRoomId = department.roomIdentity("PAD", "P-B-4F", "13");
  const labMeeting = { crn: "101", days: ["Monday"], isLab: true, startMinutes: 480, endMinutes: 590,
    room: "13", building: "P-B-4F", campus: "PAD", instructor: "Staff 0",
    labInstructorId: "I0", instructors: [{ id: "I0", name: "Staff 0" }], teachingInstructorIds: ["I0"] };
  const labBusy = { code: main[0].code, crn: "101", isLab: true, days: ["Monday"], start: 480, end: 590 };
  const snapshot = { type: "main", version: 5, startDate: "2026-10-20", weeks: [1, 2, 3], selectedWeek: 1,
    settings: { slotIntervalMinutes: 30, startHour: 8, endHour: 18, studentsPerRoom: 25, examDurationMinutes: 60 },
    assignments: { 1: { Monday: { "09:00": [main[0].id] } }, 2: { Tuesday: { "12:00": [main[1].id] } }, 3: {} },
    asdAssignments: { 1: { Monday: { "09:00": [asd.id] } }, 3: { Wednesday: { "10:30": [asdOnly.id] } } },
    asdExamDurations: { [asd.id]: 60, [asdOnly.id]: 90 }, hasAsdStep: true,
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
  step(tree, "resources").props.onClick();
  tree = app.render();
  button(tree, "Auto-Assign Resources").props.onClick();
  tree = app.render();
  step(tree, "export").props.onClick();
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
  assert.equal(nodes(printView()).filter((node) => node.props.className === "export-board-print").length, 1);
  assert.ok(!nodes(printView()).some((node) => node.type?.name === "PreviewTable"), "Week board printing does not fall back to the chronological table");
  assert.match(text(find(panel(), (node) => node.props.className === "export-print-layout")), /Print layout: Week board/);
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
  const printedBoards = nodes(printView()).filter((node) => node.type?.name === "PrintWeekBoard");
  assert.deepEqual(JSON.parse(JSON.stringify(printedBoards.map((board) => board.props.week))), [1, 2]);
  assert.ok(printedBoards.every((board) => nodes(board).filter((node) => node.type === "th" && node.props.scope === "col").length === 5));
  assert.match(text(printedBoards[0]), /2 primary invigilators needed/);
  assert.match(text(printedBoards[0]), /During lab time/);
  assert.match(text(printedBoards[1]), /1 primary invigilator needed/);
  assert.ok(!nodes(printedBoards[1]).some((node) => node.props.className === "export-lab"));
  assert.ok(!text(printView()).includes("PAD"));
  button(tree, "Print / Save PDF").props.onClick();
  assert.equal(app.prints(), 1);
  button(tree, "Chronological list").props.onClick();
  tree = app.render();
  assert.match(text(find(panel(), (node) => node.props.className === "export-print-layout")), /Print layout: Chronological list/);
  assert.ok(!nodes(printView()).some((node) => node.props.className === "export-board-print"));
  const printedOverviews = nodes(printView()).filter((node) => node.type?.name === "PreviewTable" && node.props.label === "Exam overview");
  const chronologicalColumns = ["#", "Date", "Day", "Time", "Course", "Course title"];
  assert.deepEqual(JSON.parse(JSON.stringify(printedOverviews.map((table) => table.props.rows[0].slice(0, 5)))),
    [[1, "19/10/2026", "Monday", "09:00-10:00", "MAIN-1000"], [1, "27/10/2026", "Tuesday", "12:00-13:00", "MAIN-2000"]]);
  assert.ok(printedOverviews.every((table) => JSON.stringify(table.props.columns) === JSON.stringify(chronologicalColumns)));
  const labList = find(preview(), (node) => node.type?.name === "PreviewTable" && node.props.label === "Exam overview preview");
  assert.equal(labList.props.className, "export-table--chronological");
  assert.deepEqual(JSON.parse(JSON.stringify(labList.props.rows[0].slice(0, 5))), [1, "19/10/2026", "Monday", "09:00-10:00", "MAIN-1000"]);
  assert.ok(!text(preview()).includes("Primary invigilators"));
  await button(tree, "Download Excel").props.onClick();
  const chronologicalZip = await JSZip.loadAsync(await app.downloads.at(-1).arrayBuffer());
  for (const week of [1, 2]) {
    const file = await chronologicalZip.file("Week_" + week + "_Exam_Overview.xlsx").async("uint8array");
    const sheet = XLSX.read(file, { type: "array", cellDates: true, cellNF: true }).Sheets["Exam overview"];
    assert.deepEqual(XLSX.utils.sheet_to_json(sheet, { header: 1 })[0], chronologicalColumns);
    assert.equal(sheet["!ref"], "A1:F2");
    assert.equal(sheet.A2.v, 1);
    assert.equal(sheet.B2.t, "d");
    assert.equal(sheet.B2.z, "dd/mm/yyyy");
  }
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
  assert.match(overviewCsv, /"#","Date","Day","Time","Course","Course title"/);
  assert.match(overviewCsv, /"1","27\/10\/2026","Tuesday","12:00-13:00","MAIN-2000"/);
  assert.ok(!overviewCsv.includes("Primary invigilators"));

  const asdToggle = () => find(find(panel(), (node) => node.type === "label" && text(node).trim() === "Include ASD exams"),
    (node) => node.type === "input");
  assert.equal(asdToggle().props.checked, false);
  weekInput(1).props.onChange();
  tree = app.render();
  const departmentStats = () => nodes(find(panel(), (node) => node.props.className === "export-stats"))
    .filter((node) => node.type === "div").slice(2).map(text);
  const statsBeforeAsd = departmentStats();
  asdToggle().props.onChange({ target: { checked: true } });
  tree = app.render();
  assert.equal(weekInput(3).props.checked, true, "ASD-only weeks are offered only when opted in");
  assert.match(text(panel()), /2 ASD reference exams in included weeks/);
  assert.deepEqual(departmentStats(), statsBeforeAsd,
    "Department exam, room and student totals are unchanged");
  button(preview(), "Week 1").props.onClick();
  button(tree, "Week board").props.onClick();
  tree = app.render();
  const referenceCard = find(preview(), (node) => node.props.className === "export-board__exam export-board__exam--asd");
  assert.match(text(referenceCard), /09:00-10:00ASD-1000/);
  assert.match(text(referenceCard), /ASD \/ reference/);
  assert.ok(!text(referenceCard).includes("invigilator"));
  assert.ok(!text(referenceCard).includes("students"));
  assert.ok(!text(printView()).includes("ASD Private Student"));
  assert.equal(nodes(printView()).filter((node) => node.props.className === "export-board__exam export-board__exam--asd").length, 2);
  button(tree, "Chronological list").props.onClick();
  tree = app.render();
  const asdList = find(preview(), (node) => node.type?.name === "PreviewTable" && node.props.label === "Exam overview preview");
  assert.equal(asdList.props.rows[1][4], "ASD-1000", "Department exam stays first at the same time");
  assert.ok(asdList.props.rows[1].at(-1).endsWith(" (ASD)"));
  assert.equal(asdList.props.rowClassNames[1], "export-row--asd");
  assert.equal(asdList.props.rows[1].length, 6);
  const asdPrintedTables = nodes(printView()).filter((node) => node.type?.name === "PreviewTable" && node.props.label === "Exam overview");
  assert.equal(asdPrintedTables.length, 3);
  assert.equal(asdPrintedTables[0].props.rowClassNames[1], "export-row--asd");
  assert.equal(asdPrintedTables[2].props.rows[0][4], "ASD-2000");
  assert.match(text(asdPrintedTables[2]), /10:30-12:00/);
  await button(tree, "Download CSV").props.onClick();
  const asdCsvZip = await JSZip.loadAsync(await app.downloads.at(-1).arrayBuffer());
  assert.deepEqual(Object.keys(asdCsvZip.files).sort(), ["Week_1_Exam_Overview.csv", "Week_2_Exam_Overview.csv", "Week_3_Exam_Overview.csv"]);
  assert.match(await asdCsvZip.file("Week_1_Exam_Overview.csv").async("string"), /"ASD-1000".* \(ASD\)"/);
  assert.match(await asdCsvZip.file("Week_3_Exam_Overview.csv").async("string"), /"10:30-12:00","ASD-2000"/);
  weekInput(1).props.onChange();
  weekInput(2).props.onChange();
  tree = app.render();
  assert.match(text(preview()), /ASD-2000/);
  assert.ok(!text(preview()).includes("MAIN-"));
  radio("Excel (.xlsx)").props.onChange();
  tree = app.render();
  await button(tree, "Download Excel").props.onClick();
  assert.equal(app.downloadNames.at(-1), "Week_3_Exam_Overview.xlsx");
  const asdWorkbook = XLSX.read(await app.downloads.at(-1).arrayBuffer(), { type: "array" });
  const asdExcelRow = XLSX.utils.sheet_to_json(asdWorkbook.Sheets["Exam overview"])[0];
  assert.equal(asdExcelRow.Course, "ASD-2000");
  assert.ok(asdExcelRow["Course title"].endsWith(" (ASD)"));
  assert.equal(asdExcelRow.Schedule, undefined);
  assert.equal(asdExcelRow.Students, undefined);
  assert.equal(asdExcelRow["Primary invigilators needed"], undefined);
  asdToggle().props.onChange({ target: { checked: false } });
  tree = app.render();
  assert.equal(button(tree, "Download Excel").props.disabled, true, "Turning ASD off with only an ASD week selected leaves no export");
  assert.ok(!nodes(panel()).some((node) => node.props["aria-label"] === "Include Week 3"));
  asdToggle().props.onChange({ target: { checked: true } });
  tree = app.render();
  weekInput(2).props.onChange();
  tree = app.render();

  reportButton("Staff duties").props.onClick();
  tree = app.render();
  assert.ok(!nodes(panel()).some((node) => node.props.className === "export-asd-option"));
  assert.ok(!nodes(panel()).some((node) => node.props["aria-label"] === "Include Week 3"));
  assert.ok(!text(printView()).includes("ASD-"));
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
  assert.ok(!nodes(panel()).some((node) => node.props.className === "export-asd-option"));
  assert.ok(!text(printView()).includes("ASD-"));
  assert.match(text(preview()), /Bob 0/);
  const search = find(preview(), (node) => node.type === "input" && node.props.type === "search");
  search.props.onChange({ target: { value: "NO-MATCH" } });
  tree = app.render();
  assert.match(text(preview()), /No matching records/);
  assert.match(text(printView()), /Bob 0/, "Preview search must not filter print output");
  button(tree, "Print / Save PDF").props.onClick();
  assert.equal(app.prints(), 2);
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
  assert.ok(!nodes(panel()).some((node) => node.props.className === "export-asd-option"));
  assert.ok(!text(preview()).includes("ASD-"));
  assert.ok(!text(printView()).includes("ASD-"));
  await button(tree, "Export Timetable").props.onClick();
  const completeAfterAsd = XLSX.read(await app.downloads.at(-1).arrayBuffer(), { type: "array" });
  assert.ok(!JSON.stringify(completeAfterAsd.Sheets).includes("ASD-"));
  assert.equal(radio("Excel (.xlsx)").props.checked, true);
  assert.equal(radio("CSV (.csv)").props.disabled, true);
  reportButton("Course seating").props.onClick();
  tree = app.render();
  assert.ok(!nodes(panel()).some((node) => node.props.className === "export-asd-option" || node.props.name === "export-format" || node.props.name === "export-packaging"));
  assert.equal(text(find(preview(), (node) => node.type === "h4")), "MAIN-1000");
  assert.match(text(preview()), /StudentIDRoom/);
  assert.ok(!text(preview()).includes("Alice"));
  assert.ok(!text(printView()).includes("ASD-"));
  const seatingTable = () => find(preview(), (node) => node.type?.name === "PreviewTable" && node.props.label === "Course seating preview");
  assert.equal(seatingTable().props.rows.length, 16);
  find(preview(), (node) => node.type === "input" && node.props.type === "search").props.onChange({ target: { value: "NO-MATCH" } });
  tree = app.render();
  assert.equal(seatingTable().props.rows.length, 0);
  assert.equal(nodes(printView()).filter((node) => node.type?.name === "PreviewTable")[0].props.rows.length, 16);
  await button(tree, "Download This Exam").props.onClick();
  assert.equal(app.downloadNames.at(-1), "MAIN-1000 - First Exam.xlsx");
  const singleSeating = XLSX.read(await app.downloads.at(-1).arrayBuffer(), { type: "array" });
  assert.deepEqual(singleSeating.SheetNames, ["CRN 101"]);
  assert.equal(XLSX.utils.sheet_to_json(singleSeating.Sheets["CRN 101"], { header: 1 }).slice(8).length, 16, "Search must not filter a seating file");
  weekInput(2).props.onChange();
  tree = app.render();
  assert.equal(nodes(printView()).filter((node) => node.props.className === "export-print__crn").length, 2);
  await button(tree, "Download Course Files").props.onClick();
  assert.equal(app.downloadNames.at(-1), "Course_Seating_Exam_Files.zip");
  const seatingZip = await JSZip.loadAsync(await app.downloads.at(-1).arrayBuffer());
  assert.deepEqual(Object.keys(seatingZip.files).sort(), ["MAIN-1000 - First Exam.xlsx", "MAIN-2000 - Second Exam.xlsx"]);
  weekInput(1).props.onChange();
  tree = app.render();
  assert.equal(text(find(preview(), (node) => node.type === "h4")), "MAIN-2000");
  await button(tree, "Download Course Files").props.onClick();
  assert.equal(app.downloadNames.at(-1), "MAIN-2000 - Second Exam.xlsx", "One exam downloads directly without a ZIP");
  button(tree, "Print / Save PDF").props.onClick();
  assert.equal(app.prints(), 3);
  weekInput(2).props.onChange();
  tree = app.render();
  assert.equal(button(tree, "Download Course Files").props.disabled, true);
  assert.equal(button(tree, "Print / Save PDF").props.disabled, true);
  assert.match(text(preview()), /Choose a week to preview/);
  weekInput(1).props.onChange();
  tree = app.render();
  reportButton("Exam overview").props.onClick();
  tree = app.render();
  assert.equal(asdToggle().props.checked, true, "The overview preference is retained without affecting another report");
  assert.equal(weekInput(3).props.checked, true);
  assert.match(text(preview()), /ASD-1000/);
});

test("course seating preview switches exams and CRN sheets, paginates IDs and prints all included rosters", async () => {
  const makeCourse = (id, count, crns) => ({ id, code: id, title: id + " title", crns, labSessions: [],
    students: Array.from({ length: count }, (_, index) => ({ id: id + "-" + String(index).padStart(4, "0"),
      name: "Private student " + index, crn: crns[index < 55 ? 0 : 1] })) });
  const courses = { A: makeCourse("A", 60, ["101", "102"]), B: makeCourse("B", 3, ["201"]), C: makeCourse("C", 2, ["301"]) };
  const sessions = resources.buildExamSessions({ 1: { Monday: { "12:00": ["A", "B"] } }, 2: { Tuesday: { "12:00": ["C"] } } }, courses, 60);
  const catalog = {
    rooms: Array.from({ length: 4 }, (_, i) => ({ id: "R" + i, name: "PAD / P-B-4F / " + (13 + i), enabled: true, busy: [] })),
    invigilators: Array.from({ length: 10 }, (_, i) => ({ id: "I" + i, name: "Staff " + i, enabled: true, busy: [] })),
  };
  const exports = [];
  const props = { sessions, catalog, plan: resources.assignResources(sessions, catalog), startDate: "2026-10-19", ready: true, isExporting: false,
    asdExams: exportReports.buildAsdOverviewExams({ assignments: { 3: { Monday: { "12:00": ["ASD"] } } }, courseLookup: { ASD: makeCourse("ASD", 1, ["999"]) } }),
    onExport(options) { exports.push(JSON.parse(JSON.stringify(options))); } };
  const app = harness("ExportStudio", props);
  let tree = app.render();
  button(tree, "Exam overview").props.onClick();
  tree = app.render();
  find(find(tree, (node) => node.props.className === "export-asd-option"), (node) => node.type === "input").props.onChange({ target: { checked: true } });
  tree = app.render();
  button(tree, "Student room lists").props.onClick();
  tree = app.render();
  find(tree, (node) => node.type === "input" && node.props.name === "export-format" && node.props.checked === false).props.onChange();
  tree = app.render();
  button(tree, "Course seating").props.onClick();
  tree = app.render();
  const preview = () => find(tree, (node) => node.props.className === "export-preview");
  const printView = () => find(tree, (node) => node.props.className === "export-print");
  const table = () => find(preview(), (node) => node.type?.name === "PreviewTable");
  const infoLabels = (root) => nodes(find(root, (node) => node.props.className === "export-seating-info"))
    .filter((node) => node.type === "dt").map(text);
  const crnButton = (crn) => find(preview(), (node) => node.type === "button" && text(node).trim().startsWith("CRN " + crn + " "));
  assert.equal(table().props.rows.length, 55);
  assert.equal(table().props.limit, 50);
  assert.deepEqual(Array.from(table().props.columns), ["StudentID", "Room"]);
  assert.deepEqual(infoLabels(preview()), ["Course", "Course title", "Date", "Day", "Time"]);
  assert.deepEqual(infoLabels(printView()), ["Course", "Course title", "Date", "Day", "Time"]);
  assert.ok(!nodes(printView()).filter((node) => node.type === "h4").some((node) => text(node).includes("CRN")));
  assert.ok(!text(preview()).includes("Private student"));
  assert.ok(!text(preview()).includes("PAD"));
  assert.ok(!nodes(tree).some((node) => node.props["aria-label"] === "Include Week 3"));
  assert.equal(nodes(printView()).filter((node) => node.props.className === "export-print__crn").length, 4);
  const printedTables = () => nodes(printView()).filter((node) => node.type?.name === "PreviewTable");
  assert.equal(printedTables().reduce((sum, item) => sum + item.props.rows.length, 0), 65);
  assert.ok(printedTables().every((item) => item.props.limit === undefined), "Printed CRN lists are never truncated to preview pages");
  button(tree, "Show 50 more (5 remaining)").props.onClick();
  tree = app.render();
  assert.equal(table().props.limit, 100);
  crnButton("102").props.onClick();
  tree = app.render();
  assert.equal(crnButton("102").props["aria-pressed"], true);
  assert.equal(table().props.rows.length, 5);
  assert.equal(table().props.limit, 50);
  find(preview(), (node) => node.type === "input" && node.props.type === "search").props.onChange({ target: { value: "NO-MATCH" } });
  tree = app.render();
  assert.equal(table().props.rows.length, 0);
  await button(tree, "Download This Exam").props.onClick();
  assert.deepEqual(exports.at(-1), { report: "seating", format: "xlsx", packaging: "course", weeks: [1, 2], includeAsd: false, examIds: ["1/Monday/12:00/A"] });
  assert.equal(printedTables()[0].props.rows.length, 55);
  find(preview(), (node) => node.type === "select" && node.props["aria-label"] === "Seating exam").props.onChange({ target: { value: "1/Monday/12:00/B" } });
  tree = app.render();
  assert.equal(text(find(preview(), (node) => node.type === "h4")), "B");
  assert.equal(crnButton("201").props["aria-pressed"], true);
  assert.equal(table().props.rows.length, 3);
  assert.equal(find(preview(), (node) => node.type === "input" && node.props.type === "search").props.value, "");
  button(tree, "Week 2").props.onClick();
  tree = app.render();
  assert.equal(text(find(preview(), (node) => node.type === "h4")), "C");
  assert.equal(crnButton("301").props["aria-pressed"], true);
  assert.equal(printedTables().length, 4, "Preview selection cannot filter print");
  props.isExporting = true;
  tree = app.render();
  assert.equal(button(tree, "Download This Exam").props.disabled, true);
  assert.equal(button(tree, "Exporting...").props.disabled, true);
  props.isExporting = false;
  props.ready = false;
  tree = app.render();
  assert.equal(button(tree, "Download Course Files").props.disabled, true);
  assert.match(text(preview()), /Confirm resources first/);
  assert.ok(!nodes(tree).some((node) => node.props.className === "export-print"));
});

test("week-board printing keeps every exam in busy days and all included weeks, independently of the active preview", () => {
  const courses = Object.fromEntries(Array.from({ length: 13 }, (_, i) => {
    const id = "EXAM-" + i;
    return [id, { id, code: id, title: "Exam " + i, crns: ["CRN-" + i], labSessions: [],
      students: [{ id: "S" + i, name: "Private Student " + i, crn: "CRN-" + i }] }];
  }));
  const slots = Object.fromEntries(Array.from({ length: 6 }, (_, i) => [
    String(8 + i).padStart(2, "0") + ":00", ["EXAM-" + (2 * i), "EXAM-" + (2 * i + 1)],
  ]));
  const sessions = resources.buildExamSessions({ 1: { Monday: slots }, 2: { Tuesday: { "12:00": ["EXAM-12"] } } }, courses, 60);
  const catalog = {
    rooms: Array.from({ length: 2 }, (_, i) => ({ id: "R" + i, name: "PAD / P-B-4F / " + (13 + i), enabled: true, busy: [] })),
    invigilators: Array.from({ length: 4 }, (_, i) => ({ id: "I" + i, name: "Staff " + i, enabled: true, busy: [] })),
  };
  const plan = resources.assignResources(sessions, catalog);
  const app = harness("ExportStudio", { sessions, catalog, plan, startDate: "2026-10-19", ready: true, isExporting: false, onExport() {} });
  let tree = app.render();
  find(tree, (node) => node.type === "button" && node.props["aria-label"] === "Exam overview").props.onClick();
  tree = app.render();
  const printView = () => find(tree, (node) => node.props.className === "export-print");
  const printedBoards = () => nodes(printView()).filter((node) => node.props.className === "export-board-print");
  const cardCodes = (board) => nodes(board).filter((node) => node.type === "h4").map(text);
  assert.equal(printedBoards().length, 2);
  assert.deepEqual(cardCodes(printedBoards()[0]), Object.keys(courses).slice(0, 12));
  assert.deepEqual(cardCodes(printedBoards()[1]), ["EXAM-12"]);
  assert.equal(nodes(printedBoards()[0]).filter((node) => node.type === "tbody")[0].children.length, 12, "A dense day is not truncated to a fixed number of cards");
  assert.equal(nodes(printedBoards()[0]).filter((node) => node.props.className === "export-board__free").length, 4);
  assert.match(text(printedBoards()[0]), /Monday19 OctTuesday20 OctWednesday21 OctThursday22 OctFriday23 Oct/);
  assert.match(text(printedBoards()[0]), /08:00-09:00/);
  assert.match(text(printedBoards()[0]), /13:00-14:00/);
  assert.ok(!text(printView()).includes("Private Student"));
  assert.ok(!text(printView()).includes("PAD"));
  const preview = () => find(tree, (node) => node.props.className === "export-preview");
  button(preview(), "Week 2").props.onClick();
  tree = app.render();
  assert.ok(!text(preview()).includes("EXAM-0"), "The screen switches weeks");
  assert.equal(printedBoards().length, 2, "The print still contains every included week");
  find(tree, (node) => node.type === "input" && node.props["aria-label"] === "Include Week 1").props.onChange();
  tree = app.render();
  assert.equal(printedBoards().length, 1);
  assert.deepEqual(cardCodes(printedBoards()[0]), ["EXAM-12"]);
  button(tree, "Print / Save PDF").props.onClick();
  assert.equal(app.prints(), 1);
  button(tree, "Chronological list").props.onClick();
  tree = app.render();
  assert.equal(printedBoards().length, 0);
  const table = find(printView(), (node) => node.type?.name === "PreviewTable");
  assert.equal(table.props.rows.length, 1);
  assert.equal(table.props.rows[0][4], "EXAM-12");
  button(tree, "Week board").props.onClick();
  tree = app.render();
  assert.equal(printedBoards().length, 1, "Switching back restores the printable board");
  find(tree, (node) => node.type === "button" && node.props["aria-label"] === "Staff duties").props.onClick();
  tree = app.render();
  assert.equal(printedBoards().length, 0, "The board mode cannot replace another report's print layout");
  assert.ok(!nodes(tree).some((node) => node.props.className === "export-print-layout"));
  find(tree, (node) => node.type === "input" && node.props["aria-label"] === "Include Week 2").props.onChange();
  tree = app.render();
  assert.equal(button(tree, "Print / Save PDF").props.disabled, true);
  assert.ok(!nodes(tree).some((node) => node.props.className === "export-print"), "No print content is generated for an empty selection");
});
