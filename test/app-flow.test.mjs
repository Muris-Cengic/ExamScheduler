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
const button = (tree, label) => find(tree, (node) => node.type === "button" && text(node) === label);
const review = (tree) => find(tree, (node) => node.type?.name === "CourseSelection");
const root = "data/input/26-27 S1/";
const enrolmentPath = root + "Students Registration 26-27 S1.xlsx";
const crnPath = root + "CRN List 26-27 S1.xlsx";
const asdPath = root + "Other Departments Schedule/Midterm Schedule 26-27 S1.xlsx";

test("resource selection precedes generation and stays editable after the resource-aware draft", async () => {
  const workbook = (rows) => {
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "First");
    const data = XLSX.write(book, { type: "array", bookType: "xlsx" });
    return { target: { files: [{ name: "input.xlsx", arrayBuffer: async () => data }], value: "input.xlsx" } };
  };
  const app = harness();
  let tree = app.render();
  const upload = find(find(tree, (node) => node.type === "label" && text(node) === "Upload Enrollment File"),
    (node) => node.type === "input");
  await upload.props.onChange(workbook([
    ["Student ID", "Student Name", "Course Code", "Course Title", "CRN"],
    ["S1", "Student One", "MAIN-1000", "Main Exam", "101"],
    ["S2", "Student Two", "MAIN-1000", "Main Exam", "102"],
  ]));
  tree = app.render();
  await review(tree).props.onUpload(workbook([
    ["Campus", "Crn No", "Course Code", "Title", "Cr", "Maximum Load", "No Of Enrolled", "Session Id", "DAYS", "Time", "Primary Instructor", "Second Instructor", "Type", "Building", "Room"],
    ["PAD", "101", "MAIN-1000", "Main Exam", 3, 25, 1, "01", "M", "0800 - 0950", "10: Alice", "20: Bob", "T", "P-B-4F", "13"],
    ["PAD", "102", "MAIN-1000", "Main Exam", 3, 25, 1, "02", "T", "0800 - 0950", "30: Carol", "", "T", "P-B-4F", "14"],
  ]));
  tree = app.render();
  button(tree, "Continue To ASD").props.onClick();
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
  const roomNames = pool().props.catalog.rooms.map((room) => room.name);
  include(roomNames[0]).props.onChange({ target: { checked: false } });
  tree = app.render();
  include(roomNames[1]).props.onChange({ target: { checked: false } });
  tree = app.render();
  assert.equal(button(tree, "Continue To Scheduling").props.disabled, true);
  include(roomNames[1]).props.onChange({ target: { checked: true } });
  include("Carol").props.onChange({ target: { checked: false } });
  include("Bob").props.onChange({ target: { checked: false } });
  tree = app.render();
  assert.equal(button(tree, "Continue To Scheduling").props.disabled, true, "One person cannot cover both an exam and backup");
  include("Carol").props.onChange({ target: { checked: true } });
  include("Bob").props.onChange({ target: { checked: true } });
  tree = app.render();
  assert.equal(button(tree, "Continue To Scheduling").props.disabled, false);
  button(tree, "Continue To Scheduling").props.onClick();
  tree = app.render();
  button(tree, "Auto-Schedule Remaining Exams").props.onClick();
  tree = app.render();
  assert.match(text(tree), /1 exams added. 0 remain unplaced/);
  button(tree, "Proceed To Resources").props.onClick();
  tree = app.render();
  const assignment = () => find(tree, (node) => node.type?.name === "ResourceAssignment").props;
  assert.equal(assignment().validation.complete, true, "The generated draft includes its checked allocation");
  assert.equal(button(tree, "Proceed To Export").props.disabled, false);
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
  assert.equal(button(tree, "Proceed To Export").props.disabled, true, "Pool edits invalidate the generated plan until review");
  assert.match(text(tree), /Confirm valid room and staff assignments to assess standby capacity/);
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
  button(tree, "Back To Scheduling").props.onClick();
  tree = app.render();
  button(tree, "Review Resource Pool").props.onClick();
  tree = app.render();
  assert.equal(include(roomNames[0]).props.checked, false);
  assert.equal(include(roomNames[1]).props.checked, true);
  assert.ok(!nodes(selection()).some((node) => node.type === "select"), "Pre-scheduling pool selection never exposes exam assignments");

});

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
  assert.match(text(tree), /Select Resources Before Scheduling/);
  assert.ok(!nodes(tree).some((node) => node.props.className === "timetable"));
  button(tree, "Continue To Scheduling").props.onClick();
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
  button(loadedTree, "Continue To Resources").props.onClick();
  loadedTree = reloaded.render();
  button(loadedTree, "Continue To Scheduling").props.onClick();
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
      const total = sizes.reduce((sum, size) => sum + size, 0);
      assert.deepEqual(sizes, examRooms.examRoomSizes(total, afterToggle.settings.studentsPerRoom));
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
  button(tree, "Proceed To Resources").props.onClick();
  tree = reloaded.render();
  assert.equal(decisionButton("Distribute across remaining rooms").props["aria-pressed"], true);
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
  assert.deepEqual(examRooms().map((room) => room.students.length), [19, 18, 15]);
  assert.equal(button(tree, "Proceed To Export").props.disabled, true, "The restored room must be assigned before export");
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
  assert.deepEqual(JSON.parse(JSON.stringify(printedOverviews.map((table) => table.props.rows[0].slice(-3)))), [[2, "Yes", "Department"], [1, "No", "Department"]]);
  assert.ok(printedOverviews.every((table) => table.props.columns.includes("Primary invigilators needed") && table.props.columns.includes("During lab time")));
  const labList = find(preview(), (node) => node.type?.name === "PreviewTable" && node.props.label === "Exam overview preview");
  assert.deepEqual(JSON.parse(JSON.stringify(labList.props.rows[0].slice(-4))), ["P-B-4F/13", 2, "Yes", "Department"]);
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
  assert.equal(asdList.props.rows[1][0], "ASD-1000", "Department exam stays first at the same time");
  assert.equal(asdList.props.rows[1].at(-1), "ASD (reference)");
  assert.equal(asdList.props.rowClassNames[1], "export-row--asd");
  assert.deepEqual(JSON.parse(JSON.stringify(asdList.props.rows[1].slice(4, 8))), ["N/A", "Not managed here", "N/A", "N/A"]);
  const asdPrintedTables = nodes(printView()).filter((node) => node.type?.name === "PreviewTable" && node.props.label === "Exam overview");
  assert.equal(asdPrintedTables.length, 3);
  assert.equal(asdPrintedTables[0].props.rowClassNames[1], "export-row--asd");
  assert.equal(asdPrintedTables[2].props.rows[0][0], "ASD-2000");
  assert.match(text(asdPrintedTables[2]), /10:30-12:00/);
  await button(tree, "Download CSV").props.onClick();
  const asdCsvZip = await JSZip.loadAsync(await app.downloads.at(-1).arrayBuffer());
  assert.deepEqual(Object.keys(asdCsvZip.files).sort(), ["Week_1_Exam_Overview.csv", "Week_2_Exam_Overview.csv", "Week_3_Exam_Overview.csv"]);
  assert.match(await asdCsvZip.file("Week_1_Exam_Overview.csv").async("string"), /"ASD-1000".*"ASD \(reference\)"/);
  assert.match(await asdCsvZip.file("Week_3_Exam_Overview.csv").async("string"), /"10:30","12:00","ASD-2000"/);
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
  assert.equal(asdExcelRow.Schedule, "ASD (reference)");
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
  reportButton("Exam overview").props.onClick();
  tree = app.render();
  assert.equal(asdToggle().props.checked, true, "The overview preference is retained without affecting another report");
  assert.equal(weekInput(3).props.checked, true);
  assert.match(text(preview()), /ASD-1000/);
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
  assert.equal(table.props.rows[0][0], "EXAM-12");
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
