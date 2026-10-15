import test from "node:test";
import assert from "node:assert/strict";
import { chooseExportDestination } from "../src/fileExports.js";

function fixture({ exists = false, confirmed = true, failure } = {}) {
  const calls = [];
  let output;
  const writable = { async write(blob) { calls.push("write"); if (failure) throw failure; output = blob; },
    async close() { calls.push("close"); }, async abort() { calls.push("abort"); } };
  const browser = { confirm(message) { calls.push(["confirm", message]); return confirmed; },
    async showDirectoryPicker(options) {
      calls.push(["picker", options]);
      return { name: "Exam exports", async getFileHandle(filename, options) {
        calls.push(["file", filename, options]);
        if (!exists && !options?.create) throw Object.assign(new Error("Not found"), { name: "NotFoundError" });
        return { async createWritable() { calls.push("stream"); return writable; } };
      } };
    } };
  return { browser, calls, output: () => output, writable };
}

test("every export opens a destination picker immediately and saves the complete blob with its generated filename", async () => {
  const f = fixture();
  const pending = chooseExportDestination({ browser: f.browser });
  assert.deepEqual(f.calls, [["picker", { id: "exam-exports", mode: "readwrite" }]], "The picker must open before the first asynchronous report task");
  const destination = await pending;
  const blob = new Blob(["Workbook content"], { type: "application/zip" });
  assert.equal(await destination.save(blob, "Exam reports.zip"), true);
  assert.equal(f.output(), blob);
  assert.deepEqual(f.calls.slice(1), [["file", "Exam reports.zip", undefined], ["file", "Exam reports.zip", { create: true }], "stream", "write", "close"]);
  await chooseExportDestination({ browser: f.browser });
  assert.equal(f.calls.filter((call) => call[0] === "picker").length, 2, "Do not cache and silently reuse the destination");
});

test("cancelling the picker saves nothing and does not fall back to an automatic browser download", async () => {
  let downloaded = false;
  const browser = { showDirectoryPicker: async () => { throw Object.assign(new Error("Cancelled"), { name: "AbortError" }); } };
  assert.equal(await chooseExportDestination({ browser, download: () => { downloaded = true; } }), null);
  assert.equal(downloaded, false);
});

test("picker failures are reported, not treated as cancellation or silently redirected to Downloads", async () => {
  const failure = Object.assign(new Error("Folder permission denied"), { name: "SecurityError" });
  const browser = { showDirectoryPicker: async () => { throw failure; } };
  await assert.rejects(chooseExportDestination({ browser, download: () => assert.fail("Must not download") }), (error) => error === failure);
});

test("existing files require replacement confirmation, and declining leaves the original untouched", async () => {
  const f = fixture({ exists: true, confirmed: false });
  const destination = await chooseExportDestination({ browser: f.browser });
  assert.equal(await destination.save(new Blob(["new"]), "Exam.xlsx"), false);
  assert.match(f.calls.find((call) => call[0] === "confirm")[1], /"Exam.xlsx".*"Exam exports".*Replace/);
  assert.ok(!f.calls.includes("stream"));
  assert.equal(f.output(), undefined);
  const approved = fixture({ exists: true });
  assert.equal(await (await chooseExportDestination({ browser: approved.browser })).save(new Blob(["new"]), "Exam.xlsx"), true);
  assert.ok(approved.calls.includes("close"));
});

test("failed writes abort the stream without committing or hiding the original error", async () => {
  const failure = new Error("Disk full");
  const f = fixture({ exists: true, failure });
  const destination = await chooseExportDestination({ browser: f.browser });
  await assert.rejects(destination.save(new Blob(["new"]), "Exam.xlsx"), (error) => error === failure);
  assert.ok(f.calls.includes("abort"));
  assert.ok(!f.calls.includes("close"));
});

test("close failures abort and handle lookup or permission errors never trigger replacement or automatic downloads", async () => {
  const f = fixture();
  const failure = new Error("Cannot commit");
  f.writable.close = async () => { throw failure; };
  f.writable.abort = async () => { throw new Error("Already closed"); };
  const destination = await chooseExportDestination({ browser: f.browser });
  await assert.rejects(destination.save(new Blob(["new"]), "Exam.xlsx"), (error) => error === failure);
  const browser = { showDirectoryPicker: async () => ({ getFileHandle: async () => { throw Object.assign(new Error("Permission lost"), { name: "NotAllowedError" }); } }) };
  await assert.rejects((await chooseExportDestination({ browser })).save(new Blob(), "Exam.xlsx"), /Permission lost/);
});

test("unsupported browsers explain the download setting and require explicit confirmation for every fallback download", async () => {
  const confirmations = [];
  const downloads = [];
  const browser = { confirm: (message) => { confirmations.push(message); return confirmations.length === 1; } };
  const download = (blob, filename) => downloads.push({ blob, filename });
  const destination = await chooseExportDestination({ browser, download });
  const blob = new Blob(["JSON"]);
  await destination.save(blob, "Timetable.json");
  assert.deepEqual(downloads, [{ blob, filename: "Timetable.json" }]);
  assert.match(confirmations[0], /Ask where to save each file before downloading/);
  assert.equal(await chooseExportDestination({ browser, download }), null);
  assert.equal(downloads.length, 1);
});
