import test from "node:test";
import assert from "node:assert/strict";
import { isFridayExamTimeAllowed } from "../src/examWindows.js";

test("Friday allows only its two session starts, with exams fully inside each window", () => {
  for (const [start, end] of [[540, 600], [630, 690], [540, 570]]) {
    assert.equal(isFridayExamTimeAllowed("Friday", start, end), true);
  }
  for (const [start, end] of [[480, 540], [570, 630], [600, 660], [660, 720], [720, 780],
    [1020, 1080], [540, 660], [630, 691], [540, 540], [540, 539]]) {
    assert.equal(isFridayExamTimeAllowed("Friday", start, end), false);
  }
  assert.equal(isFridayExamTimeAllowed("Monday", 720, 780), true);
  assert.equal(isFridayExamTimeAllowed("Thursday", 900, 1020), true);
});
