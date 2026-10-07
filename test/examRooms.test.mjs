import test from "node:test";
import assert from "node:assert/strict";
import { examInvigilatorsNeeded, examRoomLayout, examRoomSizes, readRoomDistributionChoices, roomCapacity, roomConsolidationOptions } from "../src/examRooms.js";

test("balanced room sizing uses the minimum rooms and the actual 15-student staffing threshold", () => {
  for (const [students, sizes, invigilators] of [
    [0, [], 0], [15, [15], 1], [16, [16], 2], [25, [25], 2], [26, [13, 13], 2],
    [30, [15, 15], 2], [31, [16, 15], 3], [32, [16, 16], 4], [50, [25, 25], 4],
    [51, [17, 17, 17], 6], [55, [19, 18, 18], 6], [60, [20, 20, 20], 6], [76, [19, 19, 19, 19], 8],
  ]) {
    assert.deepEqual(examRoomSizes(students), sizes, `${students} students`);
    assert.equal(examInvigilatorsNeeded(students), invigilators, `${students} students`);
  }
  assert.deepEqual(examRoomSizes(55, 15), [14, 14, 14, 13]);
  assert.equal(examInvigilatorsNeeded(55, 15), 4);
  assert.deepEqual(examRoomSizes(55, 500), [19, 18, 18]);
  assert.equal(roomCapacity(15.8), 15);
  assert.deepEqual(examRoomSizes(-5), []);
  assert.deepEqual(examRoomSizes(NaN), []);
});

test("balanced room sizing preserves every student and the capacity for all supported room limits", () => {
  for (let capacity = 1; capacity <= 25; capacity += 1) {
    for (let students = 1; students <= 300; students += 1) {
      const sizes = examRoomSizes(students, capacity);
      assert.equal(sizes.length, Math.ceil(students / capacity));
      assert.equal(sizes.reduce((sum, size) => sum + size, 0), students);
      assert.ok(sizes.every((size) => Number.isInteger(size) && size > 0 && size <= capacity));
      assert.ok(Math.max(...sizes) - Math.min(...sizes) <= 1);
    }
  }
});

test("a 27-student exception is offered only when it can remove an overflow room and must be explicitly chosen", () => {
  assert.deepEqual(roomConsolidationOptions(52).map(({ value, sizes }) => [value, sizes]), [["merge", [27, 25]], ["distribute", [26, 26]]]);
  assert.deepEqual(examRoomSizes(52), [18, 17, 17]);
  assert.deepEqual(examRoomLayout(52, 25, "merge"), { sizes: [27, 25], maxStudents: 27, choice: "merge" });
  assert.deepEqual(examRoomSizes(52, 25, "distribute"), [26, 26]);
  assert.equal(examInvigilatorsNeeded(52, 25, "distribute"), 4);
  assert.deepEqual(examRoomSizes(27, 25, "merge"), [27]);
  for (const students of [25, 28, 50, 55, 75, 82]) assert.deepEqual(roomConsolidationOptions(students), []);
  assert.deepEqual(roomConsolidationOptions(52, 20), []);
  assert.deepEqual(examRoomSizes(52, 20, "merge"), [18, 17, 17]);
  assert.deepEqual(examRoomLayout(28, 25, "merge"), { sizes: [14, 14], maxStudents: 25, choice: "standard" });
  assert.deepEqual(examRoomSizes(52, 25, "unapproved"), [18, 17, 17]);
  assert.deepEqual(readRoomDistributionChoices({ A: "merge", B: "standard" }), { A: "merge", B: "standard" });
  for (const invalid of [null, [], { A: true }, { A: "unapproved" }]) assert.throws(() => readRoomDistributionChoices(invalid), /Invalid room distribution/);
});

test("every offered consolidation keeps all students in exactly one fewer room and never exceeds 27", () => {
  for (let students = 1; students <= 500; students += 1) {
    for (const option of roomConsolidationOptions(students)) {
      assert.equal(option.sizes.length, examRoomSizes(students).length - 1);
      assert.equal(option.sizes.reduce((sum, size) => sum + size, 0), students);
      assert.ok(option.sizes.every((size) => size > 0 && size <= 27));
      if (option.value === "distribute") assert.ok(Math.max(...option.sizes) - Math.min(...option.sizes) <= 1);
    }
  }
});
