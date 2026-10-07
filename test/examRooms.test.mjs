import test from "node:test";
import assert from "node:assert/strict";
import { examInvigilatorsNeeded, examRoomLayout, examRoomSizes, readRoomDistributionChoices, roomCapacity, roomConsolidationOptions } from "../src/examRooms.js";

test("room sizing keeps the minimum rooms and a one-invigilator last room when feasible", () => {
  for (const [students, sizes, invigilators] of [
    [0, [], 0], [15, [15], 1], [16, [16], 2], [25, [25], 2], [26, [13, 13], 2],
    [30, [15, 15], 2], [31, [16, 15], 3], [32, [17, 15], 3], [40, [25, 15], 3],
    [41, [21, 20], 4], [50, [25, 25], 4],
    [51, [18, 18, 15], 5], [55, [20, 20, 15], 5], [60, [23, 22, 15], 5],
    [65, [25, 25, 15], 5], [66, [22, 22, 22], 6],
    [76, [21, 20, 20, 15], 7], [90, [25, 25, 25, 15], 7], [91, [23, 23, 23, 22], 8],
  ]) {
    assert.deepEqual(examRoomSizes(students), sizes, `${students} students`);
    assert.equal(examInvigilatorsNeeded(students), invigilators, `${students} students`);
  }
  assert.deepEqual(examRoomSizes(55, 15), [14, 14, 14, 13]);
  assert.equal(examInvigilatorsNeeded(55, 15), 4);
  assert.deepEqual(examRoomSizes(55, 500), [20, 20, 15]);
  assert.deepEqual(examRoomSizes(55, 20), [20, 20, 15]);
  assert.deepEqual(examRoomSizes(56, 20), [19, 19, 18]);
  assert.equal(roomCapacity(15.8), 15);
  assert.deepEqual(examRoomSizes(-5), []);
  assert.deepEqual(examRoomSizes(NaN), []);
});

test("room sizing preserves every student, capacity and balance except when the last room saves an invigilator", () => {
  for (let capacity = 1; capacity <= 25; capacity += 1) {
    for (let students = 1; students <= 300; students += 1) {
      const sizes = examRoomSizes(students, capacity);
      assert.equal(sizes.length, Math.ceil(students / capacity));
      assert.equal(sizes.reduce((sum, size) => sum + size, 0), students);
      assert.ok(sizes.every((size) => Number.isInteger(size) && size > 0 && size <= capacity));
      const fullyBalanced = Array.from({ length: sizes.length }, (_, i) =>
        Math.floor(students / sizes.length) + (i < students % sizes.length ? 1 : 0));
      const canSave = sizes.length > 1 && fullyBalanced.at(-1) > 15 &&
        students <= (sizes.length - 1) * capacity + 15;
      if (canSave) {
        assert.equal(sizes.at(-1), 15);
        const others = sizes.slice(0, -1);
        assert.ok(Math.max(...others) - Math.min(...others) <= 1);
      } else assert.deepEqual(sizes, fullyBalanced);
      const balancedDuties = fullyBalanced.reduce((sum, size) => sum + (size > 15 ? 2 : 1), 0);
      assert.equal(examInvigilatorsNeeded(students, capacity), balancedDuties - Number(canSave));
    }
  }
});

test("a 27-student exception is offered only when it can remove an overflow room and must be explicitly chosen", () => {
  assert.deepEqual(roomConsolidationOptions(52).map(({ value, sizes }) => [value, sizes]), [["merge", [27, 25]], ["distribute", [26, 26]]]);
  assert.deepEqual(examRoomSizes(52), [19, 18, 15]);
  assert.deepEqual(examRoomLayout(52, 25, "merge"), { sizes: [27, 25], maxStudents: 27, choice: "merge" });
  assert.deepEqual(examRoomSizes(52, 25, "distribute"), [26, 26]);
  assert.equal(examInvigilatorsNeeded(52, 25, "distribute"), 4);
  assert.deepEqual(examRoomSizes(27, 25, "merge"), [27]);
  for (const students of [25, 28, 50, 55, 75, 82]) assert.deepEqual(roomConsolidationOptions(students), []);
  assert.deepEqual(roomConsolidationOptions(52, 20), []);
  assert.deepEqual(examRoomSizes(52, 20, "merge"), [19, 18, 15]);
  assert.deepEqual(examRoomLayout(28, 25, "merge"), { sizes: [14, 14], maxStudents: 25, choice: "standard" });
  assert.deepEqual(examRoomSizes(52, 25, "unapproved"), [19, 18, 15]);
  assert.deepEqual(readRoomDistributionChoices({ A: "merge", B: "standard" }), { A: "merge", B: "standard" });
  for (const invalid of [null, [], { A: true }, { A: "unapproved" }]) assert.throws(() => readRoomDistributionChoices(invalid), /Invalid room distribution/);
});

test("every offered consolidation keeps all students in exactly one fewer room and never exceeds 27", () => {
  for (let students = 1; students <= 500; students += 1) {
    for (const option of roomConsolidationOptions(students)) {
      assert.equal(option.sizes.length, examRoomSizes(students).length - 1);
      assert.equal(option.sizes.reduce((sum, size) => sum + size, 0), students);
      assert.ok(option.sizes.every((size) => size > 0 && size <= 27));
      if (option.value === "distribute") {
        const canSave = Math.floor(students / option.sizes.length) > 15 &&
          students <= (option.sizes.length - 1) * 27 + 15;
        if (canSave) {
          assert.equal(option.sizes.at(-1), 15);
          const others = option.sizes.slice(0, -1);
          assert.ok(Math.max(...others) - Math.min(...others) <= 1);
        } else assert.ok(Math.max(...option.sizes) - Math.min(...option.sizes) <= 1);
      }
    }
  }
});

test("approved distributions may save the last room's invigilator but never change an explicit merge", () => {
  assert.deepEqual(examRoomLayout(176, 25, "distribute"), {
    sizes: [27, 27, 27, 27, 27, 26, 15], maxStudents: 27, choice: "distribute",
  });
  assert.equal(examInvigilatorsNeeded(176, 25, "distribute"), 13);
  assert.deepEqual(examRoomSizes(176, 25, "merge"), [26, 25, 25, 25, 25, 25, 25]);
  assert.equal(examInvigilatorsNeeded(176, 25, "merge"), 14);
  assert.deepEqual(examRoomSizes(178, 25, "distribute"), [26, 26, 26, 25, 25, 25, 25]);
  assert.equal(examInvigilatorsNeeded(178, 25, "distribute"), 14, "Overfilling another room to save an invigilator is forbidden");
});
