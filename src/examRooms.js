export const roomCapacity = (value) => Math.min(25, Math.max(1, Math.floor(Number(value) || 25)));

const studentTotal = (value) => Number.isFinite(Number(value)) ? Math.max(0, Math.floor(Number(value))) : 0;

function balancedSizes(count, rooms) {
  if (!rooms) return [];
  const size = Math.floor(count / rooms);
  const extra = count % rooms;
  return Array.from({ length: rooms }, (_, index) => size + (index < extra ? 1 : 0));
}

function staffingAwareSizes(count, rooms, capacity) {
  const sizes = balancedSizes(count, rooms);
  // Keep the last room at the one-invigilator threshold only when it saves a duty.
  if (rooms > 1 && sizes[rooms - 1] > 15 && count - 15 <= (rooms - 1) * capacity) {
    return [...balancedSizes(count - 15, rooms - 1), 15];
  }
  return sizes;
}

export function roomConsolidationOptions(studentCount, capacity = 25) {
  const count = studentTotal(studentCount);
  const rooms = Math.ceil(count / 25) - 1;
  const overflow = count % 25;
  if (roomCapacity(capacity) !== 25 || !overflow || rooms < 1 || count > rooms * 27) return [];
  const options = [];
  if (overflow <= 2) {
    const sizes = Array(rooms).fill(25);
    sizes[0] += overflow;
    options.push({ value: "merge", label: "Merge overflow into one room", sizes });
  }
  const sizes = staffingAwareSizes(count, rooms, 27);
  if (!options.some((option) => option.sizes.every((size, index) => size === sizes[index]))) {
    options.push({ value: "distribute", label: "Distribute across remaining rooms", sizes });
  }
  return options;
}

export function examRoomLayout(studentCount, capacity = 25, choice = "standard") {
  const count = studentTotal(studentCount);
  const approved = roomConsolidationOptions(count, capacity).find((option) => option.value === choice);
  if (approved) return { sizes: approved.sizes, maxStudents: 27, choice: approved.value };
  const limit = roomCapacity(capacity);
  return { sizes: staffingAwareSizes(count, Math.ceil(count / limit), limit), maxStudents: limit, choice: "standard" };
}

export function examRoomSizes(studentCount, capacity = 25, choice = "standard") {
  return examRoomLayout(studentCount, capacity, choice).sizes;
}

export function examInvigilatorsNeeded(studentCount, capacity = 25, choice = "standard") {
  return examRoomSizes(studentCount, capacity, choice).reduce((sum, count) => sum + (count > 15 ? 2 : 1), 0);
}

export function readRoomDistributionChoices(value) {
  if (!value || typeof value !== "object" || Array.isArray(value) || Object.values(value).some((choice) => !["standard", "merge", "distribute"].includes(choice))) {
    throw new Error("Invalid room distribution choices in snapshot.");
  }
  return value;
}
