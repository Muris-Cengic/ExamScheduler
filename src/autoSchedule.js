import { assignmentIds, labExamEndMinutes } from "./department.js";
import { examInvigilatorsNeeded, examRoomSizes, roomCapacity } from "./examRooms.js";

const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const MORNING_LAB_START = 540;
const MORNING_LAB_END = 600;
const COMMON_WINDOWS = [
  { days: DAYS, startMinutes: 720, endMinutes: 780 },
  { days: DAYS, startMinutes: 1020, endMinutes: 1080 },
  { days: ["Friday"], startMinutes: 540, endMinutes: 600 },
  { days: ["Friday"], startMinutes: 630, endMinutes: 690 },
];
const minutes = (id) => Number(id.split(":")[0]) * 60 + Number(id.split(":")[1]);
const overlaps = (a, b) => a.start < b.end && b.start < a.end;
const shareStudents = (a, b) => [...a.students].some((id) => b.students.has(id));

function staffing(studentCount, capacity, choice) {
  return { invigilators: examInvigilatorsNeeded(studentCount, capacity, choice), rooms: examRoomSizes(studentCount, capacity, choice).length };
}

function candidatesFor(course, weeks, timeSlots, duration, interval) {
  const lab = course.crns.length === 1 && course.labSessions?.length > 0;
  // Early labs use the second morning hour, including the approved 09:50-to-10:00 allowance.
  const windows = lab ? course.labSessions.map((window) => window.startMinutes < MORNING_LAB_START
    ? { ...window, startMinutes: MORNING_LAB_START, endMinutes: Math.min(labExamEndMinutes(window), MORNING_LAB_END) }
    : window) : COMMON_WINDOWS;
  const candidates = [];
  weeks.forEach((week) => windows.forEach((window) => window.days.forEach((day) => {
    timeSlots.forEach((slot) => {
      const start = minutes(slot.id);
      if (start >= window.startMinutes && start + duration <= window.endMinutes &&
        start + duration <= minutes(timeSlots[timeSlots.length - 1].id) + interval) {
        candidates.push({ week, day, slotId: slot.id, start, end: start + duration });
      }
    });
  })));
  return [...new Map(candidates.map((item) => [`${item.week}/${item.day}/${item.slotId}`, item])).values()];
}

// Add only unplaced exams. ASD participates in student rules, never main staffing.
export function autoSchedule({ courses, courseLookup, assignments, asdAssignments = {}, asdExamDurations = {}, weeks, timeSlots, settings, roomDistributionChoices = {} }) {
  const interval = settings.slotIntervalMinutes;
  const duration = Math.ceil(settings.examDurationMinutes / interval) * interval;
  const capacity = roomCapacity(settings.studentsPerRoom);
  const invigilators = settings.invigilatorCount;
  const next = Object.fromEntries(Object.entries(assignments).map(([week, dayMap]) => [week,
    Object.fromEntries(Object.entries(dayMap).map(([day, slots]) => [day,
      Object.fromEntries(Object.entries(slots).map(([slot, ids]) => [slot, [...ids]])),
    ])),
  ]));
  const entries = [];
  const addEntries = (map, locked) => Object.entries(map).forEach(([week, dayMap]) => {
    Object.entries(dayMap).forEach(([day, slots]) => Object.entries(slots).forEach(([slot, ids]) => ids.forEach((id) => {
      const course = courseLookup[id];
      if (!course) return;
      const examDuration = locked ? Math.ceil((asdExamDurations[id] || duration) / interval) * interval : duration;
      entries.push({ id, week: Number(week), day, start: minutes(slot), end: minutes(slot) + examDuration,
        students: new Set(course.students.map((student) => student.id)), locked });
    })));
  });
  addEntries(next, false);
  addEntries(asdAssignments, true);
  const scheduled = new Set([...assignmentIds(next), ...assignmentIds(asdAssignments)]);
  const pending = courses.filter((course) => !scheduled.has(course.id)).map((course) => ({
    course, candidates: timeSlots.length ? candidatesFor(course, weeks, timeSlots, duration, interval) : [],
    students: new Set(course.students.map((student) => student.id)),
  })).sort((a, b) => a.candidates.length - b.candidates.length || b.students.size - a.students.size || a.course.code.localeCompare(b.course.code));
  const placed = [];
  const unplaced = [];
  pending.forEach((item) => {
    const reasons = new Set();
    let best;
    let bestScore = Infinity;
    item.candidates.forEach((candidate) => {
      const exam = { ...candidate, students: item.students };
      const sameDay = entries.filter((entry) => entry.week === exam.week && entry.day === exam.day);
      const concurrent = sameDay.filter((entry) => overlaps(exam, entry));
      if (concurrent.some((entry) => shareStudents(exam, entry))) {
        reasons.add("student overlap with a main or ASD exam");
        return;
      }
      if ([...exam.students].some((id) => sameDay.filter((entry) => entry.students.has(id)).length >= 2)) {
        reasons.add("more than two exams per student in a day");
        return;
      }
      const main = concurrent.filter((entry) => !entry.locked);
      // Check each occupied slot, not the union of consecutive, non-overlapping exams.
      for (let time = exam.start; time < exam.end; time += interval) {
        const active = main.filter((entry) => entry.start <= time && entry.end > time);
        const demands = [item.course, ...active.map((entry) => courseLookup[entry.id])].map((course) => staffing(new Set(course.students.map((student) => student.id)).size, capacity, roomDistributionChoices[course.id]));
        const roomCount = demands.reduce((sum, demand) => sum + demand.rooms, 0);
        const backups = Math.max(1, Math.floor(roomCount * 0.4));
        if (demands.reduce((sum, demand) => sum + demand.invigilators, 0) + backups > invigilators) {
          reasons.add("insufficient invigilators, including slot backups");
          return;
        }
      }
      const dayLoad = sameDay.filter((entry) => !entry.locked).reduce((sum, entry) => sum + entry.students.size, 0);
      const weekLoad = entries.filter((entry) => !entry.locked && entry.week === exam.week).length;
      const score = dayLoad * 10 + main.reduce((sum, entry) => sum + entry.students.size, 0) * 10 + weekLoad * 2;
      // Prefer an earlier eligible time even when that slot already contains feasible exams.
      if (!best || exam.start < best.start || (exam.start === best.start && score < bestScore)) {
        best = candidate;
        bestScore = score;
      }
    });
    if (!item.students.size || !best) {
      const morningLabNote = item.course.crns.length === 1 && item.course.labSessions?.some((lab) => lab.startMinutes < MORNING_LAB_START)
        ? " Morning lab exams must fit within 09:00-10:00 and the lab hours (a 09:50 lab end is treated as 10:00); the exam duration is not shortened."
        : "";
      const reason = !item.students.size ? "No enrolment for the listed CRNs." : !item.candidates.length
        ? `No matching window fits the exam duration and timetable hours.${morningLabNote}`
        : `No valid slot: ${[...reasons].join("; ")}.`;
      unplaced.push({ courseId: item.course.id, reason });
      return;
    }
    next[best.week] ??= {};
    next[best.week][best.day] ??= {};
    next[best.week][best.day][best.slotId] ??= [];
    next[best.week][best.day][best.slotId].push(item.course.id);
    entries.push({ ...best, id: item.course.id, students: item.students, locked: false });
    placed.push({ courseId: item.course.id, ...best });
  });
  return { assignments: next, placed, unplaced };
}
