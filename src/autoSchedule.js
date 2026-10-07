import { assignmentIds, labExamEndMinutes } from "./department.js";
import { roomCapacity } from "./examRooms.js";
import { FRIDAY_EXAM_NOTICE, FRIDAY_EXAM_WINDOWS, isFridayExamTimeAllowed } from "./examWindows.js";
import { assignResources, buildExamSessions, PREFERRED_STANDBY_COUNT, resourceSessionLabel, standbyAvailability, validateResourcePlan } from "./resources.js";

const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const MORNING_LAB_START = 540;
const MORNING_LAB_END = 600;
const COMMON_WINDOWS = [
  { days: DAYS.slice(0, -1), startMinutes: 720, endMinutes: 780 },
  { days: DAYS.slice(0, -1), startMinutes: 1020, endMinutes: 1080 },
  ...FRIDAY_EXAM_WINDOWS.map((window) => ({ days: ["Friday"], startMinutes: window.start, endMinutes: window.end })),
];
const minutes = (id) => Number(id.split(":")[0]) * 60 + Number(id.split(":")[1]);
const overlaps = (a, b) => a.start < b.end && b.start < a.end;
const shareStudents = (a, b) => [...a.students].some((id) => b.students.has(id));

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
        isFridayExamTimeAllowed(day, start, start + duration) &&
        start + duration <= minutes(timeSlots[timeSlots.length - 1].id) + interval) {
        candidates.push({ week, day, slotId: slot.id, start, end: start + duration });
      }
    });
  })));
  return [...new Map(candidates.map((item) => [`${item.week}/${item.day}/${item.slotId}`, item])).values()];
}

// Add only unplaced exams. ASD participates in student rules, never main staffing.
export function autoSchedule({ courses, courseLookup, assignments, asdAssignments = {}, asdExamDurations = {}, weeks, timeSlots, settings, catalog, roomDistributionChoices = {} }) {
  const interval = settings.slotIntervalMinutes;
  const duration = Math.ceil(settings.examDurationMinutes / interval) * interval;
  const capacity = roomCapacity(settings.studentsPerRoom);
  const next = Object.fromEntries(Object.entries(assignments).map(([week, days]) => [week,
    Object.fromEntries(Object.entries(days).map(([day, slots]) => [day,
      Object.fromEntries(Object.entries(slots).map(([slot, ids]) => [slot, [...ids]])),
    ])),
  ]));
  const entries = [];
  const addEntries = (map, locked) => Object.entries(map).forEach(([week, days]) => {
    Object.entries(days).forEach(([day, slots]) => Object.entries(slots).forEach(([slot, ids]) => ids.forEach((id) => {
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
  const chronological = (a, b) => a.week - b.week || DAYS.indexOf(a.day) - DAYS.indexOf(b.day) || a.start - b.start;
  const pending = courses.filter((course) => !scheduled.has(course.id)).map((course) => ({
    course, candidates: timeSlots.length ? candidatesFor(course, weeks, timeSlots, duration, interval).sort(chronological) : [],
    students: new Set(course.students.map((student) => student.id)),
  })).sort((a, b) => a.candidates.length - b.candidates.length || b.students.size - a.students.size || a.course.code.localeCompare(b.course.code));
  const placed = [];
  const unplaced = [];
  let resourcePlan;
  pending.forEach((item) => {
    const reasons = new Set();
    let best;
    let fallback;
    for (const candidate of item.candidates) {
      if (!item.students.size) break;
      const exam = { ...candidate, students: item.students };
      const sameDay = entries.filter((entry) => entry.week === exam.week && entry.day === exam.day);
      if (sameDay.some((entry) => overlaps(exam, entry) && shareStudents(exam, entry))) {
        reasons.add("student overlap with a main or ASD exam");
        continue;
      }
      if ([...exam.students].some((id) => sameDay.filter((entry) => entry.students.has(id)).length >= 2)) {
        reasons.add("more than two exams per student in a day");
        continue;
      }
      const proposed = { ...next, [candidate.week]: { ...next[candidate.week],
        [candidate.day]: { ...next[candidate.week]?.[candidate.day],
          [candidate.slotId]: [...(next[candidate.week]?.[candidate.day]?.[candidate.slotId] || []), item.course.id] } } };
      const sessions = buildExamSessions(proposed, courseLookup, duration, capacity, roomDistributionChoices);
      const plan = assignResources(sessions, catalog);
      const validation = validateResourcePlan(sessions, catalog, plan);
      if (!validation.complete) {
        validation.issues.forEach((issue) => reasons.add(issue.context + (issue.exam ? " / " + issue.exam : "") + ": " + issue.title + ". " + issue.message));
        continue;
      }
      // Trial the actual allocation, not just pool totals, including fixed labs and overlapping starts.
      const affected = sessions.filter((session) => session.week === exam.week && session.day === exam.day && overlaps(session, exam));
      const choice = { candidate, plan };
      if (affected.every((session) => standbyAvailability(session, sessions, catalog, plan).count >= PREFERRED_STANDBY_COUNT)) {
        best = choice;
        break;
      }
      fallback ??= choice;
    }
    best ??= fallback;
    if (!item.students.size || !best) {
      const morningLabNote = item.course.crns.length === 1 && item.course.labSessions?.some((lab) => lab.startMinutes < MORNING_LAB_START)
        ? " Morning lab exams must fit within 09:00-10:00 and the lab hours (a 09:50 lab end is treated as 10:00); the exam duration is not shortened."
        : "";
      const reason = !item.students.size ? "No enrolment for the listed CRNs." : !item.candidates.length
        ? "No matching window fits the exam duration and timetable hours." + morningLabNote +
          (item.course.labSessions?.some((lab) => lab.days.includes("Friday")) ? " " + FRIDAY_EXAM_NOTICE : "")
        : "No valid slot. " + [...reasons].slice(0, 3).join("; ") +
          (reasons.size > 3 ? " Additional slots are blocked by the same resource/student constraints. Review the pool, existing timetable or add a week." : "");
      unplaced.push({ courseId: item.course.id, reason });
      return;
    }
    const { candidate } = best;
    next[candidate.week] ??= {};
    next[candidate.week][candidate.day] ??= {};
    next[candidate.week][candidate.day][candidate.slotId] ??= [];
    next[candidate.week][candidate.day][candidate.slotId].push(item.course.id);
    resourcePlan = best.plan;
    entries.push({ ...candidate, id: item.course.id, students: item.students, locked: false });
    placed.push({ courseId: item.course.id, ...candidate });
  });
  const sessions = buildExamSessions(next, courseLookup, duration, capacity, roomDistributionChoices);
  resourcePlan ??= assignResources(sessions, catalog);
  const warnings = sessions.flatMap((session) => {
    const standby = standbyAvailability(session, sessions, catalog, resourcePlan);
    return standby.count < PREFERRED_STANDBY_COUNT ? [{ sessionId: session.id,
      message: resourceSessionLabel(session) + ": only " + standby.count + " available standby invigilator(s), including assigned backups; aim for " + PREFERRED_STANDBY_COUNT + ". Review the resource pool or move an exam." }] : [];
  });
  return { assignments: next, placed, unplaced, resourcePlan, warnings };
}
