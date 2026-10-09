import { assignmentIds, labExamEndMinutes } from "./department.js";
import { examInvigilatorsNeeded, roomCapacity } from "./examRooms.js";
import { FRIDAY_EXAM_NOTICE, FRIDAY_EXAM_WINDOWS, isFridayExamTimeAllowed } from "./examWindows.js";
import { assignResources, buildExamSessions, PREFERRED_STANDBY_COUNT, resourceSessionLabel, standbyAvailability, validateResourcePlan } from "./resources.js";

const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const MORNING_LAB_START = 540;
const MORNING_LAB_END = 600;
const EVENING_START = 1020;
const PREFERRED_STUDENT_BREAK = 60;
const COMMON_WINDOWS = [
  { days: DAYS.slice(0, -1), startMinutes: 720, endMinutes: 780 },
  { days: DAYS.slice(0, -1), startMinutes: EVENING_START, endMinutes: 1080 },
  ...FRIDAY_EXAM_WINDOWS.map((window) => ({ days: ["Friday"], startMinutes: window.start, endMinutes: window.end })),
];
const minutes = (id) => Number(id.split(":")[0]) * 60 + Number(id.split(":")[1]);
const overlaps = (a, b) => a.start < b.end && b.start < a.end;
const shareStudents = (a, b) => [...a.students].some((id) => b.students.has(id));
const usesLabWindow = (course) => course.crns.length === 1 && course.labSessions?.length > 0;

function studentIssue(exam, sameDay) {
  if (sameDay.some((entry) => overlaps(exam, entry) && shareStudents(exam, entry))) {
    return "student overlap with a main or ASD exam";
  }
  if ([...exam.students].some((id) => sameDay.filter((entry) => entry.students.has(id)).length >= 2)) {
    return "more than two exams per student in a day";
  }
  return "";
}

function breakPenalty(exam, entries) {
  return entries.reduce((total, entry) => {
    if (entry.week !== exam.week || entry.day !== exam.day) return total;
    const missingBreak = Math.max(0, PREFERRED_STUDENT_BREAK - Math.max(exam.start - entry.end, entry.start - exam.end));
    return total + missingBreak * [...exam.students].filter((id) => entry.students.has(id)).length;
  }, 0);
}

const withExam = (map, candidate, id) => ({ ...map, [candidate.week]: { ...map[candidate.week],
  [candidate.day]: { ...map[candidate.week]?.[candidate.day],
    [candidate.slotId]: [...(map[candidate.week]?.[candidate.day]?.[candidate.slotId] || []), id] } } });

function candidatesFor(course, weeks, timeSlots, duration, interval) {
  const lab = usesLabWindow(course);
  // Use the same :50 lab-end allowance as manual placement and resource validation.
  const windows = lab ? course.labSessions.map((window) => window.startMinutes < MORNING_LAB_START
    ? { ...window, startMinutes: MORNING_LAB_START, endMinutes: Math.min(labExamEndMinutes(window), MORNING_LAB_END) }
    : { ...window, endMinutes: labExamEndMinutes(window) }) : COMMON_WINDOWS;
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
  let next = Object.fromEntries(Object.entries(assignments).map(([week, days]) => [week,
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
  const pending = courses.filter((course) => !scheduled.has(course.id)).map((course) => {
    const students = new Set(course.students.map((student) => student.id));
    return { course, students, lab: usesLabWindow(course), reasons: new Set(),
      invigilators: examInvigilatorsNeeded(students.size, capacity, roomDistributionChoices[course.id]),
      candidates: timeSlots.length ? candidatesFor(course, weeks, timeSlots, duration, interval).sort(chronological) : [],
    };
  }).sort((a, b) => a.candidates.length - b.candidates.length || b.invigilators - a.invigilators || b.students.size - a.students.size || a.course.code.localeCompare(b.course.code));
  const placed = [];
  const unplaced = [];
  let resourcePlan;
  const place = (item, candidates) => {
    const { reasons } = item;
    let best;
    let fallback;
    for (const candidate of candidates) {
      if (!item.students.size) break;
      const exam = { ...candidate, students: item.students };
      const sameDay = entries.filter((entry) => entry.week === exam.week && entry.day === exam.day);
      const issue = studentIssue(exam, sameDay);
      if (issue) {
        reasons.add(issue);
        continue;
      }
      const proposed = withExam(next, candidate, item.course.id);
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
    if (!best) return;
    const { candidate } = best;
    next[candidate.week] ??= {};
    next[candidate.week][candidate.day] ??= {};
    next[candidate.week][candidate.day][candidate.slotId] ??= [];
    next[candidate.week][candidate.day][candidate.slotId].push(item.course.id);
    resourcePlan = best.plan;
    scheduled.add(item.course.id);
    entries.push({ ...candidate, id: item.course.id, students: item.students, locked: false });
    placed.push({ courseId: item.course.id, ...candidate });
  };
  // Reserve fixed lab windows, then give staffing-heavy flexible exams first choice of daytime capacity.
  pending.filter((item) => item.lab).forEach((item) => place(item, item.candidates));
  const flexible = pending.filter((item) => !item.lab);
  flexible.forEach((item) => place(item, item.candidates.filter((candidate) => candidate.start < EVENING_START)));
  // Finish the daytime pass for every course before filling evenings with the lightest staffing demand first.
  flexible.filter((item) => !scheduled.has(item.course.id))
    .sort((a, b) => a.invigilators - b.invigilators || a.students.size - b.students.size || a.course.code.localeCompare(b.course.code))
    .forEach((item) => place(item, item.candidates.filter((candidate) => candidate.start >= EVENING_START)));
  pending.forEach((item) => {
    if (!scheduled.has(item.course.id)) {
      const { reasons } = item;
      const morningLabNote = item.course.crns.length === 1 && item.course.labSessions?.some((lab) => lab.startMinutes < MORNING_LAB_START)
        ? " Morning lab exams must fit within 09:00-10:00 and the lab hours (a 09:50 lab end is treated as 10:00); the exam duration is not shortened."
        : "";
      const reason = !item.students.size ? "No enrolment for the listed CRNs." : !item.candidates.length
        ? "No matching window fits the exam duration and timetable hours." + morningLabNote +
          (item.course.labSessions?.some((lab) => lab.days.includes("Friday")) ? " " + FRIDAY_EXAM_NOTICE : "")
        : "No valid slot. " + [...reasons].slice(0, 3).join("; ") +
          (reasons.size > 3 ? " Additional slots are blocked by the same resource/student constraints. Review the pool, existing timetable or add a week." : "");
      unplaced.push({ courseId: item.course.id, reason });
    }
  });
  let sessions = buildExamSessions(next, courseLookup, duration, capacity, roomDistributionChoices);
  resourcePlan ??= assignResources(sessions, catalog);
  // Once noon exams are known, improve student breaks by moving only this run's placements, labs first.
  const spacingOrder = [...pending.filter((item) => item.lab), ...flexible];
  let improved;
  do {
    improved = false;
    for (const item of spacingOrder) {
      const placement = placed.find((exam) => exam.courseId === item.course.id);
      if (!placement) continue;
      const entry = entries.find((exam) => !exam.locked && exam.id === item.course.id);
      const others = entries.filter((exam) => exam !== entry);
      const currentPenalty = breakPenalty(entry, others);
      if (!currentPenalty) continue;
      const candidates = item.candidates
        .filter((candidate) => item.lab || (candidate.start < EVENING_START) === (entry.start < EVENING_START))
        .map((candidate) => ({ ...candidate, students: item.students }))
        .filter((candidate) => !studentIssue(candidate, others.filter((exam) => exam.week === candidate.week && exam.day === candidate.day)))
        .map((candidate) => ({ candidate, penalty: breakPenalty(candidate, others) }))
        .filter((choice) => choice.penalty < currentPenalty)
        .sort((a, b) => a.penalty - b.penalty || chronological(a.candidate, b.candidate));
      if (!candidates.length) continue;
      const standbyTargets = new Map(sessions.map((session) => [session.id,
        Math.min(PREFERRED_STANDBY_COUNT, standbyAvailability(session, sessions, catalog, resourcePlan).count)]));
      const previousSessionId = `${entry.week}/${entry.day}/${entry.slotId}`;
      for (const { candidate } of candidates) {
        const withoutExam = { ...next, [entry.week]: { ...next[entry.week], [entry.day]: { ...next[entry.week][entry.day],
          [entry.slotId]: next[entry.week][entry.day][entry.slotId].filter((id) => id !== item.course.id) } } };
        const proposed = withExam(withoutExam, candidate, item.course.id);
        const trialSessions = buildExamSessions(proposed, courseLookup, duration, capacity, roomDistributionChoices);
        const plan = assignResources(trialSessions, catalog);
        if (!validateResourcePlan(trialSessions, catalog, plan).complete) continue;
        if (trialSessions.some((session) => standbyAvailability(session, trialSessions, catalog, plan).count <
          (standbyTargets.get(session.id) ?? standbyTargets.get(previousSessionId)))) continue;
        next = proposed;
        resourcePlan = plan;
        sessions = trialSessions;
        const { students: _students, ...slot } = candidate;
        Object.assign(entry, slot);
        Object.assign(placement, slot);
        improved = true;
        break;
      }
    }
  } while (improved); // Each accepted move strictly reduces the total student break shortfall.
  const warnings = sessions.flatMap((session) => {
    const standby = standbyAvailability(session, sessions, catalog, resourcePlan);
    return standby.count < PREFERRED_STANDBY_COUNT ? [{ sessionId: session.id,
      message: resourceSessionLabel(session) + ": only " + standby.count + " available standby invigilator(s), including assigned backups; aim for " + PREFERRED_STANDBY_COUNT + ". Review the resource pool or move an exam." }] : [];
  });
  return { assignments: next, placed, unplaced, resourcePlan, warnings };
}
