import { labExamEndMinutes, roomIdentity } from "./department.js";
import { examRoomLayout, roomConsolidationOptions } from "./examRooms.js";

const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const normalise = (value) => String(value ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
const overlaps = (a, b) => a.start < b.end && b.start < a.end;
const sameDay = (a, b) => a.week === b.week && a.day === b.day;
export const backupTarget = (roomCount) => Math.max(1, Math.floor(roomCount * 0.4));
export const PREFERRED_STANDBY_COUNT = 2;
export const clock = (value) => `${String(Math.floor(value / 60)).padStart(2, "0")}:${String(value % 60).padStart(2, "0")}`;
export const resourceSessionLabel = (session) => `Week ${session.week} / ${session.day} / ${clock(session.start)}-${clock(session.end)}`;
export const resourceAssignmentAnchor = (id) => `resource-assignment-${encodeURIComponent(id)}`;

export function parseResourceCatalog(selection) {
  const rooms = new Map();
  const invigilators = new Map();
  selection.courses.forEach((course) => course.meetings.forEach((meeting) => {
    if (meeting.room && !/^(TBA|TBD|ONLINE|OUTSIDE|N\/A)$/i.test(meeting.room)) {
      const id = roomIdentity(meeting.campus, meeting.building, meeting.room);
      rooms.set(id, { id, name: [meeting.campus, meeting.building, meeting.room].filter(Boolean).join(" / "), enabled: true, busy: [] });
    }
    meeting.instructors.forEach((person) => invigilators.set(person.id, { ...person, enabled: true, busy: [] }));
  }));
  const records = new Map();
  selection.courses.forEach((course) => course.meetings.forEach((meeting) => {
    if (!meeting.days.length) return;
    const unknownTime = !Number.isFinite(meeting.startMinutes);
    const entry = { code: course.code, crn: meeting.crn, isLab: meeting.isLab, days: meeting.days, start: unknownTime ? 0 : meeting.startMinutes, end: unknownTime ? 1440 : meeting.endMinutes, unknownTime };
    const key = JSON.stringify([entry, meeting.campus, meeting.building, meeting.room, [...meeting.teachingInstructorIds].sort()]);
    records.set(key, { entry, meeting });
  }));
  records.forEach(({ entry, meeting }) => {
    const room = rooms.get(roomIdentity(meeting.campus, meeting.building, meeting.room));
    if (room) room.busy.push(entry);
    meeting.teachingInstructorIds.forEach((id) => {
      const staff = invigilators.get(id);
      if (staff) staff.busy.push(entry);
    });
  });
  const sort = (values) => [...values].sort((a, b) => a.name.localeCompare(b.name));
  return { rooms: sort(rooms.values()), invigilators: sort(invigilators.values()) };
}

export function readResourceCatalog(value) {
  if (!value || !Array.isArray(value.rooms) || !Array.isArray(value.invigilators)) throw new Error("Snapshot is missing its resource catalog.");
  for (const [kind, resources] of Object.entries(value)) {
    if (!["rooms", "invigilators"].includes(kind)) throw new Error("Invalid resource catalog.");
    const ids = new Set();
    for (const resource of resources) {
      if (typeof resource.id !== "string" || !resource.id || ids.has(resource.id) || typeof resource.name !== "string" || !resource.name ||
        typeof resource.enabled !== "boolean" || !Array.isArray(resource.busy)) {
        throw new Error("Invalid resource in snapshot.");
      }
      ids.add(resource.id);
      for (const entry of resource.busy) {
        if (!Array.isArray(entry.days) || entry.days.some((day) => !DAYS.includes(day)) || typeof entry.code !== "string" || typeof entry.crn !== "string" ||
          typeof entry.isLab !== "boolean" || !Number.isFinite(entry.start) || !Number.isFinite(entry.end) || entry.start < 0 || entry.end > 1440 || entry.end <= entry.start) {
          throw new Error("Invalid resource availability in snapshot.");
        }
      }
    }
  }
  return value;
}

export function readResourcePlan(value) {
  if (value === null) return null;
  const record = (item) => item && typeof item === "object" && !Array.isArray(item);
  if (!record(value) || typeof value.fingerprint !== "string" || !record(value.allocations) || !record(value.backups) ||
    Object.values(value.allocations).some((item) => !record(item) || typeof item.roomId !== "string" || !Array.isArray(item.invigilatorIds) || item.invigilatorIds.some((id) => typeof id !== "string")) ||
    Object.values(value.backups).some((ids) => !Array.isArray(ids) || ids.some((id) => typeof id !== "string"))) {
    throw new Error("Invalid resource assignments in snapshot.");
  }
  return value;
}

export function buildExamSessions(assignments, courseLookup, duration, capacity = 25, roomDistributionChoices = {}) {
  const sessions = [];
  Object.entries(assignments).forEach(([week, days]) => Object.entries(days).forEach(([day, slots]) => Object.entries(slots).forEach(([slotId, ids]) => {
    if (!ids.length) return;
    const [hour, minute] = slotId.split(":").map(Number);
    const start = hour * 60 + minute;
    const session = { id: `${week}/${day}/${slotId}`, week: Number(week), day, slotId, start, end: start + duration, rooms: [], roomDecisions: [], replacedLabs: [], issues: [] };
    ids.forEach((courseId) => {
      const course = courseLookup[courseId];
      if (!course?.students.length) { session.issues.push(`Missing enrolment for ${courseId}.`); return; }
      const lab = course.crns.length === 1 ? course.labSessions?.find((meeting) => meeting.days.includes(day) && start >= meeting.startMinutes && session.end <= labExamEndMinutes(meeting)) : null;
      if (course.crns.length === 1 && course.labSessions?.length && !lab) session.issues.push(`${course.code}: move this single-CRN exam into its listed lab session.`);
      if (lab) session.replacedLabs.push({ code: course.code, crn: lab.crn, start: lab.startMinutes, end: lab.endMinutes });
      const students = [...new Map(course.students.map((student) => [student.id, student])).values()].sort((a, b) => a.name.localeCompare(b.name) || a.id.localeCompare(b.id));
      const layout = examRoomLayout(students.length, capacity, roomDistributionChoices[courseId]);
      const alternatives = roomConsolidationOptions(students.length, capacity);
      if (alternatives.length) session.roomDecisions.push({ courseId, code: course.code, title: course.title, students: students.length,
        overflow: students.length % 25, choice: layout.choice,
        options: [{ value: "standard", label: "Keep the extra room (normal limit)", sizes: examRoomLayout(students.length, capacity).sizes }, ...alternatives],
      });
      let offset = 0;
      layout.sizes.forEach((size) => {
        const roster = students.slice(offset, offset + size);
        session.rooms.push({
          id: `${session.id}/${courseId}/${offset}`, courseId, code: course.code, title: course.title,
          students: roster, requiredInvigilators: roster.length > 15 ? 2 : 1,
          maxStudents: layout.maxStudents, distributionChoice: layout.choice,
          instructors: [...new Set(roster.map((student) => course.crnDetails?.find((detail) => detail.crn === student.crn)?.instructor || course.primaryInstructor).filter(Boolean))],
          fixedRoomId: lab && offset === 0 && lab.room ? roomIdentity(lab.campus, lab.building, lab.room) : "",
          fixedInvigilatorId: lab && offset === 0 ? lab.labInstructorId : "",
          missingLabResource: Boolean(lab && offset === 0 && (!lab.room || !lab.labInstructorId)),
        });
        offset += size;
      });
    });
    sessions.push(session);
  })));
  return sessions.sort((a, b) => a.week - b.week || DAYS.indexOf(a.day) - DAYS.indexOf(b.day) || a.start - b.start);
}

export function resourceFingerprint(sessions, catalog) {
  return JSON.stringify({ sessions, catalog });
}

export function emptyResourcePlan(sessions, catalog) {
  return { fingerprint: resourceFingerprint(sessions, catalog), allocations: {}, backups: {} };
}

export function reconcileRoomDistribution(previousSessions, sessions, catalog, previousPlan, courseId) {
  const plan = emptyResourcePlan(sessions, catalog);
  if (!previousPlan || previousPlan.fingerprint !== resourceFingerprint(previousSessions, catalog)) return plan;
  const spareStaff = [];
  sessions.forEach((session) => {
    const previous = previousSessions.find((entry) => entry.id === session.id);
    const oldRooms = previous?.rooms.filter((room) => room.courseId === courseId) || [];
    const newRooms = session.rooms.filter((room) => room.courseId === courseId);
    const retainedStaff = new Set();
    session.rooms.forEach((room) => {
      const oldRoom = room.courseId === courseId ? oldRooms[newRooms.indexOf(room)] : room;
      const allocation = previousPlan.allocations[oldRoom?.id];
      const ids = Array.from({ length: room.requiredInvigilators }, (_, index) => allocation?.invigilatorIds[index] || "");
      plan.allocations[room.id] = { roomId: allocation?.roomId || "", invigilatorIds: ids };
      if (room.courseId === courseId) ids.filter(Boolean).forEach((id) => retainedStaff.add(id));
    });
    oldRooms.forEach((room) => (previousPlan.allocations[room.id]?.invigilatorIds || []).forEach((id) => {
      if (id && !retainedStaff.has(id)) spareStaff.push({ sessionId: session.id, id });
    }));
    plan.backups[session.id] = (previousPlan.backups[session.id] || []).slice(0, backupTarget(session.rooms.length));
  });
  // Reuse freed staff from this exam, without moving unrelated rooms or manual assignments.
  sessions.forEach((session) => session.rooms.filter((room) => room.courseId === courseId).forEach((room) => {
    const allocation = plan.allocations[room.id];
    allocation.invigilatorIds.forEach((id, index) => {
      if (id) return;
      const person = catalog.invigilators.find((candidate) => spareStaff.some((spare) => spare.sessionId === session.id && spare.id === candidate.id) &&
        !(index === 0 && room.fixedInvigilatorId && candidate.id !== room.fixedInvigilatorId) &&
        !resourceChoiceReason(candidate, "invigilator", session, sessions, plan, `${room.id}/invigilator/${index}`));
      if (person) allocation.invigilatorIds[index] = person.id;
    });
  }));
  return plan;
}

function isReplacedLab(entry, session, sessions) {
  return entry.isLab && sessions.some((exam) => sameDay(exam, session) && exam.replacedLabs.some((lab) =>
    normalise(lab.code) === normalise(entry.code) && lab.crn === entry.crn && lab.start === entry.start && lab.end === entry.end));
}

export function isTeachingTimeDuty(person, session, sessions) {
  return Boolean(person?.busy.some((entry) => entry.days.includes(session.day) && !entry.unknownTime &&
    entry.start <= session.start && isReplacedLab(entry, session, sessions) &&
    labExamEndMinutes({ startMinutes: entry.start, endMinutes: entry.end }) >= session.end));
}

function availabilityConflict(resource, session, sessions) {
  if (!resource) return { type: "missing", message: "Not found in the resource pool." };
  if (!resource.enabled) return { type: "excluded", message: "Not in the active resource pool (excluded in Review Resource Pool)." };
  const busy = resource.busy.find((entry) => {
    if (!entry.days.includes(session.day) || !overlaps(session, entry)) return false;
    return !isReplacedLab(entry, session, sessions);
  });
  return busy ? { type: "class", message: `Class ${busy.code} (CRN ${busy.crn}) on ${session.day}, ${busy.unknownTime ? "time unspecified; the whole day is blocked" : `${clock(busy.start)}-${clock(busy.end)}`}.`, entry: busy } : null;
}

export function resourceBusyReason(resource, session, sessions) {
  return availabilityConflict(resource, session, sessions)?.message || "";
}

function planBookings(sessions, plan) {
  const bookings = [];
  sessions.forEach((session) => {
    session.rooms.forEach((room) => {
      const allocation = plan?.allocations?.[room.id];
      if (allocation?.roomId) bookings.push({ ...session, resourceId: allocation.roomId, owner: room.id, kind: "room", courseCode: room.code });
      (allocation?.invigilatorIds || []).forEach((id, index) => {
        if (id) bookings.push({ ...session, resourceId: id, owner: `${room.id}/invigilator/${index}`, kind: "invigilator", courseCode: room.code });
      });
    });
    (plan?.backups?.[session.id] || []).forEach((id, index) => {
      if (id) bookings.push({ ...session, resourceId: id, owner: `${session.id}/backup/${index}`, kind: "invigilator" });
    });
  });
  return bookings;
}

function choiceConflict(resource, kind, session, sessions, plan, owner) {
  const busy = availabilityConflict(resource, session, sessions);
  if (busy) return busy;
  const booking = planBookings(sessions, plan).find((entry) => entry.kind === kind && entry.owner !== owner && entry.resourceId === resource.id && sameDay(entry, session) && overlaps(entry, session));
  return booking ? { type: "booking", booking, message: `Already assigned to ${booking.courseCode ? `the ${booking.courseCode} exam` : "slot backup duty"} on ${booking.day}, ${clock(booking.start)}-${clock(booking.end)}.` } : null;
}

export function resourceChoiceReason(resource, kind, session, sessions, plan, owner) {
  return choiceConflict(resource, kind, session, sessions, plan, owner)?.message || "";
}

export function resourceChoiceSummary(resource, kind, session, sessions, plan, owner) {
  const conflict = choiceConflict(resource, kind, session, sessions, plan, owner);
  if (!conflict) return "";
  if (conflict.type === "excluded") return "Excluded from pool";
  if (conflict.type === "class") {
    const entry = conflict.entry;
    return "Class " + entry.code + (entry.unknownTime ? " (time unknown; day blocked)" : " " + clock(entry.start) + "-" + clock(entry.end));
  }
  if (conflict.type === "booking") {
    const booking = conflict.booking;
    return (booking.courseCode ? "Exam " + booking.courseCode : "Backup duty") + " " + clock(booking.start) + "-" + clock(booking.end);
  }
  return conflict.message;
}

export function standbyAvailability(session, sessions, catalog, plan) {
  // Assigned backups are standby capacity, but bookings in other overlapping slots are not.
  const assigned = new Set((plan?.backups?.[session.id] || []).filter(Boolean));
  const occupied = new Set(planBookings(sessions, plan).filter((booking) => booking.kind === "invigilator" &&
    !booking.owner.startsWith(session.id + "/backup/") && sameDay(booking, session) && overlaps(booking, session))
    .map((booking) => booking.resourceId));
  const available = catalog.invigilators.filter((person) => !occupied.has(person.id) && !availabilityConflict(person, session, sessions));
  return { count: available.length, assigned: available.filter((person) => assigned.has(person.id)).length,
    unassigned: available.filter((person) => !assigned.has(person.id)).length };
}

export function validateResourcePlan(sessions, catalog, plan) {
  const issues = [];
  const add = (session, room, title, message, action) => {
    const courseRooms = session?.rooms.filter((entry) => entry.courseId === room?.courseId) || [];
    issues.push({ sessionId: session?.id, roomId: room?.id, title, message, action,
      context: session ? resourceSessionLabel(session) : "Timetable",
      exam: room ? `${room.code} / Exam room ${courseRooms.indexOf(room) + 1} of ${courseRooms.length} / ${room.students.length} students` : "",
    });
  };
  if (!sessions.length) add(null, null, "No exams to resource", "No main exams scheduled.", "Return to scheduling and place your department exams first.");
  if (!plan || plan.fingerprint !== resourceFingerprint(sessions, catalog)) add(null, null, "Resources need reassignment", "Resource assignments are missing or outdated because the timetable or resource pool changed.", "Click Auto-Assign Resources, then review the proposed assignments before export.");
  const rooms = new Map(catalog.rooms.map((room) => [room.id, room]));
  const staff = new Map(catalog.invigilators.map((person) => [person.id, person]));
  const check = (resource, kind, session, owner, room) => {
    const label = kind === "room" ? "Room" : room ? "Invigilator" : "Backup invigilator";
    const conflict = choiceConflict(resource, kind, session, sessions, plan, owner);
    if (!conflict) return;
    const action = conflict.type === "excluded" ? `Include ${resource.name} in Review Resource Pool, or choose another available ${kind}.` :
      conflict.type === "class" && conflict.entry.unknownTime ? `Check the missing class time in the CRN list and reimport it, or choose another available ${kind}.` :
      `Select an available ${kind} in ${room ? "the assignment" : "Slot backups"} below. If none are available, review the timetable and resource pool.`;
    add(session, room, resource ? `${label} unavailable` : `${label} not assigned`,
      resource ? `${resource.name}: ${conflict.message}` : `No ${kind} from the resource pool is assigned.`, action);
  };
  const checkFixed = (resource, selectedId, kind, session, owner, room) => {
    const label = kind === "room" ? "Original lab room" : "Lab instructor";
    const conflict = choiceConflict(resource, kind, session, sessions, plan, owner);
    const name = resource?.name || `the ${kind} listed for this lab`;
    if (conflict) {
      let action;
      if (conflict.type === "excluded") action = `Include ${name} in Review Resource Pool, then reassign resources.`;
      else if (conflict.type === "missing") action = `Check the lab ${kind} in the CRN list and reimport the corrected file.`;
      else if (conflict.type === "booking") action = `Remove the overlapping exam or backup assignment for ${name}. This lab exam must retain its original ${kind}.`;
      else action = `Check the overlapping commitments in the CRN list and resolve or correct them before reimporting. Alternatively, choose another listed lab time if available. Only this exam's lab is replaced; other classes remain blocked.`;
      add(session, room, `${label} unavailable`, `${name} is required for this exam's original lab room. ${conflict.message}`, action);
    } else if (selectedId !== resource.id) {
      add(session, room, `${label} not assigned`, `This exam replaces its lab and must use ${name}.`,
        `Select ${name} in the assignment below, or click Auto-Assign Resources.`);
    }
  };
  sessions.forEach((session) => {
    session.issues.forEach((issue) => add(session, null, "Exam scheduling issue", issue, "Return to scheduling and check this exam's enrolment and listed lab times."));
    session.rooms.forEach((room) => {
      const allocation = plan?.allocations?.[room.id];
      const ids = allocation?.invigilatorIds || [];
      const approved = room.distributionChoice === "merge" || room.distributionChoice === "distribute";
      const limit = approved && room.maxStudents === 27 ? 27 : Math.min(25, room.maxStudents || 25);
      if (room.students.length > limit) add(session, room, "Room capacity exceeded", `This room has ${room.students.length} students; the ${approved ? "approved" : "normal"} maximum is ${limit}.`, "Use more rooms, or explicitly approve an offered consolidation up to 27 students per room. Never exceed 27.");
      if (room.missingLabResource) add(session, room, "Lab details missing", "The CRN list is missing the room or teaching instructor for the replaced lab.", "Complete the lab details in the CRN list and reimport it, then reassign resources.");
      if (room.fixedRoomId) checkFixed(rooms.get(room.fixedRoomId), allocation?.roomId, "room", session, room.id, room);
      else check(rooms.get(allocation?.roomId), "room", session, room.id, room);
      if (room.fixedInvigilatorId) checkFixed(staff.get(room.fixedInvigilatorId), ids[0], "invigilator", session, `${room.id}/invigilator/0`, room);
      const missing = Array.from({ length: room.requiredInvigilators }, (_, index) => index).filter((index) => !ids[index] && !(index === 0 && room.fixedInvigilatorId)).length;
      if (missing) add(session, room, room.fixedInvigilatorId ? "Additional invigilator needed" : "Invigilators not assigned",
        `${missing} ${room.fixedInvigilatorId ? "additional " : ""}invigilator(s) still needed. ${room.students.length} students require ${room.requiredInvigilators} invigilator(s) in this room${room.students.length > 15 ? " because there are more than 15 students" : ""}.`,
        "Select available invigilators in the assignment below, or include more available staff in Review Resource Pool and reassign.");
      if (ids.some((id, index) => id && index >= room.requiredInvigilators)) add(session, room, "Too many invigilators", `This room requires ${room.requiredInvigilators} invigilator(s).`, "Remove the extra room duties by reassigning resources.");
      for (let index = 0; index < room.requiredInvigilators; index += 1) {
        const id = ids[index];
        if (id && !(index === 0 && room.fixedInvigilatorId)) check(staff.get(id), "invigilator", session, `${room.id}/invigilator/${index}`, room);
      }
    });
    const backups = plan?.backups?.[session.id] || [];
    const backupCount = backups.filter(Boolean).length;
    const maximum = backupTarget(session.rooms.length);
    if (backupCount < 1 || backupCount > maximum) add(session, null, backupCount < 1 ? "Slot backup not assigned" : "Too many slot backups",
      `This time slot has ${session.rooms.length} exam room(s) and ${backupCount} backup(s) assigned. It needs at least 1 and at most ${maximum}; backups cover the slot, not each room.`,
      backupCount < 1 ? "Select an available invigilator in Slot backups below. They cannot also cover an exam room at this time." : "Remove the extra selections in Slot backups below.");
    backups.forEach((id, index) => { if (id) check(staff.get(id), "invigilator", session, `${session.id}/backup/${index}`, null); });
  });
  return { complete: issues.length === 0, issues };
}

export function invigilatorWorkloads(catalog, sessions, plan) {
  return catalog.invigilators.map((person) => {
    let exam = 0;
    let backup = 0;
    let teaching = 0;
    const bySlot = {};
    const backupBySlot = {};
    sessions.forEach((session) => {
      const count = session.rooms.reduce((sum, room) => sum + (plan?.allocations?.[room.id]?.invigilatorIds || []).filter((id) => id === person.id).length, 0);
      const backupCount = (plan?.backups?.[session.id] || []).filter((id) => id === person.id).length;
      if (isTeachingTimeDuty(person, session, sessions)) {
        teaching += count + backupCount;
        return;
      }
      exam += count;
      if (count) bySlot[session.slotId] = (bySlot[session.slotId] || 0) + count;
      backup += backupCount;
      if (backupCount) backupBySlot[session.slotId] = (backupBySlot[session.slotId] || 0) + backupCount;
    });
    return { ...person, exam, backup, teaching, bySlot, backupBySlot };
  });
}

export function assignResources(sessions, catalog) {
  const plan = emptyResourcePlan(sessions, catalog);
  const staff = catalog.invigilators.filter((person) => person.enabled);
  const workloads = new Map(staff.map((person) => [person.id, { exam: 0, backup: 0, examSlots: {}, backupSlots: {} }]));
  const roomUse = new Map();
  const selectStaff = (session, owner, role) => staff.filter((person) => !resourceChoiceReason(person, "invigilator", session, sessions, plan, owner)).sort((a, b) => {
    const loadA = workloads.get(a.id);
    const loadB = workloads.get(b.id);
    const extraA = isTeachingTimeDuty(a, session, sessions) ? 0 : 1;
    const extraB = isTeachingTimeDuty(b, session, sessions) ? 0 : 1;
    return (loadA[role] + extraA) - (loadB[role] + extraB) ||
      ((loadA[`${role}Slots`][session.slotId] || 0) + extraA) - ((loadB[`${role}Slots`][session.slotId] || 0) + extraB) ||
      extraA - extraB || a.name.localeCompare(b.name);
  })[0];
  const record = (person, role, session) => {
    if (isTeachingTimeDuty(person, session, sessions)) return;
    const load = workloads.get(person.id);
    load[role] += 1;
    load[`${role}Slots`][session.slotId] = (load[`${role}Slots`][session.slotId] || 0) + 1;
  };
  // Reserve fixed lab resources before flexible exams, including overlapping start times.
  sessions.forEach((session) => session.rooms.forEach((room) => {
    const allocation = { roomId: "", invigilatorIds: Array(room.requiredInvigilators).fill("") };
    plan.allocations[room.id] = allocation;
    const fixedRoom = catalog.rooms.find((entry) => entry.id === room.fixedRoomId);
    if (fixedRoom && !resourceChoiceReason(fixedRoom, "room", session, sessions, plan, room.id)) allocation.roomId = fixedRoom.id;
    const lead = staff.find((entry) => entry.id === room.fixedInvigilatorId);
    if (lead && !resourceChoiceReason(lead, "invigilator", session, sessions, plan, `${room.id}/invigilator/0`)) {
      allocation.invigilatorIds[0] = lead.id;
      record(lead, "exam", session);
    }
  }));
  const ordered = [...sessions].sort((a, b) => {
    const available = (session) => staff.filter((person) => !resourceBusyReason(person, session, sessions)).length - session.rooms.reduce((sum, room) => sum + room.requiredInvigilators, 0);
    return available(a) - available(b) || a.week - b.week || DAYS.indexOf(a.day) - DAYS.indexOf(b.day) || a.start - b.start;
  });
  ordered.forEach((session) => session.rooms.forEach((room) => {
    const allocation = plan.allocations[room.id];
    if (!allocation.roomId && !room.fixedRoomId) {
      const availableRoom = catalog.rooms.filter((entry) => !resourceChoiceReason(entry, "room", session, sessions, plan, room.id))
        .sort((a, b) => (roomUse.get(a.id) || 0) - (roomUse.get(b.id) || 0) || a.name.localeCompare(b.name))[0];
      if (availableRoom) allocation.roomId = availableRoom.id;
    }
    if (allocation.roomId) roomUse.set(allocation.roomId, (roomUse.get(allocation.roomId) || 0) + 1);
    for (let index = 0; index < room.requiredInvigilators; index += 1) {
      if (allocation.invigilatorIds[index] || (index === 0 && room.fixedInvigilatorId)) continue;
      const person = selectStaff(session, `${room.id}/invigilator/${index}`, "exam");
      if (person) { allocation.invigilatorIds[index] = person.id; record(person, "exam", session); }
    }
  }));
  // Backups use their own balancing counters and cannot reuse room invigilators.
  ordered.forEach((session) => {
    plan.backups[session.id] = [];
    for (let index = 0; index < backupTarget(session.rooms.length); index += 1) {
      const person = selectStaff(session, `${session.id}/backup/${index}`, "backup");
      if (person) { plan.backups[session.id].push(person.id); record(person, "backup", session); }
    }
  });
  return plan;
}
