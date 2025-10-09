import { useEffect, useMemo, useRef, useState } from "react";

import * as XLSX from "xlsx";
import JSZip from "jszip";

import "./App.css";

const days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];

const MAX_WEEKS = 10;

const DEFAULT_SLOT_INTERVAL_MINUTES = 30;

const DEFAULT_START_HOUR = 8;

const DEFAULT_END_HOUR = 17;

const DEFAULT_STUDENTS_PER_ROOM = 25;

const DEFAULT_INVIGILATOR_PLACEHOLDER_COUNT = 15;

const XLSX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
}

function alignDateToMonday(date) {
  const result = new Date(date);
  const day = result.getDay();
  const diff = (day + 6) % 7;
  result.setDate(result.getDate() - diff);
  return result;
}

function formatDateToISO(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getDefaultStartDate() {
  const date = new Date();
  date.setHours(0, 0, 0, 0);
  date.setDate(date.getDate() + 14);
  return alignDateToMonday(date);
}

function getDefaultStartDateISO() {
  return formatDateToISO(getDefaultStartDate());
}

function parseISODateString(value) {
  if (!value) {
    return null;
  }

  const [yearStr, monthStr, dayStr] = value.split("-");
  const year = Number.parseInt(yearStr, 10);
  const month = Number.parseInt(monthStr, 10);
  const day = Number.parseInt(dayStr, 10);

  if (Number.isNaN(year) || Number.isNaN(month) || Number.isNaN(day)) {
    return null;
  }

  const date = new Date(year, month - 1, day);
  if (Number.isNaN(date.getTime())) {
    return null;
  }

  date.setHours(0, 0, 0, 0);
  return date;
}

function addDays(date, amount) {
  const result = new Date(date);
  result.setDate(result.getDate() + amount);
  return result;
}

const invigilatorDateFormatter = new Intl.DateTimeFormat(undefined, {
  weekday: "long",
  year: "numeric",
  month: "short",
  day: "numeric",
});

function formatDateForInvigilator(date) {
  return invigilatorDateFormatter.format(date);
}

function generateInvigilatorPlaceholders(
  count = DEFAULT_INVIGILATOR_PLACEHOLDER_COUNT,
) {
  return Array.from({ length: count }, (_, index) => {
    const number = String(index + 1).padStart(2, "0");
    return `Invigilator ${number}`;
  });
}

function assignInvigilatorsToRows(totalRows, placeholders, roomNames = []) {
  const primaryPool = placeholders.map((name, index) => ({
    name,
    primaryCount: 0,
    roomUsage: new Map(),
    order: index,
  }));

  const backupPool = placeholders.map((name, index) => ({
    name,
    backupCount: 0,
    primaryCount: 0,
    order: index,
  }));

  const selectPrimary = (roomName, excluded = new Set()) =>
    primaryPool
      .filter((candidate) => !excluded.has(candidate.name))
      .sort((a, b) => {
        if (a.primaryCount !== b.primaryCount) {
          return a.primaryCount - b.primaryCount;
        }

        const aRoom = a.roomUsage.get(roomName) ?? 0;
        const bRoom = b.roomUsage.get(roomName) ?? 0;
        if (aRoom !== bRoom) {
          return aRoom - bRoom;
        }

        return a.order - b.order;
      })[0] ?? null;

  const selectBackup = (excluded = new Set()) =>
    backupPool
      .filter((candidate) => !excluded.has(candidate.name))
      .sort((a, b) => {
        if (a.backupCount !== b.backupCount) {
          return a.backupCount - b.backupCount;
        }

        if (a.primaryCount !== b.primaryCount) {
          return a.primaryCount - b.primaryCount;
        }

        return a.order - b.order;
      })[0] ?? null;

  const assignments = [];

  for (let index = 0; index < totalRows; index += 1) {
    const roomName = roomNames[index] ?? "";
    const primaryExclusions = new Set();

    const primaryOne = selectPrimary(roomName, primaryExclusions);
    if (primaryOne) {
      primaryOne.primaryCount += 1;
      primaryOne.roomUsage.set(
        roomName,
        (primaryOne.roomUsage.get(roomName) ?? 0) + 1,
      );
      primaryExclusions.add(primaryOne.name);
    }

    const primaryTwo = selectPrimary(roomName, primaryExclusions);
    if (primaryTwo) {
      primaryTwo.primaryCount += 1;
      primaryTwo.roomUsage.set(
        roomName,
        (primaryTwo.roomUsage.get(roomName) ?? 0) + 1,
      );
      primaryExclusions.add(primaryTwo.name);
    }

    const backupExclusions = new Set(primaryExclusions);
    const backup = selectBackup(backupExclusions);
    if (backup) {
      backup.backupCount += 1;
    }

    assignments.push({
      primaryOne: primaryOne?.name ?? "",
      primaryTwo: primaryTwo?.name ?? "",
      backup: backup?.name ?? "",
      roomName,
    });
  }

  return { assignments };
}
function formatTimeLabel(totalMinutes) {
  const hour24 = Math.floor(totalMinutes / 60);

  const minute = totalMinutes % 60;

  const suffix = hour24 >= 12 ? "PM" : "AM";

  const hour12 = ((hour24 + 11) % 12) + 1;

  const paddedMinute = minute.toString().padStart(2, "0");

  return `${hour12}:${paddedMinute} ${suffix}`;
}

function buildTimeSlots(startHour, endHour, slotIntervalMinutes) {
  const slots = [];

  if (
    Number.isFinite(startHour) &&
    Number.isFinite(endHour) &&
    Number.isFinite(slotIntervalMinutes) &&
    slotIntervalMinutes > 0 &&
    endHour > startHour
  ) {
    for (
      let minutes = startHour * 60;
      minutes <= endHour * 60 - slotIntervalMinutes;
      minutes += slotIntervalMinutes
    ) {
      const hour = Math.floor(minutes / 60);
      const minute = minutes % 60;
      const id = `${hour.toString().padStart(2, "0")}:${minute.toString().padStart(2, "0")}`;

      slots.push({
        id,
        label: formatTimeLabel(minutes),
      });
    }
  }

  return slots;
}

const DEFAULT_TIME_SLOTS = buildTimeSlots(
  DEFAULT_START_HOUR,
  DEFAULT_END_HOUR,
  DEFAULT_SLOT_INTERVAL_MINUTES,
);

function StudentIcon() {
  return (
    <svg
      className="slot-summary__icon"
      viewBox="0 0 24 24"
      width="16"
      height="16"
      aria-hidden="true"
      focusable="false"
    >
      <circle cx="8" cy="9" r="3" fill="currentColor" />

      <circle cx="16" cy="9" r="3" fill="currentColor" fillOpacity="0.6" />

      <path d="M4 20c0-3 3.8-5.5 8-5.5s8 2.5 8 5.5v1H4z" fill="currentColor" />
    </svg>
  );
}

function RoomIcon() {
  return (
    <svg
      className="slot-summary__icon"
      viewBox="0 0 24 24"
      width="16"
      height="16"
      aria-hidden="true"
      focusable="false"
    >
      <path
        d="M4 21V10.2L12 4l8 6.2V21h-5.5v-6.5h-5V21H4z"
        fill="currentColor"
      />

      <rect x="11" y="12.5" width="2" height="3.5" fill="currentColor" />
    </svg>
  );
}

function InvigilatorIcon() {
  return (
    <svg
      className="slot-summary__icon"
      viewBox="0 0 24 24"
      width="16"
      height="16"
      aria-hidden="true"
      focusable="false"
    >
      <circle cx="12" cy="7" r="3.5" fill="currentColor" />

      <path
        d="M6.5 21v-3.2c0-3.5 2.9-6.3 5.5-6.3s5.5 2.8 5.5 6.3V21H6.5z"
        fill="currentColor"
      />

      <path
        d="M11.2 11.6h1.6l0.9 2.3-1.7 2.4-1.7-2.4 0.9-2.3z"
        fill="#ffffff"
      />
    </svg>
  );
}

function createEmptyDaySlotMap(timeSlots) {
  const map = {};

  days.forEach((day) => {
    map[day] = {};

    timeSlots.forEach((slot) => {
      map[day][slot.id] = [];
    });
  });

  return map;
}


function buildEmptyAssignments(weekList, timeSlots) {
  const empty = {};

  weekList.forEach((week) => {
    empty[week] = createEmptyDaySlotMap(timeSlots);
  });

  return empty;
}


function cloneAssignments(assignments, timeSlots) {
  const clone = {};

  Object.entries(assignments || {}).forEach(([weekKey, weekAssignments]) => {
    const week = Number(weekKey);

    clone[week] = {};

    days.forEach((day) => {
      clone[week][day] = {};

      timeSlots.forEach((slot) => {
        const sourceList = weekAssignments?.[day]?.[slot.id] ?? [];

        clone[week][day][slot.id] = [...sourceList];
      });
    });
  });

  return clone;
}

function reshapeAssignments(assignments, timeSlots) {
  const weekKeys = Object.keys(assignments || {}).map((value) => Number(value));
  const reshaped = {};

  weekKeys.forEach((week) => {
    const weekAssignments = assignments?.[week] || {};
    reshaped[week] = {};

    days.forEach((day) => {
      const dayAssignments = weekAssignments?.[day] || {};
      reshaped[week][day] = {};

      timeSlots.forEach((slot) => {
        const existing = Array.isArray(dayAssignments[slot.id])
          ? dayAssignments[slot.id]
          : [];
        reshaped[week][day][slot.id] = [...existing];
      });
    });
  });

  return reshaped;
}



function formatStudentReference(studentId, directory) {
  const rawName = (directory[studentId] || "").trim();

  if (!rawName) {
    return studentId;
  }

  const parts = rawName.split(/\s+/).filter(Boolean);

  if (!parts.length) {
    return studentId;
  }

  const first = parts[0];

  const last = parts[parts.length - 1];

  if (first === last) {
    return `${studentId} ${first}`;
  }

  return `${studentId} ${first} ${last}`;
}

const DEFAULT_INVIGILATOR_HEADER = [
  "CRN",
  "Course Code",
  "Course Title",
  "No of Students",
  "Date",
  "Time",
  "Instructor Name",
  "Invigilator room",
  "Invigilator1",
  "Invigilator2",
  "Backup Invigilator",
];

const DEFAULT_STUDENT_HEADER = [
  "CRN",
  "Code",
  "Title",
  "Student ID",
  "Student Name",
  "Class room",
  "Present/ Absent",
];

function parseSlotIdToMinutes(slotId) {
  const [hour = "0", minute = "0"] = String(slotId).split(":");

  return Number.parseInt(hour, 10) * 60 + Number.parseInt(minute, 10);
}

function formatSlotRange(slotId) {
  const startMinutes = parseSlotIdToMinutes(slotId);

  const endMinutes = startMinutes + 60;

  return `${formatTimeLabel(startMinutes)} - ${formatTimeLabel(endMinutes)}`;
}

function sortStudentsForExport(students = []) {
  return [...students].sort((a, b) => {
    const nameA = (a.name || "").toLowerCase();

    const nameB = (b.name || "").toLowerCase();

    if (nameA && nameB && nameA !== nameB) {
      return nameA.localeCompare(nameB);
    }

    const idA = String(a.id || "");

    const idB = String(b.id || "");

    return idA.localeCompare(idB);
  });
}

function resolveStudentName(student, directory) {
  const explicitName = student?.name?.trim();

  if (explicitName) {
    return explicitName;
  }

  const lookupName = directory?.[student?.id];

  return lookupName ? lookupName.trim() : "";
}

function getCourseCrnString(course) {
  if (Array.isArray(course?.crnDetails) && course.crnDetails.length) {
    const crns = course.crnDetails.map((detail) => detail.crn).filter(Boolean);
    if (crns.length) {
      return crns.join(", ");
    }
  }

  if (Array.isArray(course?.crns) && course.crns.length) {
    return course.crns.join(", ");
  }

  return "";
}
function normaliseSheetName(base) {
  if (!base) {
    return "Sheet";
  }

  return base.length <= 31 ? base : base.slice(0, 31);
}

function computeSlotSummaries(assignments, courseLookup, week, timeSlots, studentsPerRoom) {
  const weekAssignments = assignments[week] || {};

  const summary = createEmptyDaySlotMap(timeSlots);
  const studentsPerRoomSafe = Math.max(1, Number(studentsPerRoom) || 1);

  days.forEach((day) => {
    timeSlots.forEach((slot) => {
      const courses = weekAssignments[day]?.[slot.id] ?? [];

      const isStartSlot = courses.length > 0;

      const seenStudentIds = new Set();

      if (isStartSlot) {
        courses.forEach((courseId) => {
          const course = courseLookup[courseId];

          if (!course) return;

          course.students.forEach((student) => {
            seenStudentIds.add(student.id);
          });
        });
      }

      const uniqueStudentCount = seenStudentIds.size;

      const roomCount =
        isStartSlot && uniqueStudentCount > 0
          ? Math.ceil(uniqueStudentCount / studentsPerRoomSafe)
          : 0;

      const invigilatorCount = roomCount * 2;

      summary[day][slot.id] = {
        studentCount: uniqueStudentCount,

        uniqueStudents: uniqueStudentCount,

        roomCount,

        invigilatorCount,

        isStartSlot,
      };
    });
  });

  return summary;
}


function computeConflicts(
  assignments,
  courseLookup,
  studentDirectory,
  timeSlots,
  studentsPerRoom,
  availableInvigilators,
) {
  const byWeek = {};

  const overallMessages = new Set();

  const studentsPerRoomSafe = Math.max(1, Number(studentsPerRoom) || 1);
  const invigilatorCapacity = Math.max(
    0,
    Number(availableInvigilators) || 0,
  );

  const weekKeys = Object.keys(assignments || {})
    .map((value) => Number(value))
    .sort((a, b) => a - b);

  weekKeys.forEach((week) => {
    const weekConflicts = createEmptyDaySlotMap(timeSlots);

    byWeek[week] = weekConflicts;

    const studentDayCounts = {};

    const weekAssignments = assignments[week] || {};

    days.forEach((day) => {
      timeSlots.forEach((slot, slotIndex) => {
        const startCourses = weekAssignments[day]?.[slot.id] ?? [];

        if (!startCourses.length && slotIndex === 0) {
          return;
        }

        const slotStudentCourses = new Map();

        const capacityStudentIds = new Set();

        const addCourseStudents = (courseId) => {
          const course = courseLookup[courseId];

          if (!course) return;

          course.students.forEach((student) => {
            capacityStudentIds.add(student.id);
          });
        };

        const addCourseToSlotMap = (courseId) => {
          const course = courseLookup[courseId];

          if (!course) return;

          course.students.forEach((student) => {
            if (!slotStudentCourses.has(student.id)) {
              slotStudentCourses.set(student.id, new Set());
            }

            slotStudentCourses
              .get(student.id)
              .add(course.code || course.title || courseId);
          });
        };

        startCourses.forEach((courseId) => {
          addCourseToSlotMap(courseId);
          addCourseStudents(courseId);

          const course = courseLookup[courseId];

          if (!course) return;

          course.students.forEach((student) => {
            if (!studentDayCounts[student.id]) {
              studentDayCounts[student.id] = {};
            }

            studentDayCounts[student.id][day] =
              (studentDayCounts[student.id][day] || 0) + 1;
          });
        });

        if (slotIndex > 0) {
          const previousSlotId = timeSlots[slotIndex - 1].id;

          const previousCourses = weekAssignments[day]?.[previousSlotId] ?? [];

          previousCourses.forEach((courseId) => addCourseToSlotMap(courseId));
        }

        if (slotStudentCourses.size === 0) {
          return;
        }

        if (capacityStudentIds.size === 0 && slotIndex > 0) {
          const previousSlotId = timeSlots[slotIndex - 1].id;
          const previousCourses = weekAssignments[day]?.[previousSlotId] ?? [];

          previousCourses.forEach((courseId) => addCourseStudents(courseId));
        }

        if (capacityStudentIds.size > 0) {
          const roomsNeeded = Math.ceil(
            capacityStudentIds.size / studentsPerRoomSafe,
          );
          const requiredInvigilators = roomsNeeded * 2;

          if (requiredInvigilators > invigilatorCapacity) {
            const message = `Week ${week}: ${day} ${slot.label} requires ${requiredInvigilators} invigilators but only ${invigilatorCapacity} available.`;

            overallMessages.add(message);

            if (!weekConflicts[day][slot.id].includes(message)) {
              weekConflicts[day][slot.id] = [
                ...weekConflicts[day][slot.id],
                message,
              ];
            }
          }
        }

        slotStudentCourses.forEach((courseSet, studentId) => {
          if (courseSet.size > 1) {
            const studentLabel = formatStudentReference(
              studentId,
              studentDirectory,
            );

            const conflicts = Array.from(courseSet).join(", ");

            const message = `Week ${week}: Student ${studentLabel} has overlapping exams (${conflicts}) on ${day} at ${slot.label}`;

            overallMessages.add(message);

            weekConflicts[day][slot.id] = [
              ...weekConflicts[day][slot.id],
              message,
            ];
          }
        });
      });
    });

    Object.entries(studentDayCounts).forEach(([studentId, counts]) => {
      Object.entries(counts).forEach(([day, total]) => {
        if (total > 2) {
          const studentLabel = formatStudentReference(
            studentId,
            studentDirectory,
          );

          const message = `Week ${week}: Student ${studentLabel} is scheduled for ${total} exams on ${day}`;

          overallMessages.add(message);

          timeSlots.forEach((slot) => {
            const courses = weekAssignments[day]?.[slot.id] ?? [];

            if (courses.length === 0) return;

            const involved = courses.some((courseId) => {
              const course = courseLookup[courseId];

              if (!course) return false;

              return course.students.some(
                (student) => student.id === studentId,
              );
            });

            if (involved) {
              weekConflicts[day][slot.id] = [
                ...weekConflicts[day][slot.id],
                message,
              ];
            }
          });
        }
      });
    });

    days.forEach((day) => {
      timeSlots.forEach((slot) => {
        if (weekConflicts[day][slot.id].length > 1) {
          weekConflicts[day][slot.id] = Array.from(
            new Set(weekConflicts[day][slot.id]),
          );
        }
      });
    });
  });

  return {
    overall: Array.from(overallMessages),

    byWeek,
  };
}


function computeSummary(assignments, courseLookup, timeSlots, studentsPerRoom) {
  const scheduledCourseIds = new Set();

  const studentIds = new Set();

  let roomCount = 0;
  const studentsPerRoomSafe = Math.max(1, Number(studentsPerRoom) || 1);

  Object.values(assignments || {}).forEach((weekAssignments) => {
    days.forEach((day) => {
      timeSlots.forEach((slot) => {
        const courseIds = weekAssignments?.[day]?.[slot.id] ?? [];

        if (!courseIds.length) {
          return;
        }

        const slotStudentIds = new Set();

        courseIds.forEach((courseId) => {
          const course = courseLookup[courseId];

          if (!course) return;

          scheduledCourseIds.add(courseId);

          course.students.forEach((student) => {
            studentIds.add(student.id);

            slotStudentIds.add(student.id);
          });
        });

        if (slotStudentIds.size > 0) {
          roomCount += Math.ceil(slotStudentIds.size / studentsPerRoomSafe);
        }
      });
    });
  });

  const invigilators = roomCount * 2;

  return {
    totalCourses: scheduledCourseIds.size,

    totalStudents: studentIds.size,

    totalRooms: roomCount,

    totalInvigilators: invigilators,
  };
}


function App() {
  const [settings, setSettings] = useState({
    slotIntervalMinutes: DEFAULT_SLOT_INTERVAL_MINUTES,
    startHour: DEFAULT_START_HOUR,
    endHour: DEFAULT_END_HOUR,
    studentsPerRoom: DEFAULT_STUDENTS_PER_ROOM,
    invigilatorPlaceholderCount: DEFAULT_INVIGILATOR_PLACEHOLDER_COUNT,
  });

  const {
    slotIntervalMinutes,
    startHour,
    endHour,
    studentsPerRoom,
    invigilatorPlaceholderCount,
  } = settings;

  const [startDate, setStartDate] = useState(() => getDefaultStartDateISO());

  const [weeks, setWeeks] = useState(() => [1]);

  const [assignments, setAssignments] = useState(() =>
    buildEmptyAssignments([1], DEFAULT_TIME_SLOTS),
  );

  const [selectedWeek, setSelectedWeek] = useState(1);

  const [courses, setCourses] = useState([]);

  const [studentDirectory, setStudentDirectory] = useState({});

  const [uploadError, setUploadError] = useState("");

  const [courseSearch, setCourseSearch] = useState("");

  const [hoverTarget, setHoverTarget] = useState(null);

  const templateHeadersRef = useRef(null);

  const [isExporting, setIsExporting] = useState(false);

  const [exportError, setExportError] = useState("");

  const timeSlots = useMemo(
    () => buildTimeSlots(startHour, endHour, slotIntervalMinutes),
    [startHour, endHour, slotIntervalMinutes],
  );

  useEffect(() => {
    setAssignments((previous) => reshapeAssignments(previous, timeSlots));
  }, [timeSlots]);

  const studentsPerRoomCapacity = Math.max(1, Number(studentsPerRoom) || 1);

  const invigilatorPlaceholderTotal = Math.max(
    1,
    Number(invigilatorPlaceholderCount) || 1,
  );

  const courseLookup = useMemo(() => {
    const lookup = {};

    courses.forEach((course) => {
      lookup[course.id] = course;
    });

    return lookup;
  }, [courses]);

  const slotSummaries = useMemo(
    () =>
      computeSlotSummaries(
        assignments,
        courseLookup,
        selectedWeek,
        timeSlots,
        studentsPerRoomCapacity,
      ),
    [assignments, courseLookup, selectedWeek, timeSlots, studentsPerRoomCapacity],
  );

  const conflicts = useMemo(
    () =>
      computeConflicts(
        assignments,
        courseLookup,
        studentDirectory,
        timeSlots,
        studentsPerRoomCapacity,
        invigilatorPlaceholderTotal,
      ),
    [
      assignments,
      courseLookup,
      studentDirectory,
      studentsPerRoomCapacity,
      invigilatorPlaceholderTotal,
      timeSlots,
    ],
  );

  const occupiedSlotIds = useMemo(() => {
    const result = new Set();

    const weekAssignments = assignments[selectedWeek] || {};

    timeSlots.forEach((slot, index) => {
      const hasCourseInColumn = days.some((day) => {
        const dayAssignments = weekAssignments[day] || {};

        const slotCourses = dayAssignments[slot.id] || [];

        if (slotCourses.length > 0) {
          return true;
        }

        const previousSlotId = index > 0 ? timeSlots[index - 1].id : null;

        if (!previousSlotId) {
          return false;
        }

        const trailingCourses = dayAssignments[previousSlotId] || [];

        return trailingCourses.some(
          (courseId) => !slotCourses.includes(courseId),
        );
      });

      if (hasCourseInColumn) {
        result.add(slot.id);
      }
    });

    return result;
  }, [assignments, selectedWeek, timeSlots]);

  const summary = useMemo(
    () =>
      computeSummary(
        assignments,
        courseLookup,
        timeSlots,
        studentsPerRoomCapacity,
      ),
    [assignments, courseLookup, timeSlots, studentsPerRoomCapacity],
  );

  const assignedCourseIds = useMemo(() => {
    const ids = new Set();

    Object.values(assignments || {}).forEach((weekAssignments) => {
      days.forEach((day) => {
        timeSlots.forEach((slot) => {
          const courseIds = weekAssignments?.[day]?.[slot.id] ?? [];

          courseIds.forEach((courseId) => ids.add(courseId));
        });
      });
    });

    return ids;
  }, [assignments, timeSlots]);

  const orderedCourses = useMemo(() => {
    return [...courses].sort((a, b) => {
      const codeA = a.code.toLowerCase();

      const codeB = b.code.toLowerCase();

      if (codeA === codeB) {
        return a.title.localeCompare(b.title);
      }

      return codeA.localeCompare(codeB);
    });
  }, [courses]);

  const availableCourses = useMemo(
    () => orderedCourses.filter((course) => !assignedCourseIds.has(course.id)),
    [orderedCourses, assignedCourseIds],
  );

  const filteredCourses = useMemo(() => {
    const query = courseSearch.trim().toLowerCase();

    if (!query) {
      return availableCourses;
    }

    return availableCourses.filter((course) => {
      const sectionText = Array.isArray(course.sections)
        ? course.sections.join(" ")
        : course.section || "";

      const crnText = Array.isArray(course.crns) ? course.crns.join(" ") : "";

      const haystack =
        `${course.code} ${course.title} ${sectionText} ${crnText}`.toLowerCase();

      return haystack.includes(query);
    });
  }, [availableCourses, courseSearch]);

  const handleStartDateChange = (event) => {
    const rawValue = event.target.value;
    const parsed = parseISODateString(rawValue);
    const aligned = parsed ? alignDateToMonday(parsed) : getDefaultStartDate();
    setStartDate(formatDateToISO(aligned));
  };

  const handleNumericSettingChange = (key, options = {}) => (event) => {
    const rawValue = Number.parseInt(event.target.value, 10);

    if (Number.isNaN(rawValue)) {
      return;
    }

    const { min, max } = options;

    let value = rawValue;

    if (typeof min === "number") {
      value = Math.max(min, value);
    }

    if (typeof max === "number") {
      value = Math.min(max, value);
    }

    setSettings((previous) => {
      const next = { ...previous, [key]: value };

      if (key === "startHour" && value >= next.endHour) {
        next.endHour = Math.min(23, value + 1);
      }

      if (key === "endHour" && value <= next.startHour) {
        next.startHour = Math.max(0, value - 1);
      }

      if (key === "slotIntervalMinutes" && value <= 0) {
        next.slotIntervalMinutes = DEFAULT_SLOT_INTERVAL_MINUTES;
      }

      if (key === "studentsPerRoom" && value <= 0) {
        next.studentsPerRoom = 1;
      }

      if (key === "invigilatorPlaceholderCount" && value <= 0) {
        next.invigilatorPlaceholderCount = 1;
      }

      return next;
    });
  };

  const handleFileUpload = async (event) => {
    const file = event.target.files?.[0];

    if (!file) return;

    setUploadError("");

    try {
      let workbook;

      if (file.name.toLowerCase().endsWith(".csv")) {
        const text = await file.text();

        workbook = XLSX.read(text, { type: "string" });
      } else {
        const arrayBuffer = await file.arrayBuffer();

        workbook = XLSX.read(arrayBuffer, { type: "array" });
      }

      const [firstSheetName] = workbook.SheetNames;

      const sheet = workbook.Sheets[firstSheetName];

      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      if (!rows.length) {
        throw new Error("The selected file does not contain any rows.");
      }

      const coursesMap = new Map();

      const studentNames = {};

      rows.forEach((row) => {
        const studentIdRaw =
          row.SPRIDEN_ID ?? row.Student_ID ?? row["Student ID"] ?? "";

        const studentId = String(studentIdRaw).trim();

        if (!studentId) return;

        const studentNameRaw =
          row.STUDENT_NAME ?? row.Student_Name ?? row["Student Name"] ?? "";

        const studentName = String(studentNameRaw).trim() || studentId;

        studentNames[studentId] = studentName;

        const instructorRaw =
          row.CF_INSTRUCTOR ??
          row.Instructor_Name ??
          row.Instructor ??
          row["Instructor Name"] ??
          row["Instructor"];

        const instructorName = String(instructorRaw ?? "").trim();

        const courseCodeRaw =
          row.SCBCRSE_SUBJ_CODE_SCBCRSE_CRSE ??
          (row.SSBSECT_SUBJ_CODE && row.SSBSECT_CRSE_NUMB
            ? `${row.SSBSECT_SUBJ_CODE}-${row.SSBSECT_CRSE_NUMB}`
            : (row.Course_Code ?? row["Course Code"] ?? ""));

        const courseTitleRaw =
          row.SCBCRSE_TITLE ??
          row.Course_Title ??
          row["Course Title"] ??
          "Untitled Course";

        const courseSection = String(
          row.SSBSECT_SEQ_NUMB ?? row.Section ?? row["Section"] ?? "",
        ).trim();

        const courseCrn = String(
          row.SSBSECT_CRN ?? row.CRN ?? row["CRN"] ?? "",
        ).trim();

        const courseCode = String(courseCodeRaw || "Course").trim();

        const courseId =
          courseCrn ||
          `${courseCode}${courseSection ? `-${courseSection}` : ""}` ||
          courseCode;

        if (!courseId) return;

        if (!coursesMap.has(courseId)) {
          coursesMap.set(courseId, {
            id: courseId,

            crn: courseCrn,

            code: courseCode,

            title: String(courseTitleRaw).trim() || courseCode,

            section: courseSection,

            students: [],

            studentMap: new Map(),

            instructors: new Set(),
          });
        }

        const existing = coursesMap.get(courseId);

        existing.code = courseCode || existing.code;

        existing.title = String(courseTitleRaw).trim() || existing.title;

        if (courseSection && !existing.section) {
          existing.section = courseSection;
        }

        if (courseCrn) {
          existing.crn = courseCrn;
        }

        if (instructorName) {
          existing.instructors.add(instructorName);
        }

        const existingStudent = existing.studentMap.get(studentId);

        if (!existingStudent) {
          const studentEntry = {
            id: studentId,

            name: studentName,

            crn: courseCrn || courseId,

            instructor: instructorName,
          };

          existing.studentMap.set(studentId, studentEntry);

          existing.students.push(studentEntry);
        } else {
          if (!existingStudent.name && studentName) {
            existingStudent.name = studentName;
          }

          if (!existingStudent.crn && (courseCrn || courseId)) {
            existingStudent.crn = courseCrn || courseId;
          }

          if (!existingStudent.instructor && instructorName) {
            existingStudent.instructor = instructorName;
          }
        }
      });

      const parsedCourses = Array.from(coursesMap.values()).map((course) => {
        const students = course.students.slice();

        students.sort(
          (a, b) => a.name.localeCompare(b.name) || a.id.localeCompare(b.id),
        );

        const instructors = Array.from(course.instructors ?? []).filter(
          Boolean,
        );

        const studentCount = students.length;

        const roomsNeeded = studentCount
          ? Math.max(1, Math.ceil(studentCount / studentsPerRoomCapacity))
          : 0;

        return {
          id: course.id,

          crn: course.crn || "",

          code: course.code,

          title: course.title,

          section: course.section,

          students,

          studentCount,

          roomsNeeded,

          instructors,

          primaryInstructor: instructors[0] || "",
        };
      });

      if (!parsedCourses.length) {
        throw new Error("No courses were found in the provided file.");
      }

      const groupedCoursesMap = new Map();

      parsedCourses.forEach((course) => {
        const key = course.code || course.id;

        if (!key) return;

        if (!groupedCoursesMap.has(key)) {
          groupedCoursesMap.set(key, {
            id: key,

            code: course.code || key,

            title: course.title,

            sections: new Set(),

            crns: new Set(),

            students: [],

            studentMap: new Map(),

            instructors: new Set(),

            crnDetails: new Map(),
          });
        }

        const groupedCourse = groupedCoursesMap.get(key);

        if (course.section) {
          groupedCourse.sections.add(course.section);
        }

        if (course.crn) {
          groupedCourse.crns.add(course.crn);
        }

        if (Array.isArray(course.instructors)) {
          course.instructors.forEach((name) => {
            if (name) {
              groupedCourse.instructors.add(name);
            }
          });
        }

        if (course.primaryInstructor) {
          groupedCourse.instructors.add(course.primaryInstructor);
        }

        const crnKey = course.crn || course.id;

        if (crnKey) {
          groupedCourse.crns.add(crnKey);

          if (!groupedCourse.crnDetails.has(crnKey)) {
            groupedCourse.crnDetails.set(crnKey, {
              instructor:
                course.primaryInstructor ||
                (Array.isArray(course.instructors)
                  ? course.instructors.find(Boolean) || ""
                  : ""),

              students: [],
            });
          }
        }

        const crnDetail = groupedCourse.crnDetails.get(crnKey);

        if (!crnDetail.instructor) {
          const fallbackInstructor =
            course.primaryInstructor ||
            (Array.isArray(course.instructors)
              ? course.instructors.find(Boolean) || ""
              : "");

          if (fallbackInstructor) {
            crnDetail.instructor = fallbackInstructor;
          }
        }

        course.students.forEach((student) => {
          const studentId = student.id;

          if (!studentId) return;

          const studentCrn = student.crn || crnKey || "";

          const studentNameValue =
            student.name || studentNames[studentId] || studentId;

          crnDetail.students.push({
            id: studentId,

            name: studentNameValue,

            crn: studentCrn,
          });

          if (!groupedCourse.studentMap.has(studentId)) {
            groupedCourse.studentMap.set(studentId, {
              id: studentId,

              name: studentNameValue,

              crn: studentCrn,
            });

            groupedCourse.students.push(
              groupedCourse.studentMap.get(studentId),
            );
          } else {
            const existingStudent = groupedCourse.studentMap.get(studentId);

            if (!existingStudent.crn && studentCrn) {
              existingStudent.crn = studentCrn;
            }

            if (!existingStudent.name && studentNameValue) {
              existingStudent.name = studentNameValue;
            }
          }
        });
      });

      const groupedCourses = Array.from(groupedCoursesMap.values()).map(
        (group) => {
          const sections = Array.from(group.sections).filter(Boolean).sort();

          const crns = Array.from(group.crns).filter(Boolean).sort();

          const instructors = Array.from(group.instructors)
            .filter(Boolean)
            .sort();

          const students = group.students

            .map((student) => ({ ...student }))

            .sort(
              (a, b) =>
                a.name.localeCompare(b.name) || a.id.localeCompare(b.id),
            );

          const studentCount = students.length;

          const roomsNeeded = studentCount
            ? Math.max(1, Math.ceil(studentCount / studentsPerRoomCapacity))
            : 0;

          const crnDetails = Array.from(group.crnDetails.entries())

            .map(([crn, detail]) => ({
              crn,

              instructor: detail.instructor || "",

              students: sortStudentsForExport(detail.students),
            }))

            .sort((a, b) => a.crn.localeCompare(b.crn));

          return {
            id: group.id,

            code: group.code,

            title: group.title,

            sections,

            crns,

            students,

            studentCount,

            roomsNeeded,

            instructors,

            primaryInstructor: instructors[0] || "",

            crnDetails,
          };
        },
      );

      const initialWeeks = [1];

      setWeeks(initialWeeks);

      setCourses(groupedCourses);

      setCourseSearch("");

      setStudentDirectory(studentNames);

      setAssignments(buildEmptyAssignments(initialWeeks, timeSlots));

      setSelectedWeek(initialWeeks[0]);
    } catch (error) {
      console.error(error);

      setUploadError(error.message || "Failed to read the provided file.");
    }
  };

  const updateHoverTarget = (day, slotIndex) => {
    if (slotIndex > timeSlots.length - 2) {
      setHoverTarget(null);

      return;
    }

    setHoverTarget((previous) => {
      if (
        previous &&
        previous.day === day &&
        previous.slotIndex === slotIndex
      ) {
        return previous;
      }

      return { day, slotIndex };
    });
  };

  const handleDragEnterSlot = (event, day, slotIndex) => {
    updateHoverTarget(day, slotIndex);
  };

  const handleDragOverSlot = (event, day, slotIndex) => {
    event.preventDefault();

    updateHoverTarget(day, slotIndex);
  };

  const handleDragLeaveSlot = (event, day, slotIndex) => {
    const relatedTarget = event.relatedTarget;

    if (relatedTarget) {
      const relatedCell =
        typeof relatedTarget.closest === "function"
          ? relatedTarget.closest("td[data-slot-index]")
          : null;

      if (relatedCell) {
        const relatedDay = relatedCell.getAttribute("data-day");

        const relatedSlotIndex = Number(
          relatedCell.getAttribute("data-slot-index"),
        );

        if (
          relatedDay === day &&
          (relatedSlotIndex === slotIndex || relatedSlotIndex === slotIndex + 1)
        ) {
          return;
        }
      }
    }

    setHoverTarget((previous) => {
      if (!previous) {
        return previous;
      }

      if (previous.day === day && previous.slotIndex === slotIndex) {
        return null;
      }

      return previous;
    });
  };

  const handleDrop = (day, slotId, slotIndex, event) => {
    event.preventDefault();

    setHoverTarget(null);

    const courseId = event.dataTransfer.getData("text/plain");

    if (!courseId || !courseLookup[courseId]) return;

    if (slotIndex === undefined || slotIndex > timeSlots.length - 2) return;

    setAssignments((previous) => {
      const next = cloneAssignments(previous, timeSlots);

      Object.keys(next).forEach((weekKey) => {
        const weekAssignments = next[weekKey];

        days.forEach((existingDay) => {
          timeSlots.forEach((slot) => {
            const list = weekAssignments[existingDay][slot.id];

            const index = list.indexOf(courseId);

            if (index !== -1) {
              list.splice(index, 1);
            }
          });
        });
      });

      const targetWeek = next[selectedWeek] || createEmptyDaySlotMap(timeSlots);

      if (!next[selectedWeek]) {
        next[selectedWeek] = targetWeek;
      }

      const targetList = targetWeek[day][slotId];

      if (!targetList.includes(courseId)) {
        targetList.push(courseId);
      }

      return next;
    });
  };

  const handleRemoveCourse = (day, slotId, courseId) => {
    setAssignments((previous) => {
      const next = cloneAssignments(previous, timeSlots);

      next[selectedWeek][day][slotId] = next[selectedWeek][day][slotId].filter(
        (id) => id !== courseId,
      );

      return next;
    });
  };

  const getTemplateHeaders = async () => {
    if (templateHeadersRef.current) {
      return templateHeadersRef.current;
    }

    try {
      const templateUrl = new URL(
        "../data/ReportReference/Report Template.xlsx",
        import.meta.url,
      );
      const response = await fetch(templateUrl);
      if (!response.ok) {
        throw new Error(
          `Failed to fetch template workbook: ${response.status}`,
        );
      }

      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array" });

      const weekSheetName =
        workbook.SheetNames.find((name) =>
          name.toLowerCase().includes("week"),
        ) ?? workbook.SheetNames[0];
      const daySheetName =
        workbook.SheetNames.find((name) =>
          name.toLowerCase().includes("day"),
        ) ?? workbook.SheetNames[1];

      const invigilatorHeaderRow =
        XLSX.utils.sheet_to_json(workbook.Sheets[weekSheetName], {
          header: 1,
        })[0] || DEFAULT_INVIGILATOR_HEADER;
      const studentHeaderRow =
        XLSX.utils.sheet_to_json(workbook.Sheets[daySheetName], {
          header: 1,
        })[0] || DEFAULT_STUDENT_HEADER;

      const normalisedInvigilatorHeader = (
        Array.isArray(invigilatorHeaderRow) && invigilatorHeaderRow.length
          ? invigilatorHeaderRow
          : DEFAULT_INVIGILATOR_HEADER
      ).map((value) =>
        value === undefined || value === null ? "" : String(value),
      );

      const normalisedStudentHeaderInitial = (
        Array.isArray(studentHeaderRow) && studentHeaderRow.length
          ? studentHeaderRow
          : DEFAULT_STUDENT_HEADER
      ).map((value) =>
        value === undefined || value === null ? "" : String(value),
      );

      const normalisedStudentHeader = normalisedStudentHeaderInitial
        .map((value) => value.trim())
        .filter((value) => value && value.toLowerCase() !== "course");

      templateHeadersRef.current = {
        invigilatorHeader: normalisedInvigilatorHeader,
        studentHeader:
          normalisedStudentHeader.length === DEFAULT_STUDENT_HEADER.length
            ? normalisedStudentHeader
            : [...DEFAULT_STUDENT_HEADER],
      };
    } catch (error) {
      console.error("Failed to load template workbook", error);
      templateHeadersRef.current = {
        invigilatorHeader: [...DEFAULT_INVIGILATOR_HEADER],
        studentHeader: [...DEFAULT_STUDENT_HEADER],
      };
    }

    return templateHeadersRef.current;
  };

  const buildWorkbookForWeek = (week, templateHeaders, baseStartDate) => {
    const weekAssignments = assignments[week];
    if (!weekAssignments) {
      return null;
    }

    const safeStartDate = alignDateToMonday(
      baseStartDate ?? getDefaultStartDate(),
    );
    const weekStartDate = addDays(safeStartDate, (week - 1) * 7);

    const invigilatorRows = [];
    const dayRowsMap = new Map();

    days.forEach((day, dayIndex) => {
      const dayAssignments = weekAssignments[day] || {};
      const dayDate = addDays(weekStartDate, dayIndex);
      const formattedDate = formatDateForInvigilator(dayDate);

      timeSlots.forEach((slot) => {
        const courseIds = dayAssignments[slot.id] || [];
        if (!courseIds.length) {
          return;
        }

        const timeRange = formatSlotRange(slot.id);
        const timeSortKey = parseSlotIdToMinutes(slot.id);

        const sortedCourseIds = [...courseIds].sort((a, b) => {
          const aCount = courseLookup[a]?.students?.length || 0;
          const bCount = courseLookup[b]?.students?.length || 0;
          return bCount - aCount;
        });

        const roomInputCourses = [];

        sortedCourseIds.forEach((courseId) => {
          const course = courseLookup[courseId];
          if (
            !course ||
            !Array.isArray(course.students) ||
            !course.students.length
          ) {
            return;
          }

          const crnGroups =
            Array.isArray(course.crnDetails) && course.crnDetails.length
              ? course.crnDetails
              : [
                  {
                    crn: getCourseCrnString(course) || course.id || "",
                    instructor:
                      course.primaryInstructor ||
                      (Array.isArray(course.instructors)
                        ? course.instructors.find((name) => name) || ""
                        : ""),
                    students: course.students,
                  },
                ];

          const seenStudentIds = new Set();
          const courseStudents = [];

          crnGroups.forEach((crnGroup) => {
            const resolvedInstructor =
              crnGroup.instructor ||
              course.primaryInstructor ||
              (Array.isArray(course.instructors)
                ? course.instructors.find((name) => name) || ""
                : "");

            const sortedStudents = sortStudentsForExport(
              crnGroup.students || [],
            )
              .map((student) => {
                const studentId = String(student.id ?? "").trim();
                if (!studentId) {
                  return null;
                }

                if (seenStudentIds.has(studentId)) {
                  return null;
                }

                const studentNameValue =
                  resolveStudentName(student, studentDirectory) || studentId;
                const studentCrn =
                  student.crn || crnGroup.crn || course.id || "";

                return {
                  id: studentId,
                  name: studentNameValue,
                  crn: studentCrn,
                  instructor: student.instructor || resolvedInstructor,
                  courseId: course.id,
                  courseCode: course.code || "",
                  courseTitle: course.title ? course.title.trim() : "",
                };
              })
              .filter(Boolean);

            sortedStudents.forEach((studentEntry) => {
              seenStudentIds.add(studentEntry.id);
              courseStudents.push(studentEntry);
            });
          });

          if (courseStudents.length) {
            roomInputCourses.push({
              course,
              students: courseStudents,
            });
          }
        });

        if (!roomInputCourses.length) {
          return;
        }

        const rooms = [];
        let currentRoom = null;

        const ensureRoom = () => {
          if (
            !currentRoom ||
            currentRoom.students.length >= studentsPerRoomCapacity
          ) {
            currentRoom = {
              students: [],
              courseCodes: new Set(),
              courseTitles: new Set(),
              crns: new Set(),
              instructors: new Set(),
            };
            rooms.push(currentRoom);
          }
        };

        roomInputCourses.forEach(({ course, students }) => {
          const remaining = [...students];

          while (remaining.length) {
            ensureRoom();

            const availableSpots =
              studentsPerRoomCapacity - currentRoom.students.length;
            const portion = remaining.splice(0, availableSpots);

            portion.forEach((studentEntry) => {
              currentRoom.students.push(studentEntry);
              if (studentEntry.courseCode) {
                currentRoom.courseCodes.add(studentEntry.courseCode);
              }
              if (studentEntry.courseTitle) {
                currentRoom.courseTitles.add(studentEntry.courseTitle);
              }
              if (studentEntry.crn) {
                currentRoom.crns.add(studentEntry.crn);
              }
              const instructorValue =
                studentEntry.instructor ||
                course.primaryInstructor ||
                (Array.isArray(course.instructors)
                  ? course.instructors.find((name) => name) || ""
                  : "");
              if (instructorValue) {
                currentRoom.instructors.add(instructorValue);
              }
            });
          }
        });

        rooms.forEach((room, roomIndex) => {
          const roomName = `Room ${roomIndex + 1}`;
          const crnLabel = Array.from(room.crns).sort().join(", ");
          const courseCodeLabel = Array.from(room.courseCodes)
            .sort()
            .join(", ");
          const courseTitleLabel = Array.from(room.courseTitles)
            .sort()
            .join("; ");
          const instructorLabel = Array.from(room.instructors)
            .sort()
            .join(", ");

          invigilatorRows.push([
            crnLabel || courseCodeLabel || "",
            courseCodeLabel,
            courseTitleLabel,
            String(room.students.length),
            formattedDate,
            timeRange,
            instructorLabel,
            roomName,
            "",
            "",
            "",
          ]);

          if (!dayRowsMap.has(day)) {
            dayRowsMap.set(day, []);
          }

          const rowsForDay = dayRowsMap.get(day);
          room.students.forEach((studentEntry) => {
            rowsForDay.push({
              sortKey: timeSortKey,
              roomName,
              courseCode: studentEntry.courseCode || "",
              studentId: studentEntry.id,
              row: [
                studentEntry.crn || "",
                studentEntry.courseCode || "",
                studentEntry.courseTitle
                  ? `${studentEntry.courseTitle} (${timeRange})`
                  : timeRange,
                studentEntry.id,
                studentEntry.name,
                roomName,
                "",
              ],
            });
          });
        });
      });
    });

    if (!invigilatorRows.length) {
      return null;
    }

    const placeholders = generateInvigilatorPlaceholders(invigilatorPlaceholderTotal);
    const roomNamesForRows = invigilatorRows.map((row) => row[7] || "");
    const { assignments: invigilatorAssignments } = assignInvigilatorsToRows(
      invigilatorRows.length,
      placeholders,
      roomNamesForRows,
    );

    const poolSheetName = normaliseSheetName(`Week ${week} Invigilator Pool`);
    const placeholderRowMap = new Map();
    placeholders.forEach((name, index) => {
      placeholderRowMap.set(name, index + 2);
    });

    const workbook = XLSX.utils.book_new();

    const invSheetData = [
      [...templateHeaders.invigilatorHeader],
      ...invigilatorRows.map((row) =>
        row.map((value) =>
          value === undefined || value === null ? "" : value,
        ),
      ),
    ];
    const invSheet = XLSX.utils.aoa_to_sheet(invSheetData);

    const invSheetName = normaliseSheetName(`Week ${week} Invigilators`);
    const roomPoolSheetName = normaliseSheetName(`Week ${week} Room Pool`);

    const uniqueRoomNames = Array.from(
      new Set(roomNamesForRows.filter((room) => room && room.trim())),
    ).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
    const roomRowMap = new Map();
    uniqueRoomNames.forEach((roomName, index) => {
      roomRowMap.set(roomName, index + 2);
    });

    invigilatorAssignments.forEach((assignment, index) => {
      const sheetRowIndex = index + 1;

      const setFormula = (columnIndex, name) => {
        if (!name) {
          return;
        }

        const rowNumber = placeholderRowMap.get(name);
        if (!rowNumber) {
          return;
        }

        const cellAddress = XLSX.utils.encode_cell({
          c: columnIndex,
          r: sheetRowIndex,
        });

        invSheet[cellAddress] = {
          t: "s",
          v: name,
          f: `'${poolSheetName}'!A${rowNumber}`,
        };
      };

      setFormula(8, assignment.primaryOne);
      setFormula(9, assignment.primaryTwo);
      setFormula(10, assignment.backup);

      const roomName = assignment.roomName;
      if (roomName) {
        const roomRowNumber = roomRowMap.get(roomName);
        if (roomRowNumber) {
          const roomCellAddress = XLSX.utils.encode_cell({
            c: 7,
            r: sheetRowIndex,
          });
          invSheet[roomCellAddress] = {
            t: "s",
            v: roomName,
            f: `='${roomPoolSheetName}'!A${roomRowNumber}`,
          };
        }
      }
    });

    XLSX.utils.book_append_sheet(workbook, invSheet, invSheetName);

    const primaryColumnLetterOne = XLSX.utils.encode_col(8);
    const primaryColumnLetterTwo = XLSX.utils.encode_col(9);
    const backupColumnLetter = XLSX.utils.encode_col(10);

    const placeholderSheetData = [
      ["Invigilator", "Primary assignments", "Backup assignments"],
      ...placeholders.map((name, index) => {
        const rowNumber = index + 2;
        const primaryFormula = `=COUNTIF('${invSheetName}'!$${primaryColumnLetterOne}:$${primaryColumnLetterOne},A${rowNumber})+COUNTIF('${invSheetName}'!$${primaryColumnLetterTwo}:$${primaryColumnLetterTwo},A${rowNumber})`;
        const backupFormula = `=COUNTIF('${invSheetName}'!$${backupColumnLetter}:$${backupColumnLetter},A${rowNumber})`;

        return [name, { f: primaryFormula }, { f: backupFormula }];
      }),
    ];
    const placeholderSheet = XLSX.utils.aoa_to_sheet(placeholderSheetData);

    XLSX.utils.book_append_sheet(workbook, placeholderSheet, poolSheetName);

    const roomColumnLetter = XLSX.utils.encode_col(7);

    const roomPoolData = [
      ["Room", "Assignments"],
      ...uniqueRoomNames.map((roomName, index) => {
        const rowNumber = index + 2;
        const formula = `=COUNTIF('${invSheetName}'!$${roomColumnLetter}:$${roomColumnLetter},A${rowNumber})`;

        return [roomName, { f: formula }];
      }),
    ];
    const roomPoolSheet = XLSX.utils.aoa_to_sheet(roomPoolData);

    XLSX.utils.book_append_sheet(workbook, roomPoolSheet, roomPoolSheetName);

    dayRowsMap.forEach((rows, day) => {
      const sortedRows = rows
        .sort((a, b) => {
          if (a.sortKey !== b.sortKey) {
            return a.sortKey - b.sortKey;
          }

          if (a.roomName !== b.roomName) {
            return a.roomName.localeCompare(b.roomName);
          }

          if (a.courseCode !== b.courseCode) {
            return a.courseCode.localeCompare(b.courseCode);
          }

          return a.studentId.localeCompare(b.studentId);
        })
        .map((entry) =>
          entry.row.map((value) =>
            value === undefined || value === null ? "" : value,
          ),
        );

      const sheetData = [[...templateHeaders.studentHeader], ...sortedRows];
      const sheet = XLSX.utils.aoa_to_sheet(sheetData);

      sortedRows.forEach((row, rowIndex) => {
        const roomName = row[5];
        if (!roomName) {
          return;
        }

        const roomRowNumber = roomRowMap.get(roomName);
        if (!roomRowNumber) {
          return;
        }

        const cellAddress = XLSX.utils.encode_cell({ c: 5, r: rowIndex + 1 });
        sheet[cellAddress] = {
          t: "s",
          v: roomName,
          f: `='${roomPoolSheetName}'!A${roomRowNumber}`,
        };
      });

      XLSX.utils.book_append_sheet(
        workbook,
        sheet,
        normaliseSheetName(`Week ${week} ${day}`),
      );
    });

    return workbook;
  };
  const handleExportSchedule = async () => {
    setExportError("");

    if (!summary.totalCourses) {
      setExportError("No scheduled exams available to export.");

      return;
    }

    setIsExporting(true);

    try {
      const templateHeaders = await getTemplateHeaders();
      const parsedStartDate = parseISODateString(startDate);
      const baseStartDate = alignDateToMonday(
        parsedStartDate ?? getDefaultStartDate(),
      );

      const exportedFiles = [];

      for (const week of weeks) {
        const workbook = buildWorkbookForWeek(
          week,
          templateHeaders,
          baseStartDate,
        );

        if (!workbook) {
          continue;
        }

        const arrayBuffer = XLSX.write(workbook, {
          bookType: "xlsx",
          type: "array",
          compression: true,
        });

        exportedFiles.push({
          filename: `Week_${week}_Exam_Schedule.xlsx`,
          arrayBuffer,
        });
      }

      if (!exportedFiles.length) {
        setExportError("No scheduled exams available to export.");
        return;
      }

      if (exportedFiles.length === 1) {
        const { filename, arrayBuffer } = exportedFiles[0];
        const blob = new Blob([arrayBuffer], { type: XLSX_MIME_TYPE });
        downloadBlob(blob, filename);
        return;
      }

      const zip = new JSZip();

      exportedFiles.forEach(({ filename, arrayBuffer }) => {
        zip.file(filename, arrayBuffer);
      });

      const zipBlob = await zip.generateAsync({
        type: "blob",
        compression: "DEFLATE",
        compressionOptions: { level: 9 },
      });

      const now = new Date();
      const datePart = now.toISOString().slice(0, 10);
      const zipFilename = `Exam_Schedules_${datePart}.zip`;

      downloadBlob(zipBlob, zipFilename);
    } catch (error) {
      console.error("Failed to export schedule", error);

      setExportError("Failed to export the schedule. Please try again.");
    } finally {
      setIsExporting(false);
    }
  };

  const addWeek = () => {
    setWeeks((previousWeeks) => {
      if (previousWeeks.length >= MAX_WEEKS) {
        return previousWeeks;
      }

      const nextWeekNumber = previousWeeks.length
        ? Math.max(...previousWeeks) + 1
        : 1;

      if (previousWeeks.includes(nextWeekNumber)) {
        return previousWeeks;
      }

      setAssignments((previousAssignments) => {
        if (previousAssignments?.[nextWeekNumber]) {
          return previousAssignments;
        }

        return {
          ...previousAssignments,

          [nextWeekNumber]: createEmptyDaySlotMap(timeSlots),
        };
      });

      setSelectedWeek(nextWeekNumber);

      return [...previousWeeks, nextWeekNumber];
    });
  };

  const resetSchedule = () => {
    setAssignments(buildEmptyAssignments(weeks, timeSlots));

    setSelectedWeek(weeks[0] ?? 1);
  };

  const renderWeekTabs = (position) => (
    <div className={`week-tabs week-tabs--${position}`}>
      <div className="week-tabs__list">
        {weeks.map((week) => (
          <button
            key={week}
            type="button"
            className={week === selectedWeek ? "is-active" : ""}
            onClick={() => setSelectedWeek(week)}
          >
            Week {week}
          </button>
        ))}
      </div>

      {position === "top" ? (
        <button
          type="button"
          className="week-tabs__add"
          onClick={addWeek}
          disabled={weeks.length >= MAX_WEEKS}
        >
          + Add Week
        </button>
      ) : null}
    </div>
  );

  const totalUniqueStudentsAcrossCourses = useMemo(() => {
    const ids = new Set();

    courses.forEach((course) => {
      course.students.forEach((student) => ids.add(student.id));
    });

    return ids.size;
  }, [courses]);

  return (
    <div className="app">
      <header className="app__header">
        <div>
          <h1>Exam Scheduling Helper</h1>

          <p>
            Upload student enrolment file to start building the exam timetable.
          </p>
        </div>

        <div className="app__actions">
          <div className="start-date-control">
            <label htmlFor="start-date-input">Select exam start date:</label>
            <input
              id="start-date-input"
              type="date"
              value={startDate}
              onChange={handleStartDateChange}
            />
          </div>

          <label className="file-input">
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleFileUpload}
            />

            <span>Select .xlsx or .csv</span>
          </label>

          <button
            type="button"
            onClick={handleExportSchedule}
            disabled={isExporting || !summary.totalCourses}
          >
            {isExporting ? "Exporting..." : "Export Timetable"}
          </button>

          <button
            type="button"
            onClick={resetSchedule}
            disabled={!courses.length || isExporting}
          >
            Clear Timetable
          </button>

          <details className="settings-panel">
            <summary>Settings</summary>

            <div className="settings-panel__grid">
              <label>
                <span>Slot interval</span>
                <select
                  value={slotIntervalMinutes}
                  onChange={handleNumericSettingChange(
                    "slotIntervalMinutes",
                    { min: 30, max: 60 },
                  )}
                >
                  <option value={30}>30 minutes</option>
                  <option value={60}>1 hour</option>
                </select>
              </label>

              <label>
                <span>Start hour</span>
                <input
                  type="number"
                  min="0"
                  max="22"
                  value={startHour}
                  onChange={handleNumericSettingChange(
                    "startHour",
                    { min: 0, max: 22 },
                  )}
                />
              </label>

              <label>
                <span>End hour</span>
                <input
                  type="number"
                  min="1"
                  max="23"
                  value={endHour}
                  onChange={handleNumericSettingChange(
                    "endHour",
                    { min: 1, max: 23 },
                  )}
                />
              </label>

              <label>
                <span>#Students per room</span>
                <input
                  type="number"
                  min="1"
                  max="500"
                  value={studentsPerRoom}
                  onChange={handleNumericSettingChange(
                    "studentsPerRoom",
                    { min: 1, max: 500 },
                  )}
                />
              </label>

              <label>
                <span>#Invigilators</span>
                <input
                  type="number"
                  min="1"
                  max="200"
                  value={invigilatorPlaceholderCount}
                  onChange={handleNumericSettingChange(
                    "invigilatorPlaceholderCount",
                    { min: 1, max: 200 },
                  )}
                />
              </label>
            </div>
          </details>
        </div>
      </header>

      {uploadError ? (
        <div className="alert alert--error">{uploadError}</div>
      ) : null}

      {exportError ? (
        <div className="alert alert--error">{exportError}</div>
      ) : null}

      {courses.length ? (
        <section className="overview">
          <div>
            <strong>#Courses:</strong> {courses.length}
          </div>

          <div>
            <strong>#Students:</strong> {totalUniqueStudentsAcrossCourses}
          </div>
        </section>
      ) : (
        <section className="placeholder">
          No courses loaded yet. Upload a .xlsx or .csv file with student
          courses.
        </section>
      )}

      {courses.length ? (
        <>
          <section className="conflicts conflicts--full">
            <h2>Conflicts</h2>

            {conflicts.overall.length ? (
              <ul>
                {conflicts.overall.map((message) => (
                  <li key={message}>{message}</li>
                ))}
              </ul>
            ) : (
              <p>No conflicts detected. Great job!</p>
            )}
          </section>

          <div className="layout">
            <aside className="course-list">
              <h2>Course Pool</h2>

              <p>Drag a course into a timetable slot to schedule its exam.</p>

              <div className="course-search">
                <input
                  type="search"
                  value={courseSearch}
                  onChange={(event) => setCourseSearch(event.target.value)}
                  placeholder="Search by code or title"
                  aria-label="Search courses"
                />

                {courseSearch ? (
                  <button
                    type="button"
                    onClick={() => setCourseSearch("")}
                    aria-label="Clear course search"
                  >
                    Clear
                  </button>
                ) : null}
              </div>

              <ul>
                {filteredCourses.length ? (
                  filteredCourses.map((course) => (
                    <li
                      key={course.id}
                      draggable
                      onDragStart={(event) => {
                        event.dataTransfer.setData("text/plain", course.id);

                        event.dataTransfer.effectAllowed = "move";
                      }}
                      onDragEnd={() => setHoverTarget(null)}
                    >
                      <div className="course-code">{course.code}</div>

                      <div className="course-title">{course.title}</div>

                      <div className="course-meta">
                        {course.studentCount} student
                        {course.studentCount === 1 ? "" : "s"}
                      </div>

                      {course.crns && course.crns.length ? (
                        <div className="course-meta course-meta--secondary">
                          CRNs: {course.crns.join(", ")}
                        </div>
                      ) : null}
                    </li>
                  ))
                ) : (
                  <li className="course-list__empty">
                    No available courses match your search.
                  </li>
                )}
              </ul>
            </aside>

            <main className="scheduler">
              {renderWeekTabs("top")}

              <section className="timetable">
                <table>
                  <thead>
                    <tr>
                      <th>Day / Time</th>

                      {timeSlots.map((slot) => {
                        const slotIsOccupied = occupiedSlotIds.has(slot.id);

                        const headerClassName = [
                          "slot-column",
                          slotIsOccupied
                            ? "slot-column--occupied"
                            : "slot-column--empty",
                        ].join(" ");

                        return (
                          <th key={slot.id} className={headerClassName}>
                            {slot.label}
                          </th>
                        );
                      })}
                    </tr>
                  </thead>

                  <tbody>
                    {days.map((day) => (
                      <tr key={day}>
                        <th scope="row">{day}</th>

                        {timeSlots.map((slot, slotIndex) => {
                          const weekAssignments =
                            assignments[selectedWeek] || {};

                          const dayAssignments = weekAssignments[day] || {};

                          const slotCourses = dayAssignments[slot.id] || [];

                          const previousSlotId =
                            slotIndex > 0 ? timeSlots[slotIndex - 1].id : null;

                          const trailingCourses = previousSlotId
                            ? dayAssignments[previousSlotId] || []
                            : [];

                          const trailingOnlyCourses = trailingCourses.filter(
                            (courseId) => !slotCourses.includes(courseId),
                          );

                          const hasAnyCourses =
                            slotCourses.length > 0 ||
                            trailingOnlyCourses.length > 0;

                          const slotIsOccupied = occupiedSlotIds.has(slot.id);

                          const conflictMessages =
                            conflicts.byWeek?.[selectedWeek]?.[day]?.[
                              slot.id
                            ] ?? [];

                          const {
                            studentCount,
                            roomCount,
                            invigilatorCount,
                            isStartSlot,
                          } = slotSummaries[day][slot.id];

                          const cellClassNames = [
                            "slot-column",

                            slotIsOccupied
                              ? "slot-column--occupied"
                              : "slot-column--empty",
                          ];

                          if (hoverTarget && hoverTarget.day === day) {
                            const isHoverStart =
                              hoverTarget.slotIndex === slotIndex;

                            const isHoverContinuation =
                              hoverTarget.slotIndex < timeSlots.length - 1 &&
                              hoverTarget.slotIndex + 1 === slotIndex;

                            if (isHoverStart || isHoverContinuation) {
                              cellClassNames.push("is-hovered");
                            }
                          }

                          if (conflictMessages.length) {
                            cellClassNames.push("has-conflict");
                          }

                          return (
                            <td
                              key={slot.id}
                              data-day={day}
                              data-slot-index={slotIndex}
                              onDragOver={(event) =>
                                handleDragOverSlot(event, day, slotIndex)
                              }
                              onDragEnter={(event) =>
                                handleDragEnterSlot(event, day, slotIndex)
                              }
                              onDragLeave={(event) =>
                                handleDragLeaveSlot(event, day, slotIndex)
                              }
                              onDrop={(event) =>
                                handleDrop(day, slot.id, slotIndex, event)
                              }
                              className={cellClassNames.join(" ")}
                              title={conflictMessages.join("\n")}
                            >
                              <div className="slot-content">
                                <div className="slot-summary">
                                  {isStartSlot ? (
                                    <>
                                      <span
                                        className="slot-summary__item slot-summary__item--students"
                                        title="Students starting in this slot"
                                      >
                                        <StudentIcon />

                                        {studentCount}
                                      </span>

                                      <span
                                        className="slot-summary__item slot-summary__item--rooms"
                                        title="Rooms needed for this slot"
                                      >
                                        <RoomIcon />

                                        {roomCount}
                                      </span>

                                      <span
                                        className="slot-summary__item slot-summary__item--invigilators"
                                        title="Invigilators needed for this slot"
                                      >
                                        <InvigilatorIcon />

                                        {invigilatorCount}
                                      </span>
                                    </>
                                  ) : hasAnyCourses ? (
                                    <span
                                      className="slot-summary__status"
                                      title="Exam continues from the previous slot"
                                    >
                                      Exam in progress
                                    </span>
                                  ) : (
                                    <span className="slot-summary__empty">
                                      Drop course here
                                    </span>
                                  )}
                                </div>

                                <div className="slot-courses">
                                  {hasAnyCourses ? (
                                    <>
                                      {slotCourses.map((courseId) => {
                                        const course = courseLookup[courseId];

                                        if (!course) return null;

                                        return (
                                          <article
                                            key={courseId}
                                            className="scheduled-course"
                                          >
                                            <header>
                                              <span className="course-code">
                                                {course.code}
                                              </span>

                                              <button
                                                type="button"
                                                onClick={() =>
                                                  handleRemoveCourse(
                                                    day,
                                                    slot.id,
                                                    courseId,
                                                  )
                                                }
                                                aria-label={`Remove ${course.code} from ${day} at ${slot.label}`}
                                              >
                                                X
                                              </button>
                                            </header>

                                            <p>{course.title}</p>

                                            <footer>
                                              <span>
                                                {course.studentCount} students
                                              </span>
                                            </footer>
                                          </article>
                                        );
                                      })}

                                      {trailingOnlyCourses.map((courseId) => {
                                        const course = courseLookup[courseId];

                                        if (!course) return null;

                                        return (
                                          <article
                                            key={`${courseId}-ghost-${slot.id}`}
                                            className="scheduled-course scheduled-course--ghost"
                                          >
                                            <header>
                                              <span className="course-code">
                                                {course.code}
                                              </span>
                                            </header>

                                            <p>{course.title}</p>

                                            <footer>
                                              <span>
                                                {course.studentCount} students
                                              </span>
                                            </footer>
                                          </article>
                                        );
                                      })}
                                    </>
                                  ) : null}
                                </div>
                              </div>
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </section>

              {renderWeekTabs("bottom")}
            </main>
          </div>
        </>
      ) : null}
    </div>
  );
}

export default App;

