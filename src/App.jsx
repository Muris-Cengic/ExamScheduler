import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
const MAX_WEEKS = 10
const SLOT_INTERVAL_MINUTES = 30
const START_HOUR = 8
const END_HOUR = 17
const STUDENTS_PER_ROOM = 25

function formatTimeLabel(totalMinutes) {
  const hour24 = Math.floor(totalMinutes / 60)
  const minute = totalMinutes % 60
  const suffix = hour24 >= 12 ? 'PM' : 'AM'
  const hour12 = ((hour24 + 11) % 12) + 1
  const paddedMinute = minute.toString().padStart(2, '0')
  return `${hour12}:${paddedMinute} ${suffix}`
}

const timeSlots = []
for (let minutes = START_HOUR * 60; minutes <= (END_HOUR * 60) - SLOT_INTERVAL_MINUTES; minutes += SLOT_INTERVAL_MINUTES) {
  const hour = Math.floor(minutes / 60)
  const minute = minutes % 60
  const id = `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}`
  timeSlots.push({
    id,
    label: formatTimeLabel(minutes),
  })
}

function StudentIcon() {
  return (
    <svg className="slot-summary__icon" viewBox="0 0 24 24" width="16" height="16" aria-hidden="true" focusable="false">
      <circle cx="8" cy="9" r="3" fill="currentColor" />
      <circle cx="16" cy="9" r="3" fill="currentColor" fillOpacity="0.6" />
      <path d="M4 20c0-3 3.8-5.5 8-5.5s8 2.5 8 5.5v1H4z" fill="currentColor" />
    </svg>
  )
}

function RoomIcon() {
  return (
    <svg className="slot-summary__icon" viewBox="0 0 24 24" width="16" height="16" aria-hidden="true" focusable="false">
      <path d="M4 21V10.2L12 4l8 6.2V21h-5.5v-6.5h-5V21H4z" fill="currentColor" />
      <rect x="11" y="12.5" width="2" height="3.5" fill="currentColor" />
    </svg>
  )
}

function InvigilatorIcon() {
  return (
    <svg className="slot-summary__icon" viewBox="0 0 24 24" width="16" height="16" aria-hidden="true" focusable="false">
      <circle cx="12" cy="7" r="3.5" fill="currentColor" />
      <path d="M6.5 21v-3.2c0-3.5 2.9-6.3 5.5-6.3s5.5 2.8 5.5 6.3V21H6.5z" fill="currentColor" />
      <path d="M11.2 11.6h1.6l0.9 2.3-1.7 2.4-1.7-2.4 0.9-2.3z" fill="#ffffff" />
    </svg>
  )
}

function createEmptyDaySlotMap() {
  const map = {}
  days.forEach((day) => {
    map[day] = {}
    timeSlots.forEach((slot) => {
      map[day][slot.id] = []
    })
  })
  return map
}

function buildEmptyAssignments(weekList) {
  const empty = {}
  weekList.forEach((week) => {
    empty[week] = createEmptyDaySlotMap()
  })
  return empty
}

function cloneAssignments(assignments) {
  const clone = {}

  Object.entries(assignments || {}).forEach(([weekKey, weekAssignments]) => {
    const week = Number(weekKey)
    clone[week] = {}
    days.forEach((day) => {
      clone[week][day] = {}
      timeSlots.forEach((slot) => {
        const sourceList = weekAssignments?.[day]?.[slot.id] ?? []
        clone[week][day][slot.id] = [...sourceList]
      })
    })
  })

  return clone
}

function formatStudentReference(studentId, directory) {
  const rawName = (directory[studentId] || '').trim()
  if (!rawName) {
    return studentId
  }
  const parts = rawName.split(/\s+/).filter(Boolean)
  if (!parts.length) {
    return studentId
  }
  const first = parts[0]
  const last = parts[parts.length - 1]
  if (first === last) {
    return `${studentId} ${first}`
  }
  return `${studentId} ${first} ${last}`
}

function computeSlotSummaries(assignments, courseLookup, week) {
  const weekAssignments = assignments[week] || {}
  const summary = createEmptyDaySlotMap()

  days.forEach((day) => {
    timeSlots.forEach((slot, slotIndex) => {
      const courses = weekAssignments[day]?.[slot.id] ?? []
      const isStartSlot = courses.length > 0

      const seenStudentIds = new Set()

      if (isStartSlot) {
        courses.forEach((courseId) => {
          const course = courseLookup[courseId]
          if (!course) return
          course.students.forEach((student) => {
            seenStudentIds.add(student.id)
          })
        })
      }

      const uniqueStudentCount = seenStudentIds.size
      const roomCount = isStartSlot && uniqueStudentCount > 0 ? Math.ceil(uniqueStudentCount / STUDENTS_PER_ROOM) : 0
      const invigilatorCount = roomCount * 2

      summary[day][slot.id] = {
        studentCount: uniqueStudentCount,
        uniqueStudents: uniqueStudentCount,
        roomCount,
        invigilatorCount,
        isStartSlot,
      }
    })
  })

  return summary
}

function computeConflicts(assignments, courseLookup, studentDirectory) {
  const byWeek = {}
  const overallMessages = new Set()

  const weekKeys = Object.keys(assignments || {}).map((value) => Number(value)).sort((a, b) => a - b)

  weekKeys.forEach((week) => {
    const weekConflicts = createEmptyDaySlotMap()
    byWeek[week] = weekConflicts
    const studentDayCounts = {}
    const weekAssignments = assignments[week] || {}

    days.forEach((day) => {
      timeSlots.forEach((slot, slotIndex) => {
        const startCourses = weekAssignments[day]?.[slot.id] ?? []
        if (!startCourses.length && slotIndex === 0) {
          return
        }

        const slotStudentCourses = new Map()

        const addCourseToSlotMap = (courseId) => {
          const course = courseLookup[courseId]
          if (!course) return
          course.students.forEach((student) => {
            if (!slotStudentCourses.has(student.id)) {
              slotStudentCourses.set(student.id, new Set())
            }
            slotStudentCourses.get(student.id).add(course.code || course.title || courseId)
          })
        }

        startCourses.forEach((courseId) => {
          addCourseToSlotMap(courseId)
          const course = courseLookup[courseId]
          if (!course) return
          course.students.forEach((student) => {
            if (!studentDayCounts[student.id]) {
              studentDayCounts[student.id] = {}
            }
            studentDayCounts[student.id][day] = (studentDayCounts[student.id][day] || 0) + 1
          })
        })

        if (slotIndex > 0) {
          const previousSlotId = timeSlots[slotIndex - 1].id
          const previousCourses = weekAssignments[day]?.[previousSlotId] ?? []
          previousCourses.forEach((courseId) => addCourseToSlotMap(courseId))
        }

        if (slotStudentCourses.size === 0) {
          return
        }

        slotStudentCourses.forEach((courseSet, studentId) => {
          if (courseSet.size > 1) {
            const studentLabel = formatStudentReference(studentId, studentDirectory)
            const conflicts = Array.from(courseSet).join(', ')
            const message = `Week ${week}: Student ${studentLabel} has overlapping exams (${conflicts}) on ${day} at ${slot.label}`
            overallMessages.add(message)
            weekConflicts[day][slot.id] = [...weekConflicts[day][slot.id], message]
          }
        })
      })
    })

    Object.entries(studentDayCounts).forEach(([studentId, counts]) => {
      Object.entries(counts).forEach(([day, total]) => {
        if (total > 2) {
          const studentLabel = formatStudentReference(studentId, studentDirectory)
          const message = `Week ${week}: Student ${studentLabel} is scheduled for ${total} exams on ${day}`
          overallMessages.add(message)

          timeSlots.forEach((slot) => {
            const courses = weekAssignments[day]?.[slot.id] ?? []
            if (courses.length === 0) return
            const involved = courses.some((courseId) => {
              const course = courseLookup[courseId]
              if (!course) return false
              return course.students.some((student) => student.id === studentId)
            })
            if (involved) {
              weekConflicts[day][slot.id] = [...weekConflicts[day][slot.id], message]
            }
          })
        }
      })
    })

    days.forEach((day) => {
      timeSlots.forEach((slot) => {
        if (weekConflicts[day][slot.id].length > 1) {
          weekConflicts[day][slot.id] = Array.from(new Set(weekConflicts[day][slot.id]))
        }
      })
    })
  })

  return {
    overall: Array.from(overallMessages),
    byWeek,
  }
}

function computeSummary(assignments, courseLookup) {
  const scheduledCourseIds = new Set()
  const studentIds = new Set()
  let roomCount = 0

  Object.values(assignments || {}).forEach((weekAssignments) => {
    days.forEach((day) => {
      timeSlots.forEach((slot) => {
        const courseIds = weekAssignments?.[day]?.[slot.id] ?? []
        if (!courseIds.length) {
          return
        }
        const slotStudentIds = new Set()
        courseIds.forEach((courseId) => {
          const course = courseLookup[courseId]
          if (!course) return
          scheduledCourseIds.add(courseId)
          course.students.forEach((student) => {
            studentIds.add(student.id)
            slotStudentIds.add(student.id)
          })
        })
        if (slotStudentIds.size > 0) {
          roomCount += Math.ceil(slotStudentIds.size / STUDENTS_PER_ROOM)
        }
      })
    })
  })

  const invigilators = roomCount * 2

  return {
    totalCourses: scheduledCourseIds.size,
    totalStudents: studentIds.size,
    totalRooms: roomCount,
    totalInvigilators: invigilators,
  }
}

function App() {
  const [weeks, setWeeks] = useState(() => [1])
  const [assignments, setAssignments] = useState(() => buildEmptyAssignments([1]))
  const [selectedWeek, setSelectedWeek] = useState(1)
  const [courses, setCourses] = useState([])
  const [studentDirectory, setStudentDirectory] = useState({})
  const [uploadError, setUploadError] = useState('')
  const [courseSearch, setCourseSearch] = useState('')
  const [hoverTarget, setHoverTarget] = useState(null)

  const courseLookup = useMemo(() => {
    const lookup = {}
    courses.forEach((course) => {
      lookup[course.id] = course
    })
    return lookup
  }, [courses])

  const slotSummaries = useMemo(
    () => computeSlotSummaries(assignments, courseLookup, selectedWeek),
    [assignments, courseLookup, selectedWeek]
  )

  const conflicts = useMemo(
    () => computeConflicts(assignments, courseLookup, studentDirectory),
    [assignments, courseLookup, studentDirectory]
  )

  const occupiedSlotIds = useMemo(() => {
    const result = new Set()
    const weekAssignments = assignments[selectedWeek] || {}

    timeSlots.forEach((slot, index) => {
      const hasCourseInColumn = days.some((day) => {
        const dayAssignments = weekAssignments[day] || {}
        const slotCourses = dayAssignments[slot.id] || []
        if (slotCourses.length > 0) {
          return true
        }
        const previousSlotId = index > 0 ? timeSlots[index - 1].id : null
        if (!previousSlotId) {
          return false
        }
        const trailingCourses = dayAssignments[previousSlotId] || []
        return trailingCourses.some((courseId) => !slotCourses.includes(courseId))
      })

      if (hasCourseInColumn) {
        result.add(slot.id)
      }
    })

    return result
  }, [assignments, selectedWeek])
  const summary = useMemo(
    () => computeSummary(assignments, courseLookup),
    [assignments, courseLookup]
  )

  const assignedCourseIds = useMemo(() => {
    const ids = new Set()

    Object.values(assignments || {}).forEach((weekAssignments) => {
      days.forEach((day) => {
        timeSlots.forEach((slot) => {
          const courseIds = weekAssignments?.[day]?.[slot.id] ?? []
          courseIds.forEach((courseId) => ids.add(courseId))
        })
      })
    })

    return ids
  }, [assignments])

  const orderedCourses = useMemo(() => {
    return [...courses].sort((a, b) => {
      const codeA = a.code.toLowerCase()
      const codeB = b.code.toLowerCase()
      if (codeA === codeB) {
        return a.title.localeCompare(b.title)
      }
      return codeA.localeCompare(codeB)
    })
  }, [courses])

  const availableCourses = useMemo(() => orderedCourses.filter((course) => !assignedCourseIds.has(course.id)), [orderedCourses, assignedCourseIds])

  const filteredCourses = useMemo(() => {
    const query = courseSearch.trim().toLowerCase()
    if (!query) {
      return availableCourses
    }
    return availableCourses.filter((course) => {
      const sectionText = Array.isArray(course.sections) ? course.sections.join(' ') : course.section || ''
      const crnText = Array.isArray(course.crns) ? course.crns.join(' ') : ''
      const haystack = `${course.code} ${course.title} ${sectionText} ${crnText}`.toLowerCase()
      return haystack.includes(query)
    })
  }, [availableCourses, courseSearch])


  const handleFileUpload = async (event) => {
    const file = event.target.files?.[0]
    if (!file) return

    setUploadError('')

    try {
      let workbook
      if (file.name.toLowerCase().endsWith('.csv')) {
        const text = await file.text()
        workbook = XLSX.read(text, { type: 'string' })
      } else {
        const arrayBuffer = await file.arrayBuffer()
        workbook = XLSX.read(arrayBuffer, { type: 'array' })
      }

      const [firstSheetName] = workbook.SheetNames
      const sheet = workbook.Sheets[firstSheetName]
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' })

      if (!rows.length) {
        throw new Error('The selected file does not contain any rows.')
      }

      const coursesMap = new Map()
      const studentNames = {}

      rows.forEach((row) => {
        const studentIdRaw = row.SPRIDEN_ID ?? row.Student_ID ?? row['Student ID'] ?? ''
        const studentId = String(studentIdRaw).trim()
        if (!studentId) return

        const studentNameRaw = row.STUDENT_NAME ?? row.Student_Name ?? row['Student Name'] ?? ''
        const studentName = String(studentNameRaw).trim() || studentId
        studentNames[studentId] = studentName

        const courseCodeRaw =
          row.SCBCRSE_SUBJ_CODE_SCBCRSE_CRSE ??
          (row.SSBSECT_SUBJ_CODE && row.SSBSECT_CRSE_NUMB
            ? `${row.SSBSECT_SUBJ_CODE}-${row.SSBSECT_CRSE_NUMB}`
            : row.Course_Code ?? row['Course Code'] ?? '')
        const courseTitleRaw = row.SCBCRSE_TITLE ?? row.Course_Title ?? row['Course Title'] ?? 'Untitled Course'
        const courseSection = String(row.SSBSECT_SEQ_NUMB ?? row.Section ?? row['Section'])
        const courseCrn = String(row.SSBSECT_CRN ?? row.CRN ?? row['CRN'] ?? '').trim()

        const courseCode = String(courseCodeRaw || 'Course').trim()
        const courseId = courseCrn || `${courseCode}${courseSection ? `-${courseSection}` : ''}` || courseCode

        if (!courseId) return

        const existing = coursesMap.get(courseId) || {
          id: courseId,
          code: courseCode,
          title: String(courseTitleRaw).trim() || courseCode,
          section: courseSection,
          students: [],
          studentSet: new Set(),
        }

        if (!existing.studentSet.has(studentId)) {
          existing.studentSet.add(studentId)
          existing.students.push({ id: studentId, name: studentName })
        }

        existing.code = courseCode || existing.code
        existing.title = String(courseTitleRaw).trim() || existing.title

        coursesMap.set(courseId, existing)
      })

      const parsedCourses = Array.from(coursesMap.values()).map((course) => {
        const studentCount = course.students.length
        const roomsNeeded = studentCount ? Math.max(1, Math.ceil(studentCount / STUDENTS_PER_ROOM)) : 0
        return {
          id: course.id,
          code: course.code,
          title: course.title,
          section: course.section,
          students: course.students,
          studentCount,
          roomsNeeded,
        }
      })

      if (!parsedCourses.length) {
        throw new Error('No courses were found in the provided file.')
      }

      const groupedCoursesMap = new Map()

      parsedCourses.forEach((course) => {
        const key = course.code || course.id
        if (!key) return
        if (!groupedCoursesMap.has(key)) {
          groupedCoursesMap.set(key, {
            id: key,
            code: course.code || key,
            title: course.title,
            sections: new Set(),
            crns: new Set(),
            students: [],
            studentSet: new Set(),
          })
        }

        const groupedCourse = groupedCoursesMap.get(key)

        if (course.section) {
          groupedCourse.sections.add(course.section)
        }
        if (course.id && course.id !== key) {
          groupedCourse.crns.add(course.id)
        }

        course.students.forEach((student) => {
          if (!groupedCourse.studentSet.has(student.id)) {
            groupedCourse.studentSet.add(student.id)
            groupedCourse.students.push(student)
          }
        })
      })

      const groupedCourses = Array.from(groupedCoursesMap.values()).map((group) => {
        const sections = Array.from(group.sections).filter(Boolean).sort()
        const crns = Array.from(group.crns).filter(Boolean).sort()
        const students = group.students.slice()
        students.sort((a, b) => a.name.localeCompare(b.name))
        const studentCount = students.length
        const roomsNeeded = studentCount ? Math.max(1, Math.ceil(studentCount / STUDENTS_PER_ROOM)) : 0

        return {
          id: group.id,
          code: group.code,
          title: group.title,
          sections,
          crns,
          students,
          studentCount,
          roomsNeeded,
        }
      })

      const initialWeeks = [1]
      setWeeks(initialWeeks)
      setCourses(groupedCourses)
      setCourseSearch('')
      setStudentDirectory(studentNames)
      setAssignments(buildEmptyAssignments(initialWeeks))
      setSelectedWeek(initialWeeks[0])
    } catch (error) {
      console.error(error)
      setUploadError(error.message || 'Failed to read the provided file.')
    }
  }

  const updateHoverTarget = (day, slotIndex) => {
    if (slotIndex > timeSlots.length - 2) {
      setHoverTarget(null)
      return
    }
    setHoverTarget((previous) => {
      if (previous && previous.day === day && previous.slotIndex === slotIndex) {
        return previous
      }
      return { day, slotIndex }
    })
  }

  const handleDragEnterSlot = (event, day, slotIndex) => {
    updateHoverTarget(day, slotIndex)
  }

  const handleDragOverSlot = (event, day, slotIndex) => {
    event.preventDefault()
    updateHoverTarget(day, slotIndex)
  }

  const handleDragLeaveSlot = (event, day, slotIndex) => {
    const relatedTarget = event.relatedTarget
    if (relatedTarget) {
      const relatedCell =
        typeof relatedTarget.closest === 'function'
          ? relatedTarget.closest('td[data-slot-index]')
          : null
      if (relatedCell) {
        const relatedDay = relatedCell.getAttribute('data-day')
        const relatedSlotIndex = Number(relatedCell.getAttribute('data-slot-index'))
        if (
          relatedDay === day &&
          (relatedSlotIndex === slotIndex || relatedSlotIndex === slotIndex + 1)
        ) {
          return
        }
      }
    }
    setHoverTarget((previous) => {
      if (!previous) {
        return previous
      }
      if (previous.day === day && previous.slotIndex === slotIndex) {
        return null
      }
      return previous
    })
  }

  const handleDrop = (day, slotId, slotIndex, event) => {
    event.preventDefault()
    setHoverTarget(null)
    const courseId = event.dataTransfer.getData('text/plain')
    if (!courseId || !courseLookup[courseId]) return
    if (slotIndex === undefined || slotIndex > timeSlots.length - 2) return

    setAssignments((previous) => {
      const next = cloneAssignments(previous)

      Object.keys(next).forEach((weekKey) => {
        const weekAssignments = next[weekKey]
        days.forEach((existingDay) => {
          timeSlots.forEach((slot) => {
            const list = weekAssignments[existingDay][slot.id]
            const index = list.indexOf(courseId)
            if (index !== -1) {
              list.splice(index, 1)
            }
          })
        })
      })

      const targetWeek = next[selectedWeek] || createEmptyDaySlotMap()
      if (!next[selectedWeek]) {
        next[selectedWeek] = targetWeek
      }
      const targetList = targetWeek[day][slotId]
      if (!targetList.includes(courseId)) {
        targetList.push(courseId)
      }

      return next
    })
  }

  const handleRemoveCourse = (day, slotId, courseId) => {
    setAssignments((previous) => {
      const next = cloneAssignments(previous)
      next[selectedWeek][day][slotId] = next[selectedWeek][day][slotId].filter((id) => id !== courseId)
      return next
    })
  }

  const addWeek = () => {
    setWeeks((previousWeeks) => {
      if (previousWeeks.length >= MAX_WEEKS) {
        return previousWeeks
      }

      const nextWeekNumber = previousWeeks.length ? Math.max(...previousWeeks) + 1 : 1
      if (previousWeeks.includes(nextWeekNumber)) {
        return previousWeeks
      }

      setAssignments((previousAssignments) => {
        if (previousAssignments?.[nextWeekNumber]) {
          return previousAssignments
        }
        return {
          ...previousAssignments,
          [nextWeekNumber]: createEmptyDaySlotMap(),
        }
      })

      setSelectedWeek(nextWeekNumber)
      return [...previousWeeks, nextWeekNumber]
    })
  }

  const resetSchedule = () => {
    setAssignments(buildEmptyAssignments(weeks))
    setSelectedWeek(weeks[0] ?? 1)
  }

  const renderWeekTabs = (position) => (
    <div className={`week-tabs week-tabs--${position}`}>
      <div className="week-tabs__list">
        {weeks.map((week) => (
          <button
            key={week}
            type="button"
            className={week === selectedWeek ? 'is-active' : ''}
            onClick={() => setSelectedWeek(week)}
          >
            Week {week}
          </button>
        ))}
      </div>
      {position === 'top' ? (
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
  )

  const totalUniqueStudentsAcrossCourses = useMemo(() => {
    const ids = new Set()
    courses.forEach((course) => {
      course.students.forEach((student) => ids.add(student.id))
    })
    return ids.size
  }, [courses])

  return (
    <div className="app">
      <header className="app__header">
        <div>
          <h1>Exam Scheduling Helper</h1>
          <p>Upload your student enrolment file to start building the exam timetable.</p>
        </div>
        <div className="app__actions">
          <label className="file-input">
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} />
            <span>Select .xlsx or .csv</span>
          </label>
          <button type="button" onClick={resetSchedule} disabled={!courses.length}>
            Clear Timetable
          </button>
        </div>
      </header>

      {uploadError ? <div className="alert alert--error">{uploadError}</div> : null}

      {courses.length ? (
        <section className="overview">
          <div>
            <strong>Courses loaded:</strong> {courses.length}
          </div>
          <div>
            <strong>Unique students:</strong> {totalUniqueStudentsAcrossCourses}
          </div>
        </section>
      ) : (
        <section className="placeholder">No courses loaded yet. Upload a .xlsx or .csv file with student courses.</section>
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
                  <button type="button" onClick={() => setCourseSearch('')} aria-label="Clear course search">
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
                        event.dataTransfer.setData('text/plain', course.id)
                        event.dataTransfer.effectAllowed = 'move'
                      }}
                      onDragEnd={() => setHoverTarget(null)}
                    >
                      <div className="course-code">{course.code}</div>
                      <div className="course-title">{course.title}</div>
                      <div className="course-meta">
                        {course.studentCount} student{course.studentCount === 1 ? '' : 's'}
                      </div>
                      { course.crns && course.crns.length ? (
                        <div className="course-meta course-meta--secondary">
                          CRNs: {course.crns.join(', ')}
                        </div>
                      ) : null}
                    </li>
                  ))
                ) : (
                  <li className="course-list__empty">No available courses match your search.</li>
                )}
              </ul>
            </aside>

            <main className="scheduler">
              {renderWeekTabs('top')}
              <section className="timetable">
                <table>
                  <thead>
                    <tr>
                      <th>Day / Time</th>
                      {timeSlots.map((slot) => {
                        const slotIsOccupied = occupiedSlotIds.has(slot.id)
                        const headerClassName = ['slot-column', slotIsOccupied ? 'slot-column--occupied' : 'slot-column--empty'].join(' ')
                        return (
                          <th key={slot.id} className={headerClassName}>
                            {slot.label}
                          </th>
                        )
                      })}
                    </tr>
                  </thead>
                  <tbody>
                    {days.map((day) => (
                      <tr key={day}>
                        <th scope="row">{day}</th>
                        {timeSlots.map((slot, slotIndex) => {
                          const weekAssignments = assignments[selectedWeek] || {}
                          const dayAssignments = weekAssignments[day] || {}
                          const slotCourses = dayAssignments[slot.id] || []
                          const previousSlotId = slotIndex > 0 ? timeSlots[slotIndex - 1].id : null
                          const trailingCourses = previousSlotId ? dayAssignments[previousSlotId] || [] : []
                          const trailingOnlyCourses = trailingCourses.filter((courseId) => !slotCourses.includes(courseId))
                          const hasAnyCourses = slotCourses.length > 0 || trailingOnlyCourses.length > 0
                          const slotIsOccupied = occupiedSlotIds.has(slot.id)
                          const conflictMessages = conflicts.byWeek?.[selectedWeek]?.[day]?.[slot.id] ?? []
                          const { studentCount, roomCount, invigilatorCount, isStartSlot } = slotSummaries[day][slot.id]
                          const cellClassNames = [
                            'slot-column',
                            slotIsOccupied ? 'slot-column--occupied' : 'slot-column--empty',
                          ]
                          if (hoverTarget && hoverTarget.day === day) {
                            const isHoverStart = hoverTarget.slotIndex === slotIndex
                            const isHoverContinuation =
                              hoverTarget.slotIndex < timeSlots.length - 1 &&
                              hoverTarget.slotIndex + 1 === slotIndex
                            if (isHoverStart || isHoverContinuation) {
                              cellClassNames.push('is-hovered')
                            }
                          }
                          if (conflictMessages.length) {
                            cellClassNames.push('has-conflict')
                          }
                          return (
                            <td
                              key={slot.id}
                              data-day={day}
                              data-slot-index={slotIndex}
                              onDragOver={(event) => handleDragOverSlot(event, day, slotIndex)}
                              onDragEnter={(event) => handleDragEnterSlot(event, day, slotIndex)}
                              onDragLeave={(event) => handleDragLeaveSlot(event, day, slotIndex)}
                              onDrop={(event) => handleDrop(day, slot.id, slotIndex, event)}
                              className={cellClassNames.join(' ')}
                              title={conflictMessages.join('\n')}
                            >
                              <div className="slot-content">
                                <div className="slot-summary">
                                  {isStartSlot ? (
                                    <>
                                      <span className="slot-summary__item slot-summary__item--students" title="Students starting in this slot">
                                        <StudentIcon />
                                        {studentCount}
                                      </span>
                                      <span className="slot-summary__item slot-summary__item--rooms" title="Rooms needed for this slot">
                                        <RoomIcon />
                                        {roomCount}
                                      </span>
                                      <span className="slot-summary__item slot-summary__item--invigilators" title="Invigilators needed for this slot">
                                        <InvigilatorIcon />
                                        {invigilatorCount}
                                      </span>
                                    </>
                                  ) : hasAnyCourses ? (
                                    <span className="slot-summary__status" title="Exam continues from the previous slot">Exam in progress</span>
                                  ) : (
                                    <span className="slot-summary__empty">Drop course here</span>
                                  )}
                                </div>
                                <div className="slot-courses">
                                  {hasAnyCourses ? (
                                    <>
                                      {slotCourses.map((courseId) => {
                                        const course = courseLookup[courseId]
                                        if (!course) return null
                                        return (
                                          <article key={courseId} className="scheduled-course">
                                            <header>
                                              <span className="course-code">{course.code}</span>
                                              <button
                                                type="button"
                                                onClick={() => handleRemoveCourse(day, slot.id, courseId)}
                                                aria-label={`Remove ${course.code} from ${day} at ${slot.label}`}
                                              >
                                                X
                                              </button>
                                            </header>
                                            <p>{course.title}</p>
                                            <footer>
                                              <span>{course.studentCount} students</span>
                                            </footer>
                                          </article>
                                        )
                                      })}
                                      {trailingOnlyCourses.map((courseId) => {
                                        const course = courseLookup[courseId]
                                        if (!course) return null
                                        return (
                                          <article key={`${courseId}-ghost-${slot.id}`} className="scheduled-course scheduled-course--ghost">
                                            <header>
                                              <span className="course-code">{course.code}</span>
                                            </header>
                                            <p>{course.title}</p>
                                            <footer>
                                              <span>{course.studentCount} students</span>
                                            </footer>
                                          </article>
                                        )
                                      })}
                                    </>
                                  ) : null}
                                </div>
                              </div>
                            </td>
                          )
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </section>

              {renderWeekTabs('bottom')}

              <section className="report">
                <h2>Schedule Report</h2>
                <div className="report-grid">
                  <div>
                    <span>Total courses scheduled</span>
                    <strong>{summary.totalCourses}</strong>
                  </div>
                  <div>
                    <span>Total students scheduled</span>
                    <strong>{summary.totalStudents}</strong>
                  </div>
                  <div>
                    <span>Total rooms needed</span>
                    <strong>{summary.totalRooms}</strong>
                  </div>
                  <div>
                    <span>Invigilators required</span>
                    <strong>{summary.totalInvigilators}</strong>
                  </div>
                </div>
              </section>
            </main>
        </div>
        </>
      ) : null}
    </div>
  )
}

export default App
