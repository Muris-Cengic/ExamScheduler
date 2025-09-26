import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
const weeks = [1, 2, 3, 4, 5]

const timeSlots = Array.from({ length: 9 }, (_, index) => {
  const hour = 8 + index
  const nextHour = hour + 1
  return {
    id: `${String(hour).padStart(2, '0')}:00`,
    label: `${formatHour(hour)} - ${formatHour(nextHour)}`,
  }
})

function formatHour(hour) {
  const displayHour = hour % 12 === 0 ? 12 : hour % 12
  const suffix = hour >= 12 ? 'PM' : 'AM'
  return `${displayHour}:00 ${suffix}`
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

function buildEmptyAssignments() {
  const empty = {}
  weeks.forEach((week) => {
    empty[week] = createEmptyDaySlotMap()
  })
  return empty
}

function cloneAssignments(assignments) {
  const clone = {}
  weeks.forEach((week) => {
    clone[week] = {}
    days.forEach((day) => {
      clone[week][day] = {}
      timeSlots.forEach((slot) => {
        clone[week][day][slot.id] = [...(assignments[week]?.[day]?.[slot.id] ?? [])]
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
    timeSlots.forEach((slot) => {
      const courses = weekAssignments[day]?.[slot.id] ?? []
      let studentCount = 0
      let roomCount = 0
      const seenStudentIds = new Set()

      courses.forEach((courseId) => {
        const course = courseLookup[courseId]
        if (!course) return
        studentCount += course.studentCount
        course.students.forEach((student) => {
          seenStudentIds.add(student.id)
        })
        roomCount += course.roomsNeeded
      })

      const invigilatorCount = roomCount * 2

      summary[day][slot.id] = {
        studentCount,
        uniqueStudents: seenStudentIds.size,
        roomCount,
        invigilatorCount,
      }
    })
  })

  return summary
}
function computeConflicts(assignments, courseLookup, studentDirectory) {
  const byWeek = {}
  const overallMessages = new Set()

  weeks.forEach((week) => {
    const weekConflicts = createEmptyDaySlotMap()
    byWeek[week] = weekConflicts
    const studentDayCounts = {}
    const weekAssignments = assignments[week] || {}

    days.forEach((day) => {
      timeSlots.forEach((slot) => {
        const courses = weekAssignments[day]?.[slot.id] ?? []
        if (courses.length === 0) return

        const slotStudentCourses = new Map()

        courses.forEach((courseId) => {
          const course = courseLookup[courseId]
          if (!course) return
          course.students.forEach((student) => {
            if (!slotStudentCourses.has(student.id)) {
              slotStudentCourses.set(student.id, new Set())
            }
            slotStudentCourses.get(student.id).add(course.code || course.title || courseId)

            if (!studentDayCounts[student.id]) {
              studentDayCounts[student.id] = {}
            }
            studentDayCounts[student.id][day] = (studentDayCounts[student.id][day] || 0) + 1
          })
        })

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



  weeks.forEach((week) => {

    const weekAssignments = assignments[week] || {}

    days.forEach((day) => {

      timeSlots.forEach((slot) => {

        const courseIds = weekAssignments[day]?.[slot.id] ?? []

        courseIds.forEach((courseId) => {

          const course = courseLookup[courseId]

          if (!course) return

          scheduledCourseIds.add(courseId)

          course.students.forEach((student) => {

            studentIds.add(student.id)

          })

          roomCount += course.roomsNeeded

        })

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
  const [courses, setCourses] = useState([])
  const [studentDirectory, setStudentDirectory] = useState({})
  const [assignments, setAssignments] = useState(() => buildEmptyAssignments())
  const [selectedWeek, setSelectedWeek] = useState(weeks[0])
  const [uploadError, setUploadError] = useState('')
  const [courseSearch, setCourseSearch] = useState('')

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

  const summary = useMemo(
    () => computeSummary(assignments, courseLookup),
    [assignments, courseLookup]
  )

  const assignedCourseIds = useMemo(() => {
    const ids = new Set()
    weeks.forEach((week) => {
      const weekAssignments = assignments[week] || {}
      days.forEach((day) => {
        timeSlots.forEach((slot) => {
          const courseIds = weekAssignments[day]?.[slot.id] ?? []
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
        const roomsNeeded = studentCount ? Math.max(1, Math.ceil(studentCount / 25)) : 0
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
        const roomsNeeded = studentCount ? Math.max(1, Math.ceil(studentCount / 25)) : 0

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

      setCourses(groupedCourses)
      setCourseSearch('')
      setStudentDirectory(studentNames)
      setAssignments(buildEmptyAssignments())
      setSelectedWeek(weeks[0])
    } catch (error) {
      console.error(error)
      setUploadError(error.message || 'Failed to read the provided file.')
    }
  }

  const handleDrop = (day, slotId, event) => {
    event.preventDefault()
    const courseId = event.dataTransfer.getData('text/plain')
    if (!courseId || !courseLookup[courseId]) return

    setAssignments((previous) => {
      const next = cloneAssignments(previous)

      weeks.forEach((existingWeek) => {
        days.forEach((existingDay) => {
          timeSlots.forEach((slot) => {
            const list = next[existingWeek][existingDay][slot.id]
            const index = list.indexOf(courseId)
            if (index !== -1) {
              list.splice(index, 1)
            }
          })
        })
      })

      const targetWeek = next[selectedWeek] || createEmptyDaySlotMap()
      next[selectedWeek] = targetWeek
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

  const resetSchedule = () => {
    setAssignments(buildEmptyAssignments())
    setSelectedWeek(weeks[0])
  }

  const renderWeekTabs = (position) => (
    <div className={`week-tabs week-tabs--${position}`}>
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
                    >
                      <div className="course-code">{course.code}</div>
                      <div className="course-title">{course.title}</div>
                      <div className="course-meta">
                        {course.studentCount} student{course.studentCount === 1 ? '' : 's'} | {course.roomsNeeded} room{course.roomsNeeded === 1 ? '' : 's'}
                      </div>
                      {course.sections && course.sections.length ? (
                        <div className="course-meta course-meta--secondary">
                          Sections: {course.sections.join(', ')}
                        </div>
                      ) : null}
                      {!course.sections?.length && course.crns && course.crns.length ? (
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
                      {timeSlots.map((slot) => (
                        <th key={slot.id}>{slot.label}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {days.map((day) => (
                      <tr key={day}>
                        <th scope="row">{day}</th>
                        {timeSlots.map((slot) => {
                          const weekAssignments = assignments[selectedWeek] || {}
                          const dayAssignments = weekAssignments[day] || {}
                          const slotCourses = dayAssignments[slot.id] || []
                          const conflictMessages = conflicts.byWeek?.[selectedWeek]?.[day]?.[slot.id] ?? []
                          const { studentCount, roomCount, invigilatorCount } = slotSummaries[day][slot.id]
                          const slotBadgeTitle = `Week ${selectedWeek}: ${studentCount} students | ${roomCount} rooms | ${invigilatorCount} invigilators`

                          return (
                            <td
                              key={slot.id}
                              onDragOver={(event) => event.preventDefault()}
                              onDrop={(event) => handleDrop(day, slot.id, event)}
                              className={conflictMessages.length ? 'has-conflict' : ''}
                              title={conflictMessages.join('\n')}
                            >
                              <div className="slot-content">
                                <div className="slot-badge" title={slotBadgeTitle}>
                                  <span className="slot-badge__item" aria-label={`${studentCount} students`}>
                                    <span className="slot-badge__icon slot-badge__icon--students" aria-hidden="true" />
                                    <span className="slot-badge__value">{studentCount}</span>
                                  </span>
                                  <span className="slot-badge__item" aria-label={`${roomCount} rooms`}>
                                    <span className="slot-badge__icon slot-badge__icon--rooms" aria-hidden="true" />
                                    <span className="slot-badge__value">{roomCount}</span>
                                  </span>
                                  <span className="slot-badge__item" aria-label={`${invigilatorCount} invigilators`}>
                                    <span className="slot-badge__icon slot-badge__icon--invigilators" aria-hidden="true" />
                                    <span className="slot-badge__value">{invigilatorCount}</span>
                                  </span>
                                </div>
                                <div className="slot-courses">
                                  {slotCourses.length ? (
                                    slotCourses.map((courseId) => {
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
                                            <span>{course.roomsNeeded} room{course.roomsNeeded === 1 ? '' : 's'}</span>
                                          </footer>
                                        </article>
                                      )
                                    })
                                  ) : (
                                    <p className="slot-placeholder">Drop course here</p>
                                  )}
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
