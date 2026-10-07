import { useState } from "react";
import { defaultHasExam } from "./department.js";

const clock = (minutes) => `${String(Math.floor(minutes / 60)).padStart(2, "0")}:${String(minutes % 60).padStart(2, "0")}`;

export default function CourseSelection({ courses, selection, onUpload, examChoices, onExamChange, missingCrns }) {
  const [search, setSearch] = useState("");
  const visible = courses.filter((course) => `${course.code} ${course.title} ${course.crns.join(" ")}`.toLowerCase().includes(search.toLowerCase()));
  const examCount = courses.filter((course) => examChoices[course.id] ?? defaultHasExam(course)).length;
  return (
    <section className="course-review">
      <div className="course-review__heading">
        <div>
          <h2>Select Department Exams</h2>
          <p>Upload your department CRN list. Only its first sheet is used for courses, rooms, instructors and teaching availability.</p>
        </div>
        <label className="file-input">
          <input type="file" accept=".xlsx,.xls,.csv" onChange={onUpload} />
          <span>{selection ? "Replace CRN List" : "Upload CRN List"}</span>
        </label>
      </div>
      {selection ? (
        <>
          <div className="course-review__controls">
            <p><strong>First sheet:</strong> {selection.sheetName}. Other sheets are ignored.</p>
            <label>Find a course
              <input type="search" value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Code, title or CRN" />
            </label>
            <p><strong>{examCount}</strong> exams selected / {courses.length} department courses</p>
          </div>
          <p className="course-review__hint">Multiple CRNs: 12:00-13:00 or 17:00-18:00, plus Friday 09:00-10:00 and 10:30-11:30. Single CRN: its lab meeting; if there is no lab, use the common windows. Project, internship and OCT courses default to no exam; change these choices below if needed.</p>
          {missingCrns.length > 0 ? <p className="alert alert--info">{missingCrns.length} listed CRNs have no loaded enrolment: {missingCrns.join(", ")}. Empty courses cannot be scheduled.</p> : null}
          <div className="course-review__table-wrap">
            <table className="course-review__table">
              <thead><tr><th scope="col">Exam?</th><th scope="col">Course</th><th scope="col">CRNs / Students</th><th scope="col">Lab Meetings</th></tr></thead>
              <tbody>
                {visible.map((course) => (
                  <tr key={course.id} className={(examChoices[course.id] ?? defaultHasExam(course)) ? "" : "course-review__excluded"}>
                    <td><input type="checkbox" aria-label={`Exam for ${course.code}`} checked={examChoices[course.id] ?? defaultHasExam(course)} onChange={(event) => onExamChange(course.id, event.target.checked)} /></td>
                    <td><strong>{course.code}</strong><br />{course.title}</td>
                    <td>{course.crns.join(", ")}<br /><small>{course.studentCount} students</small></td>
                    <td>{course.labSessions.length ? course.labSessions.map((lab, index) => (
                      <div key={index}><strong>{lab.crn}</strong>: {lab.days.map((day) => day.slice(0, 3)).join(", ")} {clock(lab.startMinutes)}-{clock(lab.endMinutes)}{lab.room ? ` (${lab.room})` : ""}</div>
                    )) : <small>No lab listed; common windows</small>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {!visible.length ? <p>No courses match your search.</p> : null}
          </div>
        </>
      ) : <p className="course-review__empty">Choose the CRN List workbook, then review exam eligibility before continuing. The first sheet is used regardless of its name.</p>}
    </section>
  );
}
