import { useMemo, useState } from "react";
import { buildCourseSeating, buildEnvelopeCourses, buildExportModel, buildStaffDistribution, buildStaffSchedules, buildStaffSessionRosters, courseSeatingInfo, envelopeCoverRows, envelopeSessionInfo, formatChronologicalDate, formatReportDate, reportDate, reportFileName, reportTable, REPORT_DAYS, REPORT_VIEWS, reportTimeRange, staffScheduleSlots, staffSlotEntries } from "./exportReports.js";
import { clock } from "./resources.js";
import "./ExportStudio.css";

function ReportIcon({ kind }) {
  const paths = {
    complete: "M5 3h10l4 4v14H5V3Zm10 0v5h4M8 12h8M8 16h8",
    overview: "M3 5h18v16H3V5Zm0 5h18M8 3v4M16 3v4M8 10v11M16 10v11",
    staff: "M16 21v-3a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v3M21 21v-3a4 4 0 0 0-3-4M18 3a4 4 0 0 1 0 8",
    students: "M4 3h16v18H4V3ZM8 7h8M8 11h8M8 15h4",
    seating: "M3 4h18v16H3V4ZM3 9h18M11 9v11M6 13h2M6 16h2M14 13h4M14 16h4",
    envelopes: "M3 5h18v14H3V5Zm0 0 9 7 9-7M3 19l6-6m12 6-6-6",
  };
  return <svg viewBox="0 0 24 24" width="24" height="24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <path d={paths[kind]} />{kind === "staff" ? <circle cx="9" cy="7" r="4" /> : null}
  </svg>;
}

function PreviewTable({ columns, rows, rowClassNames = [], limit = Infinity, label, className = "" }) {
  return <div className="export-table-wrap" tabIndex="0" aria-label={label}>
    <table className={"export-table" + (className ? " " + className : "")}>
      <caption className="sr-only">{label}</caption>
      <thead><tr>{columns.map((column) => <th key={column} scope="col">{column}</th>)}</tr></thead>
      <tbody>{rows.slice(0, limit).map((row, index) => <tr key={index} className={rowClassNames[index]}>{row.map((cell, column) => <td key={column}>{cell}</td>)}</tr>)}</tbody>
    </table>
    {!rows.length ? <p className="export-empty">No matching records in this preview.</p> : null}
  </div>;
}

const ledgerRows = (model) => model.roomRows.map((room) => [
  <span className="export-course"><strong>{room.code}</strong><small>{room.title}</small></span>,
  room.day + " / " + reportTimeRange(room.start, room.end), room.roomName, room.studentCount,
  room.invigilators.join(" / "), room.distributionChoice === "standard" ? "Up to " + room.maxStudents : <span className="export-exception">Approved up to 27</span>,
]);
const staffRows = (model, overall = false) => model.duties.map((duty) => [duty.name,
  (overall ? "Week " + duty.week + " / " + formatReportDate(duty.date) : duty.day) + " / " + reportTimeRange(duty.start, duty.end),
  <span className="export-duty">{duty.role}</span>,
  duty.code, duty.roomName, duty.teaching ? <span className="export-teaching">Teaching hours / 0 load</span> : "1 extra duty"]);
const studentRows = (model) => model.students.map((student) => [student.id, student.name, student.code, student.crn,
  student.day + " / " + reportTimeRange(student.start, student.end), student.roomName]);
const loadRows = (model, includeIdle = false) => model.workloads.filter((person) => includeIdle || person.exam + person.teaching > 0)
  .map((person) => [person.name, person.exam, person.teaching]);
const overviewColumns = ["#", "Date", "Day", "Time", "Course", "Course title"];
const overviewRows = (model) => reportTable("overview", model, "table").rows
  .map((row) => row.map((value, column) => column === 1 ? formatChronologicalDate(value) : value));
const overviewRowClasses = (model) => model.overviewExams.map((exam, index) => [
  exam.isAsd ? "export-row--asd" : "",
  index > 0 && exam.date !== model.overviewExams[index - 1].date ? "export-row--new-day" : "",
].filter(Boolean).join(" "));
const overviewNote = "Primary invigilators are exam-room duties, including the lab instructor; slot backups are excluded. Lab time means this exam replaces its listed lab session.";
const asdReferenceNote = "ASD exams are reference only; their students, rooms and invigilators are not managed or counted here.";

function ExamCard({ exam }) {
  return <article className={"export-board__exam" + (exam.isAsd ? " export-board__exam--asd" : "")}>
    <span className="export-board__time">{reportTimeRange(exam.start, exam.end)}</span>
    <h4>{exam.code}</h4><p>{exam.title}</p>
    {exam.isAsd ? <span className="export-asd">ASD / reference</span> : <>
      <small>{exam.studentCount} students / {exam.roomNames.length} {exam.roomNames.length === 1 ? "room" : "rooms"}</small>
      <small className="export-board__staffing">{exam.primaryInvigilatorsNeeded} primary {exam.primaryInvigilatorsNeeded === 1 ? "invigilator" : "invigilators"} needed</small>
      {exam.duringLab ? <span className="export-lab">During lab time</span> : null}
      <ul>{exam.roomNames.map((room) => <li key={room}>{room}</li>)}</ul>
    </>}
  </article>;
}

function WeekBoard({ model, startDate, week }) {
  return <div className="export-board">
    {REPORT_DAYS.map((day) => <section key={day} className="export-board__day">
      <header><strong>{day.slice(0, 3)}</strong><span>{formatReportDate(reportDate(startDate, week, day))}</span></header>
      {model.overviewExams.filter((exam) => exam.day === day).map((exam) => <ExamCard key={exam.id} exam={exam} />)}
      {!model.overviewExams.some((exam) => exam.day === day) ? <p className="export-board__free">No exams</p> : null}
    </section>)}
  </div>;
}

function PrintWeekBoard({ model, startDate, week }) {
  const columns = REPORT_DAYS.map((day) => model.overviewExams.filter((exam) => exam.day === day));
  const rowCount = Math.max(1, ...columns.map((exams) => exams.length));
  // Table rows paginate between cards and repeat the weekday headers on long boards.
  return <table className="export-board-print">
    <caption className="sr-only">{"Week " + week + " exam board"}</caption>
    <thead><tr>{REPORT_DAYS.map((day) => <th key={day} scope="col"><header>
      <strong>{day}</strong><span>{formatReportDate(reportDate(startDate, week, day))}</span>
    </header></th>)}</tr></thead>
    <tbody>{Array.from({ length: rowCount }, (_, index) => <tr key={index}>
      {columns.map((exams, column) => <td key={REPORT_DAYS[column]}>
        {exams[index] ? <ExamCard exam={exams[index]} /> : index === 0 && !exams.length ? <p className="export-board__free">No exams</p> : null}
      </td>)}
    </tr>)}</tbody>
  </table>;
}

function WorkbookMap({ model, week, packaging }) {
  const combined = packaging === "combined";
  const sheets = [
    ...(combined ? [["Schedule Index", model.selectedWeeks.length + " included weeks"]] : []),
    [combined ? "Invigilator Duties" : "Week " + week + " Invigilators", model.roomRows.length + " room assignments"],
    [(combined ? "" : "Week " + week + " ") + "Invigilator Pool", model.workloads.length + " named invigilators"],
    [(combined ? "" : "Week " + week + " ") + "Room Pool", "Confirmed room references"],
    ...model.selectedWeeks.flatMap((currentWeek) => REPORT_DAYS.filter((day) => model.sessions.some((session) => session.week === currentWeek && session.day === day))
      .map((day) => ["Week " + currentWeek + " " + day, model.students.filter((student) => student.week === currentWeek && student.day === day).length + " student sittings"])),
  ];
  return <div className="export-sheet-map" aria-label="Workbook sheet preview">
    {sheets.map(([name, detail], index) => <div key={name}><span>{String(index + 1).padStart(2, "0")}</span>
      <strong>{name}</strong><small>{detail}</small></div>)}
    <p>{combined ? "Duties and pools cover all included weeks once. Student sheets stay separate by week and day." : "Each weekly workbook has its own duties, pools and daily student sheets."} Room and invigilator references remain linked inside Excel.</p>
  </div>;
}

function SeatingSheet({ exam, sheet, rows = sheet.rows, limit, label }) {
  return <article className="export-seating-sheet">
    <header><span className="export-kicker">Student seating</span><h4>{exam.code}</h4></header>
    <dl className="export-seating-info">{courseSeatingInfo(exam).map(([name, value]) => <div key={name}>
      <dt>{name}</dt><dd>{name === "Date" ? formatReportDate(value, true) : value}</dd>
    </div>)}</dl>
    <PreviewTable label={label} columns={["StudentID", "Room"]} rows={rows} limit={limit} />
  </article>;
}

function SeatingPreview({ exams, exam, sheet, onExamChange, onCrnChange, search, onSearch, limit, onShowMore, onDownload, disabled }) {
  const rows = sheet.rows.filter((row) => row.join(" ").toLowerCase().includes(search.toLowerCase()));
  return <>
    <div className="export-seating-controls">
      <label><span>Exam / course</span><select aria-label="Seating exam" value={exam.id} disabled={disabled}
        onChange={(event) => onExamChange(event.target.value)}>
        {exams.map((item) => <option key={item.id} value={item.id}>{item.code}: {item.title} / {item.day} {reportTimeRange(item.start, item.end)}</option>)}
      </select></label>
      <button type="button" disabled={disabled} onClick={onDownload}>Download This Exam</button>
      <div className="export-seating-crns" aria-label="CRN sheets">{exam.crnSheets.map((item) => <button type="button" key={item.crn}
        aria-pressed={sheet.crn === item.crn} onClick={() => onCrnChange(item.crn)}>
        CRN {item.crn || "Unspecified"} <small>{item.rows.length} students</small>
      </button>)}</div>
    </div>
    <label className="export-search"><span>Search this CRN</span><input type="search" value={search} placeholder="Student ID or room..."
      onChange={(event) => onSearch(event.target.value)} /><small>Preview only. Downloads include every CRN for the exam; print includes all selected exams.</small></label>
    <SeatingSheet exam={exam} sheet={sheet} rows={rows} limit={limit} label="Course seating preview" />
    {rows.length > limit ? <button type="button" className="export-show-more" onClick={onShowMore}>Show 50 more ({rows.length - limit} remaining)</button> : null}
  </>;
}

function EnvelopeCover({ page }) {
  return <article className="export-envelope-cover" aria-label={page.code + " envelope cover for " + page.roomName}>
    <h4>{page.code} Envelope Cover</h4>
    <div className="export-envelope-cover__fields">{envelopeCoverRows(page).map((row) => <dl key={row[0][0]}
      className={"export-envelope-cover__row" + (row.length > 1 ? " export-envelope-cover__row--counts" : "")}>
      {row.map(([label, value, note]) => <div key={label}>
        <dt>{label}</dt><dd>{value || "\u00a0"}</dd>{note ? <small>{note}</small> : null}
      </div>)}
    </dl>)}</div>
    <dl className="export-envelope-cover__session">{envelopeSessionInfo(page).map(([label, value]) => <div key={label}>
      <dt>{label}</dt><dd>{value}</dd>
    </div>)}</dl>
  </article>;
}

function EnvelopePreview({ courses, course, page, onCourseChange, onRoomChange, onDownload, disabled }) {
  return <>
    <div className="export-envelope-controls">
      <label><span>Course</span><select aria-label="Envelope course" value={course.code} disabled={disabled}
        onChange={(event) => onCourseChange(event.target.value)}>{courses.map((item) => <option key={item.code} value={item.code}>
          {item.code}: {item.title}</option>)}</select></label>
      <label><span>Room / exam</span><select aria-label="Envelope room" value={page.id} disabled={disabled}
        onChange={(event) => onRoomChange(event.target.value)}>{course.pages.map((item) => <option key={item.id} value={item.id}>
          {formatReportDate(item.date)} / {reportTimeRange(item.start, item.end)} / {item.roomName} / {item.studentCount} students</option>)}</select></label>
      <button type="button" onClick={onDownload} disabled={disabled}>Download This Course</button>
    </div>
    <EnvelopeCover page={page} />
    <details className="export-envelope-packing" open><summary>Session packing list / this week</summary>
      <PreviewTable label="Envelope packing list" columns={["Date", "Time", "Course", "Room", "CRNs", "Students"]}
        rows={courses.flatMap((item) => item.pages).sort((a, b) => a.date.localeCompare(b.date) || a.start - b.start || a.code.localeCompare(b.code))
          .map((item) => [formatChronologicalDate(item.date), reportTimeRange(item.start, item.end), item.code, item.roomName, item.crns.join(", "), item.studentCount])} />
    </details>
  </>;
}

function StaffDistribution({ model, query = "", sort = "name" }) {
  const distribution = buildStaffDistribution(model);
  const search = query.trim().toLowerCase();
  const rows = distribution.rows.filter((person) => !search || [person.name,
    ...person.weeks.flatMap((week) => week.days.flatMap((day) => day.duties.flatMap((duty) => [duty.code, duty.roomName])))
  ].join(" ").toLowerCase().includes(search)).sort((a, b) =>
    (sort === "name" ? 0 : b[sort] - a[sort]) || a.name.localeCompare(b.name));
  return <section className="staff-distribution" aria-label="Invigilator session distribution">
    <div className="staff-distribution__legend">
      <span><i className="staff-distribution__dot staff-distribution__dot--exam" aria-hidden="true">1</i>Extra exam load</span>
      <span><i className="staff-distribution__dot staff-distribution__dot--teaching" aria-hidden="true">0</i>Lab/tutorial hours</span>
      <span><i className="staff-distribution__dot staff-distribution__dot--none" aria-hidden="true" />No exam duty</span>
      <span><i className="staff-distribution__dot staff-distribution__dot--off" aria-hidden="true" />No exam session</span>
    </div>
    <p className="staff-distribution__note">Daily start/end includes remaining teaching, CyberSkills and exam duties, including gaps. Tutorials, unchecked courses, replaced lab classes and backups do not extend the span. Days over eight hours are highlighted. CyberSkills adds no invigilation load.</p>
    <div className="staff-distribution__scroll" tabIndex="0" aria-label="Session distribution table">
      <table className="staff-distribution__table">
        <caption className="sr-only">Exam session distribution and daily duty hours for included invigilators</caption>
        <thead><tr><th scope="col">Invigilator</th><th scope="col">Extra<br />load</th><th scope="col">Lab/tutorial<br />duties</th>
          {distribution.weeks.map((week) => <th key={week.week} scope="col" className="staff-distribution__week-heading">
            <strong>{"Week " + week.week}</strong><div className="staff-distribution__grid staff-distribution__dates"><span>Exam time</span>
              {week.days.map((day) => <span key={day.day}>{day.day.slice(0, 3)}<small>{formatReportDate(day.date)}</small></span>)}
            </div>
          </th>)}
        </tr></thead>
        <tbody>{rows.map((person) => <tr key={person.id}>
          <th scope="row"><strong>{person.name}</strong><small>{person.total} exam-room {person.total === 1 ? "duty" : "duties"}</small></th>
          <td className="staff-distribution__total">{person.exam}</td><td className="staff-distribution__total staff-distribution__total--teaching">{person.teaching}</td>
          {person.weeks.map((week, index) => <td key={week.week} className="staff-distribution__week-cell">
            <div className="staff-distribution__grid">
              <div className="staff-distribution__times">{distribution.weeks[index].slots.map((slot) => <span key={slot.key}>{reportTimeRange(slot.start, slot.end)}</span>)}<small>Daily span</small></div>
              {week.days.map((day) => <div key={day.day} className="staff-distribution__day">
                {day.slots.map((slot, slotIndex) => {
                  const timing = distribution.weeks[index].slots[slotIndex];
                  const status = slot.duty ? slot.duty.teaching ? "teaching" : "exam" : slot.scheduled ? "none" : "off";
                  const label = person.name + " / " + day.day + " " + formatReportDate(day.date) + " / " + reportTimeRange(timing.start, timing.end) + ": " +
                    (slot.duty ? slot.duty.code + " / " + slot.duty.roomName + (slot.duty.teaching ? " / lab or tutorial hours, zero extra load" : " / one extra exam load") : slot.scheduled ? "No exam duty" : "No exam session");
                  return <span key={slot.key} className={"staff-distribution__dot staff-distribution__dot--" + status} title={label} aria-label={label}>
                    {slot.duty ? slot.duty.extraLoad : ""}</span>;
                })}
                <div className={"staff-distribution__span" + (day.span?.overrun ? " staff-distribution__span--long" : "")}
                  aria-label={person.name + " / " + day.day + " " + formatReportDate(day.date) + " daily duty hours"}>
                  <span>Start <b>{day.span?.start != null ? clock(day.span.start) : "-"}</b></span>
                  <span>End <b>{day.span?.end != null ? clock(day.span.end) : "-"}</b></span>
                  {day.span?.minutes != null ? <small>{Math.round(day.span.minutes / 60 * 10) / 10}h{day.span.overrun ? " / over 8h" : ""}</small> : null}
                  {day.span?.unknownTime ? <small className="staff-distribution__unknown">Class time unknown; span incomplete</small> : null}
                </div>
              </div>)}
            </div>
            <p className="staff-distribution__week-total">{week.exam} extra / {week.teaching} lab/tutorial</p>
          </td>)}
        </tr>)}</tbody>
      </table>
      {!rows.length ? <p className="export-empty">No matching invigilators.</p> : null}
    </div>
  </section>;
}

function StaffSessionRosters({ model, query = "" }) {
  const search = query.trim().toLowerCase();
  const days = new Map();
  buildStaffSessionRosters(model).forEach((session) => {
    const duties = session.duties.filter((duty) => !search || [duty.name, duty.code, duty.roomName].join(" ").toLowerCase().includes(search));
    if (!duties.length) return;
    if (!days.has(session.date)) days.set(session.date, []);
    days.get(session.date).push({ ...session, duties });
  });
  return <section className="staff-sessions" aria-label="Invigilators by session">
    {[...days].map(([date, sessions]) => <section key={date} className="staff-sessions__day">
      <header><span>{"Week " + sessions[0].week}</span><h4>{sessions[0].day + " / " + formatReportDate(date, true)}</h4></header>
      {sessions.map((session) => <article key={session.id} className="staff-sessions__session"
        aria-label={"Week " + session.week + " / " + session.day + " " + formatReportDate(date) + " / " + reportTimeRange(session.start, session.end) + " invigilators"}>
        <header><div><h5><time dateTime={date + "T" + clock(session.start)}>{reportTimeRange(session.start, session.end)}</time></h5>
          <p>{session.courses.join(" / ")}</p></div>
          <span>{session.invigilatorCount} {session.invigilatorCount === 1 ? "invigilator" : "invigilators"} / {session.roomCount} {session.roomCount === 1 ? "room" : "rooms"}</span>
        </header>
        {search ? <p className="staff-sessions__matches">Showing {session.duties.length} matching {session.duties.length === 1 ? "duty" : "duties"}; session totals above are unchanged.</p> : null}
        <PreviewTable label={"Week " + session.week + " " + session.day + " " + reportTimeRange(session.start, session.end) + " invigilator roster"}
          className="export-table--session-roster" columns={["Invigilator", "Course", "Room", "Start", "End", "Load"]}
          rows={session.duties.map((duty) => [<strong>{duty.name}</strong>, duty.code, duty.roomName, clock(duty.start), clock(duty.end),
            duty.teaching ? <span className="export-teaching">Lab/tutorial hours / 0 load</span> : "1 extra duty"])} />
      </article>)}
    </section>)}
    {!days.size ? <p className="export-empty">No matching session duties.</p> : null}
  </section>;
}

function StaffTimetable({ person, week }) {
  const slots = staffScheduleSlots(week);
  const duties = week.days.flatMap((day) => day.entries).filter((entry) => entry.kind === "exam");
  return <article className="staff-timetable" aria-label={person.name + " / Week " + week.week + " schedule"}>
    <header><div><span className="export-kicker">{"Week " + week.week}</span><h4>{person.name}</h4></div>
      <small>{duties.reduce((sum, duty) => sum + duty.extraLoad, 0)} extra exam load / {duties.filter((duty) => duty.teaching).length} lab/tutorial duties</small></header>
    <div className="staff-timetable__scroll" tabIndex="0" aria-label={person.name + " weekly timetable"}>
      <table className="staff-timetable__table">
        <caption className="sr-only">{person.name + " / Week " + week.week + " commitments"}</caption>
        <thead><tr><th scope="col">Time</th>{week.days.map((day) => <th key={day.day} scope="col">
          {day.day.slice(0, 3)}<small>{formatReportDate(day.date)}</small></th>)}</tr></thead>
        <tbody>{slots.map((slot) => <tr key={slot.key}><th scope="row">{slot.unknownTime ? "Time not listed" : reportTimeRange(slot.start, slot.end)}</th>
          {week.days.map((day) => <td key={day.day}>{staffSlotEntries(day, slot).map((entry) => <div key={entry.id}
            className={"staff-timetable__entry staff-timetable__entry--" + entry.kind}>
            <span>{entry.label}</span><strong>{entry.code || entry.title}</strong>
            {entry.kind === "exam" && entry.title ? <small>{entry.title}</small> : null}
            {entry.crn ? <small>CRN {entry.crn}</small> : null}
            {entry.roomName ? <small>{entry.roomName}</small> : null}
            {entry.kind === "exam" ? <small>{entry.teaching ? "Lab/tutorial hours / 0 load" : "1 extra duty"}</small> : null}
          </div>)}</td>)}
        </tr>)}{!slots.length ? <tr><td colSpan={6} className="staff-timetable__empty">No scheduled duties this week.</td></tr> : null}</tbody>
        <tfoot><tr><th scope="row">Shown hours</th>{week.days.map((day) => <td key={day.day}>
          {day.start !== null ? reportTimeRange(day.start, day.end) : "-"}{day.unknownTime ? <small>Class time not listed</small> : null}
        </td>)}</tr></tfoot>
      </table>
    </div>
  </article>;
}

function StaffSchedules({ people, selectedId, query, onSelect, onDownload, disabled }) {
  const search = query.trim().toLowerCase();
  const filteredPeople = people.filter((person) => !search || [person.name, ...person.weeks.flatMap((week) => week.days.flatMap((day) =>
    day.entries.flatMap((entry) => [entry.code, entry.title, entry.roomName, entry.label])))].join(" ").toLowerCase().includes(search));
  const selected = filteredPeople.find((person) => person.id === selectedId) || filteredPeople[0];
  const displayed = selectedId === "all" ? filteredPeople : selected ? [selected] : [];
  return <section className="staff-schedule-browser" aria-label="Staff schedules">
    <nav className="staff-schedule-browser__list" aria-label="Browse invigilators">
      <button type="button" aria-pressed={selectedId === "all"} onClick={() => onSelect("all")}>All staff <small>{people.length} invigilators</small></button>
      {filteredPeople.map((person) => <button type="button" key={person.id} aria-pressed={selectedId !== "all" && selected?.id === person.id}
        onClick={() => onSelect(person.id)}><strong>{person.name}</strong><small>{person.duties.length} exam {person.duties.length === 1 ? "duty" : "duties"}</small></button>)}
    </nav>
    <div className="staff-schedule-browser__detail">
      {displayed.length && selectedId !== "all" ? <div className="staff-schedule-browser__actions"><button type="button" disabled={disabled}
        onClick={() => onDownload(selected.id)}>Download This Invigilator</button><small>All included weeks, not just the preview.</small></div> : null}
      {displayed.flatMap((person) => person.weeks.map((week) => <StaffTimetable key={person.id + "/" + week.week} person={person} week={week} />))}
      {!displayed.length ? <p className="export-empty">No matching invigilators.</p> : null}
    </div>
  </section>;
}

function PrintReport({ report, overviewMode, staffView = "duties", includeTeaching, includeCyberSkills, model, startDate, sessions, catalog, plan, asdExams }) {
  if (report === "staff" && staffView === "schedules") return <div className="export-print" data-report="staff-schedules">
    {buildStaffSchedules(model, { includeTeaching, includeCyberSkills }).flatMap((person) => person.weeks.map((week) => <section
      className="export-print__staff" key={person.id + "/" + week.week}>
      <header><span>Staff schedules / {includeTeaching ? "Teaching included / " : ""}{includeCyberSkills ? "CyberSkills included / " : ""}No backups</span></header>
      <StaffTimetable person={person} week={week} />
    </section>))}
  </div>;
  if (report === "envelopes") return <div className="export-print" data-report="envelopes">
    {buildEnvelopeCourses(model).flatMap((course) => course.pages.map((page) => <section className="export-print__envelope" key={page.id}>
      <EnvelopeCover page={page} />
    </section>))}
  </div>;
  if (report === "seating") return <div className="export-print" data-report="seating">
    {buildCourseSeating(model).flatMap((exam) => exam.crnSheets.map((sheet) => <section className="export-print__crn" key={exam.id + "/" + sheet.crn}>
      <SeatingSheet exam={exam} sheet={sheet} label={exam.code + " CRN " + (sheet.crn || "Unspecified") + " seating"} />
      <footer>ASD exams are excluded. Contains student IDs; share only with authorised recipients.</footer>
    </section>))}
  </div>;
  return <div className="export-print">
    {model.selectedWeeks.map((week) => {
      const current = buildExportModel({ sessions, catalog, plan, startDate, weeks: [week], asdExams, includeAsd: model.includeAsd });
      return <section className="export-print__week" key={week}>
        <header><span>Department exams / confirmed resources</span><h1>{REPORT_VIEWS.find((view) => view.id === report).label}</h1>
          <p>{"Week " + week + " / " + formatReportDate(reportDate(startDate, week, "Monday"), true) + " - " + formatReportDate(reportDate(startDate, week, "Friday"), true)}</p></header>
        {report === "overview" ? <>{overviewMode === "board" ? <p className="export-overview-note">{overviewNote}</p> : null}
          {overviewMode === "board" ? <PrintWeekBoard model={current} startDate={startDate} week={week} />
            : <PreviewTable label="Exam overview" className="export-table--chronological" columns={overviewColumns} rows={overviewRows(current)} rowClassNames={overviewRowClasses(current)} />}</> : null}
        {report === "complete" ? <><h2>Room assignments</h2><PreviewTable label="Room assignments" columns={["Course", "Day / time", "Room", "Students", "Invigilators", "Room limit"]} rows={ledgerRows(current)} /></> : null}
        {report === "staff" && staffView === "distribution" ? <><h2>Session distribution</h2><StaffDistribution model={current} /></> : null}
        {report === "staff" && staffView === "sessions" ? <><h2>Invigilators by session</h2><StaffSessionRosters model={current} /></> : null}
        {report === "complete" || (report === "staff" && staffView === "duties") ? <><h2>Staff duties</h2><PreviewTable label="Staff duties" columns={["Invigilator", "Day / time", "Duty", "Course", "Room", "Load"]} rows={staffRows(current)} />
          <h2>Workload summary</h2><PreviewTable label="Workload summary" columns={["Invigilator", "Exam load", "Teaching duties"]} rows={loadRows(current)} /></> : null}
        {report === "complete" || report === "students" ? <><h2>Student room lists</h2><PreviewTable label="Student room lists" columns={["Student ID", "Name", "Course", "CRN", "Day / time", "Room"]} rows={studentRows(current)} /></> : null}
        <footer>{model.includeAsd ? asdReferenceNote : "ASD exams are excluded."} Teaching-time duties do not add extra invigilation load. {report === "complete" || report === "students" ? "Contains student personal data." : ""}</footer>
      </section>;
    })}
  </div>;
}

export default function ExportStudio({ sessions, catalog, plan, startDate, examStartDate = startDate, asdExams = [], ready, isExporting, onExport }) {
  const [report, setReport] = useState("complete");
  const [format, setFormat] = useState("xlsx");
  const [packaging, setPackaging] = useState("combined");
  const [excludedWeeks, setExcludedWeeks] = useState([]);
  const [previewWeek, setPreviewWeek] = useState(null);
  const [mode, setMode] = useState("board");
  const [staffView, setStaffView] = useState("duties");
  const [staffSort, setStaffSort] = useState("name");
  const [previewStaff, setPreviewStaff] = useState("");
  const [includeTeaching, setIncludeTeaching] = useState(false);
  const [includeCyberSkills, setIncludeCyberSkills] = useState(false);
  const [includeCalendars, setIncludeCalendars] = useState(false);
  const [includeAsd, setIncludeAsd] = useState(false);
  const [search, setSearch] = useState("");
  const [limit, setLimit] = useState(50);
  const [previewExam, setPreviewExam] = useState(null);
  const [previewCrn, setPreviewCrn] = useState(null);
  const [previewRoom, setPreviewRoom] = useState(null);
  const withAsd = report === "overview" && includeAsd;
  const full = useMemo(() => ready ? buildExportModel({ sessions, catalog, plan, startDate, asdExams, includeAsd: withAsd }) : null, [sessions, catalog, plan, startDate, asdExams, withAsd, ready]);
  const selectedWeeks = useMemo(() => full?.selectedWeeks.filter((week) => !excludedWeeks.includes(week)) || [], [full, excludedWeeks]);
  const selected = useMemo(() => selectedWeeks.length ? buildExportModel({ sessions, catalog, plan, startDate, weeks: selectedWeeks, asdExams, includeAsd: withAsd }) : null, [sessions, catalog, plan, startDate, selectedWeeks, asdExams, withAsd]);
  const overall = report === "staff" && previewWeek === "overall";
  const activeWeek = selectedWeeks.includes(previewWeek) ? previewWeek : selectedWeeks[0];
  const current = useMemo(() => overall ? selected : activeWeek ? buildExportModel({ sessions, catalog, plan, startDate, weeks: [activeWeek], asdExams, includeAsd: withAsd }) : null, [sessions, catalog, plan, startDate, activeWeek, overall, selected, asdExams, withAsd]);
  const staffSchedules = useMemo(() => report === "staff" && staffView === "schedules" && current
    ? buildStaffSchedules(current, { includeTeaching, includeCyberSkills }) : [], [report, staffView, current, includeTeaching, includeCyberSkills]);
  const seatingExams = useMemo(() => report === "seating" && current ? buildCourseSeating(current) : [], [report, current]);
  const seatingExam = seatingExams.find((exam) => exam.id === previewExam) || seatingExams[0];
  const seatingSheet = seatingExam?.crnSheets.find((sheet) => sheet.crn === previewCrn) || seatingExam?.crnSheets[0];
  const envelopeCourses = useMemo(() => report === "envelopes" && current ? buildEnvelopeCourses(current) : [], [report, current]);
  const envelopeCourse = envelopeCourses.find((course) => course.code === previewExam) || envelopeCourses[0];
  const envelopePage = envelopeCourse?.pages.find((page) => page.id === previewRoom) || envelopeCourse?.pages[0];
  const previewLabel = overall ? "Overall" : "Week " + activeWeek;
  const view = REPORT_VIEWS.find((item) => item.id === report);
  const matches = (record) => Object.values(record).flat().filter((value) => typeof value === "string").join(" ").toLowerCase().includes(search.toLowerCase());
  const filtered = current ? { ...current, roomRows: current.roomRows.filter(matches), duties: current.duties.filter(matches), students: current.students.filter(matches) } : null;
  const records = report === "students" ? filtered?.students.length : report === "staff" ? filtered?.duties.length : filtered?.roomRows.length;
  const disabled = isExporting || !ready || !selectedWeeks.length;
  const chooseReport = (id) => {
    setReport(id);
    if (id === "staff" && report !== "staff") setPreviewWeek("overall");
    if (id === "envelopes") setFormat("pdf");
    else if (id === "complete" || id === "seating" || format === "pdf") setFormat("xlsx");
    setSearch("");
    setLimit(50);
  };
  const coursePdfs = report === "envelopes" && format === "pdf";
  const exportOptions = { report, format, packaging: report === "seating" || coursePdfs ? "course" : packaging, weeks: selectedWeeks, includeAsd: withAsd,
    ...(report === "staff" && staffView === "schedules" ? { staffView, includeTeaching, includeCyberSkills } : {}),
    ...(report === "staff" && includeCalendars ? { includeCalendars: true } : {}),
    ...(report === "overview" ? { overviewMode: mode } : {}) };
  const pdfFilename = reportFileName({ report, staffView, startDate: examStartDate, format: "pdf", overviewMode: mode,
    includeAsd: selected?.overviewExams.some((exam) => exam.isAsd), week: selectedWeeks.length === 1 ? selectedWeeks[0] : undefined });
  const handlePrint = () => {
    const previousTitle = document.title;
    document.title = pdfFilename.slice(0, -4);
    try {
      window.print();
    } finally {
      document.title = previousTitle;
    }
  };
  return <section className="export-studio" aria-label="Export workspace">
    <header className="export-hero">
      <div><h2>Export Schedule</h2></div>
      <span className={"export-ready" + (ready ? "" : " export-ready--blocked")}>{ready ? "Resources confirmed" : "Resources need attention"}</span>
      <div className="export-stats" aria-live="polite">
        <div><strong>{selectedWeeks.length}</strong><span>included weeks</span></div>
        <div><strong>{selected?.summary.exams || 0}</strong><span>department exams</span></div>
        <div><strong>{selected?.summary.roomSessions || 0}</strong><span>room sessions</span></div>
        <div><strong>{selected?.summary.students || 0}</strong><span>unique students</span></div>
      </div>
    </header>
    <fieldset className="export-weeks"><legend>Include exam weeks</legend>
      {(full?.selectedWeeks || []).map((week) => <label key={week}>
        <input type="checkbox" checked={selectedWeeks.includes(week)} disabled={isExporting} aria-label={"Include Week " + week}
          onChange={() => { setExcludedWeeks((previous) => previous.includes(week) ? previous.filter((item) => item !== week) : [...previous, week]); setSearch(""); setLimit(50); }} />
        <span><strong>{"Week " + week}</strong><small>{formatReportDate(reportDate(startDate, week, "Monday")) + " - " + formatReportDate(reportDate(startDate, week, "Friday"))}</small></span>
      </label>)}
      <p>{withAsd ? "Weeks with department or ASD exams are listed. Resource totals cover department exams only." : "Only weeks with department exams are listed. ASD is excluded."}</p>
    </fieldset>
    <div className="export-layout">
      <aside className="export-options">
        <h3>Choose your report</h3>
        <div className="export-report-choices">{REPORT_VIEWS.map((item, index) => <button type="button" key={item.id} aria-label={item.label}
          aria-pressed={report === item.id} disabled={isExporting} onClick={() => chooseReport(item.id)}>
          <ReportIcon kind={item.id} /><span><small>{String(index + 1).padStart(2, "0") + " / " + item.audience}</small><strong>{item.label}</strong><p>{item.description}</p></span>
        </button>)}</div>
        {report === "overview" ? <fieldset className="export-asd-option"><legend>ASD reference schedule</legend>
          <label><input type="checkbox" checked={includeAsd} disabled={isExporting || !asdExams.length}
            onChange={(event) => setIncludeAsd(event.target.checked)} /> <span>Include ASD exams</span></label>
          <small>{asdExams.length ? "Reference entries in both views, Excel, CSV and print/PDF. Department resource totals stay unchanged." : "No ASD schedule loaded. Add one in the ASD step to include it here."}</small>
          {withAsd && selected ? <small role="status">{selected.overviewExams.filter((exam) => exam.isAsd).length} ASD reference exams in included weeks.</small> : null}
        </fieldset> : null}
        {report === "seating" ? <div className="export-seating-package"><h3>Excel files by exam</h3>
          <p>One workbook per exam, with a sheet for each CRN. Uses confirmed room assignments.</p>
          <small>Multiple workbooks download as a ZIP. Student names and ASD exams are excluded.</small>
        </div> : <><fieldset className="export-format"><legend>File format</legend>
          {report === "envelopes" ? <label><input type="radio" name="export-format" checked={format === "pdf"} disabled={isExporting} onChange={() => setFormat("pdf")} />PDF covers (.pdf)</label> : null}
          <label><input type="radio" name="export-format" checked={format === "xlsx"} disabled={isExporting} onChange={() => setFormat("xlsx")} />Excel (.xlsx)</label>
          {report !== "envelopes" ? <label><input type="radio" name="export-format" checked={format === "csv"} disabled={isExporting || report === "complete"} onChange={() => setFormat("csv")} />CSV (.csv)</label> : null}
          {report === "envelopes" ? <small>PDF: one A4 page per room, grouped by course. Excel: packing list with one sheet per session.</small> : null}
          {report === "complete" ? <small>The complete report keeps its linked Excel sheets.</small> : report === "staff" ? <small>{staffView === "schedules"
            ? "Excel: a weekly timetable per invigilator. CSV: the selected schedule layers."
            : "Excel includes duties, workload, session distribution and daily start/end times. CSV contains exam duties only."}</small> : null}
        </fieldset>
        {report === "staff" ? <fieldset className="export-calendar-option"><legend>Calendar files</legend>
          <label><input type="checkbox" checked={includeCalendars} disabled={isExporting}
            onChange={(event) => setIncludeCalendars(event.target.checked)} />Include calendar files (.ics)</label>
          <small>Adds one file per assigned invigilator to the ZIP. Exam invigilation only, including zero-load duties. No teaching, CyberSkills or backups. Times use Asia/Dubai (UTC+4). Import into your calendar; no invitations are sent.</small>
        </fieldset> : null}
        {coursePdfs ? <div className="export-seating-package"><h3>PDF files by course</h3>
          <p>Each course PDF contains its room covers for the included weeks.</p><small>Multiple courses download as a ZIP. Instructor Name is left blank; Department is omitted.</small>
        </div> : <fieldset className="export-packaging"><legend>Package your files</legend>
          <label><input type="radio" name="export-packaging" checked={packaging === "combined"} disabled={isExporting} onChange={() => setPackaging("combined")} />
            <span><strong>{format === "xlsx" ? "One workbook" : "One CSV"}</strong><small>All included weeks in one file.</small></span></label>
          <label><input type="radio" name="export-packaging" checked={packaging === "weekly"} disabled={isExporting} onChange={() => setPackaging("weekly")} />
            <span><strong>Separate weekly files</strong><small>One file per week; multiple files download as a ZIP.</small></span></label>
        </fieldset>}
        </>}
        <div className="export-download">
          <p>{coursePdfs ? selected ? buildEnvelopeCourses(selected).length + " course PDFs / " + selected.roomRows.length + " room covers" : "Select a week to export."
            : report === "seating" ? selected?.exams.length ? selected.exams.length + (selected.exams.length === 1 ? " exam workbook" : " exam workbooks / ZIP download") : "Select a week to export."
            : selectedWeeks.length ? report === "staff" && includeCalendars ? "Report + invigilator calendars / ZIP download"
            : packaging === "weekly" && selectedWeeks.length > 1 ? selectedWeeks.length + " weekly files / ZIP download" : "1 " + (format === "xlsx" ? "Excel workbook" : "CSV file") : "Select a week to export."}</p>
          <button type="button" className="export-download__primary" disabled={disabled}
            onClick={() => onExport(exportOptions)}>{isExporting ? "Exporting..." : coursePdfs ? "Download Envelope PDFs" : report === "seating" ? "Download Course Files" : report === "complete" ? "Export Timetable" : format === "xlsx" ? "Download Excel" : "Download CSV"}</button>
          <button type="button" disabled={disabled} onClick={handlePrint}>Print / Save PDF</button>
          <small className="export-pdf-filename">PDF name: {pdfFilename}</small>
          {report === "seating" ? <small>Prints every CRN in the included weeks, starting each sheet on a new portrait page.</small> : null}
          {report === "envelopes" ? <small>Print/PDF includes every room cover in the included weeks, not just the preview.</small> : null}
          {report === "overview" ? <small className="export-print-layout" role="status">Print layout: {mode === "board" ? "Week board" : "Chronological list"}. All included weeks are printed. {mode === "board" ? "Excel and CSV use the detailed table." : "Excel and CSV match the six-column chronological list."}</small> : null}
          {report === "staff" ? <small>Print layout: {staffView === "schedules" ? "Staff schedules with selected layers" : staffView === "distribution" ? "Session distribution with daily start/end times" : staffView === "sessions" ? "Invigilators by session" : "Duty list and workload summary"}. All included weeks are printed.</small> : null}
          <small>{report === "seating" ? "Contains student IDs. Share only with authorised recipients." : report === "complete" || report === "students" ? "Contains student names and IDs. Share only with authorised recipients." : "Only the selected report and included weeks are exported."}</small>
        </div>
      </aside>
      <main className="export-preview" aria-label="Report preview">
        <div className="export-preview__toolbar">
          <span>Live preview</span>
          <div aria-label="Preview period">
            {report === "staff" ? <button type="button" aria-pressed={overall} disabled={!selectedWeeks.length}
              onClick={() => { setPreviewWeek("overall"); setSearch(""); setLimit(50); }}>Overall</button> : null}
            {selectedWeeks.map((week) => <button type="button" key={week} aria-pressed={!overall && activeWeek === week}
            onClick={() => { setPreviewWeek(week); setSearch(""); setLimit(50); }}>{"Week " + week}</button>)}</div>
        </div>
        {!current ? <div className="export-empty"><h3>{ready ? "Choose a week to preview" : "Confirm resources first"}</h3>
          <p>{ready ? "Select at least one exam week above." : "Return to Resources to resolve the assignments before exporting."}</p></div> : <>
          <header className="export-paper-heading"><span className="export-kicker">Department exams / {previewLabel}</span><h3>{view.label}</h3>
            <p>{formatReportDate(reportDate(startDate, overall ? selectedWeeks[0] : activeWeek, "Monday"), true) + " - " + formatReportDate(reportDate(startDate, overall ? selectedWeeks.at(-1) : activeWeek, "Friday"), true)}</p>
            <small>{report === "envelopes" ? "One cover per confirmed room / no student names or IDs" : report === "overview" ? "Shareable timetable / no student names or IDs" : report === "seating" ? "Student IDs and rooms / one sheet per CRN" : report === "students" ? "Room check-in register / personal data" : report === "staff" ? "Exam duties and teaching-time load / no backups" : "Confirmed allocations / linked Excel sheets"}</small>
          </header>
          {report === "seating" ? <SeatingPreview exams={seatingExams} exam={seatingExam} sheet={seatingSheet} search={search} limit={limit} disabled={disabled}
            onExamChange={(id) => { setPreviewExam(id); setPreviewCrn(null); setSearch(""); setLimit(50); }}
            onCrnChange={(crn) => { setPreviewCrn(crn); setSearch(""); setLimit(50); }}
            onSearch={(value) => { setSearch(value); setLimit(50); }} onShowMore={() => setLimit((previous) => previous + 50)}
            onDownload={() => onExport({ ...exportOptions, examIds: [seatingExam.id] })} />
            : report === "envelopes" && envelopePage ? <EnvelopePreview courses={envelopeCourses} course={envelopeCourse} page={envelopePage} disabled={disabled}
              onCourseChange={(code) => { setPreviewExam(code); setPreviewRoom(null); }} onRoomChange={setPreviewRoom}
              onDownload={() => onExport({ ...exportOptions, format: "pdf", packaging: "course", courseCodes: [envelopeCourse.code] })} />
            : report === "overview" ? <>{mode === "board" || withAsd ? <p className="export-overview-note">{mode === "board" ? overviewNote : ""} {withAsd ? asdReferenceNote : ""}</p> : null}<div className="export-view-toggle"><button type="button" aria-pressed={mode === "board"} onClick={() => setMode("board")}>Week board</button>
            <button type="button" aria-pressed={mode === "table"} onClick={() => setMode("table")}>Chronological list</button></div>
            {mode === "board" ? <WeekBoard model={current} startDate={startDate} week={activeWeek} /> : <PreviewTable label="Exam overview preview"
              className="export-table--chronological" columns={overviewColumns} rows={overviewRows(current)} rowClassNames={overviewRowClasses(current)} />}</> : <>
            {report === "complete" ? <WorkbookMap model={packaging === "combined" ? selected : current} week={activeWeek} packaging={packaging} /> : null}
            {report === "staff" ? <div className="export-view-toggle" aria-label="Staff duty views">
              <button type="button" aria-pressed={staffView === "duties"} onClick={() => { setStaffView("duties"); setSearch(""); setLimit(50); }}>Duty list</button>
              <button type="button" aria-pressed={staffView === "distribution"} onClick={() => { setStaffView("distribution"); setSearch(""); }}>Session distribution</button>
              <button type="button" aria-pressed={staffView === "sessions"} onClick={() => { setStaffView("sessions"); setSearch(""); }}>Invigilators by session</button>
              <button type="button" aria-pressed={staffView === "schedules"} onClick={() => { setStaffView("schedules"); setSearch(""); }}>Staff schedules</button>
            </div> : null}
            {report === "staff" && staffView === "schedules" ? <fieldset className="staff-schedule-layers"><legend>Include other schedules</legend>
              <label><input type="checkbox" checked={includeTeaching} disabled={isExporting}
                onChange={(event) => setIncludeTeaching(event.target.checked)} />Teaching</label>
              <label><input type="checkbox" checked={includeCyberSkills} disabled={isExporting}
                onChange={(event) => setIncludeCyberSkills(event.target.checked)} />CyberSkills</label>
              <small>Display and report layers only. Teaching includes tutorials and unchecked courses; replaced labs are omitted. Calendar files always contain exam duties only.</small>
            </fieldset> : null}
            {report === "staff" && staffView === "duties" ? <details className="export-workloads" open><summary>Workload balance / {previewLabel}</summary>
              {overall ? <p>Totals across included weeks: {selectedWeeks.map((week) => "Week " + week).join(", ")}. Include every week to match the Resource Review Pool. Invigilators with no assignments are shown too.</p> : null}
              <PreviewTable label="Workload balance" columns={["Invigilator", "Exam load", "Teaching duties"]} rows={loadRows(current, overall)} />
              <p>Exam duties fully within listed lab or tutorial hours count as zero extra load. Lecture hours do not qualify. Backups are excluded from duties and workload.</p></details> : null}
            <label className="export-search"><span>Search this preview</span><input type="search" value={search} placeholder={report === "students" ? "Student, course or room..." : report === "staff" ? "Invigilator, course or room..." : "Course, room or invigilator..."}
              onChange={(event) => { setSearch(event.target.value); setLimit(50); }} /><small>Search affects the preview only, not the download or print.</small></label>
            {report === "staff" && staffView === "schedules" ? <StaffSchedules people={staffSchedules} selectedId={previewStaff} query={search}
              onSelect={setPreviewStaff} disabled={disabled} onDownload={(id) => onExport({ ...exportOptions, staffIds: [id] })} />
            : report === "staff" && staffView === "distribution" ? <>
              <label className="staff-distribution__sort">Sort invigilators<select value={staffSort} onChange={(event) => setStaffSort(event.target.value)}>
                <option value="name">Name</option><option value="exam">Extra exam load (highest first)</option><option value="total">Exam-room duties (highest first)</option>
              </select></label><StaffDistribution model={current} query={search} sort={staffSort} />
            </> : report === "staff" && staffView === "sessions" ? <StaffSessionRosters model={current} query={search} /> : <PreviewTable label={view.label + " preview"} limit={limit}
              columns={report === "students" ? ["Student ID", "Name", "Course", "CRN", "Day / time", "Room"] : report === "staff" ? ["Invigilator", overall ? "Week / date / time" : "Day / time", "Duty", "Course", "Room", "Load"] : ["Course", "Day / time", "Room", "Students", "Invigilators", "Room limit"]}
              rows={report === "students" ? studentRows(filtered) : report === "staff" ? staffRows(filtered, overall) : ledgerRows(filtered)} />}
            {!(report === "staff" && staffView !== "duties") && records > limit ? <button type="button" className="export-show-more" onClick={() => setLimit((previous) => previous + 50)}>Show 50 more ({records - limit} remaining)</button> : null}
          </>}
          <footer className="export-preview__footer">Previewing {overall ? "all included weeks" : previewLabel}. {report === "seating" ? "Download This Exam includes only the selected exam, with all its CRNs. Course files and print include all " : "Downloads and print include all "}{selectedWeeks.length} selected {selectedWeeks.length === 1 ? "week" : "weeks"}. {withAsd ? asdReferenceNote : "ASD exams are excluded."}</footer>
        </>}
      </main>
    </div>
    {selected ? <PrintReport report={report} overviewMode={mode} staffView={staffView} includeTeaching={includeTeaching} includeCyberSkills={includeCyberSkills}
      model={selected} sessions={sessions} catalog={catalog} plan={plan} startDate={startDate} asdExams={asdExams} /> : null}
  </section>;
}
