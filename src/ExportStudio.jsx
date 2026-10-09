import { useMemo, useState } from "react";
import { buildCourseSeating, buildExportModel, courseSeatingInfo, formatChronologicalDate, formatReportDate, reportDate, reportFileName, reportTable, REPORT_DAYS, REPORT_VIEWS, reportTimeRange } from "./exportReports.js";
import "./ExportStudio.css";

function ReportIcon({ kind }) {
  const paths = {
    complete: "M5 3h10l4 4v14H5V3Zm10 0v5h4M8 12h8M8 16h8",
    overview: "M3 5h18v16H3V5Zm0 5h18M8 3v4M16 3v4M8 10v11M16 10v11",
    staff: "M16 21v-3a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v3M21 21v-3a4 4 0 0 0-3-4M18 3a4 4 0 0 1 0 8",
    students: "M4 3h16v18H4V3ZM8 7h8M8 11h8M8 15h4",
    seating: "M3 4h18v16H3V4ZM3 9h18M11 9v11M6 13h2M6 16h2M14 13h4M14 16h4",
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
  <span className={duty.role === "Backup" ? "export-duty export-duty--backup" : "export-duty"}>{duty.role}</span>,
  duty.code, duty.roomName, duty.teaching ? <span className="export-teaching">Teaching hours / 0 load</span> : "1 extra duty"]);
const studentRows = (model) => model.students.map((student) => [student.id, student.name, student.code, student.crn,
  student.day + " / " + reportTimeRange(student.start, student.end), student.roomName]);
const loadRows = (model, includeIdle = false) => model.workloads.filter((person) => includeIdle || person.exam + person.backup + person.teaching > 0)
  .map((person) => [person.name, person.exam, person.backup, person.teaching, person.exam + person.backup]);
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

function WorkbookMap({ model, week }) {
  const sheets = [
    ["Invigilators", model.roomRows.length + " room assignments"],
    ["Invigilator Pool", model.workloads.length + " named invigilators"],
    ["Room Pool", "Confirmed room references"],
    ...REPORT_DAYS.filter((day) => model.sessions.some((session) => session.day === day))
      .map((day) => [day, model.students.filter((student) => student.day === day).length + " student sittings"]),
  ];
  return <div className="export-sheet-map" aria-label="Workbook sheet preview">
    {sheets.map(([name, detail], index) => <div key={name}><span>{String(index + 1).padStart(2, "0")}</span>
      <strong>{"Week " + week + " " + name}</strong><small>{detail}</small></div>)}
    <p>Room and invigilator references remain linked inside Excel. A combined workbook adds a Schedule Index and preserves every week's sheets.</p>
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

function PrintReport({ report, overviewMode, model, startDate, sessions, catalog, plan, asdExams }) {
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
        {report === "complete" || report === "staff" ? <><h2>Staff duties</h2><PreviewTable label="Staff duties" columns={["Invigilator", "Day / time", "Duty", "Course", "Room", "Load"]} rows={staffRows(current)} />
          <h2>Workload summary</h2><PreviewTable label="Workload summary" columns={["Invigilator", "Exam load", "Backup load", "Teaching duties", "Extra duties"]} rows={loadRows(current)} /></> : null}
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
  const [includeAsd, setIncludeAsd] = useState(false);
  const [search, setSearch] = useState("");
  const [limit, setLimit] = useState(50);
  const [previewExam, setPreviewExam] = useState(null);
  const [previewCrn, setPreviewCrn] = useState(null);
  const withAsd = report === "overview" && includeAsd;
  const full = useMemo(() => ready ? buildExportModel({ sessions, catalog, plan, startDate, asdExams, includeAsd: withAsd }) : null, [sessions, catalog, plan, startDate, asdExams, withAsd, ready]);
  const selectedWeeks = useMemo(() => full?.selectedWeeks.filter((week) => !excludedWeeks.includes(week)) || [], [full, excludedWeeks]);
  const selected = useMemo(() => selectedWeeks.length ? buildExportModel({ sessions, catalog, plan, startDate, weeks: selectedWeeks, asdExams, includeAsd: withAsd }) : null, [sessions, catalog, plan, startDate, selectedWeeks, asdExams, withAsd]);
  const overall = report === "staff" && previewWeek === "overall";
  const activeWeek = selectedWeeks.includes(previewWeek) ? previewWeek : selectedWeeks[0];
  const current = useMemo(() => overall ? selected : activeWeek ? buildExportModel({ sessions, catalog, plan, startDate, weeks: [activeWeek], asdExams, includeAsd: withAsd }) : null, [sessions, catalog, plan, startDate, activeWeek, overall, selected, asdExams, withAsd]);
  const seatingExams = useMemo(() => report === "seating" && current ? buildCourseSeating(current) : [], [report, current]);
  const seatingExam = seatingExams.find((exam) => exam.id === previewExam) || seatingExams[0];
  const seatingSheet = seatingExam?.crnSheets.find((sheet) => sheet.crn === previewCrn) || seatingExam?.crnSheets[0];
  const previewLabel = overall ? "Overall" : "Week " + activeWeek;
  const view = REPORT_VIEWS.find((item) => item.id === report);
  const matches = (record) => Object.values(record).flat().filter((value) => typeof value === "string").join(" ").toLowerCase().includes(search.toLowerCase());
  const filtered = current ? { ...current, roomRows: current.roomRows.filter(matches), duties: current.duties.filter(matches), students: current.students.filter(matches) } : null;
  const records = report === "students" ? filtered?.students.length : report === "staff" ? filtered?.duties.length : filtered?.roomRows.length;
  const disabled = isExporting || !ready || !selectedWeeks.length;
  const chooseReport = (id) => { setReport(id); if (id === "complete" || id === "seating") setFormat("xlsx"); setSearch(""); setLimit(50); };
  const exportOptions = { report, format, packaging: report === "seating" ? "course" : packaging, weeks: selectedWeeks, includeAsd: withAsd,
    ...(report === "overview" ? { overviewMode: mode } : {}) };
  const pdfFilename = reportFileName({ report, startDate: examStartDate, format: "pdf", overviewMode: mode,
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
          <label><input type="radio" name="export-format" checked={format === "xlsx"} disabled={isExporting} onChange={() => setFormat("xlsx")} />Excel (.xlsx)</label>
          <label><input type="radio" name="export-format" checked={format === "csv"} disabled={isExporting || report === "complete"} onChange={() => setFormat("csv")} />CSV (.csv)</label>
          {report === "complete" ? <small>The complete report keeps its linked Excel sheets.</small> : report === "staff" && format === "csv" ? <small>CSV contains duties; the workload summary is included in Excel and print.</small> : null}
        </fieldset>
        <fieldset className="export-packaging"><legend>Package your files</legend>
          <label><input type="radio" name="export-packaging" checked={packaging === "combined"} disabled={isExporting} onChange={() => setPackaging("combined")} />
            <span><strong>{format === "xlsx" ? "One workbook" : "One CSV"}</strong><small>All included weeks in one file.</small></span></label>
          <label><input type="radio" name="export-packaging" checked={packaging === "weekly"} disabled={isExporting} onChange={() => setPackaging("weekly")} />
            <span><strong>Separate weekly files</strong><small>One file per week; multiple files download as a ZIP.</small></span></label>
        </fieldset>
        </>}
        <div className="export-download">
          <p>{report === "seating" ? selected?.exams.length ? selected.exams.length + (selected.exams.length === 1 ? " exam workbook" : " exam workbooks / ZIP download") : "Select a week to export."
            : selectedWeeks.length ? packaging === "weekly" && selectedWeeks.length > 1 ? selectedWeeks.length + " weekly files / ZIP download" : "1 " + (format === "xlsx" ? "Excel workbook" : "CSV file") : "Select a week to export."}</p>
          <button type="button" className="export-download__primary" disabled={disabled}
            onClick={() => onExport(exportOptions)}>{isExporting ? "Exporting..." : report === "seating" ? "Download Course Files" : report === "complete" ? "Export Timetable" : format === "xlsx" ? "Download Excel" : "Download CSV"}</button>
          <button type="button" disabled={disabled} onClick={handlePrint}>Print / Save PDF</button>
          <small className="export-pdf-filename">PDF name: {pdfFilename}</small>
          {report === "seating" ? <small>Prints every CRN in the included weeks, starting each sheet on a new portrait page.</small> : null}
          {report === "overview" ? <small className="export-print-layout" role="status">Print layout: {mode === "board" ? "Week board" : "Chronological list"}. All included weeks are printed. {mode === "board" ? "Excel and CSV use the detailed table." : "Excel and CSV match the six-column chronological list."}</small> : null}
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
            <small>{report === "overview" ? "Shareable timetable / no student names or IDs" : report === "seating" ? "Student IDs and rooms / one sheet per CRN" : report === "students" ? "Room check-in register / personal data" : report === "staff" ? "Exam duties, slot backups and teaching-time load" : "Confirmed allocations / linked weekly Excel sheets"}</small>
          </header>
          {report === "seating" ? <SeatingPreview exams={seatingExams} exam={seatingExam} sheet={seatingSheet} search={search} limit={limit} disabled={disabled}
            onExamChange={(id) => { setPreviewExam(id); setPreviewCrn(null); setSearch(""); setLimit(50); }}
            onCrnChange={(crn) => { setPreviewCrn(crn); setSearch(""); setLimit(50); }}
            onSearch={(value) => { setSearch(value); setLimit(50); }} onShowMore={() => setLimit((previous) => previous + 50)}
            onDownload={() => onExport({ ...exportOptions, examIds: [seatingExam.id] })} />
            : report === "overview" ? <>{mode === "board" || withAsd ? <p className="export-overview-note">{mode === "board" ? overviewNote : ""} {withAsd ? asdReferenceNote : ""}</p> : null}<div className="export-view-toggle"><button type="button" aria-pressed={mode === "board"} onClick={() => setMode("board")}>Week board</button>
            <button type="button" aria-pressed={mode === "table"} onClick={() => setMode("table")}>Chronological list</button></div>
            {mode === "board" ? <WeekBoard model={current} startDate={startDate} week={activeWeek} /> : <PreviewTable label="Exam overview preview"
              className="export-table--chronological" columns={overviewColumns} rows={overviewRows(current)} rowClassNames={overviewRowClasses(current)} />}</> : <>
            {report === "complete" ? <WorkbookMap model={current} week={activeWeek} /> : null}
            {report === "staff" ? <details className="export-workloads" open><summary>Workload balance / {previewLabel}</summary>
              {overall ? <p>Totals across included weeks: {selectedWeeks.map((week) => "Week " + week).join(", ")}. Include every week to match the Resource Review Pool. Invigilators with no assignments are shown too.</p> : null}
              <PreviewTable label="Workload balance" columns={["Invigilator", "Exam load", "Backup load", "Teaching duties", "Extra duties"]} rows={loadRows(current, overall)} />
              <p>Teaching duties count as zero extra load. Backup load is balanced separately.</p></details> : null}
            <label className="export-search"><span>Search this preview</span><input type="search" value={search} placeholder={report === "students" ? "Student, course or room..." : report === "staff" ? "Invigilator, course or room..." : "Course, room or invigilator..."}
              onChange={(event) => { setSearch(event.target.value); setLimit(50); }} /><small>Search affects the preview only, not the download or print.</small></label>
            <PreviewTable label={view.label + " preview"} limit={limit}
              columns={report === "students" ? ["Student ID", "Name", "Course", "CRN", "Day / time", "Room"] : report === "staff" ? ["Invigilator", overall ? "Week / date / time" : "Day / time", "Duty", "Course", "Room", "Load"] : ["Course", "Day / time", "Room", "Students", "Invigilators", "Room limit"]}
              rows={report === "students" ? studentRows(filtered) : report === "staff" ? staffRows(filtered, overall) : ledgerRows(filtered)} />
            {records > limit ? <button type="button" className="export-show-more" onClick={() => setLimit((previous) => previous + 50)}>Show 50 more ({records - limit} remaining)</button> : null}
          </>}
          <footer className="export-preview__footer">Previewing {overall ? "all included weeks" : previewLabel}. {report === "seating" ? "Download This Exam includes only the selected exam, with all its CRNs. Course files and print include all " : "Downloads and print include all "}{selectedWeeks.length} selected {selectedWeeks.length === 1 ? "week" : "weeks"}. {withAsd ? asdReferenceNote : "ASD exams are excluded."}</footer>
        </>}
      </main>
    </div>
    {selected ? <PrintReport report={report} overviewMode={mode} model={selected} sessions={sessions} catalog={catalog} plan={plan} startDate={startDate} asdExams={asdExams} /> : null}
  </section>;
}
