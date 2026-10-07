# Exam Scheduling Helper

React/Vite app for preparing a main exam timetable around an optional ASD schedule.

## Run locally

```sh
npm install
npm run dev
```

## Import enrolment

In **Load Data**, choose the exam start date and upload an Excel or CSV file.
Supported inputs:

- A registration report with `STUDENT_ID : ..., Student Name : ...` identity rows,
  followed by `Crn`, `Course Code` and `Title` course rows (the 26-27 S1 format).
  Subtotal rows are ignored; student identity carries into their course rows.
- A flat table with `Student ID`, `Student Name`, `Course Code`, `Course Title`
  (or `Title`), and `CRN` columns. Existing Banner column names are also supported.

Courses are grouped by code, students are deduplicated, and CRNs are retained for reports.
The new registration report contains no instructor names; the department CRN list supplies them.

## Select department exams

After enrolment, **Department Exams** accepts an Excel/CSV CRN list. For the supplied
`data/input/26-27 S1/CRN List 26-27 S1.xlsx`, the first sheet is `ISET CRNs`.
Only the **first sheet** is read, regardless of its name. All other sheets are
ignored for department courses, resource pools and teaching availability.
Both headerless Banner data and named-column layouts are supported on the first sheet.

- Only listed CRNs and their enrolled students are included in the main course pool,
  calculations and reports. Full enrolment remains available for ASD student-conflict checks.
- All `OL`/lab meetings are identified by CRN, weekday, time and room. Lectures,
  tutorials, field training and GP meetings are not treated as labs.
- Review the **Exam?** checkbox for each course. Graduation projects, capstones,
  internships and courses labelled OCT in their title default to no exam; all
  choices are editable. Turning an exam off removes
  its existing main placement, but does not alter the independent ASD timetable.
- Missing CRN enrolments are shown; courses with no enrolled students cannot be scheduled.
- Choose the exam start date and add exam weeks before scheduling. The CRN upload
  extends timetable hours to at least 18:00 to include the evening exam window.

## Select resources

The wizard runs **Load Data > Department Exams > ASD (Optional) > Select Resources
> Build Main Timetable > Review Resources > Export**. After creating/loading ASD,
or skipping it, choose the active rooms and named invigilators before generating
the main timetable. The pool comes from the CRN list's first sheet and shows each
resource's recurring class commitments. Excluded resources cannot be used.
At least one room and two invigilators must be selected to continue; this does
not imply that they are available for any particular exam.

The scheduler targets **two available standby invigilators per slot**, including
that slot's assigned backups. This is reserved capacity, not two mandatory
backup assignments. Unassigned standby people add no load and no report duties.
You can review the pool again from scheduling and adjust it after generation.

## Automatic scheduling

After selecting resources, use **Auto-Schedule Remaining Exams**:

- Multiple listed CRNs: Monday-Thursday, 12:00-13:00 or 17:00-18:00.
  Friday is restricted to **09:00-10:00** and **10:30-11:30** for all main exams,
  including lab exams. Friday lab exams must also fit their listed lab hours.
  Noon, evening and other Friday starts are blocked in automatic and manual
  scheduling. Invalid existing Friday placements must be moved before export.
  The independent ASD timetable is not restricted by this department rule.
- One listed CRN: a slot fully inside its lab meeting on that weekday. Without a
  listed lab, the common windows are used instead.
  Labs starting before 09:00 use the second-hour window, 09:00-10:00, rather than
  08:00-09:00. A listed 09:50 morning lab ending is treated as 10:00 for the exam
  and the lab instructor's teaching-time allowance. The full configured exam must
  fit; shorter labs are not extended and exam duration is not shortened.
  Other lab times are unchanged, and unrelated classes and bookings still block
  rooms and instructors throughout the full exam.
- The draft uses only the configured exam weeks, hours and exam duration. It avoids
  overlapping students across main/ASD schedules, more than two exams per student
  per day, and room/staff shortages. Every proposed placement is trial-assigned
  using the selected pool, full exam duration, class availability, fixed lab
  resources and overlapping bookings, including assigned slot backups.
- Placement prefers the **earliest week, then weekday, then time**: Monday noon
  comes before Monday evening, Tuesday noon or Friday morning. Multiple exams
  can share an earlier slot when student, room and staff constraints allow.
  Timetable load balancing no longer pushes exams to a later date. Lab exams
  still use only their permitted lab window.
- The earliest valid slot retaining two standby people is preferred. If none
  exists, the earliest slot with valid required backup coverage is used and a
  standby-shortfall warning is shown. Actual assigned-backup limits are unchanged.
  A draft includes the checked room/invigilator plan, ready for resource review.
- Existing placements are preserved. Courses that cannot fit remain in the pool,
  with reasons shown. Add a week, adjust settings or place them manually, then rerun.
  This is a greedy draft, not a guarantee of an optimal or complete timetable.
- Main JSON saves include the selected CRN list, lab metadata and exam choices.
  Older snapshots are intentionally unsupported.

## Assign resources

Before **Export**, continue to **Review Resources**. Automatic drafts already
include checked assignments; review and adjust them with the room/staff selectors.
Use **Auto-Assign Resources** after manual scheduling or pool changes. Pool edits
invalidate the old plan until it is reassigned/reviewed. Each valid slot shows
available standby capacity separately from assigned backups; capacity below two
is highlighted without changing existing backup report rules.

- Rooms and named invigilators come from the CRN workbook's first sheet. Both
  primary and second instructors enter one pool; there are no invigilator types.
  Resource-pool checkboxes can exclude rooms or staff from exam duty.
- Invigilator and backup selectors list available people first, alphabetically.
  Unavailable people appear below in a disabled group with a short reason:
  a class, another exam or backup duty, exclusion from the pool, or the required
  lab-instructor restriction. The current selection remains available in its own
  selector unless another conflict blocks it; nobody is silently reassigned.
- Availability is checked against regular classes on the **first sheet only**.
  Commitments from other sheets do not block resources. Meetings repeat weekly.
  Outside listed classes, availability is assumed within timetable hours; the file
  does not contain working-hour/leave data. Classes with days but an unspecified
  time conservatively block those days.
- Each exam is split into separate rooms of at most **25 students** (a lower limit
  can be configured). The minimum required rooms receive balanced student counts,
  except the last room stays at **15 students** when the other rooms can absorb
  the remainder within capacity and this saves an invigilator. Those other rooms
  stay balanced: 55 students use **20/20/15**, needing five room invigilators,
  instead of six for 19/18/18. If the last room must have more than 15 anyway,
  the usual balanced split is retained: 66 students use 22/22/22. Already balanced
  one-invigilator rooms, such as 13/13 for 26 students, are unchanged.
  Course rosters remain separate and every student appears exactly once.
  Every room needs one invigilator for 1-15 students, or two above 15.
- When a few students would force one more room at the normal limit, the Resources
  step asks whether to keep it, merge the overflow into one existing room, or
  distribute it among that exam's remaining rooms. Only feasible choices
  are offered: for 52 students, keep 19/18/15, merge to 27/25, or distribute 26/26.
  The default remains 25; **27 is a per-exam exception requiring an explicit choice**,
  never an automatic increase. A lower configured capacity is respected and does
  not offer this exception. Different exams are not mixed, and no room can exceed
  27. Rooms with 26-27 students still need two invigilators.
  The distribute choice uses the same last-room staffing rule within the approved
  27-student limit; an explicit merge retains its proposed merged-room layout.
  Existing room/staff assignments are retained where possible; returning to the
  normal limit may require assigning the restored room. Decisions are saved in JSON
  and shown in the report; changing one exam does not reassign unrelated exams.
- Single-CRN lab exams replace their original lab class. The first room uses the
  original lab room and the **second instructor** as its first invigilator when
  listed, otherwise the primary instructor. Additional rooms and available
  invigilators cover larger cohorts. Unrelated classes still block availability.
- Rooms and staff cannot be double-booked, including exams with different start
  times that overlap. Backups cannot also invigilate an exam room at that time.
- There is no primary/secondary distinction for exam invigilators: either course
  instructor can invigilate when available. Extra invigilation loads are balanced
  overall and by time of day; backup loads are balanced separately. An exam duty
  entirely within the instructor's replaced lab hours adds **no extra load**,
  including when assigned to another exam or as backup during those hours.
  Such duties still reserve the instructor for the full exam duration. Partial
  overlap alone does not exempt a duty. Fixed lab assignments and availability
  take precedence. Resource pools and reports show teaching-hour duties separately.
  There is at least one backup per time slot, with a maximum of 40% of the slot's
  room count (rounded down, with a minimum of one).
- Missing resources and conflicts are shown explicitly; **export stays blocked**
  until assignments are valid. Issues show the week/day/exam time, exam room and
  student count, the named resource and conflicting class/CRN or exam, and a
  suggested resolution. **Review assignment** jumps to the affected room or slot.
  A required lab instructor's other teaching is not canceled by the lab exam;
  overlapping entries in the CRN list must be resolved or corrected, or another
  listed lab time used if available. Excluded resources can be included again in
  the resource pool. Flexible room/staff conflicts can use another available resource.
- Reports use the reviewed room names and invigilators, not generated placeholders.
  ASD exams remain excluded from resource reports. JSON saves retain the pool, availability and assignments;
  changing the timetable invalidates stale resource assignments.

## Export workspace

After resources are confirmed, the **Export** step offers four reports with live
previews:

- **Complete report**: the existing linked Excel report, including weekly
  invigilators, resource pools and daily student sheets. The preview includes a
  workbook-sheet map and searchable room ledger.
- **Exam overview**: a shareable course timetable with assigned rooms and student
  counts, but no student names or IDs. Preview it as a five-day board or a
  chronological list. Both views, Excel, CSV and print include the primary
  invigilators needed across the exam's rooms (including the lab instructor,
  excluding slot backups) and whether the exam replaces its lab session. Room
  names are compact, for example `PAD / P-B-4F / 13` becomes `P-B-4F/13`.
  Resource identities and the complete report's linked room references are
  unchanged. Enable **Include ASD exams** to add the ASD reference schedule to this
  report only (off by default). ASD cards and rows are muted in both views and
  print/PDF; Excel and CSV identify them as **ASD (reference)** in the **Schedule**
  column. ASD resource/student fields are blank in files and marked not applicable
  in the preview. Department resource and student totals stay unchanged.
- **Staff duties**: each exam duty and slot-level backup, plus a separate workload
  summary in Excel and print. Teaching-time duties add zero extra load; backup
  duties remain separate. Use the **Overall** preview tab for totals and duties
  across all included weeks, or a week tab for that week's figures. Overall also
  shows invigilators with zero assignments. Include every populated week to match
  the Resource Review Pool totals.
- **Student room lists**: a searchable check-in register of student sittings,
  courses, times and confirmed rooms. Contains student personal data.

Choose which populated exam weeks to include, then select **One workbook** or
**Separate weekly files**. The combined complete report adds a navigable
**Schedule Index** and preserves all week-qualified sheet names and formula links.
Multiple weekly files download as one ZIP; a single selected week downloads
directly. Empty weeks are excluded. With **Include ASD exams** enabled, the Exam
overview also offers weeks containing only ASD exams. ASD never appears in the
Complete report, Staff duties or Student room lists, including their print/PDF
exports.

The three focused reports also support UTF-8 **CSV** as one file or weekly files.
Staff CSV contains the duties only; choose Excel for its additional workload
summary. Formula-like text is neutralised in CSV to avoid accidental spreadsheet
formula execution.

**Print / Save PDF** opens the browser's print dialog with a clean, landscape
report containing all included weeks. Choose the browser's PDF destination to
save it. Preview week tabs and search affect only the on-screen preview, not
downloads or print; the print layout includes every selected record.
For **Exam overview**, choose **Week board** or **Chronological list** in the
preview before printing. The PDF uses that layout for every included week.
The board keeps five weekday columns, compact rooms, staffing requirements and
lab-time badges. Each week starts on a new page; busy weeks continue onto further
pages with repeated day headers rather than clipping exam cards. Excel and CSV
downloads remain tabular regardless of the preview layout.
All export formats require complete, valid resource assignments. Share reports
containing student names and IDs only with authorised recipients.

## Import ASD

In **ASD (Optional)**, choose **Load ASD Excel / JSON**.
Excel/CSV schedules need `Course`, `Department`, `Day` and `Time` columns.

- Only rows with Department `ASD` are imported.
- Course titles or codes are matched to loaded courses, ignoring case, spaces and punctuation.
- Courses without loaded enrolment are skipped and listed in the import notice.
  Ambiguous matches or conflicting sessions produce an error.
- Dates such as `Monday, 19 October` use the selected exam start date's year.
  Set the correct year in Load Data before importing. Explicit years and Excel dates are supported.
- Time ranges such as `12:00 - 1:00 PM` and `10:30 - 11:30 AM` are supported.
  The timetable expands its weeks/hours and uses 30-minute slots when needed.
- Imported exam durations are retained for timetable display and cross-schedule student conflicts.
  ASD exams appear gray, leave the main course pool, and are excluded from main
  staffing totals. They can optionally be included in Exam overview exports only.
- ASD and full timetable JSON saves retain the imported durations.

## Verify and publish

```sh
npm test
npm run lint
npm run build
npm run deploy
```

GitHub Pages uses the `gh-pages` branch and the `/ExamScheduler/` base path.
Local input workbooks under `data/input/` are not bundled or published.
