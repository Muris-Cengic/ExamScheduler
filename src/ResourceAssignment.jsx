import { useState } from "react";
import { backupTarget, isTeachingTimeDuty, PREFERRED_STANDBY_COUNT, resourceAssignmentAnchor, resourceChoiceReason, resourceChoiceSummary, standbyAvailability } from "./resources.js";
import { formatReportDate, reportDate, reportTimeRange, REPORT_DAYS } from "./exportReports.js";
import ResourcePool from "./ResourcePool.jsx";

function InvigilatorOptions({ people, session, sessions, plan, owner, fixedId = "", placeholder }) {
  const choices = people.map((person) => {
    const reason = fixedId && person.id !== fixedId ? "Lab instructor required"
      : resourceChoiceSummary(person, "invigilator", session, sessions, plan, owner);
    return { person, reason, teaching: !owner.includes("/backup/") && !reason && isTeachingTimeDuty(person, session) };
  }).sort((a, b) => a.person.name.localeCompare(b.person.name));
  const available = choices.filter((choice) => !choice.reason);
  const unavailable = choices.filter((choice) => choice.reason);
  return <>
    <option value="">{placeholder}{!available.length ? " (none available)" : ""}</option>
    {available.length ? <optgroup label={"Available (" + available.length + ")"}>
      {available.map(({ person, teaching }) => <option key={person.id} value={person.id}>
        {person.name}{teaching ? " - teaching hours, no extra load" : ""}
      </option>)}
    </optgroup> : null}
    {unavailable.length ? <optgroup label={"Unavailable (" + unavailable.length + ")"} disabled>
      {unavailable.map(({ person, reason }) => <option key={person.id} value={person.id} disabled>{person.name} - {reason}</option>)}
    </optgroup> : null}
  </>;
}

export default function ResourceAssignment({ sessions, catalog, plan, validation, startDate, onAssign, onPoolChange, onAllocationChange, onBackupChange, onDistributionChange }) {
  const [weekFilter, setWeekFilter] = useState("");
  const [dayFilter, setDayFilter] = useState("");
  const [sessionFilter, setSessionFilter] = useState("");
  const weeks = [...new Set(sessions.map((session) => session.week))].sort((a, b) => a - b);
  const activeWeek = weeks.some((week) => String(week) === weekFilter) ? weekFilter : "";
  const weeklySessions = sessions.filter((session) => !activeWeek || String(session.week) === activeWeek);
  const days = REPORT_DAYS.filter((day) => weeklySessions.some((session) => session.day === day));
  const activeDay = activeWeek === weekFilter && days.includes(dayFilter) ? dayFilter : "";
  const dailySessions = weeklySessions.filter((session) => !activeDay || session.day === activeDay);
  const slotKey = (session) => `${session.start}/${session.end}`;
  const slots = [...new Map(dailySessions.map((session) => [slotKey(session), session])).values()]
    .sort((a, b) => a.start - b.start || a.end - b.end);
  const activeSession = activeWeek === weekFilter && activeDay === dayFilter && slots.some((session) => slotKey(session) === sessionFilter) ? sessionFilter : "";
  const visibleSessions = dailySessions.filter((session) => !activeSession || slotKey(session) === activeSession);
  const reviewSession = (id) => {
    const session = sessions.find((item) => item.id === id);
    if (session) { setWeekFilter(String(session.week)); setDayFilter(session.day); setSessionFilter(slotKey(session)); }
  };
  const roomOptions = (session, owner, fixedId = "") => catalog.rooms.map((resource) => {
    const reason = fixedId && resource.id !== fixedId ? "Original lab resource required" : resourceChoiceReason(resource, "room", session, sessions, plan, owner);
    return <option key={resource.id} value={resource.id} disabled={Boolean(reason)}>{resource.name}{reason ? " - " + reason : ""}</option>;
  });
  return (
    <section className="resource-panel">
      <div className="resource-panel__heading">
        <div><h2>Review Rooms &amp; Invigilators</h2><p>25 students per room. One invigilator up to 15 students; two above 15. Backups cover the slot.</p></div>
        <button type="button" className="primary-action" onClick={onAssign} disabled={!sessions.length}>Auto-Assign Resources</button>
      </div>
      <details className="schedule-rules"><summary>Room and availability rules</summary>
        <p>Students are balanced, keeping the last room at 15 when that saves an invigilator. Up to 27 per room requires approval for that exam.</p>
        <p>The first CRN sheet supplies resources and class commitments. Tutorials and unchecked CRN Info courses do not block invigilators or extend their daily span; room commitments are unchanged. Single-CRN lab exams keep their lab room and instructor; listed :50 lab endings count through the next hour. Availability outside listed classes is assumed.</p>
        <p>An exam fully within the invigilator's listed lab or tutorial hours adds no extra load, even for unchecked courses. Lecture hours do not qualify. For load credit, :50 endings count through the next hour. Other class and exam conflicts still block assignment.</p>
        <p>Imported CyberSkills sessions block overlapping exam and backup assignments on their actual dates and extend daily start/end spans. They never add invigilation load or grant teaching-hour load credit.</p>
        <p>Auto-assignment balances exam duties overall and by time slot, preferring at most eight hours from the first class or exam duty to the last, including gaps. Backups do not add workload or extend this span. Longer days are allowed when needed; fixed lab instructors remain required.</p>
      </details>
      <div className="resource-panel__summary" role="status">
        <span>{catalog.rooms.filter((room) => room.enabled).length} rooms</span>
        <span>{catalog.invigilators.filter((person) => person.enabled).length} invigilators</span>
        <strong>{validation.complete ? "Ready to export" : `${validation.issues.length} issues to resolve`}</strong>
      </div>
      <ResourcePool catalog={catalog} sessions={sessions} plan={plan} onPoolChange={onPoolChange} />
      <fieldset className="resource-session-filters"><legend>Filter sessions</legend>
        <label>Week<select aria-label="Filter assignment week" value={activeWeek} disabled={!sessions.length}
          onChange={(event) => { setWeekFilter(event.target.value); setDayFilter(""); setSessionFilter(""); }}>
          <option value="">All weeks</option>{weeks.map((week) => <option key={week} value={String(week)}>Week {week}</option>)}
        </select></label>
        <label>Day<select aria-label="Filter assignment day" value={activeDay} disabled={!weeklySessions.length}
          onChange={(event) => { setWeekFilter(activeWeek); setDayFilter(event.target.value); setSessionFilter(""); }}>
          <option value="">All days</option>{days.map((day) => <option key={day} value={day}>{day}</option>)}
        </select></label>
        <label>Session<select aria-label="Filter assignment session" value={activeSession} disabled={!dailySessions.length}
          onChange={(event) => { setWeekFilter(activeWeek); setDayFilter(activeDay); setSessionFilter(event.target.value); }}>
          <option value="">All sessions</option>{slots.map((session) => <option key={slotKey(session)} value={slotKey(session)}>
            {reportTimeRange(session.start, session.end)}</option>)}
        </select></label>
        <button type="button" disabled={!activeWeek && !activeDay && !activeSession} onClick={() => { setWeekFilter(""); setDayFilter(""); setSessionFilter(""); }}>Clear filters</button>
        <span role="status">Showing {visibleSessions.length} of {sessions.length} sessions</span>
        <small>Pool totals, issues and auto-assignment cover all sessions.</small>
      </fieldset>
      {validation.issues.length > 0 ? <details className="resource-issues" open>
        <summary>Resource Issues ({validation.issues.length})</summary>
        <p>Resolve these issues before export. Teaching commitments come only from the CRN list's first sheet; replacing an exam's lab does not cancel other classes listed there.</p>
        <ul>{validation.issues.map((issue, index) => <li key={index} className="resource-issue">
          <div className="resource-issue__context">{issue.context}</div>
          {issue.exam ? <div className="resource-issue__exam">{issue.exam}</div> : null}
          <strong>{issue.title}</strong>
          <p>{issue.message}</p>
          <p className="resource-issue__action"><strong>How to resolve: </strong>{issue.action}</p>
          {issue.sessionId ? <a href={`#${encodeURIComponent(resourceAssignmentAnchor(issue.roomId || issue.sessionId))}`} aria-label={`Review ${issue.exam || issue.context}`}
            onClick={() => reviewSession(issue.sessionId)}>Review assignment</a> : null}
        </li>)}</ul>
      </details> : null}
      <div className="resource-sessions">
        {!sessions.length ? <p>No exam sessions scheduled.</p> : null}
        {visibleSessions.map((session) => {
          const standby = validation.complete ? standbyAvailability(session, sessions, catalog, plan) : null;
          return <section className="resource-session" key={session.id} id={resourceAssignmentAnchor(session.id)}>
            <h3>{`Week ${session.week} / ${formatReportDate(reportDate(startDate, session.week, session.day))} / ${session.day} / ${reportTimeRange(session.start, session.end)}`}</h3>
            <p className={standby && standby.count < PREFERRED_STANDBY_COUNT ? "resource-standby resource-standby--short" : "resource-standby"}>
              {standby ? "Available standby capacity: " + standby.count + " (target " + PREFERRED_STANDBY_COUNT + "), including this slot's assigned backups."
                : "Confirm valid room and staff assignments to assess standby capacity."} Other standby people remain unassigned; only assigned duties appear in reports.
            </p>
            {session.roomDecisions.map((decision) => <div className="resource-room-decision" key={decision.courseId}>
              <h4>{decision.code}: Avoid an extra room?</h4>
              <p>{decision.students} students include {decision.overflow} beyond {(decision.options[0].sizes.length - 1) * 25} normal seats. Choose whether to keep the extra room or approve up to 27 per room for this exam only. Different exams are never mixed.</p>
              <div className="resource-room-decision__options" role="group" aria-label={`Room distribution for ${decision.code}`}>
                {decision.options.map((option) => <button type="button" key={option.value} aria-pressed={decision.choice === option.value}
                  aria-label={`${option.label} for ${decision.code}`} onClick={() => onDistributionChange(decision.courseId, option.value)}>
                  <strong>{option.label}</strong>
                  <span>{option.sizes.length} {option.sizes.length === 1 ? "room" : "rooms"}: {option.sizes.join(" / ")} students</span>
                  <small>{option.value === "standard" ? "No capacity exception" : "Approve a maximum of 27 per room"}</small>
                </button>)}
              </div>
              <p className="resource-room-decision__status">{decision.choice === "standard" ? "Normal room limit retained. No exception approved." : "Exception approved for this exam. Existing room/staff assignments are retained where possible; review any remaining resource issues."}</p>
            </div>)}
            <div className="resource-table-wrap"><table><thead><tr><th>Exam</th><th>Students</th><th>Room</th><th>Invigilators</th></tr></thead><tbody>
              {session.rooms.map((room) => {
                const allocation = plan?.allocations?.[room.id] || { roomId: "", invigilatorIds: [] };
                return <tr key={room.id} id={resourceAssignmentAnchor(room.id)}>
                  <td><strong>{room.code}</strong><br />{room.title}{room.fixedRoomId ? <small className="resource-lab-label">Original lab room &amp; instructor</small> : null}</td>
                  <td>{room.students.length}{room.distributionChoice !== "standard" ? <small className="resource-capacity-label">Approved up to 27</small> : null}</td>
                  <td><select aria-label={`Room for ${room.id}`} value={allocation.roomId} onChange={(event) => onAllocationChange(room.id, { ...allocation, roomId: event.target.value })}><option value="">Select room</option>{roomOptions(session, room.id, room.fixedRoomId)}</select></td>
                  <td><div className="resource-invigilators">{Array.from({ length: room.requiredInvigilators }, (_, index) => <select key={index} aria-label={`Invigilator ${index + 1} for ${room.id}`} value={allocation.invigilatorIds[index] || ""} onChange={(event) => {
                    const ids = Array.from({ length: room.requiredInvigilators }, (_, i) => allocation.invigilatorIds[i] || "");
                    ids[index] = event.target.value;
                    onAllocationChange(room.id, { ...allocation, invigilatorIds: ids });
                  }}><InvigilatorOptions people={catalog.invigilators} session={session} sessions={sessions} plan={plan}
                    owner={`${room.id}/invigilator/${index}`} fixedId={index === 0 ? room.fixedInvigilatorId : ""} placeholder={"Select invigilator " + (index + 1)} /></select>)}</div></td>
                </tr>;
              })}
            </tbody></table></div>
            <div className="resource-backups"><strong>Slot backups (at least 1, at most {backupTarget(session.rooms.length)}):</strong>{Array.from({ length: backupTarget(session.rooms.length) }, (_, index) => <select key={index} aria-label={`Backup ${index + 1} for ${session.id}`} value={plan?.backups?.[session.id]?.[index] || ""} onChange={(event) => {
              const ids = Array.from({ length: backupTarget(session.rooms.length) }, (_, i) => plan?.backups?.[session.id]?.[i] || "");
              ids[index] = event.target.value;
              onBackupChange(session.id, ids);
            }}><InvigilatorOptions people={catalog.invigilators} session={session} sessions={sessions} plan={plan}
              owner={`${session.id}/backup/${index}`} placeholder={index === 0 ? "Select required backup" : "No additional backup"} /></select>)}</div>
          </section>;
        })}
      </div>
    </section>
  );
}
