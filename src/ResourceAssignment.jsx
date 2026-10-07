import { backupTarget, clock, invigilatorWorkloads, isTeachingTimeDuty, resourceAssignmentAnchor, resourceChoiceReason, resourceSessionLabel } from "./resources.js";

export default function ResourceAssignment({ sessions, catalog, plan, validation, onAssign, onPoolChange, onAllocationChange, onBackupChange, onDistributionChange }) {
  const workloads = invigilatorWorkloads(catalog, sessions, plan);
  const options = (resources, kind, session, owner, fixedId = "") => resources.map((resource) => {
    const reason = fixedId && resource.id !== fixedId ? "Original lab resource required" : resourceChoiceReason(resource, kind, session, sessions, plan, owner);
    const teaching = kind === "invigilator" && isTeachingTimeDuty(resource, session, sessions);
    return <option key={resource.id} value={resource.id} disabled={Boolean(reason)}>{resource.name}{reason ? ` - ${reason}` : teaching ? " - teaching hours, no extra load" : ""}</option>;
  });
  return (
    <section className="resource-panel">
      <div className="resource-panel__heading">
        <div><h2>Assign Rooms &amp; Invigilators</h2><p>Normally, students are balanced across rooms of at most 25. You can explicitly approve up to 27 for an exam to avoid a small overflow room. One invigilator up to 15 students; two above 15. Backups cover the time slot, not every room.</p></div>
        <button type="button" className="primary-action" onClick={onAssign}>Auto-Assign Resources</button>
      </div>
      <p className="resource-panel__note">Only the first sheet of the CRN list supplies courses, rooms, staff and teaching availability. Its regular classes block resources for the full exam duration; all other sheets are ignored. A single-CRN exam replaces its lab, keeping that room and lab instructor (the second instructor when listed). A morning lab ending at 09:50 is treated as ending at 10:00 for the exam and teaching-time load. Extra rooms and staff cover overflow. Other classes on the first sheet remain blocked. Availability outside listed classes is assumed within timetable hours.</p>
      <div className="resource-panel__summary" role="status">
        <span>{catalog.rooms.filter((room) => room.enabled).length} rooms</span>
        <span>{catalog.invigilators.filter((person) => person.enabled).length} invigilators</span>
        <strong>{validation.complete ? "Ready to export" : `${validation.issues.length} issues to resolve`}</strong>
      </div>
      <details className="resource-pool">
        <summary>Review Resource Pool</summary>
        <p>All instructors enter the same invigilator pool. Extra invigilation duties are balanced overall and by time slot; backups are balanced separately. Duties entirely within an instructor's replaced lab hours do not add to their load.</p>
        <div className="resource-pool__tables">
          <div className="resource-table-wrap"><table><thead><tr><th>Include</th><th>Invigilator</th><th>Invigilation Load</th><th>Backup Load</th><th>During Teaching Hours</th><th>Class Times</th></tr></thead><tbody>
            {workloads.map((person) => <tr key={person.id}>
              <td><input type="checkbox" aria-label={`Include ${person.name}`} checked={person.enabled} onChange={(event) => onPoolChange("invigilators", person.id, { enabled: event.target.checked })} /></td>
              <td>{person.name}</td>
              <td>{person.exam}</td><td>{person.backup}</td><td>{person.teaching}</td>
              <td><details><summary>{person.busy.length} meetings</summary>{person.busy.map((entry, index) => <div key={index}>{entry.days.map((day) => day.slice(0, 3)).join(", ")} {entry.unknownTime ? "Time unspecified (day blocked)" : `${clock(entry.start)}-${clock(entry.end)}`}: {entry.code} (CRN {entry.crn})</div>)}</details></td>
            </tr>)}
          </tbody></table></div>
          <div className="resource-table-wrap"><table><thead><tr><th>Include</th><th>Room</th><th>Class Times</th></tr></thead><tbody>
            {catalog.rooms.map((room) => <tr key={room.id}>
              <td><input type="checkbox" aria-label={`Include ${room.name}`} checked={room.enabled} onChange={(event) => onPoolChange("rooms", room.id, { enabled: event.target.checked })} /></td>
              <td>{room.name}</td><td><details><summary>{room.busy.length} meetings</summary>{room.busy.map((entry, index) => <div key={index}>{entry.days.map((day) => day.slice(0, 3)).join(", ")} {entry.unknownTime ? "Time unspecified (day blocked)" : `${clock(entry.start)}-${clock(entry.end)}`}: {entry.code} (CRN {entry.crn})</div>)}</details></td>
            </tr>)}
          </tbody></table></div>
        </div>
      </details>
      {validation.issues.length > 0 ? <details className="resource-issues" open>
        <summary>Resource Issues ({validation.issues.length})</summary>
        <p>Resolve these issues before export. Teaching commitments come only from the CRN list's first sheet; replacing an exam's lab does not cancel other classes listed there.</p>
        <ul>{validation.issues.map((issue, index) => <li key={index} className="resource-issue">
          <div className="resource-issue__context">{issue.context}</div>
          {issue.exam ? <div className="resource-issue__exam">{issue.exam}</div> : null}
          <strong>{issue.title}</strong>
          <p>{issue.message}</p>
          <p className="resource-issue__action"><strong>How to resolve: </strong>{issue.action}</p>
          {issue.sessionId ? <a href={`#${encodeURIComponent(resourceAssignmentAnchor(issue.roomId || issue.sessionId))}`} aria-label={`Review ${issue.exam || issue.context}`}>Review assignment</a> : null}
        </li>)}</ul>
      </details> : null}
      <div className="resource-sessions">
        {sessions.map((session) => <section className="resource-session" key={session.id} id={resourceAssignmentAnchor(session.id)}>
          <h3>{resourceSessionLabel(session)}</h3>
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
                <td><select aria-label={`Room for ${room.id}`} value={allocation.roomId} onChange={(event) => onAllocationChange(room.id, { ...allocation, roomId: event.target.value })}><option value="">Select room</option>{options(catalog.rooms, "room", session, room.id, room.fixedRoomId)}</select></td>
                <td><div className="resource-invigilators">{Array.from({ length: room.requiredInvigilators }, (_, index) => <select key={index} aria-label={`Invigilator ${index + 1} for ${room.id}`} value={allocation.invigilatorIds[index] || ""} onChange={(event) => {
                  const ids = Array.from({ length: room.requiredInvigilators }, (_, i) => allocation.invigilatorIds[i] || "");
                  ids[index] = event.target.value;
                  onAllocationChange(room.id, { ...allocation, invigilatorIds: ids });
                }}><option value="">Select invigilator {index + 1}</option>{options(catalog.invigilators, "invigilator", session, `${room.id}/invigilator/${index}`, index === 0 ? room.fixedInvigilatorId : "")}</select>)}</div></td>
              </tr>;
            })}
          </tbody></table></div>
          <div className="resource-backups"><strong>Slot backups (at least 1, at most {backupTarget(session.rooms.length)}):</strong>{Array.from({ length: backupTarget(session.rooms.length) }, (_, index) => <select key={index} aria-label={`Backup ${index + 1} for ${session.id}`} value={plan?.backups?.[session.id]?.[index] || ""} onChange={(event) => {
            const ids = Array.from({ length: backupTarget(session.rooms.length) }, (_, i) => plan?.backups?.[session.id]?.[i] || "");
            ids[index] = event.target.value;
            onBackupChange(session.id, ids);
          }}><option value="">{index === 0 ? "Select required backup" : "No additional backup"}</option>{options(catalog.invigilators, "invigilator", session, `${session.id}/backup/${index}`)}</select>)}</div>
        </section>)}
      </div>
    </section>
  );
}
