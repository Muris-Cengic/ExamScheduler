import { backupTarget, isTeachingTimeDuty, PREFERRED_STANDBY_COUNT, resourceAssignmentAnchor, resourceChoiceReason, resourceChoiceSummary, resourceSessionLabel, standbyAvailability } from "./resources.js";
import ResourcePool from "./ResourcePool.jsx";

function InvigilatorOptions({ people, session, sessions, plan, owner, fixedId = "", placeholder }) {
  const choices = people.map((person) => {
    const reason = fixedId && person.id !== fixedId ? "Lab instructor required"
      : resourceChoiceSummary(person, "invigilator", session, sessions, plan, owner);
    return { person, reason, teaching: !reason && isTeachingTimeDuty(person, session, sessions) };
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

export default function ResourceAssignment({ sessions, catalog, plan, validation, onAssign, onPoolChange, onAllocationChange, onBackupChange, onDistributionChange }) {
  const roomOptions = (session, owner, fixedId = "") => catalog.rooms.map((resource) => {
    const reason = fixedId && resource.id !== fixedId ? "Original lab resource required" : resourceChoiceReason(resource, "room", session, sessions, plan, owner);
    return <option key={resource.id} value={resource.id} disabled={Boolean(reason)}>{resource.name}{reason ? " - " + reason : ""}</option>;
  });
  return (
    <section className="resource-panel">
      <div className="resource-panel__heading">
        <div><h2>Review Rooms &amp; Invigilators</h2><p>Students are balanced across rooms of at most 25, except the last room stays at 15 when that saves an invigilator without overfilling the others. You can explicitly approve up to 27 for an exam to avoid a small overflow room. One invigilator up to 15 students; two above 15. Backups cover the time slot, not every room.</p></div>
        <button type="button" className="primary-action" onClick={onAssign}>Auto-Assign Resources</button>
      </div>
      <p className="resource-panel__note">Only the first sheet of the CRN list supplies courses, rooms, staff and teaching availability. Its regular classes block resources for the full exam duration; all other sheets are ignored. A single-CRN exam replaces its lab, keeping that room and lab instructor (the second instructor when listed). A morning lab ending at 09:50 is treated as ending at 10:00 for the exam and teaching-time load. Extra rooms and staff cover overflow. Other classes on the first sheet remain blocked. Availability outside listed classes is assumed within timetable hours.</p>
      <div className="resource-panel__summary" role="status">
        <span>{catalog.rooms.filter((room) => room.enabled).length} rooms</span>
        <span>{catalog.invigilators.filter((person) => person.enabled).length} invigilators</span>
        <strong>{validation.complete ? "Ready to export" : `${validation.issues.length} issues to resolve`}</strong>
      </div>
      <ResourcePool catalog={catalog} sessions={sessions} plan={plan} onPoolChange={onPoolChange} />
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
        {sessions.map((session) => {
          const standby = validation.complete ? standbyAvailability(session, sessions, catalog, plan) : null;
          return <section className="resource-session" key={session.id} id={resourceAssignmentAnchor(session.id)}>
            <h3>{resourceSessionLabel(session)}</h3>
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
