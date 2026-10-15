import { clock, invigilatorWorkloads } from "./resources.js";

function ClassTimes({ meetings }) {
  return <details><summary>{meetings.length} meetings</summary>
    {meetings.map((entry, index) => <div key={index}>
      {entry.days.map((day) => day.slice(0, 3)).join(", ")} {entry.unknownTime ? "Time unspecified (day blocked)" : clock(entry.start) + "-" + clock(entry.end)}: {entry.code} (CRN {entry.crn})
    </div>)}
  </details>;
}

export default function ResourcePool({ catalog, sessions = [], plan, onPoolChange, selectionOnly = false }) {
  const people = selectionOnly ? catalog.invigilators : invigilatorWorkloads(catalog, sessions, plan);
  const hasExternal = people.some((person) => person.externalDuties?.length);
  return <details className="resource-pool" open={selectionOnly}>
    <summary>{selectionOnly ? "Choose Resource Pool" : "Review Resource Pool"}</summary>
    <p>{selectionOnly
      ? "Tutorials and unchecked CRN Info courses do not block invigilators; room commitments still apply. Lab exams replace their own lab only. Target: two available standby invigilators per slot."
      : "Tutorials and unchecked CRN Info courses do not block invigilators. Exam load is balanced overall and by time slot. Duties fully within listed lab or tutorial hours add no extra load, even for unchecked courses. Lecture hours do not qualify. Backups are names only, not workload."}</p>
    <div className="resource-pool__tables">
      <div className="resource-table-wrap"><table><thead><tr><th>Include</th><th>Invigilator</th>
        {!selectionOnly ? <><th>Invigilation Load</th><th>During Teaching Hours</th></> : null}
        <th>Class Times</th>{hasExternal ? <th>CyberSkills</th> : null}</tr></thead><tbody>
        {people.map((person) => <tr key={person.id}>
          <td><input type="checkbox" aria-label={"Include " + person.name} checked={person.enabled}
            onChange={(event) => onPoolChange("invigilators", person.id, { enabled: event.target.checked })} /></td>
          <td>{person.name}</td>
          {!selectionOnly ? <><td>{person.exam}</td><td>{person.teaching}</td></> : null}
          <td><ClassTimes meetings={person.busy} /></td>
          {hasExternal ? <td><details><summary>{person.externalDuties?.length || 0} sessions</summary>
            {(person.externalDuties || []).map((entry, index) => <div key={index}>{entry.date} {clock(entry.start)}-{clock(entry.end)}: {entry.label}</div>)}
          </details></td> : null}
        </tr>)}
      </tbody></table></div>
      <div className="resource-table-wrap"><table><thead><tr><th>Include</th><th>Room</th><th>Class Times</th></tr></thead><tbody>
        {catalog.rooms.map((room) => <tr key={room.id}>
          <td><input type="checkbox" aria-label={"Include " + room.name} checked={room.enabled}
            onChange={(event) => onPoolChange("rooms", room.id, { enabled: event.target.checked })} /></td>
          <td>{room.name}</td><td><ClassTimes meetings={room.busy} /></td>
        </tr>)}
      </tbody></table></div>
    </div>
  </details>;
}
