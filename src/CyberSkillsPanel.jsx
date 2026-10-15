import { clock } from "./resources.js";

export default function CyberSkillsPanel({ schedule, links, people, onUpload, onRemove, onMatch, disabled }) {
  const names = Object.keys(links);
  const unmatched = names.filter((name) => links[name] === "");
  return <section className="cyber-skills-panel" aria-labelledby="cyber-skills-title">
    <header><div><h2 id="cyber-skills-title">CyberSkills Schedule <small>Optional</small></h2>
      <p>Blocks invigilator time and extends daily duty spans. No invigilation load, student or room changes.</p></div>
      <div className="cyber-skills-panel__actions"><label className="file-input">
        <input type="file" accept=".xlsx,.xls,.csv" disabled={disabled} onChange={onUpload} />
        <span>{schedule ? "Replace CyberSkills Schedule" : "Load CyberSkills Schedule"}</span>
      </label>{schedule ? <button type="button" disabled={disabled} onClick={onRemove}>Remove CyberSkills Schedule</button> : null}</div>
    </header>
    {schedule ? <>
      <p role="status"><strong>{schedule.filename}</strong> / {schedule.sessions.length} dated duties / {names.length} trainers.
        {" "}{schedule.sessions[0].date} to {schedule.sessions.at(-1).date}.
        {unmatched.length ? " " + unmatched.length + " trainer(s) need matching." : " All trainer matches resolved."}</p>
      {!people.length ? <p>Load CRN Info to match trainers to the invigilator pool.</p> : null}
      <details open={Boolean(unmatched.length)}><summary>Trainer matches</summary>
        <p>Shortened names match only when unambiguous. Explicitly ignore trainers who are not in your invigilator pool.</p>
        <div className="resource-table-wrap"><table><thead><tr><th>CyberSkills trainer</th><th>Invigilator in resource pool</th></tr></thead><tbody>
          {names.map((name) => <tr key={name}><td>{name}</td><td><select aria-label={"Match CyberSkills trainer " + name}
            disabled={disabled} value={links[name] === null ? "__ignore__" : links[name]}
            onChange={(event) => onMatch(name, event.target.value === "__ignore__" ? null : event.target.value)}>
            <option value="">Choose invigilator</option><option value="__ignore__">Not in invigilator pool (ignore)</option>
            {people.map((person) => <option key={person.id} value={person.id}>{person.name}{person.enabled ? "" : " (excluded from pool)"}</option>)}
          </select></td></tr>)}
        </tbody></table></div>
      </details>
      <details><summary>Imported sessions</summary><div className="resource-table-wrap"><table><thead><tr><th>Date</th><th>Start</th><th>End</th><th>Trainer</th><th>Session</th></tr></thead><tbody>
        {schedule.sessions.map((entry, index) => <tr key={index}><td>{entry.date}</td><td>{clock(entry.start)}</td><td>{clock(entry.end)}</td><td>{entry.instructor}</td><td>{entry.label}</td></tr>)}
      </tbody></table></div></details>
    </> : <details><summary>Expected file format</summary>
      <p>CyberSkills weekly grids: weekday columns with dates, time blocks, and Trainer rows. Only duties within ten weeks from Exam Setup are imported; older sheets and totals are ignored. Training week numbers do not become exam week numbers.</p>
      <p>Or upload an Excel/CSV table:</p><div className="resource-table-wrap"><table><thead><tr><th>Date</th><th>Time</th><th>Instructor</th><th>Session</th></tr></thead>
        <tbody><tr><td>2026-10-19</td><td>08:00 - 09:00</td><td>Dr. Alex Smith</td><td>Computer Networks</td></tr></tbody></table></div>
    </details>}
  </section>;
}
