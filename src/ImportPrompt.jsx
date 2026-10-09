import { importExample } from "./importExamples.js";

export default function ImportPrompt({ kind, startDate, onUpload, onLoad, disabled = false, notice, onPrerequisite, children }) {
  const example = importExample(kind, startDate);
  return <section className="import-prompt" aria-labelledby={"import-" + kind + "-title"}>
    <div className="import-prompt__card">
      <span className="import-prompt__icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.4"><path d="M14 3H5v18h14V8l-5-5Zm0 0v5h5M12 17v-6m-3 3 3-3 3 3" /></svg>
      </span>
      <h2 id={"import-" + kind + "-title"}>{example.title}</h2>
      <p className="import-prompt__formats">{example.formats}</p>
      <div className="import-prompt__actions">
        {onUpload ? <label className="file-input" aria-disabled={disabled}>
          <input type="file" accept={example.accept} onChange={onUpload} disabled={disabled} />
          <span>{example.action}</span>
        </label> : <button type="button" className="primary-action" onClick={onLoad} disabled={disabled}>{example.action}</button>}
        {children}
      </div>
      {notice ? <div className="import-prompt__prerequisite" aria-label="Step prerequisites">
        <p>{notice}</p><button type="button" onClick={onPrerequisite}>Go to Student Enrollment</button>
      </div> : null}
      <div className="import-prompt__example" tabIndex="0" role="region" aria-label={example.title + " file example"}>
        <table><caption>Example spreadsheet</caption>
          <thead><tr>{example.columns.map((column) => <th key={column} scope="col">{column}</th>)}</tr></thead>
          <tbody>{example.rows.map((row, index) => <tr key={index}>{row.map((value, column) => <td key={column}>{value}</td>)}</tr>)}</tbody>
        </table>
      </div>
      <p className="import-prompt__note">{example.note}</p>
    </div>
  </section>;
}
