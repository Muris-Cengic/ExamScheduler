const escapeText = (value) => String(value ?? "").replace(/\\/g, "\\\\").replace(/\r\n|\r|\n/g, "\\n")
  .replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\p{Cc}/gu, "");
const timestamp = (date) => date.toISOString().replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z");

// Exam times are UAE local time (Asia/Dubai, UTC+4), independent of the exporter's device.
const examTimestamp = (date, minutes) => timestamp(new Date(Date.parse(date + "T00:00:00Z") + (minutes - 240) * 60000));

function foldLine(value) {
  const encoder = new TextEncoder();
  const lines = [];
  let line = "";
  let bytes = 0;
  for (const character of value) {
    const size = encoder.encode(character).length;
    if (bytes + size > 75) { lines.push(line); line = " "; bytes = 1; }
    line += character;
    bytes += size;
  }
  return [...lines, line].join("\r\n");
}

export function buildInvigilatorCalendar(person, duties, now = new Date()) {
  const lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Midterm Exam Scheduling Helper//Invigilation//EN", "CALSCALE:GREGORIAN",
    "X-WR-CALNAME:" + escapeText(person.name + " - Exam invigilation"), "X-WR-TIMEZONE:Asia/Dubai"];
  const seen = new Set();
  duties.filter((duty) => duty.invigilatorId === person.id && duty.role === "Exam")
    .sort((a, b) => a.date.localeCompare(b.date) || a.start - b.start || a.id.localeCompare(b.id)).forEach((duty) => {
      const uid = encodeURIComponent(JSON.stringify([person.id, duty.id, duty.date])) + "@midterm-scheduling-helper";
      if (seen.has(uid)) return;
      seen.add(uid);
      lines.push("BEGIN:VEVENT", "UID:" + uid, "DTSTAMP:" + timestamp(now),
        "DTSTART:" + examTimestamp(duty.date, duty.start), "DTEND:" + examTimestamp(duty.date, duty.end),
        "SUMMARY:" + escapeText("Invigilation - " + duty.code), "LOCATION:" + escapeText(duty.roomName),
        "DESCRIPTION:" + escapeText(["Invigilator: " + person.name, "Course: " + duty.code + (duty.title ? " - " + duty.title : ""),
          "Room: " + duty.roomName, "Exam times: Asia/Dubai (UTC+4)", duty.teaching ? "During lab/tutorial hours - zero extra load" : "Extra exam duty"].join("\n")),
        "STATUS:CONFIRMED", "TRANSP:OPAQUE", "END:VEVENT");
    });
  lines.push("END:VCALENDAR");
  return lines.map(foldLine).join("\r\n") + "\r\n";
}
