import * as XLSX from "xlsx/xlsx.mjs";
import { invigilatorWorkloads, isTeachingTimeDuty, validateResourcePlan } from "./resources.js";

const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const time = (minutes) => `${((Math.floor(minutes / 60) + 11) % 12) + 1}:${String(minutes % 60).padStart(2, "0")} ${minutes >= 720 ? "PM" : "AM"}`;

export function buildResourceWorkbookForWeek({ week, ...options }) {
  return buildResourceWorkbook(options, options.sessions.filter((session) => session.week === week), {
    invSheetName: `Week ${week} Invigilators`, staffSheetName: `Week ${week} Invigilator Pool`, roomSheetName: `Week ${week} Room Pool`,
  });
}

export function buildSharedResourceWorkbook({ weeks, ...options }) {
  const selected = new Set(weeks);
  const current = options.sessions.filter((session) => selected.has(session.week))
    .sort((a, b) => a.week - b.week || DAYS.indexOf(a.day) - DAYS.indexOf(b.day) || a.start - b.start || a.end - b.end);
  return buildResourceWorkbook(options, current, {
    invSheetName: "Invigilator Duties", staffSheetName: "Invigilator Pool", roomSheetName: "Room Pool",
  });
}

function buildResourceWorkbook({ sessions, catalog, plan, startDate, templateHeaders }, current, { invSheetName, staffSheetName, roomSheetName }) {
  const validation = validateResourcePlan(sessions, catalog, plan);
  if (!validation.complete) throw new Error("Complete valid resource assignments before exporting.");
  if (!current.length) return null;
  const workbook = XLSX.utils.book_new();
  const staffRows = new Map(catalog.invigilators.map((person, index) => [person.id, index + 2]));
  const roomRows = new Map(catalog.rooms.map((room, index) => [room.id, index + 2]));
  const staff = new Map(catalog.invigilators.map((person) => [person.id, person]));
  const rooms = new Map(catalog.rooms.map((room) => [room.id, room]));
  const invRows = [[...templateHeaders.invigilatorHeader.slice(0, 8), "Invigilator 1", "Invigilator 2", "Backup Invigilator",
    "Invigilator 1 extra duty", "Invigilator 2 extra duty", "Room allocation decision"]];
  const daily = new Map();
  const formatter = new Intl.DateTimeFormat("en-GB", { weekday: "long", day: "numeric", month: "short", year: "numeric", timeZone: "UTC" });
  current.forEach((session) => {
    const date = new Date(`${startDate}T00:00:00Z`);
    date.setUTCDate(date.getUTCDate() + (session.week - 1) * 7 + DAYS.indexOf(session.day));
    const range = `${time(session.start)} - ${time(session.end)}`;
    session.rooms.forEach((room, index) => {
      const allocation = plan.allocations[room.id];
      const assignedRoom = rooms.get(allocation.roomId);
      const staffCell = (id) => id ? { t: "s", v: staff.get(id).name, f: `'${staffSheetName}'!A${staffRows.get(id)}` } : "";
      const extraDuty = (id) => id ? Number(!isTeachingTimeDuty(staff.get(id), session)) : "";
      const rowNumber = invRows.length + 1;
      const crns = [...new Set(room.students.map((student) => student.crn).filter(Boolean))].sort().join(", ");
      invRows.push([
        crns, room.code, room.title, room.students.length, formatter.format(date), range, room.instructors.join(", "),
        { t: "s", v: assignedRoom.name, f: `'${roomSheetName}'!A${roomRows.get(assignedRoom.id)}` },
        staffCell(allocation.invigilatorIds[0]), staffCell(allocation.invigilatorIds[1]), staffCell(plan.backups[session.id]?.[index]),
        extraDuty(allocation.invigilatorIds[0]), extraDuty(allocation.invigilatorIds[1]),
        room.distributionChoice === "standard" ? `Normal limit: ${room.maxStudents}` : `Approved up to 27: ${room.distributionChoice === "merge" ? "merge overflow" : "staffing-aware distribution"}`,
      ]);
      const dailySheetName = `Week ${session.week} ${session.day}`;
      if (!daily.has(dailySheetName)) daily.set(dailySheetName, [templateHeaders.studentHeader]);
      room.students.forEach((student) => daily.get(dailySheetName).push([
        student.crn || "", room.code, `${room.title} (${range})`, student.id, student.name,
        { t: "s", v: assignedRoom.name, f: `'${invSheetName}'!H${rowNumber}` }, "",
      ]));
    });
  });
  const invSheet = XLSX.utils.aoa_to_sheet(invRows);
  invSheet["!cols"] = [18, 16, 42, 16, 30, 26, 30, 35, 30, 30, 30].map((wch) => ({ wch }));
  // Hidden duty flags let Excel recalculate loads without charging replaced teaching hours.
  invSheet["!cols"].push(...Array.from({ length: 2 }, () => ({ wch: 22, hidden: true })));
  invSheet["!cols"].push({ wch: 40 });
  XLSX.utils.book_append_sheet(workbook, invSheet, invSheetName);
  const workloads = invigilatorWorkloads(catalog, current, plan);
  const staffData = [["Invigilator", "Invigilation load", "During teaching hours", "Included in pool"],
    ...workloads.map((person, index) => {
      const row = index + 2;
      return [person.name,
        { t: "n", v: person.exam, f: `SUMIF('${invSheetName}'!I:I,A${row},'${invSheetName}'!L:L)+SUMIF('${invSheetName}'!J:J,A${row},'${invSheetName}'!M:M)` },
        { t: "n", v: person.teaching, f: `COUNTIF('${invSheetName}'!I:I,A${row})+COUNTIF('${invSheetName}'!J:J,A${row})-B${row}` },
        person.enabled ? "Yes" : "No"];
    })];
  const staffSheet = XLSX.utils.aoa_to_sheet(staffData);
  staffSheet["!cols"] = [32, 22, 24, 20].map((wch) => ({ wch }));
  XLSX.utils.book_append_sheet(workbook, staffSheet, staffSheetName);
  const roomData = [["Room", "Assignments", "Normal maximum students", "Largest approved exam limit", "Included in pool"], ...catalog.rooms.map((room, index) => [
    room.name, { t: "n", v: current.reduce((sum, session) => sum + session.rooms.filter((examRoom) => plan.allocations[examRoom.id].roomId === room.id).length, 0), f: `COUNTIF('${invSheetName}'!H:H,A${index + 2})` }, 25,
    Math.max(25, ...current.flatMap((session) => session.rooms.filter((examRoom) => plan.allocations[examRoom.id].roomId === room.id).map((examRoom) => examRoom.maxStudents))), room.enabled ? "Yes" : "No",
  ])];
  const roomSheet = XLSX.utils.aoa_to_sheet(roomData);
  roomSheet["!cols"] = [40, 18, 26, 28, 20].map((wch) => ({ wch }));
  XLSX.utils.book_append_sheet(workbook, roomSheet, roomSheetName);
  daily.forEach((rows, name) => {
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    sheet["!cols"] = [14, 16, 56, 18, 40, 40, 20].map((wch) => ({ wch }));
    XLSX.utils.book_append_sheet(workbook, sheet, name);
  });
  return workbook;
}
