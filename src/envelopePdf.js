import { jsPDF } from "jspdf";
import { envelopeCoverRows, envelopeSessionInfo } from "./exportReports.js";

// Fit all fields before drawing so long titles, CRN lists and rooms never overlap.
export function envelopeCoverLayout(doc, page) {
  const width = doc.internal.pageSize.getWidth() - 88;
  const height = doc.internal.pageSize.getHeight() - 88;
  for (let scale = 1; scale >= 0.55; scale -= 0.05) {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(30 * scale);
    const title = doc.splitTextToSize(page.code + " Envelope Cover", width);
    const measure = ([label, value, note], fieldWidth, minimumHeight) => {
      doc.setFont("helvetica", "bold");
      doc.setFontSize(14 * scale);
      const labelLines = doc.splitTextToSize(label, fieldWidth - 40);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(20 * scale);
      const lines = doc.splitTextToSize(value || " ", fieldWidth - 40);
      doc.setFontSize(10 * scale);
      const noteLines = note ? doc.splitTextToSize(note, fieldWidth - 40) : [];
      const valueY = (54 + (labelLines.length - 1) * 17) * scale;
      const height = Math.max(minimumHeight * scale, valueY + (lines.length - 1) * 24 * scale + 20 * scale +
        (noteLines.length ? (noteLines.length * 12 + 12) * scale : 0));
      return { label, labelLines, lines, noteLines, valueY, height };
    };
    const fieldRows = envelopeCoverRows(page).map((row) => {
      const fieldWidth = (width - (row.length - 1) * 16) / row.length;
      const fields = row.map((field) => measure(field, fieldWidth, 85));
      return { fields, fieldWidth, height: Math.max(...fields.map((field) => field.height)), valueY: Math.max(...fields.map((field) => field.valueY)) };
    });
    const session = envelopeSessionInfo(page).map((field) => measure(field, width / 3, 104));
    const sessionHeight = Math.max(...session.map((field) => field.height));
    const titleHeight = title.length * 36 * scale + 24 * scale;
    if (titleHeight + fieldRows.reduce((sum, row) => sum + row.height + 22 * scale, 0) + sessionHeight <= height) {
      return { width, title, titleHeight, fields: fieldRows.flatMap((row) => row.fields), fieldRows, session, sessionHeight, scale };
    }
  }
  throw new Error("Envelope cover text is too long to fit on A4. Shorten the course title, CRN list or room name.");
}

export function envelopeCoverPdf(pages) {
  if (!pages.length) throw new Error("No envelope covers to export.");
  const doc = new jsPDF({ unit: "pt", format: "a4", compress: true });
  doc.setProperties({ title: pages[0].code + " - Envelope Covers", subject: "Exam room envelope covers" });
  pages.forEach((page, index) => {
    if (index) doc.addPage();
    const layout = envelopeCoverLayout(doc, page);
    const { width, scale } = layout;
    let y = 44;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(30 * scale);
    doc.setTextColor(31, 41, 55);
    doc.text(layout.title, 44 + width / 2, y + 33 * scale, { align: "center", lineHeightFactor: 1.2 });
    y += layout.titleHeight;
    doc.setDrawColor(190, 194, 201);
    doc.setLineWidth(0.75);
    layout.fieldRows.forEach((row) => {
      row.fields.forEach((field, column) => {
        const x = 44 + column * (row.fieldWidth + 16);
        doc.roundedRect(x, y, row.fieldWidth, row.height, 8, 8);
        doc.setFont("helvetica", "bold");
        doc.setFontSize(14 * scale);
        doc.text(field.labelLines, x + 20, y + 24 * scale, { lineHeightFactor: 17 / 14 });
        doc.setFont("helvetica", "normal");
        doc.setFontSize(20 * scale);
        doc.text(field.lines, x + 20, y + row.valueY, { lineHeightFactor: 1.2 });
        if (field.noteLines.length) {
          doc.setFontSize(10 * scale);
          doc.setTextColor(100, 116, 139);
          doc.text(field.noteLines, x + 20, y + row.height - (12 + (field.noteLines.length - 1) * 12) * scale, { lineHeightFactor: 1.2 });
          doc.setTextColor(31, 41, 55);
        }
      });
      y += row.height + 22 * scale;
    });
    doc.roundedRect(44, y, width, layout.sessionHeight, 8, 8);
    layout.session.forEach((field, column) => {
      const x = 44 + width * (column + 0.5) / 3;
      doc.setFont("helvetica", "bold");
      doc.setFontSize(14 * scale);
      doc.text(field.label, x, y + 28 * scale, { align: "center" });
      doc.setFont("helvetica", "normal");
      doc.setFontSize(20 * scale);
      doc.text(field.lines, x, y + 58 * scale, { align: "center", lineHeightFactor: 1.2 });
    });
  });
  return doc.output("arraybuffer");
}
