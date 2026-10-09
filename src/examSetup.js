export function inferAcademicTerm(startDate) {
  if (typeof startDate !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(startDate)) return null;
  const date = new Date(startDate + "T00:00:00Z");
  if (Number.isNaN(date.getTime()) || date.toISOString().slice(0, 10) !== startDate) return null;
  const month = date.getUTCMonth() + 1;
  const year = date.getUTCFullYear();
  // August belongs to Fall where the supplied Fall and Summer ranges overlap.
  const semester = month >= 8 ? "Fall" : month >= 6 ? "Summer" : "Spring";
  const academicYearStart = month >= 8 ? year : year - 1;
  return { semester, academicYear: academicYearStart + "-" + (academicYearStart + 1) };
}
