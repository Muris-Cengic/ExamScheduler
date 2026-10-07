export const FRIDAY_EXAM_WINDOWS = [{ start: 540, end: 600 }, { start: 630, end: 690 }];
export const FRIDAY_EXAM_NOTICE = "Friday exams must use 09:00-10:00 or 10:30-11:30.";

export function isFridayExamTimeAllowed(day, start, end) {
  return day !== "Friday" || FRIDAY_EXAM_WINDOWS.some((window) => start === window.start && end > start && end <= window.end);
}
