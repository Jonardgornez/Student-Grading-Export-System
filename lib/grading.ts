import { saveAs } from "file-saver";

export type Student = {
  id: number;
  full_name: string;
  gender?: string;
};

export type AttendanceRecord = {
  student_id: number;
  attendance_date: string;
  status: "Present" | "Absent" | string;
};

export type ScoreEntry = {
  student_id: number;
  score: number;
};

export type Activity = {
  id: number;
  title: string;
  activity_type?: string;
  total_score: number;
  activity_date: string;
  scores: ScoreEntry[];
};

export type Participation = {
  id: number;
  title: string;
  total_score: number;
  participation_date: string;
  scores: ScoreEntry[];
};

export type Exam = {
  id: number;
  exam_type: string;
  title: string;
  total_score: number;
  exam_date: string;
  scores: ScoreEntry[];
};

export type GradingSettings = {
  attendance_percent: number;
  activities_percent: number;
  participation_percent: number;
  midterm_percent: number;
  final_percent: number;
};

export type GradingData = {
  meta?: {
    source_teacher_name?: string;
    app_name?: string;
    exported_at?: string;
  };
  subject?: {
    subject_code?: string;
    subject_name?: string;
    subject_course?: string;
    subject_section?: string;
    schedule_days?: string;
    schedule_time?: string;
  };
  grading_settings?: Partial<GradingSettings>;
  students: Student[];
  attendance: AttendanceRecord[];
  activities: Activity[];
  participations: Participation[];
  exams: Exam[];
};

export type ExportPreview = {
  title: string;
  description: string;
  columns: string[];
  rows: string[][];
};

const DEFAULT_SETTINGS: GradingSettings = {
  attendance_percent: 10,
  activities_percent: 20,
  participation_percent: 10,
  midterm_percent: 30,
  final_percent: 30,
};

function getSettings(data: GradingData): GradingSettings {
  return {
    ...DEFAULT_SETTINGS,
    ...data.grading_settings,
  };
}

function scoreToPercent(score: number, total: number): number {
  if (total <= 0) {
    return 0;
  }

  return (score / total) * 100;
}

function attendancePercentForStudent(
  data: GradingData,
  studentId: number,
): number {
  const records = data.attendance.filter(
    (item) => item.student_id === studentId,
  );

  if (records.length === 0) {
    return 0;
  }

  const presentCount = records.filter(
    (item) => item.status === "Present",
  ).length;
  return (presentCount / records.length) * 100;
}

function averageComponentPercentForStudent(
  items: Array<{ total_score: number; scores: ScoreEntry[] }>,
  studentId: number,
): number {
  if (items.length === 0) {
    return 0;
  }

  const totalPercent = items.reduce((sum, item) => {
    const score =
      item.scores.find((entry) => entry.student_id === studentId)?.score ?? 0;
    return sum + scoreToPercent(score, item.total_score);
  }, 0);

  return totalPercent / items.length;
}

function getExamAverageByType(
  data: GradingData,
  studentId: number,
  examType: "Midterm" | "Final",
) {
  return averageComponentPercentForStudent(
    data.exams.filter((exam) => exam.exam_type === examType),
    studentId,
  );
}

function getWeightedScore(percent: number, weight: number) {
  return (percent * weight) / 100;
}

function formatWholeNumber(value: number) {
  return Math.round(value).toString();
}

function getFinalGradeDetails(data: GradingData, studentId: number) {
  const settings = getSettings(data);
  const attendance = attendancePercentForStudent(data, studentId);
  const activities = averageComponentPercentForStudent(
    data.activities,
    studentId,
  );
  const participation = averageComponentPercentForStudent(
    data.participations,
    studentId,
  );
  const midterm = getExamAverageByType(data, studentId, "Midterm");
  const finalExam = getExamAverageByType(data, studentId, "Final");

  const attendanceWeighted = getWeightedScore(
    attendance,
    settings.attendance_percent,
  );
  const activitiesWeighted = getWeightedScore(
    activities,
    settings.activities_percent,
  );
  const participationWeighted = getWeightedScore(
    participation,
    settings.participation_percent,
  );
  const hasMidterm = data.exams.some((exam) => exam.exam_type === "Midterm");
  const hasFinal = data.exams.some((exam) => exam.exam_type === "Final");
  const midtermWeighted = hasMidterm
    ? getWeightedScore(midterm, settings.midterm_percent)
    : 0;
  const finalExamWeighted = hasFinal
    ? getWeightedScore(finalExam, settings.final_percent)
    : 0;
  const finalGrade =
    attendanceWeighted +
    activitiesWeighted +
    participationWeighted +
    midtermWeighted +
    finalExamWeighted;

  return {
    attendance,
    activities,
    participation,
    midterm,
    finalExam,
    attendanceWeighted,
    activitiesWeighted,
    participationWeighted,
    midtermWeighted,
    finalExamWeighted,
    finalGrade,
  };
}

function getFinalGrade(data: GradingData, studentId: number): string {
  return formatWholeNumber(getFinalGradeDetails(data, studentId).finalGrade);
}

function applyWorksheetTheme(
  worksheet: import("exceljs").Worksheet,
  title: string,
  subtitle?: string,
) {
  worksheet.insertRow(1, [title]);
  worksheet.getRow(1).font = {
    bold: true,
    size: 16,
    color: { argb: "FFF9FAFB" },
  };
  worksheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF0F172A" },
  };
  worksheet.mergeCells(1, 1, 1, worksheet.columnCount || 1);

  if (subtitle) {
    worksheet.insertRow(2, [subtitle]);
    worksheet.getRow(2).font = { italic: true, color: { argb: "FF334155" } };
    worksheet.mergeCells(2, 1, 2, worksheet.columnCount || 1);
    worksheet.insertRow(3, []);
  } else {
    worksheet.insertRow(2, []);
  }

  const headerRow = worksheet.getRow(subtitle ? 4 : 3);
  headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
  headerRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF2563EB" },
  };
}

function autoSizeColumns(worksheet: import("exceljs").Worksheet) {
  worksheet.columns.forEach((column) => {
    let maxLength = 12;

    if (!column.eachCell) {
      return;
    }

    column.eachCell({ includeEmpty: true }, (cell) => {
      const value = cell.value == null ? "" : String(cell.value);
      maxLength = Math.max(maxLength, value.length + 2);
    });

    column.width = Math.min(maxLength, 30);
  });
}

function getAttendanceDates(data: GradingData) {
  return [...new Set(data.attendance.map((item) => item.attendance_date))].sort(
    (left, right) => left.localeCompare(right),
  );
}

function getAttendanceCode(status: string) {
  const normalized = status.trim().toLowerCase();

  if (normalized === "present") {
    return "P";
  }

  if (normalized === "absent") {
    return "A";
  }

  return status.trim().slice(0, 4).toUpperCase();
}

function formatShortDate(date: string) {
  return new Date(date).toLocaleDateString("en-US", {
    day: "numeric",
    month: "short",
  });
}

async function createWorkbook() {
  const ExcelJS = (await import("exceljs")).default;
  return new ExcelJS.Workbook();
}

async function downloadWorkbook(
  workbook: import("exceljs").Workbook,
  fileName: string,
) {
  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), fileName);
}

export function validateGradingData(data: unknown): data is GradingData {
  if (!data || typeof data !== "object") {
    return false;
  }

  const candidate = data as Record<string, unknown>;

  return (
    Array.isArray(candidate.students) &&
    Array.isArray(candidate.attendance) &&
    Array.isArray(candidate.activities) &&
    Array.isArray(candidate.participations) &&
    Array.isArray(candidate.exams)
  );
}

export function getDashboardStats(data: GradingData) {
  const uniqueAttendanceDates = new Set(
    data.attendance.map((item) => item.attendance_date),
  ).size;

  return {
    studentCount: data.students.length,
    attendanceDays: uniqueAttendanceDates,
    activityCount: data.activities.length,
    participationCount: data.participations.length,
    examCount: data.exams.length,
  };
}

export function getAttendanceExportPreview(data: GradingData): ExportPreview {
  const dates = getAttendanceDates(data);

  return {
    title: "Attendance Preview",
    description:
      "This preview matches the attendance export layout: one row per student and one column per attendance date.",
    columns: ["Name", ...dates.map(formatShortDate)],
    rows: data.students.map((student) => [
      student.full_name,
      ...dates.map((date) => {
        const record = data.attendance.find(
          (item) =>
            item.student_id === student.id && item.attendance_date === date,
        );

        return record ? getAttendanceCode(record.status) : "";
      }),
    ]),
  };
}

export function getActivitiesExportPreview(data: GradingData): ExportPreview {
  return {
    title: "Activities Preview",
    description:
      "This preview matches the activities export layout: one row per student and one column per activity.",
    columns: [
      "Student Name",
      ...data.activities.map((activity) => activity.title),
    ],
    rows: data.students.map((student) => [
      student.full_name,
      ...data.activities.map((activity) => {
        const score =
          activity.scores.find((entry) => entry.student_id === student.id)
            ?.score ?? 0;

        return String(score);
      }),
    ]),
  };
}

export function getParticipationExportPreview(
  data: GradingData,
): ExportPreview {
  return {
    title: "Participation Preview",
    description:
      "This preview matches the participation export layout for each student.",
    columns: ["Student Name", "Score"],
    rows: data.students.map((student) => {
      const participationAverage = averageComponentPercentForStudent(
        data.participations,
        student.id,
      );

      return [student.full_name, formatWholeNumber(participationAverage)];
    }),
  };
}

export function getFinalExportPreview(data: GradingData): ExportPreview {
  const settings = getSettings(data);

  return {
    title: "Final Grades Preview",
    description:
      "This preview matches the final grades export layout, including weighted component scores.",
    columns: [
      "Student Name",
      `Attendance (${settings.attendance_percent}%)`,
      `Activities (${settings.activities_percent}%)`,
      `Participation (${settings.participation_percent}%)`,
      `Midterm (${settings.midterm_percent}%)`,
      `Final (${settings.final_percent}%)`,
      "Final Grade",
    ],
    rows: data.students.map((student) => {
      const details = getFinalGradeDetails(data, student.id);

      return [
        student.full_name,
        formatWholeNumber(details.attendanceWeighted),
        formatWholeNumber(details.activitiesWeighted),
        formatWholeNumber(details.participationWeighted),
        formatWholeNumber(details.midtermWeighted),
        formatWholeNumber(details.finalExamWeighted),
        formatWholeNumber(details.finalGrade),
      ];
    }),
  };
}

export async function exportAttendance(data: GradingData) {
  const workbook = await createWorkbook();
  const worksheet = workbook.addWorksheet("Attendance");
  const dates = getAttendanceDates(data);

  worksheet.columns = [
    { header: "Name", key: "name", width: 36 },
    ...dates.map((date) => ({
      header: formatShortDate(date),
      key: date,
      width: 12,
    })),
  ];

  data.students.forEach((student) => {
    const row: Record<string, string> = {
      name: student.full_name,
    };

    dates.forEach((date) => {
      const record = data.attendance.find(
        (item) =>
          item.student_id === student.id && item.attendance_date === date,
      );

      row[date] = record ? getAttendanceCode(record.status) : "";
    });

    worksheet.addRow(row);
  });

  applyWorksheetTheme(
    worksheet,
    "Attendance Export",
    data.subject?.subject_name,
  );

  dates.forEach((_date, index) => {
    const column = worksheet.getColumn(index + 2);
    column.alignment = { horizontal: "center", vertical: "middle" };
  });

  worksheet.getColumn(1).alignment = {
    horizontal: "left",
    vertical: "middle",
  };
  autoSizeColumns(worksheet);

  await downloadWorkbook(workbook, "attendance.xlsx");
}

export async function exportActivities(data: GradingData) {
  const workbook = await createWorkbook();
  const worksheet = workbook.addWorksheet("Activities");

  worksheet.columns = [
    { header: "Student Name", key: "name" },
    ...data.activities.map((activity) => ({
      header: activity.title,
      key: `activity_${activity.id}`,
    })),
  ];

  data.students.forEach((student) => {
    const row: Record<string, string | number> = { name: student.full_name };

    data.activities.forEach((activity) => {
      const score =
        activity.scores.find((entry) => entry.student_id === student.id)
          ?.score ?? 0;

      row[`activity_${activity.id}`] = score;
    });

    worksheet.addRow(row);
  });

  applyWorksheetTheme(
    worksheet,
    "Activities Export",
    data.subject?.subject_name,
  );
  autoSizeColumns(worksheet);

  await downloadWorkbook(workbook, "activities.xlsx");
}

export async function exportParticipation(data: GradingData) {
  const workbook = await createWorkbook();
  const worksheet = workbook.addWorksheet("Participation");

  worksheet.columns = [
    { header: "Student Name", key: "name" },
    { header: "Score", key: "score" },
  ];

  data.students.forEach((student) => {
    const participationAverage = averageComponentPercentForStudent(
      data.participations,
      student.id,
    );

    worksheet.addRow({
      name: student.full_name,
      score: formatWholeNumber(participationAverage),
    });
  });

  applyWorksheetTheme(
    worksheet,
    "Participation Export",
    data.subject?.subject_name,
  );
  autoSizeColumns(worksheet);

  await downloadWorkbook(workbook, "participation.xlsx");
}

export async function exportFinal(data: GradingData) {
  const workbook = await createWorkbook();
  const worksheet = workbook.addWorksheet("Final Grades");
  const settings = getSettings(data);

  worksheet.columns = [
    { header: "Student Name", key: "name" },
    {
      header: `Attendance (${settings.attendance_percent}%)`,
      key: "attendance",
    },
    {
      header: `Activities (${settings.activities_percent}%)`,
      key: "activities",
    },
    {
      header: `Participation (${settings.participation_percent}%)`,
      key: "participation",
    },
    { header: `Midterm (${settings.midterm_percent}%)`, key: "midterm" },
    { header: `Final (${settings.final_percent}%)`, key: "finalExam" },
    { header: "Final Grade", key: "final" },
  ];

  data.students.forEach((student) => {
    const details = getFinalGradeDetails(data, student.id);

    worksheet.addRow({
      name: student.full_name,
      attendance: formatWholeNumber(details.attendanceWeighted),
      activities: formatWholeNumber(details.activitiesWeighted),
      participation: formatWholeNumber(details.participationWeighted),
      midterm: formatWholeNumber(details.midtermWeighted),
      finalExam: formatWholeNumber(details.finalExamWeighted),
      final: getFinalGrade(data, student.id),
    });
  });

  applyWorksheetTheme(
    worksheet,
    "Final Grades Export",
    data.subject?.subject_name,
  );
  autoSizeColumns(worksheet);

  await downloadWorkbook(workbook, "final-grades.xlsx");
}
