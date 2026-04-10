"use client";

import { useState, useTransition } from "react";
import {
  exportActivities,
  exportAttendance,
  exportFinal,
  exportParticipation,
  type GradingData,
  getActivitiesExportPreview,
  getAttendanceExportPreview,
  getDashboardStats,
  getFinalExportPreview,
  getParticipationExportPreview,
  validateGradingData,
} from "@/lib/grading";

type ExportStep =
  | "load-data"
  | "review-data"
  | "export-attendance"
  | "export-activities"
  | "export-participation"
  | "export-final";

const exportButtons: Array<{
  action: ExportStep;
  label: string;
  description: string;
  fileName: string;
  run: (data: GradingData) => Promise<void>;
}> = [
  {
    action: "export-attendance",
    label: "Export Attendance",
    description:
      "Create the attendance workbook with all recorded dates and statuses.",
    fileName: "attendance.xlsx",
    run: exportAttendance,
  },
  {
    action: "export-activities",
    label: "Export Activities",
    description: "Build a score sheet where every activity becomes a column.",
    fileName: "activities.xlsx",
    run: exportActivities,
  },
  {
    action: "export-participation",
    label: "Export Participation",
    description: "Generate the participation sheet for each student.",
    fileName: "participation.xlsx",
    run: exportParticipation,
  },
  {
    action: "export-final",
    label: "Export Final Grades",
    description:
      "Combine attendance, activities, participation, and exam results.",
    fileName: "final-grades.xlsx",
    run: exportFinal,
  },
];

const stepLabels: Record<ExportStep, string> = {
  "load-data": "Load grading JSON",
  "review-data": "Review dataset summary",
  "export-attendance": "Export attendance workbook",
  "export-activities": "Export activities workbook",
  "export-participation": "Export participation workbook",
  "export-final": "Export final grades workbook",
};

export default function Home() {
  const [data, setData] = useState<GradingData | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [message, setMessage] = useState(
    "Upload a grading JSON file to preview the dataset and unlock the export buttons.",
  );
  const [completedSteps, setCompletedSteps] = useState<ExportStep[]>([]);
  const [activeExport, setActiveExport] = useState<ExportStep | null>(null);
  const [showNoDataModal, setShowNoDataModal] = useState(false);
  const [previewExportStep, setPreviewExportStep] = useState<ExportStep | null>(
    null,
  );
  const [isPending, startTransition] = useTransition();

  const stats = data
    ? getDashboardStats(data)
    : {
        studentCount: 0,
        attendanceDays: 0,
        activityCount: 0,
        participationCount: 0,
        examCount: 0,
      };
  const progressPercent = Math.round((completedSteps.length / 6) * 100);

  function markComplete(step: ExportStep) {
    setCompletedSteps((current) =>
      current.includes(step) ? current : [...current, step],
    );
  }

  function handleUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];

    if (!file) {
      return;
    }

    startTransition(async () => {
      try {
        const text = await file.text();
        const parsed = JSON.parse(text) as unknown;

        if (!validateGradingData(parsed)) {
          setData(null);
          setFileName(null);
          setCompletedSteps([]);
          setMessage(
            "The uploaded file is missing one or more required sections: students, attendance, activities, participations, exams.",
          );
          return;
        }

        setData(parsed);
        setFileName(file.name);
        setCompletedSteps(["load-data", "review-data"]);
        setMessage(
          `Loaded ${file.name}. Click any export button to preview the exact layout before downloading.`,
        );
      } catch {
        setData(null);
        setFileName(null);
        setCompletedSteps([]);
        setMessage(
          "The selected file is not valid JSON. Please upload a valid grading export.",
        );
      }
    });
  }

  function openExportPreview(step: ExportStep) {
    if (!data) {
      setShowNoDataModal(true);
      return;
    }

    setPreviewExportStep(step);
  }

  function runExport(
    step: ExportStep,
    exporter: (data: GradingData) => Promise<void>,
  ) {
    if (!data) {
      setShowNoDataModal(true);
      return;
    }

    startTransition(async () => {
      try {
        setActiveExport(step);
        setPreviewExportStep(null);
        setMessage(`Preparing ${stepLabels[step].toLowerCase()}...`);
        await exporter(data);
        markComplete(step);
        setMessage(
          `${stepLabels[step]} complete. Your download should start automatically.`,
        );
      } catch (error) {
        const detail =
          error instanceof Error ? error.message : "Unexpected export failure.";
        setMessage(`Export failed: ${detail}`);
      } finally {
        setActiveExport(null);
      }
    });
  }

  return (
    <main className="min-h-screen bg-[radial-gradient(circle_at_top,#e0f2fe_0%,#f8fafc_35%,#dbeafe_100%)] px-5 py-8 text-slate-900 sm:px-8 lg:px-12">
      <div className="mx-auto flex w-full max-w-7xl flex-col gap-8">
        <section className="overflow-hidden rounded-[2rem] border border-white/70 bg-slate-950 text-white shadow-[0_30px_100px_-40px_rgba(15,23,42,0.85)]">
          <div className="grid gap-8 px-6 py-8 sm:px-8 lg:grid-cols-[1.4fr_0.9fr] lg:px-10">
            <div className="space-y-5">
              <p className="inline-flex rounded-full border border-cyan-400/30 bg-cyan-400/10 px-3 py-1 text-xs font-semibold uppercase tracking-[0.25em] text-cyan-200">
                Student Grading Export System
              </p>
              <div className="space-y-3">
                <h1 className="max-w-3xl text-4xl font-semibold tracking-tight sm:text-5xl">
                  Export attendance, activities, participation, and final grades
                  to Excel.
                </h1>
                <p className="max-w-2xl text-base leading-7 text-slate-300 sm:text-lg">
                  This dashboard follows the README requirements and lets you
                  upload your own grading JSON before generating Excel exports.
                </p>
              </div>
              <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
                <SummaryCard label="Students" value={stats.studentCount} />
                <SummaryCard
                  label="Attendance Days"
                  value={stats.attendanceDays}
                />
                <SummaryCard label="Activities" value={stats.activityCount} />
                <SummaryCard
                  label="Participation / Exams"
                  value={`${stats.participationCount} / ${stats.examCount}`}
                />
              </div>
            </div>

            <aside className="rounded-[1.5rem] border border-white/10 bg-white/8 p-5 backdrop-blur">
              <p className="text-sm font-medium text-slate-300">
                Current dataset
              </p>
              <h2 className="mt-2 text-2xl font-semibold text-white">
                {data?.subject?.subject_name ?? "No dataset uploaded yet"}
              </h2>
              <div className="mt-5 space-y-3 text-sm text-slate-300">
                <p>File: {fileName ?? "No file uploaded"}</p>
                <p>Teacher: {data?.meta?.source_teacher_name ?? "-"}</p>
                <p>
                  Section: {data?.subject?.subject_course ?? "-"}{" "}
                  {data?.subject?.subject_section ?? ""}
                </p>
                <p>
                  Schedule: {data?.subject?.schedule_days ?? "-"}{" "}
                  {data?.subject?.schedule_time ?? ""}
                </p>
              </div>
            </aside>
          </div>
        </section>

        <section className="grid gap-8 lg:grid-cols-[1.2fr_0.8fr]">
          <div className="space-y-6 rounded-[2rem] border border-slate-200/80 bg-white/85 p-6 shadow-[0_24px_80px_-48px_rgba(15,23,42,0.7)] backdrop-blur sm:p-8">
            <div className="flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
              <div>
                <p className="text-sm font-semibold uppercase tracking-[0.22em] text-sky-700">
                  Upload Your Data
                </p>
                <h2 className="mt-2 text-3xl font-semibold tracking-tight text-slate-950">
                  JSON-powered export workflow
                </h2>
              </div>
              <label className="inline-flex cursor-pointer items-center justify-center rounded-full bg-slate-950 px-5 py-3 text-sm font-medium text-white transition hover:bg-slate-800">
                Upload JSON
                <input
                  className="hidden"
                  type="file"
                  accept=".json,application/json"
                  onChange={handleUpload}
                />
              </label>
            </div>

            <div className="rounded-[1.5rem] border border-sky-100 bg-sky-50 px-5 py-4 text-sm text-sky-950">
              {isPending ? "Working on your request..." : message}
            </div>

            <div className="grid gap-4">
              {exportButtons.map((button) => {
                const isRunning = activeExport === button.action;

                return (
                  <div
                    key={button.action}
                    className="rounded-[1.5rem] border border-slate-200 bg-white p-5 shadow-sm"
                  >
                    <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
                      <div className="space-y-1">
                        <h3 className="text-xl font-semibold text-slate-950">
                          {button.label}
                        </h3>
                        <p className="max-w-2xl text-sm leading-6 text-slate-600">
                          {button.description}
                        </p>
                      </div>
                      <button
                        type="button"
                        onClick={() => openExportPreview(button.action)}
                        disabled={isPending}
                        className="inline-flex min-w-48 items-center justify-center rounded-full bg-linear-to-r from-blue-600 to-cyan-500 px-5 py-3 text-sm font-semibold text-white transition hover:brightness-110 disabled:cursor-not-allowed disabled:opacity-60"
                      >
                        {isRunning ? "Exporting..." : "Preview Export"}
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          <div className="space-y-6">
            <section className="rounded-[2rem] border border-slate-200/80 bg-white/90 p-6 shadow-[0_24px_80px_-48px_rgba(15,23,42,0.7)] backdrop-blur sm:p-8">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-semibold uppercase tracking-[0.22em] text-emerald-700">
                    Todo List
                  </p>
                  <h2 className="mt-2 text-3xl font-semibold tracking-tight text-slate-950">
                    Progress tracker
                  </h2>
                </div>
                <span className="rounded-full bg-emerald-100 px-3 py-1 text-sm font-semibold text-emerald-800">
                  {progressPercent}% done
                </span>
              </div>

              <div className="mt-5 h-3 overflow-hidden rounded-full bg-slate-200">
                <div
                  className="h-full rounded-full bg-linear-to-r from-emerald-500 to-cyan-500 transition-all"
                  style={{ width: `${progressPercent}%` }}
                />
              </div>

              <div className="mt-6 space-y-3">
                {(Object.keys(stepLabels) as ExportStep[]).map((step) => {
                  const complete = completedSteps.includes(step);

                  return (
                    <div
                      key={step}
                      className={`flex items-center gap-3 rounded-2xl border px-4 py-3 ${
                        complete
                          ? "border-emerald-200 bg-emerald-50 text-emerald-900"
                          : "border-slate-200 bg-slate-50 text-slate-600"
                      }`}
                    >
                      <span
                        className={`inline-flex h-7 w-7 items-center justify-center rounded-full text-sm font-bold ${
                          complete
                            ? "bg-emerald-600 text-white"
                            : "bg-slate-200 text-slate-600"
                        }`}
                      >
                        {complete ? "✓" : "•"}
                      </span>
                      <span className="text-sm font-medium">
                        {stepLabels[step]}
                      </span>
                    </div>
                  );
                })}
              </div>
            </section>

            <section className="rounded-[2rem] border border-slate-200/80 bg-slate-950 p-6 text-slate-100 shadow-[0_24px_80px_-48px_rgba(15,23,42,0.9)] sm:p-8">
              <p className="text-sm font-semibold uppercase tracking-[0.22em] text-cyan-300">
                Export Guide
              </p>
              <div className="mt-4 space-y-3 text-sm leading-7 text-slate-300">
                <p>1. Upload a grading JSON file.</p>
                <p>2. Review the detected subject, teacher, and schedule.</p>
                <p>
                  3. Export the workbook you need and track progress on the
                  right.
                </p>
              </div>
            </section>
          </div>
        </section>
      </div>

      {showNoDataModal ? (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/55 px-5">
          <div className="w-full max-w-md rounded-[1.75rem] border border-slate-200 bg-white p-6 shadow-2xl">
            <p className="text-sm font-semibold uppercase tracking-[0.22em] text-rose-600">
              Upload required
            </p>
            <h2 className="mt-3 text-2xl font-semibold text-slate-950">
              No dataset available for export
            </h2>
            <p className="mt-3 text-sm leading-6 text-slate-600">
              Please upload a valid grading JSON file first. Export buttons stay
              blocked until data has been loaded into the dashboard.
            </p>
            <div className="mt-6 flex justify-end">
              <button
                type="button"
                onClick={() => setShowNoDataModal(false)}
                className="rounded-full bg-slate-950 px-5 py-3 text-sm font-semibold text-white transition hover:bg-slate-800"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      ) : null}

      {previewExportStep ? (
        <ExportPreviewModal
          data={data}
          exportButton={
            exportButtons.find(
              (button) => button.action === previewExportStep,
            ) ?? null
          }
          isPending={isPending}
          onClose={() => setPreviewExportStep(null)}
          onConfirm={(step, exporter) => runExport(step, exporter)}
        />
      ) : null}
    </main>
  );
}

function SummaryCard({
  label,
  value,
}: {
  label: string;
  value: number | string;
}) {
  return (
    <div className="rounded-[1.25rem] border border-white/10 bg-white/8 px-4 py-4">
      <p className="text-sm text-slate-300">{label}</p>
      <p className="mt-2 text-3xl font-semibold text-white">{value}</p>
    </div>
  );
}

function ExportPreviewModal({
  data,
  exportButton,
  isPending,
  onClose,
  onConfirm,
}: {
  data: GradingData | null;
  exportButton: (typeof exportButtons)[number] | null;
  isPending: boolean;
  onClose: () => void;
  onConfirm: (
    step: ExportStep,
    exporter: (data: GradingData) => Promise<void>,
  ) => void;
}) {
  if (!data || !exportButton) {
    return null;
  }

  const preview = getExportPreview(data, exportButton.action);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/60 p-4">
      <div className="flex max-h-[90vh] w-full max-w-6xl flex-col overflow-hidden rounded-[1.75rem] border border-slate-200 bg-white shadow-2xl">
        <div className="border-b border-slate-200 px-6 py-5">
          <p className="text-sm font-semibold uppercase tracking-[0.22em] text-cyan-700">
            Preview First
          </p>
          <h2 className="mt-2 text-3xl font-semibold text-slate-950">
            {preview.title}
          </h2>
          <p className="mt-2 text-sm leading-6 text-slate-600">
            {preview.description} This is the same table structure that will be
            used in {exportButton.fileName}.
          </p>
        </div>

        <div className="overflow-auto px-6 py-5">
          <div className="overflow-hidden rounded-[1.25rem] border border-slate-200">
            <table className="min-w-full table-fixed divide-y divide-slate-200 text-sm">
              <colgroup>
                {preview.columns.map((column, index) => (
                  <col
                    key={column}
                    style={{ width: index === 0 ? "20rem" : "7rem" }}
                  />
                ))}
              </colgroup>
              <thead className="sticky top-0 bg-slate-950 text-left text-white">
                <tr>
                  {preview.columns.map((column, index) => (
                    <th
                      key={column}
                      className={`px-4 py-3 font-semibold ${
                        index === 0 ? "text-left" : "text-center"
                      }`}
                    >
                      {column}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 bg-white">
                {preview.rows.map((row) => (
                  <tr key={row.join("|")} className="hover:bg-slate-50">
                    {row.map((value, index) => (
                      <td
                        key={`${preview.columns[index]}-${value}`}
                        className={`px-4 py-3 text-slate-700 ${
                          index === 0
                            ? "whitespace-nowrap text-left"
                            : "text-center"
                        }`}
                      >
                        {value || "-"}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        <div className="flex items-center justify-end gap-3 border-t border-slate-200 px-6 py-5">
          <button
            type="button"
            onClick={onClose}
            className="rounded-full border border-slate-300 px-5 py-3 text-sm font-semibold text-slate-700 transition hover:bg-slate-50"
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={() => onConfirm(exportButton.action, exportButton.run)}
            disabled={isPending}
            className="rounded-full bg-slate-950 px-5 py-3 text-sm font-semibold text-white transition hover:bg-slate-800 disabled:cursor-not-allowed disabled:opacity-60"
          >
            {isPending
              ? "Exporting..."
              : `Export ${exportButton.label.replace("Export ", "")}`}
          </button>
        </div>
      </div>
    </div>
  );
}

function getExportPreview(data: GradingData, step: ExportStep) {
  switch (step) {
    case "export-attendance":
      return getAttendanceExportPreview(data);
    case "export-activities":
      return getActivitiesExportPreview(data);
    case "export-participation":
      return getParticipationExportPreview(data);
    case "export-final":
      return getFinalExportPreview(data);
    case "load-data":
    case "review-data":
      return {
        title: "No Export Preview",
        description: "This step does not have an export preview.",
        columns: [],
        rows: [],
      };
  }
}
