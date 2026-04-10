# 📊 Student Grading Export System

This project allows you to upload student data and export it into Excel files using **ExcelJS**, built with **Next.js 16** and styled with **Tailwind CSS v4**.

You can export:

- ✅ Attendance
- ✅ Activities
- ✅ Participation
- ✅ Final Grades

---

## 🚀 Tech Stack

- Next.js 16 (App Router)
- Tailwind CSS v4
- ExcelJS

---

## 📁 Data Format

Your system expects a JSON file like this:

```json
{
  "students": [],
  "attendance": [],
  "activities": [],
  "participations": [],
  "exams": []
}
```

Each section contains student-related records:

- **students** → list of students
- **attendance** → date + status (Present/Absent)
- **activities** → scores per activity
- **participations** → participation scores
- **exams** → midterm/final scores

---

## ⚙️ Installation

```bash
npm install exceljs file-saver
```

---

## 📤 Export Functions

### 1. Export Attendance

Creates a sheet showing:

- Student Name
- Dates
- Present / Absent

```ts
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

export async function exportAttendance(data) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Attendance");

  ws.columns = [
    { header: "Student Name", key: "name" },
    { header: "Date", key: "date" },
    { header: "Status", key: "status" },
  ];

  data.attendance.forEach((a) => {
    const student = data.students.find((s) => s.id === a.student_id);

    ws.addRow({
      name: student?.full_name,
      date: a.attendance_date,
      status: a.status,
    });
  });

  const buffer = await wb.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), "attendance.xlsx");
}
```

---

### 2. Export Activities

Each activity becomes a column.

```ts
export async function exportActivities(data) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Activities");

  const columns = [
    { header: "Student Name", key: "name" },
    ...data.activities.map((a) => ({
      header: a.title,
      key: `activity_${a.id}`,
    })),
  ];

  ws.columns = columns;

  data.students.forEach((student) => {
    const row = { name: student.full_name };

    data.activities.forEach((activity) => {
      const score = activity.scores.find((s) => s.student_id === student.id);

      row[`activity_${activity.id}`] = score?.score ?? 0;
    });

    ws.addRow(row);
  });

  const buffer = await wb.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), "activities.xlsx");
}
```

---

### 3. Export Participation

```ts
export async function exportParticipation(data) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Participation");

  ws.columns = [
    { header: "Student Name", key: "name" },
    { header: "Score", key: "score" },
  ];

  data.students.forEach((student) => {
    const score = data.participations[0]?.scores.find(
      (s) => s.student_id === student.id,
    );

    ws.addRow({
      name: student.full_name,
      score: score?.score ?? 0,
    });
  });

  const buffer = await wb.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), "participation.xlsx");
}
```

---

### 4. Export Final Grades

Combine everything (attendance + activities + participation + exams)

```ts
export async function exportFinal(data) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Final Grades");

  ws.columns = [
    { header: "Student Name", key: "name" },
    { header: "Final Grade", key: "final" },
  ];

  data.students.forEach((student) => {
    const activityAvg =
      data.activities.reduce((sum, a) => {
        const score = a.scores.find((s) => s.student_id === student.id);
        return sum + (score?.score ?? 0);
      }, 0) / data.activities.length;

    const participation =
      data.participations[0]?.scores.find((s) => s.student_id === student.id)
        ?.score ?? 0;

    const exam =
      data.exams[0]?.scores.find((s) => s.student_id === student.id)?.score ??
      0;

    const final = activityAvg * 0.2 + participation * 0.1 + exam * 0.7;

    ws.addRow({
      name: student.full_name,
      final: final.toFixed(2),
    });
  });

  const buffer = await wb.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), "final-grades.xlsx");
}
```

---

## 🖥️ Example UI (Next.js + Tailwind)

```tsx
<button onClick={() => exportAttendance(data)} className="btn">
  Export Attendance
</button>

<button onClick={() => exportActivities(data)} className="btn">
  Export Activities
</button>

<button onClick={() => exportParticipation(data)} className="btn">
  Export Participation
</button>

<button onClick={() => exportFinal(data)} className="btn">
  Export Final Grades
</button>
```

---

## 💡 Tips

- Always validate your JSON before exporting
- You can add:
  - colors (Excel styles)
  - bold headers
  - auto column width

- You can also create **one file with multiple sheets** instead of separate files

---

## ✅ Summary

This system lets you:

- Upload grading data
- Process it in Next.js
- Export clean Excel reports
- Separate exports per category

---

Happy coding 🚀
