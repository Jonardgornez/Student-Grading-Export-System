# Student Grading Export System

This site lets you upload a grading JSON file, preview the export layout, and download Excel reports for:

- Attendance
- Activities
- Participation
- Final Grades

It is built with `Next.js 16`, `Tailwind CSS v4`, and `ExcelJS`.

## Requirements

- Node.js installed
- A valid grading JSON file

## Install And Run

1. Install dependencies:

```bash
npm install
```

2. Start the development server:

```bash
npm run dev
```

3. Open the site in your browser:

```text
http://localhost:3000
```

## How To Use The Site

1. Open the homepage.
2. Click `Upload JSON`.
3. Select your grading JSON file.
4. Wait for the dataset panel to show your subject, teacher, section, and schedule.
5. Click any `Preview Export` button.
6. Review the preview modal first.
7. Click the export button inside the modal to download the Excel file.

If no JSON file is uploaded, export is blocked and the site shows a warning modal.

## Export Flow

The site now uses a preview-first workflow:

1. Upload data
2. Click `Preview Export`
3. Review the exact table layout
4. Confirm export
5. Download `.xlsx`

This means the preview layout matches the exported Excel structure.

## What Each Export Contains

### Attendance

- One row per student
- One column per attendance date
- Attendance codes like `P` and `A`

### Activities

- One row per student
- One column per activity
- Raw activity scores

### Participation

- One row per student
- Participation score column
- Whole-number output

### Final Grades

- One row per student
- Weighted grading columns based on your grading settings
- Separate `Midterm` and `Final` columns
- Whole-number output with no decimals

## JSON Format

The uploaded file must contain these sections:

```json
{
  "students": [],
  "attendance": [],
  "activities": [],
  "participations": [],
  "exams": []
}
```

Optional metadata and grading settings can also be included, such as:

- `meta`
- `subject`
- `grading_settings`

## Example Grading Settings

```json
{
  "grading_settings": {
    "attendance_percent": 10,
    "activities_percent": 20,
    "participation_percent": 10,
    "midterm_percent": 30,
    "final_percent": 30
  }
}
```

These percentages are used in the Final Grades export.

## Notes

- The site starts with no dataset loaded.
- Users must upload a JSON file first.
- Preview is shown before exporting.
- Final grade component values are exported as whole numbers.

## Scripts

```bash
npm run dev
npm run build
npm run start
npm run lint
npm run format
```
