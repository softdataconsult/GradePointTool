# GradePointTool

R toolkit for turning exam scores into **Grade Points** and **Letter Grades**, producing polished **Excel workbooks** with formatted tables, **percent summaries**, and **attractive viridis plots**. Works with messy headers, supports single files and **batch processing per course**, and includes a Shiny-ready path for quick UI use.

---

## Features

* Reads **Excel/CSV** and tolerates headers like `S/N, NAME, MATRIC NO, PRAC (50), CA (10), EXAM (40), TOTAL (100)`
* Computes **GradePoint** (rounded to **2 decimals**) and **Grade** using your exact scale
* Builds `TOTAL (100)` from PRAC + CA + EXAM if needed
* Writes a single workbook with **Excel Tables**:

  * `Grades` (student-level results)
  * `Summary_Fine` (A+, A, B+, …, F with Count and Percent)
  * `Summary_AtoF` (A–F with Count and Percent)
  * `Plots` (embedded PNGs)
* Exports **bar charts** and **histograms** using colorblind-friendly **viridis** palettes
* **Batch wrapper** creates one folder per course: `results/<COURSE_NAME>/...`
* Robust against duplicate or odd column names; fixes case-insensitive clashes automatically

---

## Install

```r
install.packages(c("readxl","openxlsx","dplyr","janitor","readr","ggplot2","tidyr"))
```

Clone the repo and source the main script:

```r
source("R/process_exam_sheet.R")
```

---

## Quick Start

**Single file**

```r
res <- process_exam_sheet(
  in_path  = "data/exam_scores.xlsx",
  out_path = "results/exam_scores/exam_scores_with_grades.xlsx",
  save_plots = TRUE,
  write_excel_tables = TRUE
)

res$summary_fine
res$plot_bar_fine
```

**Batch (many courses, one folder per course)**

```r
files <- list.files("data", pattern = "\\.xlsx?$", full.names = TRUE)
process_many_courses(files, out_root = "results")
```

Folder output example:

```
results/
└── exam_scores/
    ├── exam_scores_with_grades.xlsx
    ├── exam_scores_bar_fine.png
    ├── exam_scores_bar_broad.png
    └── exam_scores_hist.png
```

---

## What You Get in the Workbook

* **Grades**: `Name, Matric_No, Total_100, GradePoint, Grade`
  GradePoint is numeric with **two decimals** in R and **Excel formatted to 0.00**.
* **Summary_Fine**: A+, A, B+, B, C+, C, D, E, F with **Count** and **Percent**
* **Summary_AtoF**: A–F with **Count** and **Percent**
* **Plots**: three images inserted on one sheet

---

## Grade Scale (default)

| Score  | GP   | Grade |
| ------ | ---- | ----- |
| 0–19   | 0.00 | F     |
| 20–29  | 1.00 | F     |
| 30–39  | 1.50 | F     |
| 40–44  | 2.00 | E     |
| 45–49  | 2.25 | D     |
| 50–54  | 2.50 | C     |
| 55–59  | 2.75 | C+    |
| 60–64  | 3.00 | B     |
| 65–69  | 3.50 | B+    |
| 70–74  | 3.75 | A     |
| 75–100 | 4.00 | A+    |

**Change it:** edit `grade_point()` and `letter_grade()` in `R/process_exam_sheet.R`.

---

## Headers the Tool Detects

After `janitor::clean_names()`, the script looks for:

* **Name**: `name, student_name`
* **Matric**: `matric_no, reg_no, matric`
* **PRAC**: `prac_50, practical_50, prac`
* **CA**: `ca_10, continuous_assessment_10, ca`
* **EXAM**: `exam_40, exam`
* **TOTAL**: `total_100, total`

If `TOTAL (100)` is missing, it uses PRAC + CA + EXAM.

---

## Shiny Hook (optional)

You can wrap `process_exam_sheet()` in a minimal Shiny uploader to drag-and-drop sheets, preview results, and download the final Excel. The codebase includes clear entry points for this.

---

## Troubleshooting

* **“colNames must be a unique vector (case sensitive)”**
  The script enforces unique names even when lower-cased with `fix_names_excel()`. If you still hit this, print `names(out)` and remove collisions in the source.

* **“Cannot overwrite existing table with another table.”**
  Do not write to the same sheet twice. This repo writes each table once and applies styles in place.

* **Strange score values (e.g., “78%” or “ 40 ”)**
  The parser uses `readr::parse_number()` which cleans these to numeric.

---

## Contributing

Pull requests are welcome. Please:

1. Keep functions pure and vectorized where possible.
2. Add examples and update README when behavior changes.
3. Run `lintr` or keep a tidy style.

---

## License

MIT. See `LICENSE`.

---

## Acknowledgments

Built with gratitude to the R community and the authors of **readxl**, **openxlsx**, **dplyr**, **janitor**, **readr**, **ggplot2**, and **tidyr**.
