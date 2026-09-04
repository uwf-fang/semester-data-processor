# Pivot Table Reporting Guide for `assessment_results.csv`

This guide explains how to open `assessment_results.csv` in spreadsheet software (**Microsoft Excel** or **Google Sheets**) and create interactive Pivot Tables for ABET, SACSCOC, and departmental accreditation reporting.

---

## 1. Overview of `assessment_results.csv` Data Fields

When `assessment-data-processor` runs, it generates `assessment_results.csv` containing detailed per-student assessment outcomes. To make Pivot Table aggregation effortless, two **numeric helper columns** (`Pass Count` and `Total Student Count`) are automatically included:

| Field Name | Type | Sample Value | Role in Pivot Table |
| :--- | :--- | :--- | :--- |
| `Accreditation` | Text | `ABET`, `SACSCOC` | **Filter / Slicer** |
| `Assessment Name` | Text | `ABET LO.PI 1.1 - Language Concepts` | **Row Header / Slicer** |
| `Program` | Text | `CS`, `SD`, `Cyber`, `CE`, `EE` | **Filter / Column Header** |
| `Full Major` | Text | `Computer Science` | **Filter / Slicer** |
| `Section ID` | Text | `16197`, `17089`, `18451` | **Row Header** |
| `Assessed Score` | Number/Text | `100`, `95`, `60`, `N/A` | **Value (Max / Average)** |
| `Pass/Fail Status` | Text | `Pass`, `Fail` | **Column Header** |
| `Pass Count` | Number | `1` (if Pass), `0` (if Fail) | **Value (Sum)** |
| `Total Student Count` | Number | `1` (for every student) | **Value (Sum)** |

---

## 2. Step-by-Step Instructions

### Microsoft Excel

1. **Open the Data File**:
   - Open `assessment_results.csv` in Microsoft Excel.

2. **Insert Pivot Table**:
   - Select all data (`Ctrl+A` on Windows, `Cmd+A` on Mac).
   - Go to the ribbon tab: **Insert** $\rightarrow$ **PivotTable** $\rightarrow$ **From Table/Range**.
   - Select **New Worksheet** and click **OK**.

3. **Configure Pivot Table Fields**:
   - Drag **`Accreditation`** and **`Program`** into the **Filters** area (top-left).
   - Drag **`Assessment Name`** and **`Section ID`** into the **Rows** area.
   - Drag **`Pass/Fail Status`** into the **Columns** area.
   - Drag **`Pass Count`** and **`Total Student Count`** into the **Values** area (summarized as **SUM**).

4. **Add Pass Rate % (Calculated Field)**:
   - Click anywhere inside the Pivot Table.
   - Go to **PivotTable Analyze** (or **Options**) $\rightarrow$ **Fields, Items, & Sets** $\rightarrow$ **Calculated Field...**
   - **Name**: `Pass Rate %`
   - **Formula**: `='Pass Count' / 'Total Student Count'`
   - Click **OK**.
   - Select the `Pass Rate %` column in the Pivot Table, right-click $\rightarrow$ **Number Format...** $\rightarrow$ **Percentage** (`0.0%`).

---

### Google Sheets

1. **Open the Data File**:
   - Import or open `assessment_results.csv` in Google Sheets.

2. **Insert Pivot Table**:
   - Click **Insert** $\rightarrow$ **Pivot table**.
   - Choose **New sheet** and click **Create**.

3. **Configure Pivot Table Fields**:
   - **Rows**: Click **Add** $\rightarrow$ select `Assessment Name`, then click **Add** $\rightarrow$ select `Section ID`.
   - **Columns**: Click **Add** $\rightarrow$ select `Pass/Fail Status`.
   - **Values**: Click **Add** $\rightarrow$ select `Pass Count` (Summarize by: **SUM**). Add `Total Student Count` (Summarize by: **SUM**).
   - **Filters**: Click **Add** $\rightarrow$ select `Program` (Filter by Status $\rightarrow$ select `CS`), then add `Accreditation` (select `ABET`).

4. **Add Calculated Pass Rate % Field**:
   - Under **Values**, click **Add** $\rightarrow$ **Calculated field**.
   - Set **Formula**: `='Pass Count' / 'Total Student Count'`.
   - Format the numbers as **Format** $\rightarrow$ **Number** $\rightarrow$ **Percent**.

---

## 3. Recommended Pivot Table Report Layouts

### Layout A: ABET / SACSCOC Section Breakdown Table
Generates section pass/fail ratios and pass percentages for targeted majors:

- **Filter**: `Program` = `CS`, `Accreditation` = `ABET`
- **Rows**: `Assessment Name`, `Section ID`
- **Columns**: `Pass/Fail Status` (`Pass`, `Fail`)
- **Values**: `Sum of Pass Count`, `Sum of Total Student Count`, `Pass Rate %`

**Output Preview**:

```
Assessment Name                   Section ID   Pass   Fail   Total Enrolled   Pass Rate %
========================================================================================
ABET LO.PI 1.1 - Language Concepts 16197        16     1      17               94.1%
                                  17089         5     0       5              100.0%
----------------------------------------------------------------------------------------
Total ABET LO.PI 1.1                           21     1      22               95.5%
```

---

### Layout B: Multi-Major Program Comparison Matrix
Compares pass rates across different degree programs (`CS`, `SD`, `Cyber`, `CE`, `EE`):

- **Filter**: `Accreditation`, `Assessment Name`
- **Rows**: `Program`
- **Columns**: `Pass/Fail Status`
- **Values**: `Sum of Total Student Count` (Displayed as **% of Row Total**)

**Output Preview**:

```
Program   Pass (%)   Fail (%)   Total Students
==============================================
CS        95.5%      4.5%       22
SD        100.0%     0.0%       3
Cyber     100.0%     0.0%       2
----------------------------------------------
Total     96.3%      3.7%       27
```

---

### Layout C: Multi-Outcome Course Dashboard
Summarizes performance across multiple learning outcomes / PIs assessed in the same course:

- **Filter**: `Program` = `CS`
- **Rows**: `Assessment Name`
- **Columns**: `Section ID`
- **Values**: `Pass Rate %`

**Output Preview**:

```
Assessment Name                           Section 16197   Section 17089   Total Pass Rate %
===========================================================================================
ABET LO.PI 1.1 - Language Concepts        94.1%           100.0%          95.5%
ABET LO.PI 1.2 - Program Implementation   94.1%           100.0%          95.5%
```
