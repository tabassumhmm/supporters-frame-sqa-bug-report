# Supporters Frame SQA Bug Report

Prepared by **Sumaiya Nasrin** for the Supporters Frame website:
<https://supportersframe.com/>

This submission contains a Python desktop application, an Excel bug report, a
Windows executable, and the screenshot evidence referenced by the report.

## Deliverables

| File | Description |
|---|---|
| `Bug_Report_Sumaiya_Nasrin.py` | Python source for the report viewer and Excel exporter |
| `Bug_Report_Sumaiya_Nasrin.xlsx` | Generated report workbook |
| `dist/Bug_Report_Sumaiya_Nasrin.exe` | Single-file Windows desktop executable |
| `evidence/` | Screenshot files referenced by the bug findings |
| `Bug_Report_Sumaiya_Nasrin_Submission.zip` | Packaged source, workbook, executable, and evidence folder |

The submission ZIP contains the three deliverables at its root and all evidence
screenshots under `evidence/`.

## Run the Windows application

On Windows, double-click:

```text
dist\Bug_Report_Sumaiya_Nasrin.exe
```

The executable is a PyInstaller one-file build and includes its runtime
dependencies and evidence screenshots. Python does not need to be installed to
run this executable. Windows may show a security or SmartScreen prompt for an
unsigned application; only run it if you trust the source of the file.

The application lets you filter the findings by severity and status, inspect a
selected finding, view the summary, check screenshot references, and export an
Excel report.

## Run from Python

Use Python 3 on Windows. The source requires `openpyxl`; Pillow is optional but
is needed to embed evidence images in the workbook. Tkinter is needed for the
desktop interface and is usually included with standard Windows Python
installations.

Install the Python packages if needed:

```powershell
python -m pip install openpyxl Pillow
```

Open the desktop application from the project directory:

```powershell
python .\Bug_Report_Sumaiya_Nasrin.py
```

Generate an Excel report without opening the desktop application:

```powershell
python .\Bug_Report_Sumaiya_Nasrin.py --export
```

To specify an output path:

```powershell
python .\Bug_Report_Sumaiya_Nasrin.py --export "C:\Reports\Bug_Report.xlsx"
```

Without an explicit output path, the exporter writes
`Bug_Report_Sumaiya_Nasrin.xlsx` beside the script (or beside the executable).
If that workbook is locked, for example because it is open in Excel, the
exporter tries an unused numbered filename such as
`Bug_Report_Sumaiya_Nasrin_1.xlsx`.

## Workbook contents

The workbook contains six sheets:

1. **Bug Report** - 11 findings with modules, preconditions, reproduction
   steps, expected and actual results, severity, reproducibility, status, and
   evidence filenames.
2. **Summary** - finding counts by severity and status, Lighthouse figures,
   observations not counted as bugs, and the current verification limitation.
3. **Test Log** - recorded checks and their results.
4. **Evidence** - screenshots embedded from the `evidence/` folder.
5. **Coverage** - areas tested, partial or untested areas, and limitations.
6. **Retest Plan** - prioritized cases for follow-up testing.

## Evidence files

The evidence folder should be named `evidence` and placed beside the Python
script. The executable contains its evidence images internally. The current
submission includes 14 screenshots referenced by the findings. When replacing
or adding evidence, keep filenames consistent with the references in the
script; missing or unreadable files are reported during export.

## Test scope and limitations

The findings document the homepage and campaign-creation checks performed in
Chrome, including frame upload, campaign fields, navigation, Lighthouse
accessibility observations, and network/performance observations. The workbook
also makes untested areas and a follow-up test plan explicit.

During a later verification attempt on **3 November 2026**, the homepage
returned an inactive-store/forbidden (403) page and the campaign-creation route
returned 404. Therefore the original findings are supported by the saved
evidence from the earlier test runs, but could not be independently retested at
that time. Do not treat the pending retest cases as passed. Dedicated
Edge/Firefox, tablet/real-device, authenticated publishing, and manual
screen-reader testing were not completed.

## Reproducibility and validation

The Excel report is generated from the bug data in the Python source. The
executable is built from that same source and includes the evidence images.
Validation performed for this submission confirmed that the script and
executable can generate the workbook and that the workbook contains the
documented findings and embedded evidence. Microsoft Excel desktop was not
available for a direct GUI-open check; the `.xlsx` file was checked as a valid
workbook package and loaded successfully with `openpyxl`.
