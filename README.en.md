# SGAM — attendance reconciliation

[Español](README.md)

A Python desktop application that matches a staff catalog, duty groups and recorded absences against check-in/check-out records. It provides a per-person calendar and Excel report exports.

It targets administrative attendance review for medical interns and residents. Processing uses local files; this repository does not demonstrate direct biometric-device integration or institutional usage outcomes.

## Implemented scope

- Loading and normalizing an Excel template and scanner reports.
- Identifying records outside the staff catalog and validation warnings.
- Evaluating attendance, lateness, absences, duty and post-duty days.
- Applying recorded absences and holidays.
- Per-person views, filters and analysis panels.
- Individual, filtered and consolidated Excel exports.

## Run

Python with Tk/Tcl is required for the GUI. Processing was checked with Python 3.12; the repository does not pin a Python version.

```bash
git clone https://github.com/fabrizzio2901/SGAM.git
cd SGAM
python -m venv .venv
```

```powershell
# Windows PowerShell
.\.venv\Scripts\Activate.ps1
```

```bash
# macOS / Linux
source .venv/bin/activate
```

```bash
python -m pip install -r requirements.txt
python main.py
```

Declared dependencies: CustomTkinter, pandas, openpyxl, Matplotlib, Pillow, xlrd and PyInstaller. No API keys are needed.

## Current input format

The active loader is `cargar_plantilla` in `ingestion.py`. The first five worksheets must exist; the three absence worksheets may be empty.

| Worksheet | Required columns |
|---|---|
| `1_Catalogo_Personal` | `ID`, `Nombre completo`, `Estatus`, `Tipo de personal` |
| `2_Rol_Guardias` | `ID`, `Rotación` |
| `3_Registro_Incidencias` | For nonempty data: `ID`, `Ausencia Justificada`, `Fecha Inicio`, `Fecha Termino` |
| `4_Vacaciones` | Same absence columns |
| `5_Rotaciones` | Same absence columns |
| Optional `6_Dias_Festivos` | Column A: reason; column B: date |

Additional catalog fields include `Universidad`, `Especialidad`, `Subespecialidad`, `Alta Especialidad`, `Periodo ingreso` and a photo column. Duty groups in `Rotación` are A, B, C or D.

For a GUI demonstration, use an **.xlsx** scanner file with `ID`, `Fecha`, `Entrada` and `Salida`. At least one valid dated check-in must match an ID in the catalog. The loader supports CSV, but the GUI routes a selected CSV through `pd.ExcelFile` and may fail; that path needs correction.

## Fictional example

Save the following as `demo_local.py` in a working copy and run `python demo_local.py`. It writes two demonstration files; run it where those filenames do not hold existing work.

```python
import pandas as pd

with pd.ExcelWriter("demo-maestra.xlsx", engine="openpyxl") as writer:
    pd.DataFrame([{
        "ID": "DEMO001", "Nombre completo": "PERSONA FICTICIA",
        "Estatus": "activo", "Tipo de personal": "interno"
    }]).to_excel(writer, sheet_name="1_Catalogo_Personal", index=False)
    pd.DataFrame([{"ID": "DEMO001", "Rotación": "A"}]).to_excel(
        writer, sheet_name="2_Rol_Guardias", index=False)
    for sheet in ["3_Registro_Incidencias", "4_Vacaciones", "5_Rotaciones"]:
        pd.DataFrame(columns=[
            "ID", "Ausencia Justificada", "Fecha Inicio", "Fecha Termino"
        ]).to_excel(writer, sheet_name=sheet, index=False)

pd.DataFrame([{
    "ID": "DEMO001", "Fecha": "2026-01-02",
    "Entrada": "07:00", "Salida": "15:00"
}]).to_excel("demo-scanner.xlsx", index=False)
```

Open the app, load `demo-maestra.xlsx` and `demo-scanner.xlsx`, select **A** as the month's starting group and process. Select the fictional person and export their report.

The existing `generar_datos_ejemplo.py` says that it uses identifiers from a real scanner. Verify anonymization and authorization before sharing repository data or screenshots. The example above uses an identifier created for this documentation.

## Rules and architecture

`main.py` starts `ui.py`, which loads files through `ingestion.py`, invokes `core.py` and exports through `export.py`. `utils.py` provides statistics.

The engine selects the most frequent month among applicable scanner records, processes active staff and constructs an A–B–C–D cycle. The GUI lets the user choose the initial group. It currently passes an empty DataFrame to `extraer_reglas`, so it uses defaults rather than loading rules from a `1_Reglas_ID` worksheet.

`SGAM_Proyecto_1.2.py` is an alternative implementation also present in the repository. This guide documents the `main.py` entry point.

## Verification and pending work

On September 11, 2026, the **fictional Excel → loading → processing → individual export** path was checked with Python 3.12.14 and pandas 3.0.5. It produced 31 January rows and a readable Excel workbook. This is a technical check with one fictional person, not a measure of administrative accuracy or time saved.

All nine root Python files passed syntax parsing. The GUI, every scanner format and executable packaging were not tested.

Pending work includes anonymizing sample data, month-boundary and duty tests, GUI CSV loading, consistent versioning and packaging preparation. `build_exe.py` references `assets/` and `data/`, which are absent from the reviewed checkout, so packaging is not presented as a verified installation command.

