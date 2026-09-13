# SGAM — conciliación de asistencias

[English](README.en.md)

Aplicación de escritorio en Python para cruzar un catálogo de personal, roles de guardia e incidencias con registros de entrada y salida. Presenta un calendario por persona y permite exportar reportes Excel.

Está orientada a la revisión administrativa de asistencias de internos y residentes. El procesamiento utiliza archivos locales; este repositorio no demuestra una integración directa con un dispositivo biométrico ni resultados de uso institucional.

## Funciones presentes en el código

- Lectura y normalización de una plantilla Excel y reportes de escáner.
- Detección de registros ajenos al catálogo y advertencias de validación.
- Evaluación de asistencias, retardos, faltas, guardias y postguardias.
- Aplicación de incidencias y días festivos.
- Consulta por persona, filtros y paneles de análisis.
- Exportación individual, por grupo y consolidada a Excel.

## Ejecutar

Requiere Python con Tk/Tcl para la interfaz. El procesamiento se verificó con Python 3.12; no hay una versión de Python fijada en el repositorio.

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

Dependencias declaradas: CustomTkinter, pandas, openpyxl, Matplotlib, Pillow, xlrd y PyInstaller. No requiere claves de API.

## Formato de entrada actual

El lector activo es `cargar_plantilla` en `ingestion.py`. Las primeras cinco hojas deben existir, aunque las tres de incidencias pueden estar vacías.

| Hoja | Columnas necesarias |
|---|---|
| `1_Catalogo_Personal` | `ID`, `Nombre completo`, `Estatus`, `Tipo de personal` |
| `2_Rol_Guardias` | `ID`, `Rotación` |
| `3_Registro_Incidencias` | Si tiene registros: `ID`, `Ausencia Justificada`, `Fecha Inicio`, `Fecha Termino` |
| `4_Vacaciones` | Las mismas columnas de incidencias |
| `5_Rotaciones` | Las mismas columnas de incidencias |
| `6_Dias_Festivos`, opcional | Columna A: motivo; columna B: fecha |

El catálogo admite información adicional, como `Universidad`, `Especialidad`, `Subespecialidad`, `Alta Especialidad`, `Periodo ingreso` y una columna de foto. `Rotación` utiliza grupos A, B, C o D.

Para una prueba desde la interfaz, utiliza un escáner **.xlsx** con columnas `ID`, `Fecha`, `Entrada` y `Salida`. Debe incluir al menos una entrada con fecha válida y un ID del catálogo. El módulo de lectura admite CSV, pero la selección de un CSV en la interfaz pasa por `pd.ExcelFile` y puede fallar; ese flujo queda pendiente de corrección.

## Ejemplo con datos ficticios

Guarda el siguiente contenido como `demo_local.py` en una copia de trabajo y ejecuta `python demo_local.py`. Crea dos archivos nuevos para la demostración; ejecútalo en una carpeta donde esos nombres no contengan trabajo previo.

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

Abre la app, carga `demo-maestra.xlsx` y `demo-scanner.xlsx`, elige **A** como letra inicial del mes y procesa. Selecciona la persona ficticia y exporta su reporte.

El generador existente `generar_datos_ejemplo.py` declara que usa identificadores procedentes de un escáner real. Antes de compartir datos o capturas del repositorio, confirma su anonimización y autorización. El ejemplo anterior utiliza un identificador creado para esta documentación.

## Reglas y arquitectura

`main.py` inicia la interfaz de `ui.py`. Esta carga archivos mediante `ingestion.py`, llama a `core.py` y exporta con `export.py`; `utils.py` aporta estadísticas.

El motor obtiene el mes más frecuente de los registros aplicables al catálogo, evalúa al personal activo y construye un ciclo A–B–C–D. La interfaz permite elegir la letra inicial del mes. Actualmente llama a `extraer_reglas` con un DataFrame vacío y utiliza valores predeterminados: no carga las reglas desde una hoja `1_Reglas_ID`.

Existe también `SGAM_Proyecto_1.2.py` como implementación alternativa. El procedimiento documentado aquí utiliza `main.py`.

## Verificación y pendientes

El 11 de septiembre de 2026 se comprobó la ruta **Excel ficticio → lectura → procesamiento → exportación individual** con Python 3.12.14 y pandas 3.0.5. Se generaron 31 filas correspondientes a enero y un archivo Excel legible. Es una prueba técnica con una persona ficticia; no mide precisión administrativa ni ahorro de tiempo.

La sintaxis de los nueve archivos Python de la raíz pasó la revisión. No se probaron la interfaz gráfica, todos los formatos de escáner ni el empaquetado.

Pendientes: anonimización de archivos de muestra, pruebas de límites de mes y guardias, validación de CSV en la interfaz, unificación de versiones y preparación del empaquetado. `build_exe.py` referencia carpetas `assets/` y `data/` que no están en la copia revisada, por lo que no se presenta como un comando de instalación verificado.

## Mi participación

Mi rol fue de desarrollo integral de la aplicación: interfaz, procesamiento de datos y exportación de reportes. Es una aplicación de escritorio dentro de mi portafolio de desarrollo full stack.
