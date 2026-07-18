# AsiToPix

AsiToPix is a Windows PowerShell toolkit for organizing astrophotography data and preparing [PixInsight](https://pixinsight.com/) Weighted Batch Preprocessing (WBPP) projects.

It turns capture folders into a predictable archive, builds compact processing projects from symbolic links (or copies), records the project-to-calibration mapping in `project_meta.json`, and can export the masters produced by WBPP back into a reusable calibration library.

The original capture files are treated as read-only. The tools copy them or create links to them; they do not rename, move, edit, or delete them.

## Workflow

```text
capture folders
    |
    +-- ImportSession.ps1 / ImportAll.ps1
    |       -> AstroPhoto\ASIAir archive (Good/Trash, filter, night)
    |
    +-- ImportCalibration.ps1
            -> AstroPhoto\Calibration\<camera>\Source

ASIAir archive + calibration library
    |
    +-- CreateProject.ps1
            -> Processing\<object>\<setup>\Source + Pix + project_meta.json
                         |
                         +-- PixInsight WBPP -> Pix\master
                                                   |
                                                   +-- ExportMasters.ps1
                                                           -> Calibration\<camera>\Master
```

`Get-ImportReport.ps1` summarizes light frames before import. `CombineSeasons.ps1` creates a shared WBPP input tree from several existing processing projects.

## Components

| Component | Purpose |
| --- | --- |
| `ImportSession.ps1` | Imports one light session into the canonical ASIAir archive using copies or file symlinks. |
| `ImportAll.ps1` | Finds and imports every light session under an `Import` staging tree. |
| `ImportCalibration.ps1` | Copies bias, dark, and flat frames into the canonical calibration `Source` tree. |
| `Get-ImportReport.ps1` | Reports frame counts, per-night exposure expressions, and integration time from staging folders; can emit TSV. |
| `CreateProject.ps1` | Selects lights and matching calibration data, then creates a WBPP-ready project with directory symlinks or copied files. |
| `ExportMasters.ps1` | Plans and interactively copies, renames, or replaces WBPP `.xisf` masters in the calibration `Master` tree. |
| `CombineSeasons.ps1` | Combines the `Source` trees of selected processing projects into `Combined\Source`. |
| `Init.cmd` | Configures the current-user PowerShell execution policy and, when elevated, Windows symlink evaluation and long-path support. |
| `Run-CreateProject.cmd` | Launches `CreateProject.ps1` with `pwsh`, falling back to Windows PowerShell. |
| `src\*.psm1` | Reusable path, environment, import, report, metadata, and master-export logic. |
| `tests\*.Tests.ps1` | Pester tests for the reusable modules and script conventions. |

## Requirements

- Windows.
- Windows PowerShell 5.1 or PowerShell 7+.
- PixInsight with WBPP for the processing stage.
- `robocopy`, included with Windows, if project data is copied instead of linked.
- Permission to create symbolic links when using Symlink mode. An elevated shell works; Windows Developer Mode can also permit local symlink creation without elevation.
- Network symlink evaluation enabled when links cross local and network volumes.

Paths may contain spaces and Unicode characters, but the current scripts warn about Cyrillic characters in paths because those paths are not supported by the workflow.

## Initial setup

Clone or download the repository, open PowerShell in its directory, and create an `AstroPhoto` root on any available filesystem drive. Most commands automatically search for `*:\AstroPhoto`; if there is no unique match, they prompt for a path. Commands that expose `-AstroPhotoRoot` can be given the path explicitly.

The minimum useful root is:

```text
D:\AstroPhoto\
├── Import\
├── Calibration\
└── Processing\
```

To perform the Windows setup interactively:

```bat
Init.cmd
```

`Init.cmd` sets the current-user PowerShell execution policy to `RemoteSigned`. If you approve elevation, it also enables local/network symlink evaluation and the machine-wide `LongPathsEnabled` registry value. It does not install dependencies.

If only the execution policy is blocking local scripts, either run `Enable-PowerShellScripts.cmd` once or use PowerShell directly:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Copy mode does not require symlink privileges.

## Directory conventions

### Light staging and archive

The batch importer expects this staging layout:

```text
AstroPhoto\Import\
└── <setup>\
    └── Light\
        └── <object>\
            └── *.fit / *.fits / supported camera RAW files
```

`ImportAll.ps1` also recognizes a `Lights` folder. `Get-ImportReport.ps1` currently reads the singular `Light` layout and FIT/FITS files only.

Imported lights are organized as:

```text
AstroPhoto\ASIAir\
└── <object>\
    └── <season>\
        └── <scope> @ <camera>\
            ├── Good\
            │   └── <filter>\
            │       └── yy.MM.dd[-<exposure>]\
            └── Trash\
                └── <filter>\
                    └── yy.MM.dd[-<exposure>]\
```

The exposure suffix is added only when a filter/night contains mixed exposure lengths. Existing filenames found anywhere under `Good` or `Trash` are not imported again; a filename already in `Trash` remains excluded.

FITS metadata is read from ASIAir-style filenames. Common camera RAW formats are also accepted for lights, with file timestamps used when capture timestamps are unavailable. Missing object, setup, camera, or filter information is requested interactively.

### Night dates

All lights and calibration frames use a local noon-to-noon observing night:

- A capture before 12:00 belongs to the previous calendar date.
- A capture at or after 12:00 belongs to the current calendar date.
- Night folders use the start date in `yy.MM.dd` format.

For example, captures at `2026-07-19 02:30` and `2026-07-18 22:00` both belong to night `26.07.18`.

### Calibration library

`ImportCalibration.ps1` expects its source root to contain `bias`/`biases`, `dark`/`darks`, or `flat`/`flats` folders directly below it. FIT, FITS, and ARW files are supported.

The generated library follows this shape:

```text
AstroPhoto\Calibration\<camera>\
├── Source\
│   ├── biases\gain<gain>\<temperature>C\yy.MM\
│   ├── darks\gain<gain>\<temperature>C\<exposure>sec\yy.MM\
│   └── flats\<setup>\yy.MM.dd <filter> [<angle>deg]\
└── Master\
    ├── biases\...
    ├── darks\...
    ├── flats\...
    └── flat-darks\...
```

Temperatures are rounded to 5-degree folders. Metadata not present in filenames is requested before the import plan is applied. Calibration files are copied, never overwritten, and same-name/different-size conflicts stop that entry.

### Processing project

By default, `CreateProject.ps1` creates a setup below `AstroPhoto\Processing\<object>`:

```text
AstroPhoto\Processing\<object>\
└── <season>_<scope>_<camera>\
    ├── Source\
    │   ├── lights\
    │   ├── biases\
    │   ├── darks\
    │   ├── flats\
    │   └── flat-darks\
    ├── Pix\
    └── project_meta.json
```

Entries below `Source` carry WBPP metadata in their generated names and point to the selected archive/calibration directories. `project_meta.json` records schema version 2 metadata, WBPP paths, cameras, filters, targets, calibration sources, and the intended master destinations.

## Usage

All scripts are interactive when required values are omitted. Run them from the repository root. Quote paths containing spaces.

### 1. Inspect staged lights

```powershell
.\Get-ImportReport.ps1 -ImportPath 'D:\AstroPhoto\Import'
```

Write a pipeline-friendly TSV report:

```powershell
.\Get-ImportReport.ps1 -ImportPath 'D:\AstroPhoto\Import' -Tsv |
    Set-Content -LiteralPath '.\import-report.tsv' -Encoding UTF8
```

Multiple import roots can be supplied as an array:

```powershell
.\Get-ImportReport.ps1 -ImportPath 'D:\AstroPhoto\Import', 'N:\AstroPhoto\Import'
```

### 2. Import lights

Preview and import all staging sessions using the default Copy mode:

```powershell
.\ImportAll.ps1 `
    -AstroPhotoRoot 'D:\AstroPhoto' `
    -ImportRoot 'D:\AstroPhoto\Import' `
    -SeasonName '2026' `
    -ImportMode Copy `
    -WhatIf

.\ImportAll.ps1 `
    -AstroPhotoRoot 'D:\AstroPhoto' `
    -ImportRoot 'D:\AstroPhoto\Import' `
    -SeasonName '2026' `
    -ImportMode Copy
```

Import one session instead:

```powershell
.\ImportSession.ps1 `
    -SourcePath 'E:\ASIAIR\SQA55\Light\M 31' `
    -AstroPhotoRoot 'D:\AstroPhoto' `
    -ObjectName 'M 31' `
    -SeasonName '2026' `
    -TelescopeName 'SQA55 @ 1.0x' `
    -ImportMode Symlink
```

`SourcePath` may be a folder, a supported light file, or an object name that can be resolved under an `AstroPhoto\Import` tree. Copy is the interactive default.

### 3. Import calibration frames

The destination `AstroPhoto\Calibration` directory must already exist.

```powershell
.\ImportCalibration.ps1 `
    -SourcePath 'E:\ASIAIR\SQA55' `
    -AstroPhotoRoot 'D:\AstroPhoto' `
    -WhatIf

.\ImportCalibration.ps1 `
    -SourcePath 'E:\ASIAIR\SQA55' `
    -AstroPhotoRoot 'D:\AstroPhoto'
```

Optional fallback parameters include `-CameraName`, `-Gain`, `-TemperatureC`, `-DarkExposureSeconds`, `-FilterName`, and `-AngleDegrees`. Filename metadata takes precedence; fallback values fill only missing data.

### 4. Create a PixInsight project

```powershell
.\CreateProject.ps1
```

Or use `Run-CreateProject.cmd`.

The script:

1. Finds the `AstroPhoto` root.
2. Accepts a light folder, a FITS file, or an object name from the ASIAir archive.
3. Confirms the detected object, season, scope, and project path.
4. Scans all filter/night folders for that object and setup.
5. Selects matching master or source calibration folders and asks for flat choices where necessary.
6. Shows the complete project tree and asks whether to create directory symlinks or copy the data.
7. Writes `project_meta.json` beside `Source` and `Pix`.

Use `-WhatIf` to preview filesystem changes. If the generated `Source` tree already exists, read the cleanup prompt carefully: accepting it replaces that generated input tree, not the archived source data.

### 5. Run WBPP

In PixInsight WBPP:

- Load data from the generated `Source` subfolders.
- Use the generated `Pix` directory as the output location.
- Configure WBPP grouping keywords printed by `CreateProject.ps1`: `CAM`, `FILTER`, `TARGET`, `GAIN`, `TEMP`, and `EXP`.
- Keep integration masters under `Pix\master`, which is the path recorded in `project_meta.json`.

The script prints the recommended pre/post keyword placement and a camera/filter mapping summary at the end of project creation.

### 6. Export WBPP masters

Preview the export:

```powershell
.\ExportMasters.ps1 `
    -MetaPath 'D:\AstroPhoto\Processing\M_31\2026_SQA55_1.0x_ASI2600\project_meta.json' `
    -WhatIf
```

Apply it interactively:

```powershell
.\ExportMasters.ps1 `
    -MetaPath 'D:\AstroPhoto\Processing\M_31\2026_SQA55_1.0x_ASI2600\project_meta.json'
```

The exporter reads `.xisf` bias, dark, and flat masters from `Pix\master`, strips generated WBPP tags from canonical filenames, deduplicates masters that map to the same destination, and asks before each copy, rename, or replacement. Conflicts are reported rather than overwritten.

`-AstroPhotoRoot` is normally unnecessary for schema-version-2 metadata, but it can be supplied for older projects that require archive link discovery.

### 7. Combine processing seasons

```powershell
.\CombineSeasons.ps1
```

Choose a root whose immediate child directories each contain a `Source` directory. The script asks which children to include and creates:

```text
<selected root>\Combined\
├── Source\
└── Pix\
```

This script is currently experimental: it is interactive, does not support `-WhatIf`, does not generate merged `project_meta.json`, and reports same-name link conflicts without resolving them. If `Combined` already exists, accepting the cleanup prompt recursively removes that generated directory before rebuilding it.

## Safety notes

- Source capture directories are read-only from AsiToPix's perspective.
- Import plans never overwrite ordinary destination files.
- Calibration import always copies; light import and project creation offer Copy/Symlink modes as documented above.
- Existing project links are validated before replacement. Ordinary files or directories are not silently replaced by links.
- Review `-WhatIf` output before large imports, project rebuilds, or master exports.
- `Init.cmd` changes user/machine Windows settings as described in [Initial setup](#initial-setup); it is optional for Copy-only workflows.
- `CombineSeasons.ps1` has the limitations listed in its usage section and should be used only with a clearly identified generated `Combined` directory.

## Development

Reusable functions live under `src`, while executable workflows remain in the top-level scripts. Tests use Pester:

```powershell
Invoke-Pester -Path .\tests
```

Run PSScriptAnalyzer when it is installed:

```powershell
Invoke-ScriptAnalyzer -Path . -Recurse
```

The project targets both Windows PowerShell 5.1 and PowerShell 7+, so changes should avoid syntax or APIs unavailable in Windows PowerShell 5.1.
