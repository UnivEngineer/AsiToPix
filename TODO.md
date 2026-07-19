# TODO

Заметки по результатам чтения `CreateProject.ps1`, `ExportMasters.ps1` и `CombineSeasons.ps1`.

## Критичные риски безопасности данных

- [ ] `CreateProject.ps1`: перед `Remove-Item $sourcePath -Recurse -Force` проверять, что удаляемая папка действительно создана проектом и находится внутри ожидаемого project root.
- [ ] `CombineSeasons.ps1`: перед `Remove-Item $combineRoot -Recurse -Force` валидировать resolved path и использовать marker/metadata созданного проектом каталога.
- [ ] `CreateProject.ps1`: перед созданием или заменой target path проверять, что существующий объект является symbolic link, а не обычным файлом или папкой.
- [x] `ExportMasters.ps1`: заменить безусловный `Copy-Item -Force` на интерактивное безопасное копирование и замену мастеров.
- [ ] Добавить `-WhatIf` через `[CmdletBinding(SupportsShouldProcess)]` для операций создания, удаления, копирования и замены файлов/каталогов.
- [x] `ExportMasters.ps1`: поддержать `-WhatIf` для копирования, переименования и замены мастеров.
- [ ] Не менять системные настройки (`fsutil`, `HKLM:\...\LongPathsEnabled`) без явного параметра/подтверждения и документации требований администратора.

## Ошибки и пограничные случаи

- [ ] `CreateProject.ps1`: проверять существование `$lightsRoot` до `Get-ChildItem`, выдавать ошибку с полным путем.
- [ ] `CreateProject.ps1`: сбрасывать `$fFound` и `$foundIn` для каждой сессии, чтобы выбор flats из предыдущей сессии не протекал в следующую.
- [ ] `CreateProject.ps1`: обработать случай, когда flat-папки есть, но ни одна не совпала по фильтру; сейчас возможен выбор из пустого списка.
- [ ] `CreateProject.ps1`: защитить `Substring(0,8)` для коротких или нестандартных имен flat-папок.
- [ ] `CreateProject.ps1`: привести парсинг дат к одному формату; код местами допускает 4-значный год, но `ParseDate` принимает только `yy.MM.dd`.
- [ ] `CreateProject.ps1`: уточнить, какой формат даты используется в ASIAir-папках: `yy.MM.dd` или `dd.MM.yy`; сейчас комментарии и примеры могут читаться неоднозначно.
- [ ] `CreateProject.ps1`: исправить расхождение комментария и поведения для mono `L`: комментарий говорит `L -> L/RGB`, код делает `L -> L/L`.
- [ ] `CreateProject.ps1`: добавить `-ErrorAction Stop` и `try/catch` вокруг создания symbolic links, чтобы не игнорировать ошибки.
- [ ] `CreateProject.ps1`: при существующем target path не просто пропускать, а проверять тип и target существующего link.
- [ ] `CreateProject.ps1`: проверить regex извлечения фильтра из имени файла, особенно выражение с escaped camera id.
- [ ] `CreateProject.ps1`: проверить fallback defaults (`gain=120`, `exp=300s`, `temp=-20`) и явно предупреждать, когда metadata не распознана.
- [ ] `CreateProject.ps1`: проверить подсчет total size, когда `Measure-Object -Sum` возвращает `$null`.
- [x] Не трактовать фильтр `S` в именах flat-папок как единицу измерения экспозиции.
- [x] `ExportMasters.ps1`: валидировать структуру и обязательные поля `project_meta.json` (`PixPath`, `Scope`, `Cameras[].Name`).
- [x] `ExportMasters.ps1`: не классифицировать любой `masterDark` с экспозицией `< 10s` как flat-dark без дополнительного признака.
- [x] `ExportMasters.ps1`: убрать hardcoded resolution `6248x4176` или получать его из metadata/имени файла.
- [x] `ExportMasters.ps1`: сделать matching камеры точнее, чем `*$camFull*`, чтобы избежать ложных совпадений.
- [x] `ExportMasters.ps1`: обрабатывать отсутствие подходящих `.xisf` мастеров с понятным сообщением.
- [x] `CreateProject.ps1`: сохранять в `project_meta.json` точные `CalibrationSources` и destination-папки, необходимые экспортёру.
- [x] `ExportMasters.ps1`: читать `CalibrationSources` из metadata с fallback на project Source links для старых проектов.
- [x] Использовать каноническую project-папку `Source\flat-darks`, сохранив чтение старой `Source\FlatDarks`.
- [ ] `CombineSeasons.ps1`: добавить стратегию конфликтов одинаковых имен symlink-папок вместо сообщения "shouldn't happen".
- [ ] `CombineSeasons.ps1`: создавать или объединять `project_meta.json` для `Combined`.
- [ ] `CombineSeasons.ps1`: если входной элемент не symlink, явно решать, допустимо ли ссылаться на обычную папку внутри сезонного `Source`.
- [ ] `CombineSeasons.ps1`: при выводе existing target безопасно обрабатывать обычные папки, у которых нет `Target`.

## Недочеты структуры скриптов

- [ ] Добавить во все executable scripts `[CmdletBinding()]`, explicit `param()`, `Set-StrictMode -Version Latest`, `$ErrorActionPreference = 'Stop'`.
- [x] Заменить hardcoded `Z:\AstroPhoto` и `Z:\AstroPhoto\Calibration` на автоматический поиск `*:\AstroPhoto` с ручным fallback.
- [ ] Удалить неиспользуемый `$localBase` из `CreateProject.ps1` или оформить как параметр, если он нужен.
- [ ] Вынести повторяемые функции и правила в модуль под `src/`.
- [ ] Вынести интерактивные подтверждения `Y/n` в общий helper.
- [ ] Вынести безопасное создание каталогов, symlink и copy operations в общий helper.
- [ ] Разделить scan/plan/apply: сначала строить план действий, потом применять его после подтверждения.
- [x] `ExportMasters.ps1`: реализовать scan/plan/apply, подтверждение каждого изменения, исправление legacy-имён и транзакционную замену мастеров.
- [ ] Заменить ручную конкатенацию путей на `Join-Path`, `Resolve-Path` и `[System.IO.Path]` там, где это еще не сделано.
- [ ] Сделать сообщения об ошибках actionable и включать affected path.
- [ ] Привести PowerShell function names к approved verbs.
- [ ] Убрать объявление `FlatFolderMatchesFilter` и `ParseDate` из внутреннего цикла сессий.
- [ ] Упростить и покрыть тестами алгоритм выбора flats по дате/углу/Origin.

## Тесты и проверка

- [x] Добавить выровненный и TSV-отчет по сабам из `*:\AstroPhoto\Import` с разбивкой по сетапам, объектам, фильтрам, ночам и экспозициям.
- [x] Добавить Pester tests под `tests/`.
- [x] Покрыть тестами очистку имён WBPP masters, дедупликацию dark/bias/flat и выбор dark/flat-dark по project Source links.
- [x] Покрыть тестами schema v2 `project_meta.json` и экспорт без чтения project Source links.
- [ ] Покрыть тестами filter mapping для OSC и Mono.
- [ ] Покрыть тестами формирование project paths и symlink tag names.
- [ ] Покрыть тестами поиск calibration paths (`Master` перед `Source`, temperature rounding, exposure exact/prefix match).
- [ ] Покрыть тестами парсинг имен lights и WBPP masters.
- [ ] Покрыть тестами edge cases: пустые folders, короткие flat folder names, unknown filter, missing gain/temp/exp.
- [ ] Добавить проверку синтаксиса всех PowerShell-файлов.
- [ ] Запускать PSScriptAnalyzer, если он доступен.
- [ ] После каждого изменения показывать `git diff`.
## Recent fixes

- [x] Match `+`-separated catalog compositions by the complete object list in import selection and TSV catalog/name resolution, so `M 8` and `M 8 + M 20` are not ambiguous.
- [x] `CreateProject.ps1`: deduplicate flat links and project metadata by the case-insensitive canonical physical source path; remove light session/exposure from flat tags and add optional stable `FLATSET` identity for compatibility collisions.
- [x] `Get-ImportReport.ps1`: emit full Google Sheets TSV columns, including characteristic `Exposure`, with `=` formulas; resolve catalog/name pairs from the sibling ASIAir library, warn and fall back on missing or ambiguous matches, and return one clipboard-safe multiline string.
- [x] `CreateProject.ps1`: normalize fractional temperatures independently of the Windows decimal separator so values such as `-9.8C` still select the `-10C` dark/bias folders.
- [x] `CreateProject.ps1`: convert millisecond flat exposures to seconds before matching flat-darks (`490ms` -> `0.49s`).
- [x] `CreateProject.ps1`: show the number of supported light frames for every session in the final project tree.
- [x] `CreateProject.ps1`: map unfiltered OSC lights to `Filter_RGB_Target_RGB`; reserve `Filter_L_Target_RGB` for actual OSC luminance filters such as IRC and Trib.
- [x] Use format-neutral scan messages now that imports support FITS, XISF, RAW, and bitmap images.
- [x] Ignore ASIAir `*_thn.jpg` thumbnail files in all shared image scans.
- [x] Share PixInsight image extensions and case-insensitive singular/plural frame-folder conventions across all import/report/project/export scripts.
- [x] `ImportSession.ps1`/`ImportAll.ps1`: sanitize user-entered destination path segments, including invisible spreadsheet characters such as tabs in object names.
- [x] `ImportSession.ps1`: normalize ASI setup camera names by removing `MM`/`MC` suffixes and map OSC `None` lights to the `RGB` filter folder.
- [x] `ImportAll.ps1`: add a batch light importer that scans `Import\<setup>\Light(s)\<object>`, caches season/setup/object choices, and applies shared import plans.
- [x] `ImportSession.ps1`: split single-session import into reusable find/plan/show/apply functions for batch reuse.
- [x] `ImportSession.ps1`: add Copy/Symlink import mode selection with Copy as the interactive default.
- [x] `Get-ImportReport.ps1`: print the overall integration summary after the per-object integration table.
- [x] `ImportCalibration.ps1`: group flat and other date-based calibration paths by the noon-to-noon night start date.
- [x] `ImportCalibration.ps1`: show counts and example relative paths before prompting for missing flat filter metadata.
- [x] `ImportCalibration.ps1`: accept partially or completely missing metadata collections when choosing interactive defaults.
- [x] Recover exposure and missing filter metadata for lights captured with Dark/Bias/Flat prefixes without changing FITS filenames or incremental-import identity.
- [x] `ImportCalibration.ps1`: import FIT/FITS/ARW bias, dark, and flat frames into canonical camera `Source` calibration trees with interactive RAW metadata, repeat-import detection, unusual-addition warnings, and `-WhatIf` support.
- [x] `ImportSession.ps1`: support camera RAW light files such as `.ARW`, using the source folder for object/setup and file timestamps for night dates.
- [x] `Get-ImportReport.ps1`: add an aligned per-object/filter integration summary in H:MM format.
- [x] `Get-ImportReport.ps1`: express every exposure group as `(night counts)*exposure` without dividing by the characteristic exposure.
- [x] `Get-ImportReport.ps1`: group repeated alternate-exposure multipliers across nights to shorten report expressions.
- [x] `ImportSession.ps1`: require exact catalog identifiers for fuzzy object matching so `M 16` does not match `M 17`.
- [x] `ImportSession.ps1`: use the source object folder name instead of unreliable ASIAir FITS object names such as `FOV` or stale autorun targets.
- [x] `Get-ImportReport.ps1`: use RGB/L/H/O/S order, report discovered Import folders and total integration time, and colorize interactive output.
- [x] `CreateProject.ps1`: reduce WBPP calibration tag noise and deduplicate repeated dark/bias/flat-dark source links.
- [x] `CreateProject.ps1`: deduplicate repeated flat source links by actual path and remove light-session-specific flat tags.
- [x] `CreateProject.ps1`: tested `DUMMY` calibration tags and reverted them because WBPP needs matching keyword values.
- [x] `CreateProject.ps1`: restore full WBPP calibration tags and report repeated calibration sources without changing project links.
- [x] `CreateProject.ps1`: suppress repeated calibration warnings for Master folders because WBPP can reuse the same master file.
- [x] `CreateProject.ps1`: require exact normalized exposure matches so 60s lights cannot select 600sec dark folders.
- [x] `ImportSession.ps1`: allow the first source prompt to accept an object name and fuzzy-search matching default Import sessions.
- [x] `CreateProject.ps1`: allow the first lights prompt to accept an object name and fuzzy-search matching ASIAir archive projects.
- [x] `ExportMasters.ps1`: allow the first metadata prompt to accept an object name and fuzzy-search matching projects under `AstroPhoto\Processing`.
- [x] `ExportMasters.ps1`: show compact camera/type/setup trees, resolve differing logical duplicates as conflicts, and execute only new copies after one plan confirmation without overwriting existing masters.
- [x] `CreateProject.ps1`: warn when an ASIAir light session folder contains mixed exposures without changing project links.
- [x] `ImportSession.ps1`/`ImportAll.ps1`: split imported light folders by exposure suffix when one filter/night contains mixed exposures, for example `26.07.10-180s`.
