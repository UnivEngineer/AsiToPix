# TODO

Заметки по результатам чтения `CreateProject.ps1`, `ExportMasters.ps1` и `CombineSeasons.ps1`.

## Критичные риски безопасности данных

- [ ] `CreateProject.ps1`: перед `Remove-Item $sourcePath -Recurse -Force` проверять, что удаляемая папка действительно создана проектом и находится внутри ожидаемого project root.
- [ ] `CombineSeasons.ps1`: перед `Remove-Item $combineRoot -Recurse -Force` валидировать resolved path и использовать marker/metadata созданного проектом каталога.
- [ ] `CreateProject.ps1`: перед созданием или заменой target path проверять, что существующий объект является symbolic link, а не обычным файлом или папкой.
- [ ] `ExportMasters.ps1`: убрать безусловный `Copy-Item -Force` или явно оформить безопасную политику overwrite.
- [ ] Добавить `-WhatIf` через `[CmdletBinding(SupportsShouldProcess)]` для операций создания, удаления, копирования и замены файлов/каталогов.
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
- [ ] `ExportMasters.ps1`: валидировать структуру и обязательные поля `project_meta.json` (`PixPath`, `Scope`, `Cameras[].Name`).
- [ ] `ExportMasters.ps1`: не классифицировать любой `masterDark` с экспозицией `< 10s` как flat-dark без дополнительного признака.
- [ ] `ExportMasters.ps1`: убрать hardcoded resolution `6248x4176` или получать его из metadata/имени файла.
- [ ] `ExportMasters.ps1`: сделать matching камеры точнее, чем `*$camFull*`, чтобы избежать ложных совпадений.
- [ ] `ExportMasters.ps1`: обрабатывать отсутствие подходящих `.xisf` мастеров с понятным сообщением.
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
- [ ] Заменить ручную конкатенацию путей на `Join-Path`, `Resolve-Path` и `[System.IO.Path]` там, где это еще не сделано.
- [ ] Сделать сообщения об ошибках actionable и включать affected path.
- [ ] Привести PowerShell function names к approved verbs.
- [ ] Убрать объявление `FlatFolderMatchesFilter` и `ParseDate` из внутреннего цикла сессий.
- [ ] Упростить и покрыть тестами алгоритм выбора flats по дате/углу/Origin.

## Тесты и проверка

- [ ] Добавить Pester tests под `tests/`.
- [ ] Покрыть тестами filter mapping для OSC и Mono.
- [ ] Покрыть тестами формирование project paths и symlink tag names.
- [ ] Покрыть тестами поиск calibration paths (`Master` перед `Source`, temperature rounding, exposure exact/prefix match).
- [ ] Покрыть тестами парсинг имен lights и WBPP masters.
- [ ] Покрыть тестами edge cases: пустые folders, короткие flat folder names, unknown filter, missing gain/temp/exp.
- [ ] Добавить проверку синтаксиса всех PowerShell-файлов.
- [ ] Запускать PSScriptAnalyzer, если он доступен.
- [ ] После каждого изменения показывать `git diff`.
