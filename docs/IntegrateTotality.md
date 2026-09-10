# Съёмка и обработка полной фазы солнечного затмения: SER → HDR

---

## Съёмка полной фазы в SharpCap

### 1. Конфигурация камеры

Для ASI2600MC Duo + SQA55:

```text
Colour Space  = RAW16
ROI           = 4096×3072
Bin           = 1
Gain          = 0
Cooler        = OFF
Debayer       = OFF
White Balance = 50 / 50
```

Важно: для RAW-съёмки ZWO лучше заранее выставить:

```text
WB_R = 50
WB_B = 50
```

и проверить это в `CameraSettings.txt`.

---

### 2. Экспозиционная лестница

Базовый HDR-цикл:

```text
1 ms
10 ms
100 ms
```

Назначение:

```text
1 ms
→ Baily beads
→ хромосфера
→ протуберанцы
→ самый внутренний участок короны

10 ms
→ внутренняя и средняя корона
→ основной универсальный exposure

100 ms
→ средняя и внешняя корона
```

На 100 ms хромосфера и протуберанцы могут насыщаться — это нормально. В итоговом HDR эта область берётся из 1/10 ms.

Перед затмением обязательно проверить экспозиции на реальном Солнце. При другой высоте Солнца универсальные `1/10/100 ms` могут оказаться слишком длинными или короткими.

---

### 3. Sequence Editor

В SharpCap использовать `Sequence Editor`.

Перед циклом:

```text
Put Camera in Live Mode
Set Resolution      = 4096×3072
Set Colour Space    = RAW16
Set Binning         = 1
Set Gain            = 0
Set Output Format   = SER
```

Далее повторяющийся цикл:

```text
Repeat:

    Set Exposure = 0.001 s
    Capture 20 live view frames

    Set Exposure = 0.010 s
    Capture 25 live view frames

    Set Exposure = 0.100 s
    Capture 30 live view frames
```

Числа `20 / 25 / 30` не принципиальны. Практический диапазон:

```text
1 ms   → 20–30 кадров
10 ms  → 20–30 кадров
100 ms → 20–30 кадров
```

Для 100 ms больше кадров особенно полезно из-за слабой внешней короны.

---

### 4. Почему блоками

Каждая ступень записывается отдельным SER.

Получается последовательность примерно такого вида:

```text
Block 01:
    1ms  × 20
    10ms × 25
    100ms × 30

Block 02:
    1ms  × 20
    10ms × 25
    100ms × 30

...
```

Преимущества:

* легко отдельно интегрировать каждый временной блок;
* можно собирать HDR для каждого момента;
* можно сделать анимацию;
* потеря одного SER не уничтожает всю totality;
* не получается один гигантский файл.

---

### 5. Переходные кадры

При смене exposure первые 1–2 кадра нового SER могут быть неправильными:

* один кадр может фактически иметь предыдущую экспозицию;
* следующий может иметь промежуточный уровень;
* гистограмма отличается от основной серии.

Поэтому после съёмки первые кадры каждого SER надо проверить.

Неоднозначные переходные кадры лучше исключать и записывать их индексы в:

```json
"SkipIndices": [...]
```

Не пытаться использовать промежуточные кадры в HDR.

Если SharpCap позволяет вставить задержку или пропуск нескольких live-view frames после `Set Exposure`, это желательно сделать.

---

### 6. Когда запускать sequence

Последовательность лучше запустить примерно за:

```text
5–10 секунд до C2
```

пока ND5 ещё установлен.

Далее:

1. sequence уже работает;
2. в C2 снимается ND5;
3. всю totality ничего в SharpCap не трогать;
4. сразу после C3 вернуть ND5;
5. остановить sequence.

Это минимизирует работу с ноутбуком во время полной фазы.

---

### 7. Режим ноутбука

Перед totality:

```text
Windows Power Mode = Balanced
```

Не использовать энергосберегающий режим: в тестах он давал dropped frames.

Проверить заранее:

* запись без dropped frames;
* SSD выдерживает sustained write;
* свободного места достаточно;
* ноутбук находится в тени;
* экран можно приглушить.

---

### 8. Формат записи

Использовать:

```text
SER
RAW16
```

Не:

```text
RAW8
RGB24
MONO8
```

SER сохраняет исходный Bayer RAW16 и хорошо подходит для высокой частоты кадров.

---

### 9. Контрольная карта перед C2

Проверить:

```text
[ ] RAW16
[ ] Gain 0
[ ] WB_R = 50
[ ] WB_B = 50
[ ] ROI 4096×3072
[ ] Bin1
[ ] SER
[ ] Cooler OFF
[ ] Balanced power mode
[ ] dropped frames = 0
[ ] фокус проверен
[ ] tracking работает
[ ] sequence готов
[ ] ND5 надёжно установлен до C2
```

---

## Обработка

### 1. Общая схема

Рабочий pipeline:

```text
SER
  ↓
FITS CFA
  ↓
Calibration
  ↓
DefectMap
  ↓
Debayer
  ↓
Registration
  ↓
IntegrateTotality.ps1
  ↓
Block masters 1/10/100 ms
  ↓
единая геометрия masters между экспозициями
  ↓
AlignTotality.ps1
  ↓
HDR\Aligned
  ↓
HDRTotality.ps1
  ↓
HDR\Composed
  ↓
NormalizeTotality.ps1
  ↓
HDR\Calibrated
  ↓
финальная обработка
```

Обрабатывать `1 ms`, `10 ms` и `100 ms` независимо до стадии HDR.

Последние четыре стадии автоматизированы скриптами из корня проекта AsiToPix:

```text
IntegrateTotality.ps1 → интеграция временных блоков одной экспозиции
AlignTotality.ps1     → ChannelMatch RGB + поворот/flip всех block masters
HDRTotality.ps1       → HDRComposition одинаковых номеров блоков
NormalizeTotality.ps1 → нормализация яркости и фиксированная коррекция цвета
```

---

### 2. Структура каталогов

Рабочий корень для автоматизированного pipeline:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\
```

Скрипты `AlignTotality.ps1`, `HDRTotality.ps1` и `NormalizeTotality.ps1` ищут
корни по алиасам `*:\AstroPhoto` и `*:\Astro`. Если найдено несколько корней,
скрипт предлагает выбрать нужный. Полные пути также можно передать параметрами.

Например, для 100 ms:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\
└── Total-100ms\
    ├── Frames\
    │   ├── raw\
    │   ├── calibrated\
    │   ├── corrected\
    │   ├── debayered\
    │   ├── registered\
    │   ├── integrated\
    │   └── manifest.json
    └── SourceSER\
```

Аналогично:

```text
Total-1ms\
Total-10ms\
Total-100ms\
```

Калибровочные данные удобно держать отдельно:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\Calibration\
├── Flats\
│   ├── raw\
│   └── calibrated\
├── FlatDarks\
│   └── raw\
├── masterFlatDark.xisf
└── masterFlat.xisf
```

После автоматизированной обработки добавляется дерево:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\HDR\
├── channel_match_offsets.json
├── Aligned\
│   └── Block_005_100ms_Totality_00123_00150_ca_r180.xisf
├── Composed\
│   └── Block-005-HDR.xisf
└── Calibrated\
    ├── normalization_metrics.csv
    ├── Block-005-HDR_norm.xisf
    └── Block-005-HDR_norm_cc.xisf
```

---

### 3. SER → FITS

Исходные SER сначала сохранить неизменными как архив.

В **Siril → Conversion**:

* Input: соответствующий `.ser`
* Output: отдельные FITS
* **Debayer OFF**
* не менять уровни, WB и гамму.

Получаем CFA-файлы:

```text
...\Frames\raw\
    totality_00001.fits
    totality_00002.fits
    ...
```

То же самое проделать с SER flats и flat-darks.

#### Переходные кадры

После переключения экспозиции SharpCap/ZWO выдавал первые 1–2 ненормальных кадра:

* иногда явно с предыдущей экспозицией;
* иногда с промежуточным уровнем.

Промежуточные кадры **не пытаться спасать**. Их индексы занести в `SkipIndices` в manifest.json.

Например:

```json
"SkipIndices": [1, 2, 31, 61, 62, 91, 92]
```

Это надёжнее, чем пытаться распределять сомнительные кадры между 1/10/100 ms.

---

### 4. Создание master flat

Для текущей съёмки flats и flat-darks сняты:

* Gain 0
* тем же ROI
* тем же bin
* тем же положением камеры
* **WB_R=..., WB_B=...**, как и totality lights
* без ND5, поскольку totality тоже снималась без ND5.

Важно: неправильный acquisition WB не исправлять до flat calibration. Flat должен иметь **тот же WB**, что и lights.

#### 4.1 Master flat-dark

Flat-darks интегрировать через `ImageIntegration`.

Это обычный master dark по типу данных, хотя функционально здесь он является **master flat-dark**.

Пример:

```text
50 flat-darks
    ↓ ImageIntegration
masterFlatDark.xisf
```

#### 4.2 Calibration flats

`Process → ImageCalibration`

Input:

```text
Calibration\Flats\raw\*.fits
```

Использовать:

```text
Master Dark = masterFlatDark.xisf
Bias        = OFF
Optimize    = OFF
```

Результат:

```text
Calibration\Flats\calibrated\
```

#### 4.3 Integration flats

`ImageIntegration`:

```text
Combination       = Average
Normalization     = Multiplicative
Pixel rejection   = Winsorized Sigma Clipping
```

Полученный master сохранить как:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\Calibration\masterFlat.xisf
```

---

### 5. Калибровка

Каждую экспозицию обрабатывать отдельно.

Например:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\Total-100ms\Frames\raw\
```

через:

```text
Process → ImageCalibration
```

Использовать:

```text
Master Flat = masterFlat.xisf
Master Dark = OFF
Master Bias = OFF
```

Обычные darks для `1/10/100 ms` не используются: dark current за эти экспозиции ничтожен.

Output:

```text
...\Frames\calibrated\
```

`Signal Evaluation`, `Noise Evaluation` — выключить, они не нужны, но сокращают время обработки значительно.

На этой стадии данные всё ещё **CFA, один канал**.

---

### 6. Коррекция горячих пикселей

Поскольку totality lights калибруются без обычных dark frames, в calibrated CFA остаются несколько стабильных hot pixels.

Их нужно исправлять **до Debayer**. Иначе один горячий CFA-пиксель после дебайеризации превращается в цветное пятно или крест из нескольких пикселей.

Рабочий pipeline:

```text
calibrated CFA
    ↓
CosmeticCorrection
    ↓
corrected CFA
    ↓
Debayer
```

#### 6.1. Почему не использовать Auto Detect

`CosmeticCorrection → Auto Detect` оказался недостаточно надёжным: один и тот же hot pixel определялся не на всех кадрах.

Поскольку стабильных дефектов всего несколько, лучше один раз определить их координаты и добавить вручную в постоянный `Defect List`.

Для поиска всех дефектов удобно предварительно создать диагностическую бинарную карту.

---

#### 6.2. Создание диагностической defect map

Карту строить по:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\Calibration\masterFlatDark.xisf
```

а не по солнечному light-кадру.

В `masterFlatDark` нет хромосферы, протуберанцев и других реальных ярких деталей, поэтому hot pixels можно выделить простым порогом.

В `PixelMath`:

```text
iif($T > threshold, 0, 1)
```

где:

```text
0 = найденный hot pixel
1 = нормальный пиксель
```

`threshold` подобрать по значениям `masterFlatDark`: он должен выделять известные горячие пиксели, но не обычный шум.

Создать новое изображение:

```text
Create new image = ON
Color space      = Gray
Sample format    = 32-bit floating point
```

Размер карты должен совпадать с кадрами:

```text
4096 × 3072
```

Сохранить, например:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\Calibration\DefectMap.xisf
```

Эта карта используется **только для поиска и проверки дефектов**. Пакетно применять `DefectMap` к light frames не требуется.

---

#### 6.3. Проверка количества найденных пикселей

Если `DefectMap.xisf` строго бинарная:

```text
нормальный пиксель = 1
дефектный пиксель  = 0
```

то точное количество дефектов можно получить через `Statistics`.

Для ROI:

```text
4096 × 3072
```

общее число пикселей:

```text
N = 4096 × 3072
  = 12 582 912
```

Если `mean` — среднее значение бинарной карты, то:

```text
Ndefects = N × (1 - mean)
```

Например, если результат равен:

```text
4
```

значит PixelMath нашёл ровно четыре чёрных пикселя.

Удобный альтернативный вариант — временно инвертировать карту:

```text
1 - $T
```

Тогда:

```text
дефект = 1
фон    = 0
```

и количество дефектов:

```text
Ndefects = N × mean
```

Для текущих данных таким способом обнаружено **4 стабильных hot pixels**.

---

#### 6.4. Поиск координат

Открыть `DefectMap.xisf`, сильно увеличить изображение и найти все чёрные точки.

Для каждого из четырёх пикселей записать его координаты.

Количество найденных вручную координат должно совпадать с результатом `Statistics`:

```text
Statistics: 4 дефекта
Defect List: 4 дефекта
```

Это позволяет убедиться, что ни один редкий hot pixel не пропущен.

---

#### 6.5. CosmeticCorrection

После определения координат использовать:

```text
Process → ImageCalibration → CosmeticCorrection
```

Настройки:

```text
Auto Detect      = OFF
Master Dark      = OFF
Defect List      = ON
CFA              = ON
```

В `Defect List` вручную добавить все четыре найденных пикселя.

Сохранить настроенный экземпляр `CosmeticCorrection` как process icon, чтобы один и тот же список использовался для всех экспозиций.

Input:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\Total-100ms\Frames\calibrated\
```

Output:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\Total-100ms\Frames\corrected\
```

Аналогично обработать:

```text
Total-1ms
Total-10ms
Total-100ms
```

Поскольку дефекты принадлежат матрице камеры и ROI одинаковый, используется один и тот же список координат.

---

#### 6.6. Проверка результата

Сначала прогнать несколько кадров.

Проверить:

* исчезли все hot pixels;
* вокруг исправленных точек нет цветных артефактов (после Debayer);
* остальные детали изображения не затронуты.

После этого запустить `CosmeticCorrection` на всём массиве.

---

### 7. Debayer

Критически важно проверить правильный CFA pattern. Обычно это — **RGGB**.

Хотя после SER → FITS в Siril в header могло записаться:

```text
BAYERPAT = GBRG
```

это значение оказалось неверным (баг в Siril?). Эксперимент показал:

```text
RGGB → розовые/красные протуберанцы   ← правильно
GBRG → зелёные протуберанцы
GRBG → зелёные протуберанцы
BGGR → синие протуберанцы
```

Поэтому:

```text
Process → Debayer

Bayer pattern = RGGB
Method        = Bilinear
```

`SuperPixel` тоже даёт правильный цвет и устраняет CFA-артефакты, но уменьшает разрешение:

```text
4096×3072 → 2048×1536
```

Поэтому основной workflow:

```text
RGGB + Bilinear
```

Output:

```text
...\Frames\debayered\
```

`Signal Evaluation`, `Noise Evaluation` — выключить: PSF-анализ звёзд для солнечных кадров бессмысленен.

---

### 8. Registration

Регистрировать каждую экспозицию независимо:

```text
Total-1ms
Total-10ms
Total-100ms
```

Рабочая схема:

```text
Script → Utilities → FFTRegistration
```

Reference должен быть обычным чистым кадром полной фазы:

* не переходный;
* не Baily beads;
* желательно около середины totality;
* индекс фиксируется в `manifest.json`.

Например:

```json
"RegistrationIndex": 160
```

Все кадры данной экспозиции регистрируются к одному reference.

Output:

```text
...\Frames\registered\
```

#### Единая геометрия для HDR

Хотя экспозиции обрабатываются раздельно, к моменту интеграции их reference
frames должны задавать одну геометрию. Предпочтительный reference — хороший
центральный кадр 10 ms. Если один reference нельзя надёжно применить ко всем
экспозициям, сначала совместить отдельные reference frames между собой, а
затем регистрировать каждую серию в соответствующую, но уже согласованную
геометрию.

Перед интеграцией проверить, что зарегистрированные 1/10/100 ms имеют один
размер и одинаковое положение солнечного диска. `AlignTotality.ps1` на более
поздней стадии исправляет RGB ChannelMatch и ориентацию, но не заменяет эту
межэкспозиционную регистрацию.

Перед дальнейшей обработкой обязательно проверить через Blink:

* корона должна стоять;
* протуберанцы должны стоять;
* Луна может медленно смещаться относительно солнечной короны — это физически правильно;
* деревья/облака должны двигаться относительно солнечной системы координат.

Если FFTRegistration на 100 ms работает плохо, не интегрировать слепо: проверить блоки отдельно и при необходимости регистрировать block masters через `DynamicAlignment`.

Итоговые registered-файлы могут иметь вид:

```text
*_c_d_r.xisf
*_c_cc_d_r.xisf
```

`IntegrateTotality.ps1` определяет кадр по части имени `*_<экспозиция>_<индекс>` и игнорирует весь суффикс после индекса.

---

### 9. C2/C3 Baily beads

Baily beads не выбрасывать.

Их preprocessing такой же:

```text
Calibration
→ DefectMap
→ Debayer RGGB/Bilinear
→ Registration
```

Но они логически отделяются от обычной totality.

Например, в данных:

```text
1 ms   → есть C2 и C3
10 ms  → есть C2
100 ms → есть C3
```

Для регистрации сначала можно попробовать тот же solar reference.

Если FFTRegistration ломается из-за яркой фотосферы:

* C2 регистрировать относительно ближайшего чистого кадра сразу после C2;
* C3 — относительно ближайшего чистого кадра перед C3;
* при необходимости использовать `DynamicAlignment`.

Beads не включать в обычные totality HDR-блоки.

Они пригодятся:

* для отдельных изображений контактов;
* как начало/конец анимации.

---

### 10. manifest.json

Для каждой экспозиции свой manifest.

Пример для 100 ms:

```json
{
    "BlockLength": 30,
    "BlockStartIndex": 1,
    "C2Index": null,
    "C3Index": 287,
    "RegistrationIndex": 160,
    "SkipIndices": [
        1, 2,
        31,
        61, 62,
        91, 92,
        121, 122,
        151, 152,
        181, 182,
        211,
        241,
        271, 272
    ]
}
```

Путь:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\Total-100ms\Frames\manifest.json
```

#### Поля

```text
BlockLength
```

Число кадров, которое SharpCap снимал за один блок данной экспозиции.

Например, реальные серии могут отличаться:

```text
1 ms   → 20
10 ms  → 25
100 ms → 30
```

```text
BlockStartIndex
```

Первый индекс массива согласно схеме скрипта.

```text
C2Index / C3Index
```

Индексы контактов. Если соответствующая серия отсутствует:

```json
null
```

```text
RegistrationIndex
```

Индекс выбранного reference frame.

```text
SkipIndices
```

Переходные/невалидные кадры, которые нельзя включать в интеграцию.

---

### 11. Автоматическая block integration

После того как:

```text
...\Frames\registered\
или
...\Frames\debayered\
```

содержит единый массив обработанных файлов, запускать:

```powershell
$totalityRoot = "Z:\AstroPhoto\SharpCap\Totality-2026"
.\IntegrateTotality.ps1 -ProcessingRoot $totalityRoot
```

У `IntegrateTotality.ps1` исторический путь по умолчанию —
`C:\AstroPhoto\Processing\Totality`. Поэтому для дерева
`*:\AstroPhoto\SharpCap\Totality-2026` параметр `-ProcessingRoot` нужно
передавать явно.

Скрипт запросит исходную папку:

```text
Input folder:
  [1] Registered (default)
  [2] Debayered
```

Для автоматического запуска можно передать `-InputStage Registered` или
`-InputStage Debayered`. Вариант `Debayered` полезен для диагностики, если
регистрация сделала последовательность менее стабильной, но он не создаёт
единую геометрию между экспозициями. Перед `HDRTotality.ps1` такие masters всё
равно придётся геометрически совместить. Для основного автоматизированного
pipeline использовать `Registered`.

Если PixInsight уже запущен, скрипт также предложит режим выполнения:

```text
PixInsight execution mode:
  [1] Reuse a running instance and leave it open (default)
  [2] Start a dedicated instance and close it when finished
```

`Reuse` передаёт первому доступному экземпляру через IPC `--execute` временную копию PJSR-скрипта с едиными окончаниями строк CRLF. Это предотвращает ошибки препроцессора PixInsight на файлах со смешанными CRLF/LF. Скрипт находит лежащий рядом временный `IntegrationPlan.json` через `#__FILE__`; зависящий от пути JavaScript-код не внедряется. Ключи `-n` и `--force-exit` в этом режиме не используются, поэтому существующий PixInsight и открытые в нём изображения после интеграции остаются открытыми. Пока работает ImageIntegration, этот экземпляр PixInsight занят; если нужно параллельно продолжать ручную работу, выберите `Dedicated`. Если запущено несколько экземпляров, PixInsight выбирает первый доступный IPC slot.

В Windows вспомогательный GUI-процесс IPC может завершиться раньше, чем станет доступен его `ExitCode`. Пустой код завершения не считается ошибкой: результат определяется только по JSON status-файлу, который пишет PJSR. После отправки задания проверьте окно PixInsight — для локального неподписанного скрипта он может запросить подтверждение. Если status-файл не появился за одну минуту, скрипт сохраняет диагностическую папку и предлагает проверить Process Console PixInsight.

Для неинтерактивного запуска:

```powershell
.\IntegrateTotality.ps1 `
    -ProcessingRoot "Z:\AstroPhoto\SharpCap\Totality-2026" `
    -Exposure Total-100ms `
    -InputStage Debayered `
    -PixInsightMode Reuse
```

Скрипт читает `manifest.json`, режет массив на исходные временные блоки и вызывает PixInsight `ImageIntegration`.

Output:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\Total-100ms\Frames\integrated\
```

Пример:

```text
Block_003_100ms_Totality_00063_00090.xisf
```

Скрипт обрабатывает одну экспозиционную папку за запуск. Поэтому его нужно
повторить для всех ступеней:

```powershell
foreach ($exposure in @("Total-1ms", "Total-10ms", "Total-100ms")) {
    .\IntegrateTotality.ps1 `
        -ProcessingRoot "Z:\AstroPhoto\SharpCap\Totality-2026" `
        -Exposure $exposure `
        -InputStage Registered `
        -PixInsightMode Reuse
}
```

Перед пакетным циклом PixInsight должен быть уже запущен. Если выбран
`Dedicated`, каждый запуск создаёт отдельный автоматизированный экземпляр и
закрывает его после интеграции.

---

### 12. Рабочие настройки ImageIntegration

Для totality blocks:

```text
Combination                   = Average
Normalization                 = No normalization
Weights                       = Don't care

Pixel Rejection               = Winsorized Sigma Clipping
Pixel Rejection Normalization = Scale + zero offset

Reject Low                    = true
Reject High                   = false

Sigma Low                     = 3.5

Large-scale low rejection     = ON
Layers                        = 2
Growth                        = 3

Large-scale high rejection    = OFF

Generate rejection maps       = ON
```

Эти параметры дали рабочий результат:

* настоящая корона в `rejection_low` почти чёрная;
* дерево сильно rejected, до ~0.8–0.9;
* фон показывает цветной статистический шум rejection;
* CFA/debayer-шахматка в rejection map сама по себе не критична, если её нет в master.

Для 1 ms, если дерево ещё не мешает, large-scale low можно отключить.

Для 100 ms large-scale low особенно полезен.

---

### 13. Проверка rejection map

`rejection_low` показывает долю кадров, признанных слишком тёмными.

Например:

```text
0.00 → не rejected ни один кадр
0.10 → rejected примерно 10 %
0.50 → примерно половина
0.90 → почти все
```

Для дерева высокий rejection — хорошо.

Для стабильной короны желательно:

```text
≈ 0
```

Если настоящая структура короны начинает массово появляться в rejection map, rejection слишком агрессивный.

---

### 14. Почему интеграция идёт по временным блокам

Не надо складывать всю минуту totality в один master.

Луна движется относительно Солнца, поэтому при интеграции всей серии:

* солнечная корона остаётся на месте;
* лунный лимб постепенно смещается;
* край Луны размывается.

Внутри одного блока `20–30` кадров прошло всего несколько секунд, поэтому лимб остаётся достаточно резким.

Это даёт серию временных masters:

```text
Block-001
Block-002
Block-003
...
Block-010
```

---

### 15. Выравнивание каналов и ориентации: AlignTotality.ps1

После интеграции каждой экспозиции получится примерно:

```text
Total-1ms\Frames\integrated\
    Block_003_1ms_Totality_00043_00060.xisf

Total-10ms\Frames\integrated\
    Block_003_10ms_Totality_00052_00075.xisf

Total-100ms\Frames\integrated\
    Block_003_100ms_Totality_00063_00090.xisf
```

Перед HDR все экспозиции одного блока должны иметь:

```text
одинаковый размер изображения
одинаковую ориентацию
одинаковое положение солнечного диска и короны
```

Важно: `AlignTotality.ps1` выполняет `ChannelMatch` между R/G/B внутри каждого
изображения и опциональный `FastRotation`. Он **не регистрирует 1/10/100 ms
masters друг относительно друга**. Геометрию между экспозициями надо привести
к одному reference ещё на стадии Registration. Если masters были получены с
разной геометрией, сначала исправить регистрацию и повторить интеграцию.

Запуск из корня AsiToPix:

```powershell
.\AlignTotality.ps1
```

По умолчанию скрипт:

1. находит `*:\AstroPhoto` и `*:\Astro`;
2. выбирает `SharpCap\Totality-2026` под выбранным корнем;
3. ищет все папки `Total-<экспозиция>\Frames\integrated`;
4. предлагает поворот `0/90/180/270` градусов;
5. предлагает `None/Horizontal/Vertical` mirror;
6. сохраняет результат в `HDR\Aligned`.

Если нужен явный корень:

```powershell
.\AlignTotality.ps1 `
    -TotalityPath "Z:\AstroPhoto\SharpCap\Totality-2026"
```

#### Получение offsets из ChannelMatch

При первом запуске или при выборе новых offsets скрипт запускает PixInsight,
если он ещё не открыт, и показывает рекомендуемый reference: центральный блок
со средней найденной экспозицией. На нём должны быть хорошо видны лунный лимб
и хромосфера без сильного пересвета.

Во время этой стадии должен работать ровно один экземпляр PixInsight. Если
открыто несколько экземпляров, лишние надо закрыть до запуска плана.

В PixInsight:

1. открыть предложенный XISF;
2. открыть `ChannelMatch`;
3. подобрать X/Y offsets каналов R/G/B;
4. не менять linear correction factors;
5. перетащить синий треугольник **New Instance** из ChannelMatch на свободное
   место PixInsight workspace;
6. вернуться в PowerShell и нажать Enter.

Одного открытого интерфейса ChannelMatch недостаточно: PJSR читает параметры
из созданного process icon. Скрипт запоминает список старых ChannelMatch icons
и принимает ровно один новый icon, поэтому его надо создать на свободном месте,
не заменяя старый.

Из icon сохраняются `enabled`, `dx` и `dy` всех трёх каналов. Linear correction
factor принудительно устанавливается в `1.0`, поэтому эта стадия не меняет
яркость и цветовой баланс.

Offsets записываются в:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\HDR\channel_match_offsets.json
```

При следующем запуске скрипт спрашивает, использовать ли этот JSON. Режим можно
задать явно:

```powershell
# Использовать сохранённые offsets без настройки ChannelMatch
.\AlignTotality.ps1 -ChannelMatchMode Saved

# Принудительно получить новые offsets и перезаписать JSON
.\AlignTotality.ps1 -ChannelMatchMode Capture
```

#### Поворот и flip

После ChannelMatch к каждому файлу сразу применяется выбранный `FastRotation`:

```powershell
.\AlignTotality.ps1 `
    -Rotation 180 `
    -Flip None
```

Допустимые значения:

```text
Rotation = 0, 90, 180, 270
Flip     = None, Horizontal, Vertical
```

`90` означает 90° по часовой стрелке, `270` — 90° против часовой. Если выбраны
и поворот, и mirror, сначала выполняется поворот, затем mirror.

Операции отражаются в имени:

```text
_ca       → ChannelMatch применён
_r180     → поворот 180°
_fh       → horizontal mirror
_fv       → vertical mirror
```

Примеры:

```text
Block_005_100ms_Totality_00123_00150_ca.xisf
Block_005_100ms_Totality_00123_00150_ca_r180.xisf
Block_005_100ms_Totality_00123_00150_ca_r90_fh.xisf
```

Выходные XISF на этой стадии не перезаписываются. Существующие файлы помечаются
в плане как `Skip: output exists`. Для повторной обработки с другими offsets
или transforms старые файлы из `HDR\Aligned` надо предварительно перенести или
удалить вручную после проверки пути.

План без запуска PixInsight и без создания файлов:

```powershell
.\AlignTotality.ps1 `
    -Rotation 180 `
    -Flip None `
    -ChannelMatchMode Capture `
    -WhatIf
```

---

### 16. HDR-композиция блоков: HDRTotality.ps1

Input:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\HDR\Aligned\
```

Скрипт группирует файлы по номеру блока и экспозиции. Для трёх найденных
экспозиций предлагает все непрерывные HDR-лестницы, для которых есть хотя бы
один полный блок:

```text
1 ms + 10 ms + 100 ms
1 ms + 10 ms
10 ms + 100 ms
```

Обычный интерактивный запуск:

```powershell
.\HDRTotality.ps1
```

Явный корень без выбора алиаса:

```powershell
.\HDRTotality.ps1 -AstroPhotoRoot "Z:\AstroPhoto"
```

Лестницу и режим PixInsight можно задать явно:

```powershell
.\HDRTotality.ps1 `
    -Exposure @("1ms", "10ms", "100ms") `
    -PixInsightMode Reuse
```

Доступны два режима:

```text
Reuse     → использовать запущенный PixInsight и оставить его открытым
Dedicated → запустить отдельный PixInsight и закрыть после обработки
```

В `HDRComposition.images` файлы передаются от длинной экспозиции к короткой:

```text
100 ms
10 ms
1 ms
```

Если в используемом экземпляре PixInsight есть process icon с точным именем
`HDRComposition`, скрипт берёт его настройки и заменяет только список images.
Это позволяет заранее настроить, например, `Reject black pixels`. Если icon не
найден, используется новый `HDRComposition` с настройками PixInsight по
умолчанию. Для использования собственного icon выбирать `Reuse` в том
экземпляре, где этот icon создан.

Output:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\HDR\Composed\
    Block-001-HDR.xisf
    Block-002-HDR.xisf
    ...
```

Блоки без полного набора выбранных экспозиций пропускаются с перечислением
недостающих ступеней. Существующие `Block-<номер>-HDR.xisf` не
перезаписываются.

Проверка плана:

```powershell
.\HDRTotality.ps1 `
    -Exposure @("10ms", "100ms") `
    -PixInsightMode Dedicated `
    -WhatIf
```

Роли экспозиций:

```text
1 ms
→ лимб, Baily beads, хромосфера, протуберанцы

10 ms
→ внутренняя и средняя корона

100 ms
→ средняя и внешняя корона
```

Пересвет протуберанцев на 100 ms не имеет значения: HDR должен заменить эту
область короткими экспозициями.

---

### 17. Нормализация яркости и цвета: NormalizeTotality.ps1

Input:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\HDR\Composed\
    Block-001-HDR.xisf
    ...
```

Запуск:

```powershell
.\NormalizeTotality.ps1
```

Явный корень без выбора алиаса:

```powershell
.\NormalizeTotality.ps1 -AstroPhotoRoot "Z:\AstroPhoto"
```

Для этой интерактивной операции нужен ровно один PixInsight. Если он не
запущен, скрипт запускает его. Если открыто несколько экземпляров, лишние надо
закрыть перед запуском.

После подтверждения плана скрипт просит подготовить reference ROI:

1. открыть один из `Block-*-HDR.xisf` непосредственно из `HDR\Composed`;
2. создать Preview в области, которая присутствует на всех блоках;
3. выбрать стабильную несатурированную область без края Луны, дерева, облаков
   и других движущихся деталей;
4. активировать Preview и вернуться в PowerShell;
5. нажать Enter.

Reference обязан быть одним из входных `Block-*-HDR.xisf`. Если в окне только
один Preview, его можно не активировать отдельно; при нескольких Preview надо
активировать нужный.

Скрипт читает координаты ROI и измеряет медианы `Rref/Gref/Bref`. Для каждого
кадра в той же ROI вычисляются:

```text
kR = Rref / R
kG = Gref / G
kB = Bref / B
k  = median(kR, kG, kB)
```

Сначала всё изображение умножается на scalar `k`. Затем ко всем уже
нормализованным кадрам применяются одинаковые коэффициенты reference:

```text
cR = Gref / Rref
cG = 1
cB = Gref / Bref
```

Output:

```text
Z:\AstroPhoto\SharpCap\Totality-2026\HDR\Calibrated\
    normalization_metrics.csv
    Block-001-HDR_norm.xisf
    Block-001-HDR_norm_cc.xisf
    ...
```

Назначение файлов:

```text
*_norm.xisf    → только scalar normalization
*_norm_cc.xisf → scalar normalization + фиксированные RGB coefficients
```

CSV содержит input/output paths, reference flag, Preview ID, координаты ROI,
медианы RGB, `kR/kG/kB`, итоговый `k` и `cR/cG/cB` для каждого блока.

В отличие от предыдущих стадий, повторный запуск `NormalizeTotality.ps1`
перезаписывает `normalization_metrics.csv`, `_norm.xisf` и `_norm_cc.xisf`.

Проверка плана без запуска PixInsight:

```powershell
.\NormalizeTotality.ps1 -WhatIf
```

---

### 18. Главный статичный результат

Не обязательно смешивать всю totality.

Для главного still лучше выбрать **один из центральных блоков** или 2–3 соседних:

* beads уже закончились;
* дерево ещё не дошло до диска;
* лунный лимб чистый;
* геометрия между 1/10/100 ms почти одинаковая.

Если один центральный HDR выглядит хорошо, он может быть главным результатом вообще без сложного time-composite.

---

### 19. Time-composite всей totality

Отдельный экспериментальный результат.

Все блоки можно зарегистрировать в солнечной системе координат и комбинировать так, чтобы:

* хромосфера собиралась с разных сторон;
* протуберанцы сохранялись;
* moving Moon считалась временной маской;
* деревья/облака по возможности удалялись rejection/масками.

Это уже не снимок одного мгновения, а временной композит полной фазы. Его делать только после получения нормального single-block HDR.

---

### 20. Анимация

Block HDR отлично подходят для анимации:

```text
HDR_Block01
HDR_Block02
HDR_Block03
...
HDR_Block10
```

Перед анимацией все HDR зарегистрировать к одному солнечному reference.

Тогда:

* корона стоит;
* Луна естественно движется;
* протуберанцы появляются/скрываются;
* можно добавить C2/C3 beads как начальные и конечные кадры.

Это один из лучших способов использовать все 10–11 исходных временных серий.

---

### 21. Цвет и атмосферная дисперсия

Не пытаться исправлять цвет до calibration/debayer. Например, если съёмка была сделана с:

```text
WB_R = 23
WB_B = 99
```

то линейный цвет изначально сильно искажён.

Кроме того, если Солнце было очень низко, то присутствует сильная атмосферная дисперсия.

В автоматизированном pipeline атмосферная дисперсия исправляется до HDR:

1. на хорошем integrated master проверить R/G/B отдельно;
2. подобрать offsets в `ChannelMatch`, обычно относительно G;
3. применить одни offsets ко всем masters через `AlignTotality.ps1`;
4. проверить несколько файлов из `HDR\Aligned` до запуска HDR;
5. после `HDRTotality.ps1` выровнять яркость и базовый RGB-баланс через
   `NormalizeTotality.ps1`;
6. финальный физический/визуальный цвет и stretch делать уже по
   `*_norm_cc.xisf`.

Коэффициенты `cR/cG/cB` из `NormalizeTotality.ps1` обеспечивают одинаковую
фиксированную коррекцию всей серии, но не являются полноценной фотометрической
калибровкой цвета.

Цвет сумеречного неба не использовать как физический эталон.

---

### 22. Что не переделывать без причины

После того как получены нормальные CFA calibrated frames:

```text
Calibration
```

повторять не надо.

Если найден неправильный Bayer pattern, переделывается только:

```text
Debayer
→ Registration
→ Integration
→ AlignTotality
→ HDRTotality
→ NormalizeTotality
```

Если меняются настройки ImageIntegration, переделывается только:

```text
Integration
→ AlignTotality
→ HDRTotality
→ NormalizeTotality
```

Если меняются ChannelMatch offsets, поворот или flip:

```text
AlignTotality
→ HDRTotality
→ NormalizeTotality
```

Если меняется HDR-лестница или настройки process icon `HDRComposition`:

```text
HDRTotality
→ NormalizeTotality
```

Если меняется только reference Preview для нормализации:

```text
NormalizeTotality
```

Исходные SER и calibrated CFA всегда сохранять как неизменяемую основу.

---

### 23. Текущие особенности именно этого набора

Для этой съёмки установлено:

```text
Camera        = ASI2600MC Duo
Telescope     = SQA55
ROI           = 4096×3072
Gain          = 0
Colour Space  = RAW16
Exposure HDR  = 1 / 10 / 100 ms

Correct CFA   = RGGB
Debayer       = Bilinear

Acquisition WB:
R = 23
B = 99
```

Проблемы исходных данных:

```text
- первые 1–2 кадра некоторых SER переходные;
- часть кадров C2/C3 содержит Baily beads;
- дерево постепенно наползает на корону;
- 100 ms насыщает хромосферу/протуберанцы;
- несколько стабильных hot pixels;
- сильная атмосферная дисперсия из-за высоты Солнца;
- metadata BAYERPAT=GBRG неверна.
```

Поэтому наиболее безопасная стратегия:

```text
не пытаться сделать один гигантский master всей totality
→ интегрировать исходными временными блоками
→ согласовать геометрию экспозиций
→ применить ChannelMatch и ориентацию через AlignTotality.ps1
→ собрать HDR каждого блока через HDRTotality.ps1
→ нормализовать серию через NormalizeTotality.ps1
→ выбрать центральный *_norm_cc.xisf как основной still
→ остальные нормализованные блоки использовать для time-composite и анимации
```
