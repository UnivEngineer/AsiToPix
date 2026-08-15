# Схёмка и обработка полной фазы солнечного затмения: SER → HDR

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
Alignment masters между экспозициями
  ↓
HDRComposition
  ↓
финальная обработка
```

Обрабатывать `1 ms`, `10 ms` и `100 ms` независимо до стадии HDR.

---

### 2. Структура каталогов

Например, для 100 ms:

```text
C:\AstroPhoto\Processing\Totality\
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
C:\AstroPhoto\Processing\Totality\Calibration\
├── Flats\
│   ├── raw\
│   └── calibrated\
├── FlatDarks\
│   └── raw\
├── masterFlatDark.xisf
└── masterFlat.xisf
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
C:\AstroPhoto\Processing\Totality\Calibration\masterFlat.xisf
```

---

### 5. Калибровка

Каждую экспозицию обрабатывать отдельно.

Например:

```text
C:\AstroPhoto\Processing\Totality\Total-100ms\Frames\raw\
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
C:\AstroPhoto\Processing\Totality\Calibration\masterFlatDark.xisf
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
C:\AstroPhoto\Processing\Totality\Calibration\DefectMap.xisf
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
C:\AstroPhoto\Processing\Totality\Total-100ms\Frames\calibrated\
```

Output:

```text
C:\AstroPhoto\Processing\Totality\Total-100ms\Frames\corrected\
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

Если включён `Signal Evaluation`, его надо выключить: PSF-анализ звёзд для солнечных кадров бессмысленен.

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
C:\AstroPhoto\Processing\Totality\Total-100ms\Frames\manifest.json
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
.\IntegrateTotality.ps1
```

Скрипт запросит исходную папку:

```text
Input folder:
  [1] Registered (default)
  [2] Debayered
```

Для автоматического запуска можно передать `-InputStage Registered` или `-InputStage Debayered`. Вариант `Debayered` полезен, если регистрация сделала последовательность менее стабильной.

Если PixInsight уже запущен, скрипт также предложит режим выполнения:

```text
PixInsight execution mode:
  [1] Reuse a running instance and leave it open (default)
  [2] Start a dedicated instance and close it when finished
```

`Reuse` передаёт временный PJSR-скрипт первому доступному экземпляру через IPC `--execute`. Ключи `-n` и `--force-exit` в этом режиме не используются, поэтому существующий PixInsight и открытые в нём изображения после интеграции остаются открытыми. Пока работает ImageIntegration, этот экземпляр PixInsight занят; если нужно параллельно продолжать ручную работу, выберите `Dedicated`. Если запущено несколько экземпляров, PixInsight выбирает первый доступный IPC slot.

Для неинтерактивного запуска:

```powershell
.\IntegrateTotality.ps1 `
    -Exposure Total-100ms `
    -InputStage Debayered `
    -PixInsightMode Reuse
```

Скрипт читает `manifest.json`, режет массив на исходные временные блоки и вызывает PixInsight `ImageIntegration`.

Output:

```text
C:\AstroPhoto\Processing\Totality\Total-100ms\Frames\integrated\
```

Пример:

```text
Block_003_100ms_Totality_00063_00090.xisf
```

То же самое сделать для 1 и 10 ms.

---

### 12. Рабочие настройки ImageIntegration

Для totality blocks:

```text
Combination                 = Average
Normalization               = No normalization
Weights                     = Don't care

Pixel Rejection             = Winsorized Sigma Clipping
Pixel Rejection Normalization = Scale + zero offset

Sigma Low                   = 3.5
Sigma High                  = 4.0

Large-scale low rejection   = ON
Layers                      = 2
Growth                      = 3

Large-scale high rejection  = OFF

Generate rejection maps     = ON
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

### 15. HDR каждого блока

После интеграции каждой экспозиции получится примерно:

```text
Total-1ms\Frames\integrated\
    Block_003_1ms_Totality_....

Total-10ms\Frames\integrated\
    Block_003_10ms_Totality_....

Total-100ms\Frames\integrated\
    Block_003_100ms_Totality_....
```

Для одного временного блока собрать тройку:

```text
Block03_1ms
Block03_10ms
Block03_100ms
```

Эти три masters были зарегистрированы к разным исходным reference frames, поэтому перед HDR их надо **точно совместить друг с другом**.

Практически:

```text
reference = 10 ms master
```

1 ms и 100 ms довести к нему через `DynamicAlignment`.

Реперы:

```text
1 ms  → протуберанцы / хромосфера
10 ms → reference
100 ms → структура внутренней короны
```

После этого:

```text
HDRComposition
```

Порядок:

```text
1 ms
10 ms
100 ms
```

`Reject black pixels` включить.

Роли экспозиций:

```text
1 ms
→ лимб, Baily beads, хромосфера, протуберанцы

10 ms
→ внутренняя и средняя корона

100 ms
→ средняя и внешняя корона
```

Пересвет протуберанцев на 100 ms не имеет значения: HDR должен заменить эту область короткими экспозициями.

---

### 16. Главный статичный результат

Не обязательно смешивать всю totality.

Для главного still лучше выбрать **один из центральных блоков** или 2–3 соседних:

* beads уже закончились;
* дерево ещё не дошло до диска;
* лунный лимб чистый;
* геометрия между 1/10/100 ms почти одинаковая.

Если один центральный HDR выглядит хорошо, он может быть главным результатом вообще без сложного time-composite.

---

### 17. Time-composite всей totality

Отдельный экспериментальный результат.

Все блоки можно зарегистрировать в солнечной системе координат и комбинировать так, чтобы:

* хромосфера собиралась с разных сторон;
* протуберанцы сохранялись;
* moving Moon считалась временной маской;
* деревья/облака по возможности удалялись rejection/масками.

Это уже не снимок одного мгновения, а временной композит полной фазы. Его делать только после получения нормального single-block HDR.

---

### 18. Анимация

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

### 19. Цвет и атмосферная дисперсия

Не пытаться исправлять цвет до calibration/debayer. Например, если съёмка была сделана с:

```text
WB_R = 23
WB_B = 99
```

то линейный цвет изначально сильно искажён.

Кроме того, если Солнце было очень низко, то присутствует сильная атмосферная дисперсия.

После получения masters/HDR можно:

1. проверить R/G/B отдельно;
2. через `ChannelMatch` совместить R и B относительно G по лимбу/протуберанцам;
3. только затем заниматься цветовым балансом;
4. финальный stretch делать после HDR.

Цвет сумеречного неба не использовать как физический эталон.

---

### 20. Что не переделывать без причины

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
```

Если меняются настройки ImageIntegration, переделывается только:

```text
Integration
→ HDR
```

Исходные SER и calibrated CFA всегда сохранять как неизменяемую основу.

---

### 21. Текущие особенности именно этого набора

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
→ собирать HDR каждого блока
→ выбрать центральный HDR как основной still
→ остальные использовать для time-composite и анимации
```
