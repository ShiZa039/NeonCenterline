# Разработка

Как устроен репозиторий, где что править, как собрать и проверить. Устройство плагина изнутри — [architecture.md](architecture.md), выпуск — [release.md](release.md), грабли — [pitfalls.md](pitfalls.md).

## Что нужно

| Для чего | Что |
|---|---|
| Собрать и проверить плагин | Windows, CorelDRAW с компонентом VBA (проверено в 2026), доступ к объектной модели редактора VBA (для сборщика) |
| Генераторы и проверки | Python 3 |
| Значки панели (`build_resources.py`) | Python + Pillow, `binutils-mingw-w64` (`x86_64-w64-mingw32-windres`, `…-ld`) — проще всего в Linux/WSL |
| Компиляция без CorelDRAW (`lo_compile.py`) и стенды (`lo_stand.py`) | LibreOffice. Проверено в Linux; пути в `lo_compile.py` — линуксовые |
| Выпуск | PowerShell 5.1+ (`make_release.ps1`), git, доступ к GitHub Releases |

## Карта репозитория

| Путь | Что это | Править? |
|---|---|---|
| `src/NeonUI.bas` | Кнопки «Моно» / «Курсив», окна параметров, настройки, «О программе», проверка панели, окно хода работы (`NeonJobRun`, `NeonProg*`), отчёт (`NeonReport`). Версия `NEON_VERSION`, набор кнопок `NEON_TOOLBAR_VER`, адреса `NEON_UPDATE_URL*`, `NEON_DONATE_URL` | да |
| `src/NeonTools.bas` | Инструменты: длина, выделить линии, упростить, очистить, люверсы, разбить на буквы, одинаковые, лапша, раскладка. Геометрия — в разделах `'@@…_BEGIN` / `'@@…_END` | да |
| `src/NeonCut.bas` | Тип неона, кратность реза, подгонка линий | да |
| `src/NeonUpdate.bas` | Проверка и установка обновлений | да |
| `src/frmNeonRun.code.txt` | Код окна параметров и окна хода работы (режимы 0–5, см. architecture.md). Недостающие поля окно создаёт само | да |
| `src/frmNeonSettings.code.txt` | Код окна настроек (само окно создаёт сборщик; строки упрощения и «Оставлять оригинал» окно добавляет само) | да |
| `src/ThisMacroStorage.code.txt` | События запуска CorelDRAW и открытия окна: обновление панели после обновления плагина | да |
| `core/NeonCenterline.bas`, `core/NeonScript.bas` | Рабочие модули «Моно» и «Курсив» (макросы, которые можно запускать отдельно) | да (Моно — только с нулевой регрессией) |
| `src/NeonMono.bas`, `src/NeonScript.bas` | Ядра в плагине | **нет** — `tools/gen_core.py` из `core/` |
| `tools/strings.py` | Все строки интерфейса RU/EN | да |
| `src/NeonLang.bas` | Строки в плагине | **нет** — `tools/gen_lang.py` |
| `addon/NeonCenterline/` | Панель: `AppUI.xslt`, `UserUI.xslt`, `Coreldrw.addon`, `config.xml`, `Resources<N>.dll` | да (`config.xml` и DLL — через `build_resources.py`) |
| `src/NeonAddon.bas` | Файлы панели внутри плагина + `InstallToolbar` | **нет** — `tools/gen_addon.py` |
| `icons/` | Значки кнопок (.ico и PNG 16/32) | да |
| `build/NeonBuild.bas` | Сборщик `.gms` из `src` (только ASCII — его импортируют вручную) | да |
| `proto/*.py` | Прототипы на Python — эталон для VBA: `centerline*.py`, `simplify.py`, `cleanup.py`, `grommets.py`, `letters.py`, `nesting.py`, `rasternest.py`, `noodles.py` | да |
| `tools/lo_stand.py`, `tools/stands/` | Стенд LibreOffice: раздел VBA против прототипа | да |
| `samples/` | Эталонные SVG для «Курсива» | да |
| `font/SABNeonTest.ttf` | Тестовый шрифт для «Моно» | — |
| `install.cmd`, `install.ps1` | Установщик для пользователей | да (`.ps1` — без не-ASCII в коде) |
| `tools/make_release.ps1` | Файлы выпуска, `update.json` | да |
| `dist/NeonCenterline.gms` | Сборка последней версии (запасной адрес обновления) | только `make_release.ps1` |
| `update.json` | Последняя версия — читают плагины пользователей | только `make_release.ps1`, пушить последним |
| `README.md` → `README.txt` | Описание для пользователей → текст для архива | `README.md`; `.txt` — `gen_readme_txt.py` |
| `CHANGELOG.md`, `ROADMAP.md`, `docs/` | История, планы, документация | да |
| `release/`, `src_ansi/` | Временные (в `.gitignore`) | — |

## Порядок работы над изменением

### Геометрия (новый инструмент или правка расчёта)

1. **Прототип** в `proto/` — на Python, «в стиле VBA» (простые списки, без numpy в алгоритме), чтобы порт был один в один. Проверить на образцах, нарисовать картинку.
2. **Раздел VBA** в `src/NeonTools.bas`: расчёт — в `'@@ИМЯGEO_BEGIN` / `'@@ИМЯGEO_END`, объявления уровня модуля — в `'@@ИМЯ_DECL_BEGIN` / `_END` в начале модуля. Раздел не трогает объекты Corel: на входе массивы точек, на выходе — массивы. Из раздела можно звать только `NeonProgStep` / `NeonProgPack` / `NeonProgBox` / `NeonProgCompare`.
3. **Стенд**: раздел в LibreOffice против прототипа на одних и тех же данных — образец `tools/stands/grom_hem.py` ([ниже](#стенд-libreoffice)).
4. **Команда для Corel** (раздел без `GEO`): чтение выделения в массивы → расчёт → объекты. Шаблон — [ниже](#шаблон-команды).
5. **Строки** — в `tools/strings.py` (RU и EN), затем `gen_lang.py`.
6. **Кнопка на панели** (если новая): `tools/build_resources.py` (`ITEMS`), `AppUI.xslt`, поднять `NEON_TOOLBAR_VER` и метку `[NeonToolbar N]`, значок в `icons/`, затем `build_resources.py` и `gen_addon.py`. Новую команду — в `NeonJobGo` (`NeonUI.bas`).
7. **Проверки**, **сборка**, **проверка в CorelDRAW**.
8. **Тексты**: `README.md` (RU и EN) → `gen_readme_txt.py`, `CHANGELOG.md`, `ROADMAP.md`.

### «Моно» и «Курсив»

Править `core/NeonCenterline.bas` / `core/NeonScript.bas`, затем `python tools/gen_core.py` (заменяет макрос-точку входа функцией без интерфейса, вставляет шаги окна хода работы и имя группы отмены). «Курсив» — сверка с `samples/`.

«Моно» — только при нулевой регрессии на `font/SABNeonTest.ttf`, проверка — `tools/stands/mono_font.py` (ядро в LibreOffice, весь шрифт, сравнение узел в узел):
```
python tools/stands/mono_font.py snap base.pkl                  # до правки (или --core <копия «до»>)
python tools/stands/mono_font.py snap new.pkl                   # после
python tools/stands/mono_font.py diff base.pkl new.pkl --png d.png   # «diff 0» — можно; иначе картинка изменившихся строк
```
По умолчанию — высота прописных 70 мм с фрезой 6 и 100 мм с фрезой 8 (~16 минут); `--sizes 40:4,6:1` — свои размеры, `--text` — только нужные знаки. Изменения на мелком тексте (штрих тоньше 4 мм) ожидаемы, если правка именно о нём — смотреть картинку.

«Курсив» — тот же стенд с `--script N` (ядро `core/NeonScript.bas`, режим «шрифт как образец» N: 0 — нет, 1 — до кончика штриха, как в плагине по умолчанию, 2 — с отступом на полширины). Рукописных шрифтов в репозитории нет — `--font <ttf/otf/woff/woff2>` (например, Marck Script, Bad Script, Caveat с fontsource), строки — `--texts "Привет тест|Neon sign"`, готовые контуры (дамп из CorelDRAW: список объектов, объект — список колец `[(x, y), …]` в мм) — `--polys <файл.pkl>`. Копию ядра «до» — в `--core`, иначе берётся текущее:
```
python tools/stands/mono_font.py snap s_base.pkl --core base_script.bas --script 1 --font marck.woff2 --texts "Привет тест|абвгд…" --sizes 80:8
python tools/stands/mono_font.py snap s_new.pkl --script 1 --font marck.woff2 --texts "Привет тест|абвгд…" --sizes 80:8
python tools/stands/mono_font.py diff s_base.pkl s_new.pkl --png d.png --font marck.woff2
```
«Курсив» считается заметно дольше «Моно»: набор сверки 14.11 (три шрифта, режимы 0/1/2, дамп из CorelDRAW) — около полутора часов на версию.

### Шаблон команды

```vb
Public Sub DoSomething()
    Dim sr As ShapeRange, oldUnit As cdrUnit, grpOn As Boolean, errText As String
    If ActiveDocument Is Nothing Then
        MsgBox NeonL("err_nodoc"), vbExclamation, NeonL("xx_title")
        Exit Sub
    End If
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then ... : Exit Sub
    ' 1) окно параметров — только при первом входе; затем то же самое ещё раз, уже в окне хода работы
    If Not NeonInJob("xx") Then
        If Not AskXxParams() Then Exit Sub
        NeonJobRun "xx", NeonL("xx_title")      ' + Case "xx" в NeonJobGo
        Exit Sub
    End If
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    On Error GoTo Fail
    ActiveDocument.BeginCommandGroup NeonL("xx_title")   ' всё действие — одна группа отмены
    grpOn = True
    Optimization = True: EventsEnabled = False
    NeonProgPhase NeonL("prog_objects"), sr.Count
    ' ... цикл: NeonProgStep i ...
Fin:
    EventsEnabled = True: Optimization = False
    If grpOn Then ActiveDocument.EndCommandGroup
    ActiveDocument.Unit = oldUnit
    NeonProgEnd                                         ' до любого MsgBox
    If Len(errText) > 0 Then
        MsgBox NeonLF("err_fail", errText), vbCritical, NeonL("xx_title")
        Exit Sub
    End If
    NeonReport msg, NeonL("xx_title"), changed           ' «Отмена» — только если что-то поменялось
    NeonToolbarCheck
    Exit Sub
Fail:
    errText = Err.Description
    Resume Fin
End Sub
```
(В настоящем коде `If … Then … : Exit Sub` в одну строку не пишется — только блоком; здесь сжато.)

## Генераторы и проверки

| Что менялось | Запустить |
|---|---|
| строки (`tools/strings.py`) | `python tools/gen_lang.py` |
| ядра (`core/*.bas`) | `python tools/gen_core.py` |
| файлы панели (`addon/NeonCenterline/*`) | `python tools/gen_addon.py` |
| состав панели, подписи кнопок или значки | поднять `NEON_TOOLBAR_VER` и `[NeonToolbar N]` в `AppUI.xslt` → `python tools/build_resources.py` → `python tools/gen_addon.py` |
| `README.md` | `python tools/gen_readme_txt.py` |

Перед каждой сборкой — все три проверки:

| Проверка | Что ловит | Почему не хватает компиляции |
|---|---|---|
| `python tools/shadow_check.py` | переменная с именем вызываемой процедуры («Expected array»), дубли Public-имён, объявления уровня модуля после процедур («Variable not defined») | LibreOffice это прощает |
| `python tools/call_check.py` | число аргументов вызовов, ключи строк `NeonL`/`NeonLF`, число подстановок `NeonLF` (до 8) | LibreOffice прощает лишние аргументы |
| `python tools/lo_compile.py` | компиляция всех модулей в LibreOffice Basic с заглушками объектов Corel | — |

Чего не ловит ни одна: не сокращённые `And` / `Or` с индексом (только вложенными `If`, см. [pitfalls.md](pitfalls.md#and--or-не-сокращённые--subscript-out-of-range)) и всё, что зависит от настоящего CorelDRAW.

## Стенд LibreOffice

Раздел VBA вырезается из `src/NeonTools.bas` как есть и исполняется в LibreOffice Basic (`Option VBASupport 1`) на тех же данных, что и прототип. Помощник — `tools/lo_stand.py`:

```python
import lo_stand as S
main = ['Sub Main', '    Dim g As Integer', '    g = FreeFile', f'    Open "{S.OUT}" For Output As #g',
        *S.put_segs(segs),                  # данные — туда же, куда их кладёт код для Corel
        '    GromHem 50', '    Print #g, GHN', ...,
        '    Close #g', 'End Sub']
code = S.module(S.section('SIMP_DECL'), S.section('GROM_DECL'), S.section('GROMGEO'), main)
lines = S.run(code)                         # что напечатал Main
```
- `module()` дописывает процедуры, которые раздел зовёт, а лежат вне его (`PutS`, `NormV`, `BzPt`…), и пустые `NeonProg*`;
- `run()` кладёт модуль в библиотеку `Standard` LibreOffice, запускает без экрана с таймаутом; ошибка компиляции — `StandTimeout`;
- `python tools/lo_stand.py GROMGEO NESTRASTER` — только «компилируется ли раздел».

Образец целиком — `tools/stands/grom_hem.py` (контур под подгиб против `proto/grommets.py hem()`): `python tools/stands/grom_hem.py` → `mismatches: 0`. Новый стенд — копия образца: свои входные данные, своя процедура, сравнение с допуском (`1e-4` мм для точек). Стенд ядер «Моно» и «Курсива» по целым шрифтам — `tools/stands/mono_font.py` (см. [выше](#моно-и-курсив)).

## Сборка `.gms`

1. CorelDRAW → Alt+F11 (редактор VBA) → проект `GlobalMacros` → File → Import → `build/NeonBuild.bas` (один раз; при изменении сборщика — заново).
2. Запустить макрос `NeonBuildPlugin` (в редакторе: курсор в процедуре → F5), указать папку `src` (запоминается в реестре, `Build\Src`).
3. Сборщик создаёт (или очищает) проект `NeonCenterline` в папке GMS пользователя, добавляет модули из `src` (UTF-8 читает сам и добавляет через `AddFromString`), создаёт окно настроек, вставляет код `ThisMacroStorage`.
4. В редакторе VBA выделить проект `NeonCenterline` → **Ctrl+S**. Файл — `%APPDATA%\Corel\CorelDRAW Graphics Suite <версия>\Draw\GMS\NeonCenterline.gms`.
5. Перезапустить CorelDRAW, «О программе» — нужная версия.

Нет доступа к `Application.VBE` (сборщик скажет) — собрать вручную: «Новый проект макросов» `NeonCenterline`, `python tools/to_ansi.py` → импортировать модули из `src_ansi/`; без окна настроек настройки задаются стандартными окошками.

## Проверка в CorelDRAW

Минимум после изменений в инструменте:
- работает на кривых, тексте, группах; пустое выделение — понятное сообщение;
- окно хода работы появляется и закрывается само; на большом выделении не «Не отвечает»;
- «ОК» оставляет, «Отмена» отменяет всё одним шагом; если ничего не поменялось — только «ОК»;
- «Оставить оригинал» (где есть): копии поверх, у неизменённых — копий нет;
- RU и EN (Настройки → язык): подписи влезают в окна;
- ошибка посреди команды не оставляет CorelDRAW без перерисовки и в чужих единицах.

## Git

- Пушим в `main`; выпуск — тегом `X.Y.Z` (без `v`), см. [release.md](release.md).
- Концы строк: `*.bas`, `*.txt`, `*.cmd`, `*.ps1` — CRLF (`.gitattributes`).
- На GitHub — только этот репозиторий.
