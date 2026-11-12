# Сборка и выпуск (для автора) · Build & release (for the author)

## Структура репозитория

| Путь | Что это |
|---|---|
| `src/NeonMono.bas`, `src/NeonScript.bas` | Ядра «Моно» и «Курсив» (UTF-8). **Генерируются** `tools/gen_core.py` из рабочих модулей макроса `core/NeonCenterline.bas`, `core/NeonScript.bas` — алгоритм править там, не в `src`. Генератор заменяет макрос-точку входа функцией без интерфейса и вставляет шаги окна хода работы. |
| `src/NeonUI.bas` | Кнопки «Моно» / «Курсив», окна параметров, настройки, «О программе», проверка панели, окно хода работы (`NeonJobRun`, `NeonProg*`) и отчёт «ОК / Отмена» (`NeonReport`). Версия — `NEON_VERSION`, набор кнопок — `NEON_TOOLBAR_VER`, адреса обновлений — `NEON_UPDATE_URL*`, донат — `NEON_DONATE_URL`. |
| `src/NeonTools.bas` | Инструменты: длина, выделить линии, упростить, очистить, люверсы, разбить на буквы, одинаковые, лапша, раскладка. Геометрия — в разделах `'@@…_BEGIN` / `'@@…_END` (см. ниже). |
| `src/NeonCut.bas` | Тип неона, кратность реза, подгонка линий. |
| `src/NeonUpdate.bas` | Проверка и установка обновлений. |
| `src/NeonAddon.bas` | Файлы панели внутри плагина и макрос `InstallToolbar`. **Генерируется** `tools/gen_addon.py`. |
| `src/NeonLang.bas` | Строки RU/EN. **Генерируется** `tools/gen_lang.py` из `tools/strings.py` — править строки там. |
| `src/frmNeonRun.code.txt` | Код окна перед запуском (frmNeonRun): режимы 0 — центральные линии, 1 — упрощение, 2 — люверсы, 3 — раскладка, 4 — лапша, 5 — окно хода работы. Недостающие поля окно создаёт само при открытии. |
| `src/frmNeonSettings.code.txt` | Код окна настроек (само окно создаёт сборщик; строки упрощения и «Оставлять оригинал» окно добавляет само). |
| `src/ThisMacroStorage.code.txt` | Код ThisMacroStorage: при запуске CorelDRAW обновляет панель «Неон» после обновления плагина. |
| `build/NeonBuild.bas` | Сборщик: создаёт `NeonCenterline.gms` из `src`. |
| `addon/NeonCenterline/` | Панель инструментов (CorelDRAW Addons): `AppUI.xslt`, `UserUI.xslt`, `Coreldrw.addon`, `config.xml`, `Resources<N>.dll` — значки и подписи набора кнопок N. |
| `icons/` | Значки кнопок (.ico и PNG 16/32) — для `tools/build_resources.py` и для ручной установки. |
| `proto/` | Прототипы на Python — эталон для VBA: `centerline*.py` (ядро), `simplify.py`, `cleanup.py`, `grommets.py`, `letters.py`, `nesting.py`, `rasternest.py`, `noodles.py`. |
| `samples/` | Эталонные SVG для «Курсива» (артефакты и образцы автора). |
| `font/SABNeonTest.ttf` | Тестовый шрифт для «Моно». |
| `docs/releases/<версия>.md` | Описание выпуска для GitHub Releases. |
| `dist/NeonCenterline.gms` | Сборка последней версии (запасной адрес обновления через jsDelivr). |
| `update.json` | Описание последней версии — его читают плагины пользователей. |
| `install.cmd`, `install.ps1` | Установщик для пользователей. |

## Сборка `.gms`

1. CorelDRAW → Alt+F11 → GlobalMacros → File → Import → `build/NeonBuild.bas` (один раз; при изменении сборщика — заново).
2. Запустить макрос `NeonBuildPlugin`, указать папку `src`.
3. В редакторе VBA выделить проект `NeonCenterline` → Ctrl+S.
4. Файл лежит в пользовательской папке GMS (`%APPDATA%\Corel\CorelDRAW Graphics Suite <версия>\Draw\GMS\NeonCenterline.gms`). Перезапустить CorelDRAW.

Если сборщик пишет, что нет доступа к `Application.VBE`, — собрать вручную: создать проект (Скрипты → Новый проект макросов `NeonCenterline`), импортировать модули из `src` в кодировке Windows (`python tools/to_ansi.py` → папка `src_ansi`); окно настроек можно не делать — без него настройки задаются стандартными окошками.

## После правок — генераторы и проверки

| Что менялось | Запустить |
|---|---|
| строки (`tools/strings.py`) | `python tools/gen_lang.py` |
| ядра (`core/*.bas`) | `python tools/gen_core.py` |
| файлы панели (`addon/NeonCenterline/*`) | `python tools/gen_addon.py` |
| состав панели или значки | поднять `NEON_TOOLBAR_VER` и метку `[NeonToolbar N]` в `AppUI.xslt`, затем `python tools/build_resources.py` и `python tools/gen_addon.py` |
| `README.md` | `python tools/gen_readme_txt.py` (README.txt для архива) |

Перед каждой сборкой:
- `python tools/shadow_check.py` — переменная с именем вызываемой процедуры, дубли Public-имён, объявления уровня модуля после процедур (VBA: «Expected array», «Variable not defined»);
- `python tools/call_check.py` — число аргументов вызовов, ключи строк и число подстановок `NeonLF` (до 8: `{0}`…`{7}`);
- `python tools/lo_compile.py` — компиляция всех модулей в LibreOffice Basic с заглушками объектов Corel.

LibreOffice прощает то, чего не прощает VBA: `And` / `Or` в VBA не сокращённые (`If i < n And a(i) > 0` читает `a(i)` всегда), объявления после процедур, лишние аргументы. Поэтому кроме компиляции — `shadow_check` и `call_check`, а условия с индексами — вложенными `If`.

## Устройство кода

- **Геометрия отдельно от Corel.** Расчёты инструментов — в разделах `src/NeonTools.bas` между метками `'@@ИМЯ_BEGIN` / `'@@ИМЯ_END` (`GROMGEO`, `CLEANGEO`, `SAMEGEO`, `LETTERSGEO`, `NOODLEGEO`, `NESTGEO`, `NESTRASTER`; объявления — `*_DECL`; разделы без `GEO` — код команд для Corel). Они не трогают объекты Corel и повторяют прототипы из `proto/` один в один; каждый раздел сверялся с прототипом на стенде: раздел вырезается в модуль LibreOffice Basic, получает те же входные данные и сравнивается с Python. Из разделов зовутся только `NeonProgStep` / `NeonProgPack` / `NeonProgBox` / `NeonProgCompare` (окно хода работы) — стенду нужны пустые заглушки.
- **Окно хода работы.** Команда сначала показывает своё окно параметров, затем `NeonJobRun "<задание>"` открывает frmNeonRun в режиме 5; его `Activate` вызывает `NeonJobGo`, а тот — ту же команду ещё раз, уже с `NeonInJob("<задание>") = True` (окно параметров второй раз не показывается). Шаги — `NeonProgPhase` (этап и число шагов), `NeonProgStep i` / `NeonProgTick` (обновление не чаще 5 раз в секунду, тогда же `DoEvents` — окно не становится «Не отвечает»), `NeonProgEnd` — перед отчётом или сообщением об ошибке. Без формы задание выполняется сразу.
- **Отчёт.** `NeonReport msg, заголовок, canUndo`: при `canUndo` — «ОК / Отмена», «Отмена» = `ActiveDocument.Undo` (действие — одна группа отмены `BeginCommandGroup` / `EndCommandGroup`). Отмену предлагать, только если действие что-то поменяло, иначе отменится чужой шаг.
- **Оставить оригинал.** Инструмент работает с копиями (`Duplicate`), копия без изменений удаляется; у раскладки копируется всё выделение (`ShapeRange.Duplicate`), не вставшие копии удаляются.
- **Метки объектов.** Свои объекты плагин помечает `Properties("NeonCenterline", 1)` = `same` / `sheet` / `frame` / `label` (запасной путь — начало имени). Линии центров — по имени «Неон: центр N» (не менять: по нему узнаются и линии из старых файлов).
- **Строки.** `tools/strings.py` → `NeonLang.bas`: сжатая запись (`#` переключает кириллицу, `$XXXX` — любой символ), функции по 40 строк — LibreOffice не компилирует модуль с литералами больше ~64 КБ, VBA — процедуры больше 64 КБ.
- **Панель.** Набор кнопок N: метка `[NeonToolbar N]` в `AppUI.xslt` = `NEON_TOOLBAR_VER`, значки — `Resources<N>.dll` (новое имя на каждый набор, чтобы установка не упиралась в файл, занятый CorelDRAW). После обновления плагин сравнивает номер, предлагает обновить файлы панели (права администратора) и после перезапуска сбрасывает только панель «Неон».

## Выпуск новой версии

1. Поднять `NEON_VERSION` в `src/NeonUI.bas`; в `CHANGELOG.md` — дата выпуска у версии; описание выпуска — `docs/releases/<версия>.md` (RU и EN: что нового, «Инструкция по установке — в README», значок скачиваний выпуска). Прогнать генераторы и проверки (см. выше), закоммитить, запушить.
2. `git pull`, собрать `.gms` (см. «Сборка»), проверить в CorelDRAW.
3. PowerShell в корне репозитория:
   `.\tools\make_release.ps1 -Version <версия> -NotesRu '...' -NotesEn '...'`
   — находит свежий `.gms`, **снимает заголовок CorelDRAW 2026** («GMS», 18 байт: плагины 1.0.x–1.1.1 при обновлении ждут подпись OLE в начале файла и иначе отбраковывают его; CorelDRAW 2026 файл без заголовка загружает), кладёт в `release\` и `dist\`, собирает `release\NeonCenterline-<версия>.zip` для новых пользователей, считает SHA-256, пишет `update.json`.
4. Закоммитить и запушить **только** `dist\NeonCenterline.gms` — чтобы тег встал на коммит с новой сборкой (запасной адрес jsDelivr `@<версия>/dist/…`).
5. GitHub → Releases → новый релиз: тег `<версия>` (без буквы v), на `main`; заголовок `NeonCenterline <версия>`; описание — содержимое `docs/releases/<версия>.md`; приложить `release\NeonCenterline.gms` и `release\NeonCenterline-<версия>.zip` → Publish.
6. Только после этого закоммитить и запушить `update.json` — с этого момента плагины пользователей предложат обновление.
7. Проверить: файлы в релизе на месте, SHA-256 `.gms` в релизе = `sha256` в `update.json`.

## Build & release (English)

Build: import `build/NeonBuild.bas` into GlobalMacros, run `NeonBuildPlugin`, point it to `src`, save the `NeonCenterline` project in the VBA editor (Ctrl+S). Before building run the generators that apply (`gen_lang`, `gen_core`, `gen_addon`, `build_resources`, `gen_readme_txt`) and the checks (`shadow_check`, `call_check`, `lo_compile`).

Release: bump `NEON_VERSION`, date the version in `CHANGELOG.md`, write `docs/releases/<version>.md`, push; rebuild and test; run `tools/make_release.ps1`; push `dist/NeonCenterline.gms` only; create the GitHub release tagged `<version>` (no "v") on `main` with the notes and both files (`.gms` and the zip); then push `update.json`; check that the release `.gms` matches the `sha256` in `update.json`.
