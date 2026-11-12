# Сборка и выпуск — шпаргалка · Build & release cheat sheet

Подробно — в [docs/](docs/README.md): [разработка](docs/development.md), [устройство](docs/architecture.md), [выпуск](docs/release.md), [подводные камни](docs/pitfalls.md), [если не работает](docs/troubleshooting.md).

## Где править

- Строки — `tools/strings.py` → `python tools/gen_lang.py` (не `src/NeonLang.bas`).
- «Моно» / «Курсив» — `core/*.bas` → `python tools/gen_core.py` (не `src/NeonMono.bas`, `src/NeonScript.bas`).
- Панель — `addon/NeonCenterline/*` → `python tools/gen_addon.py` (не `src/NeonAddon.bas`). Поменялся состав кнопок или значки — поднять `NEON_TOOLBAR_VER` в `src/NeonUI.bas` и `[NeonToolbar N]` в `AppUI.xslt`, затем `python tools/build_resources.py` и `gen_addon.py`.
- `README.md` → `python tools/gen_readme_txt.py`.
- Геометрия — сначала `proto/`, потом раздел `'@@…GEO` в `src/NeonTools.bas`, потом стенд (`tools/stands/grom_hem.py` — образец).

## Перед сборкой

```
python tools/shadow_check.py   -> issues 0
python tools/call_check.py     -> issues 0
python tools/lo_compile.py     -> все модули OK
```
Условия с индексом — вложенными `If` (в VBA `And` / `Or` не сокращённые).

## Сборка

CorelDRAW → Alt+F11 → GlobalMacros → File → Import → `build/NeonBuild.bas` (один раз) → макрос `NeonBuildPlugin` → папка `src` → выделить проект `NeonCenterline` → Ctrl+S → перезапустить CorelDRAW.

## Выпуск

1. `NEON_VERSION`, `CHANGELOG.md`, `docs/releases/<версия>.md` → push; собрать, проверить в CorelDRAW.
2. `.\tools\make_release.ps1 -Version <версия> -NotesRu '...' -NotesEn '...'` (снимает заголовок `.gms` CorelDRAW 2026 — так файл принимают плагины 1.0.x–1.1.1).
3. Запушить **только** `dist/NeonCenterline.gms`.
4. GitHub Release: тег `<версия>` без `v` на `main`, описание из `docs/releases/<версия>.md`, файлы `release\NeonCenterline.gms` и `release\NeonCenterline-<версия>.zip` → Publish.
5. Проверить: оба файла в выпуске, `digest` у `.gms` = `sha256` в `update.json`.
6. Только теперь запушить `update.json`.

## English

Edit sources, not generated files: strings in `tools/strings.py` (`gen_lang.py`), cores in `core/` (`gen_core.py`), toolbar in `addon/` (`gen_addon.py`; bump `NEON_TOOLBAR_VER` and `[NeonToolbar N]` when buttons change, then `build_resources.py`). Before building run `shadow_check.py`, `call_check.py`, `lo_compile.py`. Build: import `build/NeonBuild.bas` into GlobalMacros, run `NeonBuildPlugin`, point it to `src`, save the `NeonCenterline` project (Ctrl+S). Release: bump `NEON_VERSION`, date `CHANGELOG.md`, write `docs/releases/<version>.md`, push; build and test; run `tools/make_release.ps1` (it strips the CorelDRAW 2026 `.gms` header so 1.0.x–1.1.1 updaters accept the file); push `dist/NeonCenterline.gms` only; publish the GitHub release tagged `<version>` (no "v") on `main` with both files from `release\`; check the `.gms` digest equals `sha256` in `update.json`; then push `update.json`. Details (in Russian): [docs/](docs/README.md).
