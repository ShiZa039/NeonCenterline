# Сборка и выпуск (для автора) · Build & release (for the author)

## Структура репозитория

| Путь | Что это |
|---|---|
| `src/NeonMono.bas`, `src/NeonScript.bas` | Ядра «Моно» и «Курсив» (UTF-8). **Генерируются** `tools/gen_core.py` из рабочих модулей макроса — алгоритм там не править вручную. |
| `src/NeonUI.bas` | Кнопки, настройки, «О программе». Версия — `NEON_VERSION`, адрес обновлений — `NEON_UPDATE_URL`, ссылка на донат — `NEON_DONATE_URL`. |
| `src/NeonUpdate.bas` | Проверка и установка обновлений. |
| `src/NeonLang.bas` | Строки RU/EN. **Генерируется** `tools/gen_lang.py` из `tools/strings.py` — править строки там. |
| `src/frmNeonSettings.code.txt` | Код окна настроек (само окно создаёт сборщик). |
| `build/NeonBuild.bas` | Сборщик: создаёт `NeonCenterline.gms` из `src`. |
| `addon/NeonCenterline/` | Панель инструментов (CorelDRAW Addons: `AppUI.xslt`, `UserUI.xslt`, `Coreldrw.addon`). |
| `update.json` | Описание последней версии — его читают плагины пользователей. |
| `install.cmd` | Установщик для пользователей. |

## Первая сборка `.gms`

1. Адрес обновлений уже настроен на github.com/ShiZa039/NeonCenterline (`NEON_UPDATE_URL` в `src/NeonUI.bas`); при желании вписать `NEON_DONATE_URL`.
2. CorelDRAW → Alt+F11 → GlobalMacros → File → Import → `build/NeonBuild.bas`.
3. Запустить макрос `NeonBuildPlugin`, указать папку `src`.
4. В редакторе VBA выделить проект `NeonCenterline` → Ctrl+S.
5. Файл лежит в пользовательской папке GMS (`%APPDATA%\Corel\CorelDRAW Graphics Suite <версия>\Draw\GMS\NeonCenterline.gms`).

Если сборщик пишет, что нет доступа к `Application.VBE`, — собрать вручную: создать проект (Скрипты → Новый проект макросов `NeonCenterline`), импортировать модули из `src` (в кодировке Windows: `tools/to_ansi.py` или через сборщик на другой машине), окно настроек можно не делать — без него настройки задаются стандартными окошками.

## Выпуск новой версии

1. Поднять `NEON_VERSION` в `src/NeonUI.bas`, дописать `CHANGELOG.md`.
2. Собрать `.gms` (шаги 2–4 выше).
3. PowerShell в корне репозитория:
   `.\tools\make_release.ps1 -Version 1.1.0 -Owner ShiZa039 -NotesRu "..." -NotesEn "..."`
   — копирует `.gms` в `release\` и `dist\` (запасной адрес через CDN jsDelivr, если GitHub у пользователя недоступен), считает SHA-256, пишет `update.json`.
4. GitHub → Releases → новый релиз `v1.1.0`, приложить `release\NeonCenterline.gms` (и zip для новых пользователей).
5. Закоммитить и запушить `update.json` и `dist\NeonCenterline.gms` в `main` (тег `v<версия>` создаётся вместе с релизом). С этого момента плагины предложат обновление.

## Build & release (English)

First build: import `build/NeonBuild.bas` into GlobalMacros, run `NeonBuildPlugin`, point it to `src`, then save the `NeonCenterline` project in the VBA editor (Ctrl+S).
New version: bump `NEON_VERSION`, rebuild and save, run `tools/make_release.ps1`, attach `release/NeonCenterline.gms` to a GitHub release `v<version>`, commit `update.json` to `main`.
