# Выпуск новой версии

Полная процедура: от правок до проверки, что обновление дошло до пользователей. Почему порядок именно такой — в конце и в [pitfalls.md](pitfalls.md#выпуск-и-обновление).

Номер версии — `X.Y.Z` (без `v`), сравнивается по числам: 1.1.10 новее 1.1.9. Дальше в примерах — `1.2.0`.

## 0. Перед выпуском

- [ ] Все изменения проверены в CorelDRAW (список — в `ROADMAP.md`, «Ждёт проверки»).
- [ ] Генераторы прогнаны, проверки чистые ([development.md](development.md#генераторы-и-проверки)):
  ```
  python tools/gen_lang.py        (если менялись строки)
  python tools/gen_core.py        (если менялись core/*.bas)
  python tools/gen_addon.py       (если менялась панель)
  python tools/gen_readme_txt.py  (если менялся README.md)
  python tools/shadow_check.py    -> issues 0
  python tools/call_check.py      -> issues 0
  python tools/lo_compile.py      -> все модули OK
  ```

## 1. Версия и тексты

1. `src/NeonUI.bas`: `NEON_VERSION = "1.2.0"`. Если менялся состав панели — `NEON_TOOLBAR_VER` и метка `[NeonToolbar N]` в `addon/NeonCenterline/AppUI.xslt` (+ `build_resources.py`, `gen_addon.py`).
2. `CHANGELOG.md`: раздел `## 1.2.0 — ГГГГ-ММ-ДД`.
3. `docs/releases/1.2.0.md` — описание выпуска (образец — прошлый файл): значки скачиваний выпуска, «Что нового» RU и EN, «Установка и обновление» со ссылкой на README.
4. `ROADMAP.md`: «Сделано (кратко)», «Ждёт проверки».
5. Закоммитить и запушить.

## 2. Сборка и проверка в CorelDRAW

1. `git pull`.
2. Собрать `.gms` ([development.md](development.md#сборка-gms)): CorelDRAW → Alt+F11 → макрос `NeonBuildPlugin` → папка `src` → в редакторе VBA выделить проект `NeonCenterline` → Ctrl+S.
3. Перезапустить CorelDRAW, «О программе» — `1.2.0`. Прогнать то, что менялось.

## 3. Файлы выпуска

PowerShell в корне репозитория:
```powershell
.\tools\make_release.ps1 -Version 1.2.0 -NotesRu 'Коротко, что нового' -NotesEn 'Short what is new'
```
Скрипт:
- берёт самый свежий `NeonCenterline.gms` из `%APPDATA%\Corel\…`;
- **снимает заголовок CorelDRAW 2026** (18 байт «GMS») — пишет `GMS header (18 bytes) removed for older plugin updaters`; без этого плагины 1.0.x–1.1.1 обновление не примут ([почему](pitfalls.md#повреждён-контрольная-сумма-не-совпала-при-любом-обновлении--заголовок-gms-coreldraw-2026));
- кладёт файл в `release\NeonCenterline.gms` и `dist\NeonCenterline.gms`;
- собирает `release\NeonCenterline-1.2.0.zip` (плагин, установщик, панель, значки, шрифт, README, CHANGELOG, лицензия);
- считает SHA-256 и пишет `update.json` (версия, `gms_url`, `gms_url2`, `sha256`, заметки).

`-NotesRu` / `-NotesEn` — то, что пользователь увидит в окне «Доступна новая версия»: одна-две строки. Кавычки внутри — `''` (две одинарные).

## 4. Запушить только сборку

```
git add dist/NeonCenterline.gms
git commit -m "1.2.0: plugin build in dist"
git push
```
`update.json` пока **не** добавлять. Тег из шага 5 должен встать на этот коммит: запасной адрес `gms_url2` = `cdn.jsdelivr.net/gh/ShiZa039/NeonCenterline@1.2.0/dist/NeonCenterline.gms`.

## 5. Выпуск на GitHub

GitHub → Releases → **Draft a new release**:
- **Tag:** `1.2.0` — новый, без `v`; **Target:** `main`;
- **Title:** `NeonCenterline 1.2.0`;
- **Description:** содержимое `docs/releases/1.2.0.md`;
- **Assets:** `release\NeonCenterline.gms` и `release\NeonCenterline-1.2.0.zip` (оба! файл `.gms` — именно из `release\`);
- не Pre-release, **Publish release**.

## 6. Проверить до `update.json`

Файлы на месте и `.gms` в выпуске = локальный `release\NeonCenterline.gms`:
```powershell
$r = Invoke-RestMethod https://api.github.com/repos/ShiZa039/NeonCenterline/releases/tags/1.2.0
$r.assets | Select-Object name, size, digest
(Get-FileHash release\NeonCenterline.gms -Algorithm SHA256).Hash.ToLower()
(Get-Content update.json -Raw | ConvertFrom-Json).sha256
```
У `NeonCenterline.gms` `digest` = `sha256:<те же 64 символа>`, что и две строки ниже. Первые байты файла без заголовка — `D0 CF 11 E0 A1 B1 1A E1`:
```powershell
[IO.File]::ReadAllBytes("release\NeonCenterline.gms")[0..7] | ForEach-Object { $_.ToString("X2") }
```
(Если есть `gh`: `gh api repos/ShiZa039/NeonCenterline/releases/tags/1.2.0 --jq '.assets[] | {name, size, digest}'`.)

## 7. Запушить `update.json`

```
git add update.json
git commit -m "update.json: 1.2.0"
git push
```
С этого момента плагины предлагают обновление (ежедневная проверка — после расчёта центральных линий; ручная — «Проверить обновления»).

## 8. Проверка после выпуска

- [ ] `https://raw.githubusercontent.com/ShiZa039/NeonCenterline/main/update.json` — новая версия (у кого raw закрыт, плагин возьмёт `github.com/…/raw/main/update.json` или jsDelivr `@main` — тот обновится до ~12 часов; ускорить: `https://purge.jsdelivr.net/gh/ShiZa039/NeonCenterline@main/update.json`).
- [ ] `https://cdn.jsdelivr.net/gh/ShiZa039/NeonCenterline@1.2.0/dist/NeonCenterline.gms` скачивается, SHA-256 тот же.
- [ ] Значок «скачиваний» в описании выпуска и в README показывает число, а не ошибку (кэш shields.io — несколько минут).
- [ ] Живое обновление на одном компьютере: прошлая версия → «Проверить обновления» → «Да» → закрыть CorelDRAW → запустить → «О программе». Если менялась панель — вопрос «панель устарела» → «Да» → права администратора → перезапуск → кнопки нового набора.
- [ ] `ROADMAP.md`: отметить проверенное.

## Если что-то пошло не так

| Ситуация | Что делать |
|---|---|
| Выпуск опубликован без файлов | Edit release → приложить оба файла → Update. `update.json` ещё не пушить. |
| `update.json` уже запушен, а выпуска/файлов нет | Как можно быстрее опубликовать выпуск с файлами (адреса уже в `update.json`), или откатить `update.json` (`git revert`) и выпустить по порядку. |
| В выпуск приложен не тот `.gms` (из `%APPDATA%`, с заголовком, или пересобранный после `make_release`) | Удалить файл из Assets и приложить `release\NeonCenterline.gms`, чей SHA-256 = `sha256` в `update.json`. Пересобирал — заново `make_release.ps1`, и тогда см. следующую строку. |
| Сборку пришлось переделать после того, как тег создан | Новую сборку под тем же номером **не** выпускать: jsDelivr уже мог закэшировать `@1.2.0/dist` навсегда, у кого-то она уже скачана. Выпустить `1.2.1`. |
| Тег создан с `v` или не на `main` | Удалить выпуск и тег на GitHub (до пуша `update.json`!) и создать заново правильно. |
| Пользователи пишут «контрольная сумма не совпала» на новой версии | Сравнить `digest` файла в выпуске, `sha256` в `update.json` и файл в `dist/` на теге. Совпадает выпуск с `update.json` — проблема в сети пользователя (пусть ставит из архива). Не совпадает — исправить `update.json` под файл выпуска. |
| Новая версия сломана | Быстрее всего — `1.2.1` с исправлением. Временно остановить раздачу: откатить `update.json` на прошлую версию (у кого уже скачано — останется; откат у пользователя — `NeonCenterline.gms.bak`). |

## Почему порядок такой

1. **`dist/` до тега** — запасной адрес `gms_url2` смотрит в `dist/` **на теге**; jsDelivr кэширует файлы тега навсегда.
2. **Выпуск с файлами до `update.json`** — плагины читают `update.json` из `main` сразу; без файлов в выпуске у них `HTTP 404`, а ежедневная проверка в этот день больше не повторится.
3. **Файл без заголовка** — пока есть пользователи 1.0.x–1.1.1, их проверка файла понимает только подпись OLE в начале.
4. **`sha256` считается от того же файла, что в выпуске** — поэтому файл в Assets — только `release\NeonCenterline.gms`.

Механизм обновления в деталях — [architecture.md](architecture.md#обновление).
