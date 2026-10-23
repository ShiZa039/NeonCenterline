# NeonCenterline

**RU** · [EN](#english)

Плагин для CorelDRAW: строит центральную линию по контурам букв для фрезеровки канала под гибкий неон. Без DLL, чистый VBA.

- **Моно** — для моноширинных (неоновых) шрифтов: прямые штрихи остаются прямыми, дуги — минимумом кривых, восьмёрки замкнуты.
- **Курсив** — для рукописных шрифтов (Segoe Script и т. п.): линия не выходит за контур там, где штрихи сливаются, и держится от края не ближе радиуса фрезы.
- Свободный конец линии ставится на D/2 от края контура (D — диаметр фрезы, задаётся в настройках).
- Результат — кривые на слое «Неон - центр» (красная волосяная линия).

## Установка

1. Скачайте последний релиз (файл `NeonCenterline-<версия>.zip`) и распакуйте.
2. Закройте CorelDRAW и запустите `install.cmd` (попросит права администратора — для папки Addons).
   Скрипт кладёт `NeonCenterline.gms` в пользовательскую папку макросов CorelDRAW и панель инструментов — в папку Addons.
3. Запустите CorelDRAW. Появится панель **NeonCenterline** (если нет: Окно → Панели инструментов → NeonCenterline).

Вручную: `NeonCenterline.gms` → `%APPDATA%\Corel\CorelDRAW Graphics Suite <версия>\Draw\GMS\`,
папку `addon\NeonCenterline` → `<папка CorelDRAW>\Programs64\Addons\`.
Команды можно поставить на любую панель и без Addons: Инструменты → Настройка → Команды → Макросы → `NeonCenterline.NeonUI.*`.

## Работа

Выделите контуры (кривые, текст, группы) и нажмите **Моно** или **Курсив**.
В **настройках** — диаметр фрезы, спрашивать ли его при каждом запуске, язык интерфейса (русский / English / как в Windows), ежедневная проверка обновлений.

## Обновления

Раз в день (после расчёта) плагин проверяет, нет ли новой версии на GitHub. Если есть — спросит, скачает, проверит контрольную сумму, и новая версия установится сама, когда вы закроете CorelDRAW. Старая версия сохраняется рядом как `NeonCenterline.gms.bak`. Проверить вручную: Настройки → «Проверить обновления».

## Лицензия

Бесплатно для любого использования, в том числе коммерческого (изготовление вывесок, фрезеровка на заказ). Продавать и перепаковывать плагин нельзя. Полный текст — [LICENSE.txt](LICENSE.txt).

Плагин бесплатный. Если он вам помогает — можно поблагодарить автора (ссылка в окне «О программе»).

---

<a name="english"></a>
# NeonCenterline (English)

A CorelDRAW plugin that builds a centerline from letter outlines for milling flex-neon channels. Pure VBA, no DLL.

- **Mono** — for monoline (neon) fonts: straight strokes stay straight, arcs use a minimum of curves, loops are closed.
- **Script** — for handwritten fonts (Segoe Script etc.): the line stays inside the outline where strokes merge and keeps at least the cutter radius from the edge.
- A free line end is placed D/2 from the outline edge (D is the cutter diameter, set in the settings).
- Output: curves on the layer "Неон - центр" (red hairline).

## Installation

1. Download the latest release (`NeonCenterline-<version>.zip`) and unzip it.
2. Close CorelDRAW and run `install.cmd` (it asks for administrator rights for the Addons folder).
   It copies `NeonCenterline.gms` into the CorelDRAW user macro folder and the toolbar into the Addons folder.
3. Start CorelDRAW. The **NeonCenterline** toolbar appears (if not: Window → Toolbars → NeonCenterline).

Manually: `NeonCenterline.gms` → `%APPDATA%\Corel\CorelDRAW Graphics Suite <version>\Draw\GMS\`,
the `addon\NeonCenterline` folder → `<CorelDRAW folder>\Programs64\Addons\`.
Without Addons you can still put the commands on any toolbar: Tools → Customization → Commands → Macros → `NeonCenterline.NeonUI.*`.

## Usage

Select outlines (curves, text, groups) and press **Mono** or **Script**.
**Settings**: cutter diameter, whether to ask for it on every run, interface language (Russian / English / as Windows), daily update check.

## Updates

Once a day (after a run) the plugin checks GitHub for a new version. If there is one, it asks, downloads it, verifies the checksum, and the new version is installed automatically when you close CorelDRAW. The previous version is kept as `NeonCenterline.gms.bak`. Manual check: Settings → "Check for updates".

## License

Free for any use, including commercial (sign making, milling for hire). Selling or repackaging the plugin is prohibited. Full text: [LICENSE.txt](LICENSE.txt).

The plugin is free. If it helps you, you can support the author (link in the "About" window).
