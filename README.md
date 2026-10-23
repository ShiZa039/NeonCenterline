# NeonCenterline

**RU** · [EN](#english)

Плагин для CorelDRAW: строит центральную линию по контурам букв для фрезеровки канала под гибкий неон. Без DLL, чистый VBA.

- **Моно** — для моноширинных (неоновых) шрифтов: прямые штрихи остаются прямыми, дуги — минимумом кривых, восьмёрки замкнуты.
- **Курсив** ⚠️ *в разработке, работает нестабильно* — для рукописных шрифтов (Segoe Script и т. п.): линия не выходит за контур там, где штрихи сливаются, и держится от края не ближе радиуса фрезы. На стыках букв возможны изломы и лишние куски — проверяйте результат перед резкой.
- Свободный конец линии ставится на D/2 от края контура (D — диаметр фрезы, задаётся в настройках).
- Результат — кривые на слое «Неон - центр» (красная волосяная линия).

## Установка

1. Скачайте последний релиз (файл `NeonCenterline-<версия>.zip`) и распакуйте.
2. Закройте CorelDRAW и запустите `install.cmd` (попросит права администратора — для папки Addons).
   Скрипт кладёт `NeonCenterline.gms` в пользовательскую папку макросов CorelDRAW и панель инструментов — в папку Addons.
3. Запустите CorelDRAW. Появится панель **«Неон (NeonCenterline)»** (если нет: Окно → Панели инструментов → «Неон (NeonCenterline)»).

Вручную: `NeonCenterline.gms` → `%APPDATA%\Corel\CorelDRAW Graphics Suite <версия>\Draw\GMS\`,
папку `addon\NeonCenterline` → `<папка CorelDRAW>\Programs64\Addons\`.
Если панель не появилась (например, установщик нашёл папку старой версии CorelDRAW) — запустите макрос `NeonCenterline.NeonAddon.InstallToolbar` (Инструменты → Макросы → Запуск макроса): он ставит панель в текущую версию CorelDRAW.
Команды можно поставить на любую панель и без Addons: Инструменты → Настройка → Команды → Макросы → `NeonCenterline.NeonUI.*`.

## Работа

Выделите контуры (кривые, текст, группы) и нажмите **Моно** или **Курсив**.

**Длина** — отдельная кнопка: считает длину выделенного. Контуры (длина под резку) и центральные линии плагина (длина неона, число кусков, самый длинный кусок) — отдельно; результат можно скопировать в буфер обмена.
**Выделить линии** — выделяет все центральные линии плагина на странице (например, чтобы сразу посчитать длину неона или убрать их).

### Тип неона и кратность реза
В настройках: **тип неона** — светодиодный (лента внутри: кусок неона = длина паза) или **неон 2.0** (силиконовый рассеиватель в пазу, лента отдельно), и **кратность реза** (10 / 25 / 50 мм или своя).
- Галочка **«Подгонять линии под кратность реза»**: после расчёта каждый кусок укорачивается до ближайшей меньшей кратной длины — только за счёт **свободных** концов (поровну с обоих или только с конца линии — настройка). Удлинять нельзя: конец уже на D/2 от края. Замкнутые куски и куски без свободных концов не меняются — остаток виден в отчёте.
- Кнопка **«Длина»** для центральных линий показывает, сколько неона (или ленты для неона 2.0) получается кратно резу и сколько уйдёт в остатки.
- Кнопка **«Подогнать под рез»** (`NeonCenterline.NeonCut.FitToCut`) — подогнать под кратность уже готовые выделенные линии.
В **настройках** — диаметр фрезы, спрашивать ли его при каждом запуске, язык интерфейса (русский / English / как в Windows), ежедневная проверка обновлений.

## Обновления

Раз в день (после расчёта) плагин проверяет, нет ли новой версии на GitHub. Если есть — спросит, скачает, проверит контрольную сумму, и новая версия установится сама, когда вы закроете CorelDRAW. Старая версия сохраняется рядом как `NeonCenterline.gms.bak`. Проверить вручную: Настройки → «Проверить обновления».

## Лицензия

Бесплатно для любого использования, в том числе коммерческого (изготовление вывесок, фрезеровка на заказ). Продавать и перепаковывать плагин нельзя. Полный текст — [LICENSE.txt](LICENSE.txt).

Плагин бесплатный. Если он вам помогает — можно поблагодарить автора: [DONATE.md](DONATE.md) (та же ссылка — в окне «О программе»).

---

<a name="english"></a>
# NeonCenterline (English)

A CorelDRAW plugin that builds a centerline from letter outlines for milling flex-neon channels. Pure VBA, no DLL.

- **Mono** — for monoline (neon) fonts: straight strokes stay straight, arcs use a minimum of curves, loops are closed.
- **Script** ⚠️ *work in progress, not stable yet* — for handwritten fonts (Segoe Script etc.): the line stays inside the outline where strokes merge and keeps at least the cutter radius from the edge. Kinks and stray pieces may appear at letter joints — check the result before cutting.
- A free line end is placed D/2 from the outline edge (D is the cutter diameter, set in the settings).
- Output: curves on the layer "Неон - центр" (red hairline).

## Installation

1. Download the latest release (`NeonCenterline-<version>.zip`) and unzip it.
2. Close CorelDRAW and run `install.cmd` (it asks for administrator rights for the Addons folder).
   It copies `NeonCenterline.gms` into the CorelDRAW user macro folder and the toolbar into the Addons folder.
3. Start CorelDRAW. The **"Неон (NeonCenterline)"** toolbar appears (if not: Window → Toolbars → "Неон (NeonCenterline)").

Manually: `NeonCenterline.gms` → `%APPDATA%\Corel\CorelDRAW Graphics Suite <version>\Draw\GMS\`,
the `addon\NeonCenterline` folder → `<CorelDRAW folder>\Programs64\Addons\`.
If the toolbar does not appear (e.g. the installer found an old CorelDRAW version folder), run the macro `NeonCenterline.NeonAddon.InstallToolbar` (Tools → Macros → Run Macro): it installs the toolbar into the current CorelDRAW version.
Without Addons you can still put the commands on any toolbar: Tools → Customization → Commands → Macros → `NeonCenterline.NeonUI.*`.

## Usage

Select outlines (curves, text, groups) and press **Mono** or **Script**.

**Length** — a separate button: measures the selection. Outlines (cutting length) and plugin centerlines (neon length, number of pieces, longest piece) are reported separately; the result can be copied to the clipboard.
**Select lines** — selects all plugin centerlines on the page (e.g. to measure the neon or remove them).

### Neon type and cut step
Settings: **neon type** — LED neon (strip inside: neon piece = groove length) or **Neon 2.0** (silicone diffuser in a groove, separate strip), and the **cut step** (10 / 25 / 50 mm or custom).
- **"Fit lines to the cut step"**: after a run every piece is shortened to the nearest smaller multiple — only at **free** ends (equally at both or only at the line end — a setting). Lines are never extended: a free end is already D/2 from the edge. Closed pieces and pieces without free ends are left as is; the remainder is shown in the report.
- The **Length** button shows how much neon (or strip, for Neon 2.0) fits in whole cut steps and how much goes to remainders.
- The **Fit to cut** button (`NeonCenterline.NeonCut.FitToCut`) fits already selected lines to the cut step.
**Settings**: cutter diameter, whether to ask for it on every run, interface language (Russian / English / as Windows), daily update check.

## Updates

Once a day (after a run) the plugin checks GitHub for a new version. If there is one, it asks, downloads it, verifies the checksum, and the new version is installed automatically when you close CorelDRAW. The previous version is kept as `NeonCenterline.gms.bak`. Manual check: Settings → "Check for updates".

## License

Free for any use, including commercial (sign making, milling for hire). Selling or repackaging the plugin is prohibited. Full text: [LICENSE.txt](LICENSE.txt).

The plugin is free. If it helps you, you can support the author: [DONATE.md](DONATE.md) (same link in the "About" window).
