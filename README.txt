NEONCENTERLINE

RU · EN

Плагин для CorelDRAW: строит центральную линию по контурам букв для фрезеровки канала под гибкий неон. Без DLL, чистый VBA.

- Моно — для моноширинных (неоновых) шрифтов: прямые штрихи остаются прямыми, дуги — минимумом кривых, восьмёрки замкнуты.
- Курсив ⚠️ в разработке, работает нестабильно — для рукописных шрифтов (Segoe Script и т. п.): линия не выходит за контур там, где штрихи сливаются, и держится от края не ближе радиуса фрезы. На стыках букв возможны изломы и лишние куски — проверяйте результат перед резкой.
- Свободный конец линии ставится на D/2 от края контура (D — диаметр фрезы, задаётся в настройках).
- Результат — кривые на слое «Неон - центр» (красная волосяная линия).

Установка
---------

1. Скачайте последний релиз (файл NeonCenterline-<версия>.zip) и распакуйте.
2. Закройте CorelDRAW и запустите install.cmd (попросит права администратора — для папки Addons).
   Скрипт кладёт NeonCenterline.gms в пользовательскую папку макросов CorelDRAW и панель инструментов — в папку Addons. Если установлено несколько версий CorelDRAW, он ставит плагин только в самую новую — для другой версии см. ручную установку ниже.
3. Запустите CorelDRAW. Появится панель «Неон (NeonCenterline)» (если нет: Окно → Панели инструментов → «Неон (NeonCenterline)»).

Ручная установка (если установщик не сработал, несколько версий CorelDRAW или старая версия)
--------------------------------------------------------------------------------------------

1. Файл плагина. Скопируйте NeonCenterline.gms в папку макросов пользователя:
   %APPDATA%\Corel\CorelDRAW Graphics Suite <версия>\Draw\GMS\
   (вставьте путь в адресную строку Проводника; <версия> — папка вашей версии, например 2026 или 27).
2. Панель инструментов. Скопируйте папку addon\NeonCenterline целиком (5 файлов) в папку дополнений CorelDRAW:
   C:\Program Files\Corel\CorelDRAW Graphics Suite <версия>\Programs64\Addons\
   (нужны права администратора; там же обычно лежат другие дополнения).
   Или, когда плагин уже загружен, запустите макрос NeonCenterline.NeonAddon.InstallToolbar (Инструменты → Макросы → Запуск макроса) — он сам положит панель в Addons той версии CorelDRAW, из которой запущен.
3. Полностью перезапустите CorelDRAW и включите панель: Окно → Панели инструментов → «Неон (NeonCenterline)».
4. Если панели нет или кнопки пустые (бывает в старых версиях), кнопки можно поставить вручную: Инструменты → Настройка (Ctrl+J) → Команды → категория «Макросы» → перетащить на любую панель команды NeonCenterline.NeonUI.RunMono, RunScript, ShowSettings, NeonCenterline.NeonTools.MeasureLength, SelectNeonLines, SimplifyContours, NeonCenterline.NeonCut.FitToCut. Значки — в папке icons архива (вкладка «Вид» → «Импорт»).

Тестовый шрифт font\SABNeonTest.ttf для работы плагина не нужен — он для пробы (моноширинный неоновый шрифт, на нём видно, как должен работать режим «Моно»). Установить: двойной щелчок по файлу → «Установить».

Совместимость: проверено в CorelDRAW 2026 (27). Должно работать в 2017–2025; в X7–X8 расчёт, скорее всего, работает, панель — возможно, только вручную. Работает или нет в вашей версии — напишите в Issues на GitHub.

Работа
------

Выделите контуры (кривые, текст, группы) и нажмите Моно или Курсив.

Длина — отдельная кнопка: считает длину выделенного. Контуры (длина под резку) и центральные линии плагина (длина неона, число кусков, самый длинный кусок) — отдельно; результат можно скопировать в буфер обмена.
Выделить линии — выделяет все центральные линии плагина на странице (например, чтобы сразу посчитать длину неона или убрать их).
Упростить контур (NeonCenterline.NeonTools.SimplifyContours) — меньше узлов у выделенных кривых: кривая не отходит от исходной больше допуска (окно перед запуском, как у центральных линий, и Настройки; по умолчанию 0.1 мм), острые углы остаются углами, гладкие места — гладкими. Кривые меняются на месте, отмена — Ctrl+Z. Полезно для трассированных логотипов и контуров из DXF/CAM (тысячи узлов → десятки), шрифты обычно сокращаются на треть. Текст и фигуры сначала переведите в кривые (Ctrl+Q).

Тип неона и кратность реза
--------------------------
В настройках: тип неона — светодиодный (лента внутри: кусок неона = длина паза) или неон 2.0 (силиконовый рассеиватель в пазу, лента отдельно), и кратность реза (10 / 25 / 50 мм или своя).
- Галочка «Подгонять линии под кратность реза»: после расчёта каждый кусок укорачивается до ближайшей меньшей кратной длины — только за счёт свободных концов (поровну с обоих или только с конца линии — настройка). Удлинять нельзя: конец уже на D/2 от края. Замкнутые куски и куски без свободных концов не меняются — остаток виден в отчёте.
- Кнопка «Длина» для центральных линий показывает, сколько неона (или ленты для неона 2.0) получается кратно резу и сколько уйдёт в остатки.
- Кнопка «Подогнать под рез» (NeonCenterline.NeonCut.FitToCut) — подогнать под кратность уже готовые выделенные линии.
В настройках — диаметр фрезы, спрашивать ли его при каждом запуске, допуск упрощения контура, язык интерфейса (русский / English / как в Windows), ежедневная проверка обновлений.

Обновления
----------

Раз в день (после расчёта) плагин проверяет, нет ли новой версии на GitHub. Если есть — спросит, скачает, проверит контрольную сумму, и новая версия установится сама, когда вы закроете CorelDRAW. Старая версия сохраняется рядом как NeonCenterline.gms.bak. Проверить вручную: Настройки → «Проверить обновления».

Лицензия
--------

Бесплатно для любого использования, в том числе коммерческого (изготовление вывесок, фрезеровка на заказ). Продавать и перепаковывать плагин нельзя. Полный текст — LICENSE.txt (LICENSE.txt).

Плагин бесплатный. Если он вам помогает — можно поблагодарить автора: DONATE.md (DONATE.md) (та же ссылка — в окне «О программе»).

======================================================================

NEONCENTERLINE (ENGLISH)

A CorelDRAW plugin that builds a centerline from letter outlines for milling flex-neon channels. Pure VBA, no DLL.

- Mono — for monoline (neon) fonts: straight strokes stay straight, arcs use a minimum of curves, loops are closed.
- Script ⚠️ work in progress, not stable yet — for handwritten fonts (Segoe Script etc.): the line stays inside the outline where strokes merge and keeps at least the cutter radius from the edge. Kinks and stray pieces may appear at letter joints — check the result before cutting.
- A free line end is placed D/2 from the outline edge (D is the cutter diameter, set in the settings).
- Output: curves on the layer "Неон - центр" (red hairline).

Installation
------------

1. Download the latest release (NeonCenterline-<version>.zip) and unzip it.
2. Close CorelDRAW and run install.cmd (it asks for administrator rights for the Addons folder).
   It copies NeonCenterline.gms into the CorelDRAW user macro folder and the toolbar into the Addons folder. With several CorelDRAW versions installed it installs into the newest one only — for another version see manual installation below.
3. Start CorelDRAW. The "Неон (NeonCenterline)" toolbar appears (if not: Window → Toolbars → "Неон (NeonCenterline)").

Manual installation (if the installer did not work, several CorelDRAW versions, or an older version)
----------------------------------------------------------------------------------------------------

1. Plugin file. Copy NeonCenterline.gms to the user macro folder:
   %APPDATA%\Corel\CorelDRAW Graphics Suite <version>\Draw\GMS\
   (paste the path into the Explorer address bar; <version> is the folder of your version, e.g. 2026 or 27).
2. Toolbar. Copy the whole addon\NeonCenterline folder (5 files) to the CorelDRAW add-ons folder:
   C:\Program Files\Corel\CorelDRAW Graphics Suite <version>\Programs64\Addons\
   (administrator rights required; other add-ons usually live there too).
   Or, once the plugin is loaded, run the macro NeonCenterline.NeonAddon.InstallToolbar (Tools → Macros → Run Macro): it puts the toolbar into Addons of the CorelDRAW version it runs in.
3. Fully restart CorelDRAW and turn the toolbar on: Window → Toolbars → "Неон (NeonCenterline)".
4. If there is no toolbar or the buttons are blank (possible in older versions), add the buttons manually: Tools → Customization (Ctrl+J) → Commands → category "Macros" → drag to any toolbar NeonCenterline.NeonUI.RunMono, RunScript, ShowSettings, NeonCenterline.NeonTools.MeasureLength, SelectNeonLines, SimplifyContours, NeonCenterline.NeonCut.FitToCut. Icons are in the icons folder of the archive (Appearance tab → Import).

Test font font\SABNeonTest.ttf is not required by the plugin — it is for trying it out (a monoline neon font that shows how Mono mode is meant to work). To install: double-click the file → Install.

Compatibility: tested in CorelDRAW 2026 (27). Should work in 2017–2025; in X7–X8 the computation most likely works, the toolbar may need manual setup. Please report in GitHub Issues whether it works in your version.

Usage
-----

Select outlines (curves, text, groups) and press Mono or Script.

Length — a separate button: measures the selection. Outlines (cutting length) and plugin centerlines (neon length, number of pieces, longest piece) are reported separately; the result can be copied to the clipboard.
Select lines — selects all plugin centerlines on the page (e.g. to measure the neon or remove them).
Simplify outline (NeonCenterline.NeonTools.SimplifyContours) — fewer nodes on the selected curves: the curve deviates from the original by no more than the tolerance (a window before the run, like for centerlines, and Settings; 0.1 mm by default), sharp corners stay sharp, smooth parts stay smooth. Curves are changed in place, undo with Ctrl+Z. Useful for traced logos and outlines from DXF/CAM (thousands of nodes become dozens); font outlines usually lose about a third of their nodes. Convert text and shapes to curves first (Ctrl+Q).

Neon type and cut step
----------------------
Settings: neon type — LED neon (strip inside: neon piece = groove length) or Neon 2.0 (silicone diffuser in a groove, separate strip), and the cut step (10 / 25 / 50 mm or custom).
- "Fit lines to the cut step": after a run every piece is shortened to the nearest smaller multiple — only at free ends (equally at both or only at the line end — a setting). Lines are never extended: a free end is already D/2 from the edge. Closed pieces and pieces without free ends are left as is; the remainder is shown in the report.
- The Length button shows how much neon (or strip, for Neon 2.0) fits in whole cut steps and how much goes to remainders.
- The Fit to cut button (NeonCenterline.NeonCut.FitToCut) fits already selected lines to the cut step.
Settings: cutter diameter, whether to ask for it on every run, outline simplify tolerance, interface language (Russian / English / as Windows), daily update check.

Updates
-------

Once a day (after a run) the plugin checks GitHub for a new version. If there is one, it asks, downloads it, verifies the checksum, and the new version is installed automatically when you close CorelDRAW. The previous version is kept as NeonCenterline.gms.bak. Manual check: Settings → "Check for updates".

License
-------

Free for any use, including commercial (sign making, milling for hire). Selling or repackaging the plugin is prohibited. Full text: LICENSE.txt (LICENSE.txt).

The plugin is free. If it helps you, you can support the author: DONATE.md (DONATE.md) (same link in the "About" window).
