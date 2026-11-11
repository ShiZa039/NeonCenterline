NEONCENTERLINE

RU · EN

Плагин для CorelDRAW: строит центральную линию по контурам букв для фрезеровки канала под гибкий неон. Без DLL, чистый VBA.

- Моно — для моноширинных (неоновых) шрифтов: прямые штрихи остаются прямыми, дуги — минимумом кривых, восьмёрки замкнуты.
- Курсив ⚠️ бета — для рукописных шрифтов (Segoe Script и т. п.): по умолчанию линия идёт по ходу пера, как в шрифте-образце, сквозь развилки и до кончиков штрихов. Рукописные шрифты очень разные — на стыках букв возможны изломы и лишние куски, проверяйте линии перед резкой.
- Свободный конец линии ставится с отступом от края контура (по умолчанию D/2, D — диаметр фрезы; задаётся в мм в настройках).
- Результат — кривые на слое «Неон - центр» (красная волосяная линия).
- Для производства: очистить контур (убрать наложения и петли в шрифтах и фигурах), разбить на буквы (дырки остаются в своих буквах), упростить контур (меньше узлов), раскладка деталей на листы (бета: по контуру — детали входят друг в друга, с поворотом или без — или по габаритам; по цветам, длина листа — заданная или по раскладке), одинаковые — в группы, люверсы по краю баннера, лапша (полосы), длина линий, подгонка под рез.

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
4. Если панели нет или кнопки пустые (бывает в старых версиях), кнопки можно поставить вручную: Инструменты → Настройка (Ctrl+J) → Команды → категория «Макросы» → перетащить на любую панель команды NeonCenterline.NeonUI.RunMono, RunScript, ShowSettings, NeonCenterline.NeonTools.MeasureLength, SelectNeonLines, SimplifyContours, CleanContours, SplitLetters, NestSheets, GroupSame, MakeNoodles, PlaceGrommets, NeonCenterline.NeonCut.FitToCut. Значки — в папке icons архива (вкладка «Вид» → «Импорт»).
5. Если после обновления плагина на панели старые кнопки — запустите макрос NeonCenterline.NeonUI.ResetToolbar (Инструменты → Макросы → Запуск макроса): он сбросит только панель «Неон» к новому набору кнопок.

Тестовый шрифт font\SABNeonTest.ttf для работы плагина не нужен — он для пробы (моноширинный неоновый шрифт, на нём видно, как должен работать режим «Моно»). Установить: двойной щелчок по файлу → «Установить».

Совместимость: проверено в CorelDRAW 2026 (27). Должно работать в 2017–2025; в X7–X8 расчёт, скорее всего, работает, панель — возможно, только вручную. Работает или нет в вашей версии — напишите в Issues на GitHub.

Работа
------

Выделите контуры (кривые, текст, группы) и нажмите Моно или Курсив.

На панели «Неон»: Моно, Курсив | Контур ▾, Линии ▾, Раскладка ▾ | Настройки ▾. Кнопки со стрелкой открывают список: в «Контур» — упростить контур, очистить контур, разбить на буквы; в «Раскладка» — раскладка на листы, одинаковые — в группы, лапша, люверсы; в «Линии» — длина, выделить линии, подогнать под рез; в «Настройки» — настройки, проверить обновления, о программе.

После инструментов, которые меняют документ (очистить, упростить, разбить, раскладка, одинаковые, лапша, люверсы, кратность реза), показывается отчёт с кнопками ОК — оставить и Отмена — отменить всё сделанное одним шагом (как Ctrl+Z; Esc и крестик — тоже отмена). Если ничего не поменялось — только ОК.

Пока идут долгие инструменты (центральные линии, раскладка, очистить, упростить, разбить, одинаковые), открыто окно хода работы: этап, полоса, сколько из скольких и сколько секунд прошло — видно, что плагин работает, а не CorelDRAW завис. Пока оно открыто, документ трогать нельзя; закроется само.

Оставить оригинал: у раскладки — галочка в её окне (на листы встают копии деталей, детали остаются на месте); у упрощения — в его окне; у очистки, разбивки и упрощения — в настройках плагина («Оставлять оригинал: очистка, разбивка, упрощение»). Результат — копией поверх оригинала; объекты, которые не поменялись, копий не получают.

Длина (в «Линии ▾») — считает длину выделенного. Контуры (длина под резку) и центральные линии плагина (длина неона, число кусков, самый длинный кусок) — отдельно.
Выделить линии (в «Линии ▾») — выделяет все центральные линии плагина на странице (например, чтобы сразу посчитать длину неона или убрать их).
Упростить контур (в «Контур ▾», NeonCenterline.NeonTools.SimplifyContours) — меньше узлов у выделенных кривых: кривая не отходит от исходной больше допуска (окно перед запуском, как у центральных линий, и Настройки; по умолчанию 0.1 мм), острые углы остаются углами, гладкие места — гладкими. Кривые меняются на месте, отмена — Ctrl+Z. Полезно для трассированных логотипов и контуров из DXF/CAM (тысячи узлов → десятки), шрифты обычно сокращаются на треть. Текст и фигуры сначала переведите в кривые (Ctrl+Q).
Очистить контур (в «Контур ▾», NeonCenterline.NeonTools.CleanContours) — убирает наложения и петли: буквы, собранные из перекрывающихся частей (стойка поверх дуг), налезающие друг на друга буквы, контуры, которые перехлёстывают сами себя (цифры, у которых дырка — петля одного контура), «ушки» и ласточкины хвосты у фигур после «Обводка в объект» и сдвига контура. Каждый выделенный объект становится одной кривой: внутренние линии уходят, дырки остаются, мелкие петли и лоскуты (меньше 0.3 мм² или тоньше 0.15 мм) — убираются. Форма не меняется: новый контур собирается из исходных кривых, разрезанных точно в точках пересечения. Оригинал заменяется (одна операция, отмена — Ctrl+Z); текст переводится в кривые сам; прямоугольники и эллипсы пропускаются. Отдельные объекты, которые надо слить, сначала объедините (Ctrl+L).
Разбить на буквы (в «Контур ▾», NeonCenterline.NeonTools.SplitLetters) — текст в кривых или объединённые буквы (одна кривая из многих контуров) — по объекту на букву или цифру, дырки остаются в своих буквах: не нужно, как после «Разъединить» (Ctrl+K), сидеть и вычитать дырку из каждой буквы. Какой контур — дырка, решается вложенностью: контур внутри чётного числа других — наружный контур буквы, внутри нечётного — дырка ближайшего наружного. Точки и акценты остаются со своей буквой (Ё, Й, i, ї, !, ?, ÷): маленький кусок над буквой или под ней, близко по вертикали, идёт к ней. Отдельными остаются куски одного размера (двоеточие, «=», «ы» — две части одной высоты), стоящие рядом с буквой (точка после слова, кольца «%») и детали внутри дырок (буква в «®», точка в нуле моноширинного шрифта); буквы в дырке рамки к рамке не липнут. Каждая буква — копия исходной кривой (заливка, обводка, слой те же), буквы выделены — можно сразу «Одинаковые — в группы» или «Раскладка на листы». Текст переводится в кривые сам; если буквы налезают друг на друга — сначала «Очистить контур». Отмена — Ctrl+Z.
Люверсы (в «Раскладка ▾», NeonCenterline.NeonTools.PlaceGrommets) — точки под люверсы по краю выделенного баннера или подложки (прямоугольник, скруглённый прямоугольник, любой замкнутый контур): в каждом углу и равномерно по сторонам, не реже заданного расстояния; у круга и овала — равномерно по контуру. В окне перед запуском: расстояние от края до центра люверса, расстояние между люверсами (не больше), диаметр и цвет точки (из списка или любой через окно выбора цвета CorelDRAW), усиление углов — ещё по точке на обе стороны каждого угла на заданном расстоянии (по умолчанию 50 мм). Точки — на слое «Люверсы», по группе на объект; количество — в отчёте (можно скопировать).
Раскладка на листы (бета, в разработке; в «Раскладка ▾», NeonCenterline.NeonTools.NestSheets) — раскладывает выделенные детали на листы (оргстекло, ПВХ, композит): ни одна деталь не подходит к краю листа ближе поля и к соседней — ближе зазора (для фрезы зазор — не меньше её диаметра). Способ раскладки: по контуру (по умолчанию) — детали ставятся по своей форме, а не по габариту: входят друг в друга (цифры и буквы «в обнимку»), мелкие — в просветы между крупными, а с галочкой «матрёшкой» (по умолчанию включена) — ещё и в дырки букв и внутрь рамок; без неё дырки деталей считаются занятыми; можно разрешить поворот на 180° или повороты по 90° — только если материалу всё равно (без текстуры, направления волокон или плёнки). По габаритам — быстрее: прямоугольниками (MaxRects: мелкие — в щели между крупными). По контуру деталь переводится в сетку клеток (около 0.7 мм на листе 900×600, около 2.5 мм на 3050×2050) с запасом: зазор соблюдается всегда, но из-за клетки детали могут стоять чуть свободнее заданного; на сотнях деталей и больших листах раскладка по контуру идёт дольше. Контур детали должен быть замкнутым (у незамкнутого нет «внутри», и туда может встать другая деталь; концы, сошедшиеся в точку, считаются замкнутыми). В окне: длина (по горизонтали) и ширина (по вертикали) листа — любые: стандартный лист, стол лазера 900×600, заготовка (готовые размеры — в списке, свои запоминаются в нём же), поле от края, зазор, «каждый цвет — на свои листы» (цвет детали — заливка, без заливки — обводка) и «длина по раскладке»: ширина листа задана, а длина — сколько заняли детали плюс поле (вверх до 10 мм), от заданного минимума до длины листа; что не влезло — на следующий лист. Из нескольких вариантов (порядок деталей, направление заполнения) берётся тот, где меньше листов, а при равенстве — короче последний. Заполнение листа: столбцами — сверху вниз, затем вправо (по умолчанию), рядами — слева направо, затем вниз, или как плотнее (авто). Листы рисуются справа от выделения, по пять в ряд, каждый цвет — своими рядами; каждый лист — одна группа на слое деталей: рамка, подпись (лист i из n, размер, цвет, число деталей, заполнение — по контуру по площади деталей, по габаритам по габаритам; можно отключить) и детали — печатается и экспортируется вместе с ними. Деталь — выделенный объект (группа — целиком); группы «Одинаковые» и листы прежней раскладки раскладываются поштучно, одна выделенная группа (импорт SVG/DXF) — тоже. Повторный запуск по разложенному — новая раскладка, старые рамки из выделения убираются. Отчёт — по листам и цветам. Проверяйте раскладку перед резкой.
Одинаковые — в группы (в «Раскладка ▾», NeonCenterline.NeonTools.GroupSame) — находит среди выделенного объекты одной формы (контур совпадает с точностью до 0.1 мм, без поворота и зеркала; сравнивается сам контур, поэтому число и расстановка узлов, начало и направление обхода не важны — одна и та же цифра из разных переводов текста в кривые собирается вместе) и одного цвета и собирает их в группы «Одинаковые: N шт · Ш × В мм». В отчёте — ведомость форм по убыванию числа штук (можно скопировать).
Лапша (в «Раскладка ▾», NeonCenterline.NeonTools.MakeNoodles) — нарезка полос для плоттера, лазера, фрезера: задаются длина и ширина лапши, количество полос и хвосты — выход реза за край, отдельно для горизонтальных резов (за отрезы) и вертикальных отрезов (за крайние горизонтальные): лазеру и плоттеру хватает 1–2 мм, фрезе — 2–5 мм запаса. Получается N + 1 горизонтальных резов через ширину полосы и два вертикальных отреза на концах. Для фрезы — галочка «Фреза» и её диаметр: фреза режет по центральной линии и съедает по радиусу с каждой стороны, поэтому горизонтальные резы идут через ширину + диаметр, а отрезы стоят на радиус за концами полос — после фрезы полосы выходят ровно заданного размера; хвосты отсчитываются от линий резов. Порядок объектов — порядок резки: горизонтальные сверху вниз зигзагом (первый — слева направо, дальше направления чередуются — станок не ходит вхолостую), вертикальные отрезы — в самом конце списка, тоже по ходу (сначала на той стороне, где кончился последний горизонтальный). Всё — одной группой в центре страницы.

Тип неона и кратность реза
--------------------------
В настройках: тип неона — светодиодный (лента внутри: кусок неона = длина паза) или неон 2.0 (силиконовый рассеиватель в пазу, лента отдельно), и кратность реза (10 / 25 / 50 мм или своя).
- Галочка «Подгонять линии под кратность реза»: после расчёта каждый кусок укорачивается до ближайшей меньшей кратной длины — только за счёт свободных концов (поровну с обоих или только с конца линии — настройка). Удлинять нельзя: конец уже на D/2 от края. Замкнутые куски и куски без свободных концов не меняются — остаток виден в отчёте.
- Кнопка «Длина» для центральных линий показывает, сколько неона (или ленты для неона 2.0) получается кратно резу и сколько уйдёт в остатки.
- Команда «Подогнать под рез» в «Линии ▾» (NeonCenterline.NeonCut.FitToCut) — подогнать под кратность уже готовые выделенные линии.
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
- Script ⚠️ beta — for handwritten fonts (Segoe Script etc.): by default the line follows the pen path of the reference font, through junctions and up to the stroke tips. Handwritten fonts differ a lot — kinks and stray pieces may appear at letter joints, check the lines before cutting.
- A free line end is inset from the outline edge (D/2 by default, D is the cutter diameter; set in mm in the settings).
- Output: curves on the layer "Неон - центр" (red hairline).
- For production: clean outline (remove overlaps and loops in fonts and shapes), split into letters (holes stay in their letters), simplify outline (fewer nodes), nest parts on sheets (beta: by outline — parts fit into each other, with or without rotation — or by bounding boxes; by colour, sheet length fixed or to fit the parts), group identical, grommets along a banner edge, strips, line length, fit to cut.

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
4. If there is no toolbar or the buttons are blank (possible in older versions), add the buttons manually: Tools → Customization (Ctrl+J) → Commands → category "Macros" → drag to any toolbar NeonCenterline.NeonUI.RunMono, RunScript, ShowSettings, NeonCenterline.NeonTools.MeasureLength, SelectNeonLines, SimplifyContours, CleanContours, SplitLetters, NestSheets, GroupSame, MakeNoodles, PlaceGrommets, NeonCenterline.NeonCut.FitToCut. Icons are in the icons folder of the archive (Appearance tab → Import).
5. If the toolbar still shows the old buttons after a plugin update, run the macro NeonCenterline.NeonUI.ResetToolbar (Tools → Macros → Run Macro): it resets only the "Неон" toolbar to the new button set.

Test font font\SABNeonTest.ttf is not required by the plugin — it is for trying it out (a monoline neon font that shows how Mono mode is meant to work). To install: double-click the file → Install.

Compatibility: tested in CorelDRAW 2026 (27). Should work in 2017–2025; in X7–X8 the computation most likely works, the toolbar may need manual setup. Please report in GitHub Issues whether it works in your version.

Usage
-----

Select outlines (curves, text, groups) and press Mono or Script.

The "Неон" toolbar: Mono, Script | Outline ▾, Lines ▾, Layout ▾ | Settings ▾. Buttons with an arrow open a list: Outline — simplify outline, clean outline, split into letters; Layout — nest on sheets, group identical, strips, grommets; Lines — length, select lines, fit to cut; Settings — settings, check for updates, about.

After the tools that change the document (clean, simplify, split, nest, group identical, strips, grommets, cut step), a report window shows OK to keep the result and Cancel to undo all of it in one step (like Ctrl+Z; Esc and the close button also cancel). If nothing changed, there is only OK.

While the long tools run (centerlines, nesting, clean, simplify, split, group identical), a progress window shows the stage, a bar, how many of how many and the seconds elapsed — you can see the plugin is working and CorelDRAW has not hung. The document cannot be touched while it is open; it closes by itself.

Keep the original: for nesting — a box in its window (copies of the parts go onto the sheets, the parts stay in place); for simplify — in its window; for clean, split and simplify — in the plugin settings ("Keep the original: clean, split, simplify"). The result is a copy on top of the original; objects that did not change get no copy.

Length (in Lines ▾) — measures the selection. Outlines (cutting length) and plugin centerlines (neon length, number of pieces, longest piece) are reported separately.
Select lines (in Lines ▾) — selects all plugin centerlines on the page (e.g. to measure the neon or remove them).
Simplify outline (in Outline ▾, NeonCenterline.NeonTools.SimplifyContours) — fewer nodes on the selected curves: the curve deviates from the original by no more than the tolerance (a window before the run, like for centerlines, and Settings; 0.1 mm by default), sharp corners stay sharp, smooth parts stay smooth. Curves are changed in place, undo with Ctrl+Z. Useful for traced logos and outlines from DXF/CAM (thousands of nodes become dozens); font outlines usually lose about a third of their nodes. Convert text and shapes to curves first (Ctrl+Q).
Clean outline (in Outline ▾, NeonCenterline.NeonTools.CleanContours) — removes overlaps and loops: letters built from overlapping parts (a stem lying over arcs), letters running into each other, outlines that loop over themselves (digits whose hole is a loop of one outline), ears and swallowtails of shapes after "Convert outline to object" and offsetting. Each selected object becomes one curve: inner lines go, holes stay, tiny loops and slivers (under 0.3 mm² or thinner than 0.15 mm) are removed. The shape does not change: the new outline is built from the original curves cut exactly at the crossings. The original is replaced (one step, undo with Ctrl+Z); text is converted to curves automatically; rectangles and ellipses are skipped. Combine separate objects that should merge first (Ctrl+L).
Split into letters (in Outline ▾, NeonCenterline.NeonTools.SplitLetters) — text converted to curves or combined letters (one curve of many outlines) becomes one object per letter or digit, with the holes kept in their letters: no more Break Apart (Ctrl+K) and then trimming every hole out of its letter by hand. Nesting decides what is a hole: an outline inside an even number of others is a letter's outer outline, inside an odd number — a hole of the nearest outer one. Dots and accents stay with their letter (Ё, Й, i, ї, !, ?, ÷): a small piece above or below a letter and close to it goes with it. Pieces of the same size (colon, "=", the two parts of "ы"), pieces next to a letter (a full stop, the rings of "%") and parts inside holes (the letter in ®, the dot in a monospace zero) stay separate; letters inside a frame's hole do not stick to the frame. Each letter is a copy of the original curve (same fill, outline and layer), and the letters stay selected, ready for Group identical or Nest on sheets. Text is converted to curves automatically; if letters overlap, run Clean outline first. Undo with Ctrl+Z.
Grommets (in Layout ▾, NeonCenterline.NeonTools.PlaceGrommets) — points for grommets along the edge of the selected banner or backing (rectangle, rounded rectangle, any closed outline): one in every corner and evenly along the sides, never further apart than the given distance; circles and ovals get them evenly all round. The window before the run: edge to grommet centre, distance between grommets (maximum), point diameter and colour (from the list or any via the CorelDRAW colour dialog), corner reinforcement — one more point on both sides of every corner at the given distance (50 mm by default). Points go to the "Grommets" layer, one group per object; the count is in the report (can be copied).
Nest on sheets (beta, work in progress; in Layout ▾, NeonCenterline.NeonTools.NestSheets) — lays the selected parts out on sheets (acrylic, PVC, composite): no part comes closer than the margin to the sheet edge or the gap to another part (for a router, use a gap of at least the cutter diameter). Method: by outline (default) — parts are placed by their shape, not their bounding box: they fit into each other (interlocking digits and letters), small ones go into the gaps between big ones, and with the "nested" box (on by default) also into letter holes and inside frames; without it the holes of parts count as taken; you can allow rotation by 180° or in 90° steps — only if the material does not care (no texture, grain or film direction). By bounding boxes is faster: rectangles (MaxRects: small ones fill the gaps between big ones). By outline, a part becomes a grid of cells (about 0.7 mm on a 900×600 sheet, about 2.5 mm on 3050×2050) with a safety allowance: the gap is always kept, but because of the cells parts may sit a little further apart than set; with hundreds of parts on large sheets nesting by outline takes longer. A part's outline must be closed (an open one has no inside, and another part may go there; ends that meet count as closed). The window: sheet length (horizontal) and width (vertical) — any: a standard sheet, a 900×600 laser bed, an offcut (presets are in the list, your own sizes are kept there too), edge margin, gap, "each colour on its own sheets" (a part's colour is its fill, or its outline if it has no fill) and "length to fit the parts": the sheet width is fixed and the length is what the parts take plus the margin (rounded up to 10 mm), between the given minimum and the sheet length; what does not fit goes to the next sheet. Of several variants (part order, fill direction) the one with the fewest sheets wins, then the one with the shortest last sheet. Sheet fill: in columns — top to bottom, then right (default), in rows — left to right, then down, or as tight as possible (auto). Sheets are drawn to the right of the selection, five in a row, each colour in its own rows; each sheet is one group on the parts' layer: the frame, the label (sheet i of n, size, colour, part count, fill — by part area when nesting by outline, by bounding boxes otherwise; can be turned off) and the parts — printed and exported with them. A part is a selected object (a group as a whole); "Identical" groups and sheets of an earlier layout are nested one by one, and so is a single selected group (an SVG/DXF import). Running it again on a layout makes a new one and removes the old frames in the selection. The report lists sheets and colours. Check the layout before cutting.
Group identical (in Layout ▾, NeonCenterline.NeonTools.GroupSame) — finds selected objects of the same shape (the outlines match within 0.1 mm, no rotation or mirroring; the outline itself is compared, so the number and placement of nodes, the start point and direction do not matter — the same digit from different text-to-curves conversions is grouped together) and the same colour and groups them as "Identical: N pcs · W × H mm". The report lists the shapes by count (can be copied).
Strips (in Layout ▾, NeonCenterline.NeonTools.MakeNoodles) — cut lines for strips on a plotter, laser or router: strip length and width, number of strips and the tails — how far the cuts run past the edge, separately for horizontal cuts (past the end cuts) and vertical end cuts (past the outer horizontal ones): a laser or plotter needs 1–2 mm, a router 2–5 mm of margin. You get N + 1 horizontal cuts one strip width apart and two vertical end cuts. For a router tick "Router" and set the cutter diameter: the router cuts along the centre line and eats a radius on each side, so horizontal cuts are spaced by width + diameter and the end cuts sit a radius past the strip ends — after the router the strips come out exactly the size you set; tails are measured from the cut lines. The object order is the cutting order: horizontal cuts top to bottom in a zigzag (the first left to right, then alternating — no idle runs), the vertical end cuts at the very end of the list, also along the way (first on the side where the last horizontal cut ended). One group in the middle of the page.

Neon type and cut step
----------------------
Settings: neon type — LED neon (strip inside: neon piece = groove length) or Neon 2.0 (silicone diffuser in a groove, separate strip), and the cut step (10 / 25 / 50 mm or custom).
- "Fit lines to the cut step": after a run every piece is shortened to the nearest smaller multiple — only at free ends (equally at both or only at the line end — a setting). Lines are never extended: a free end is already D/2 from the edge. Closed pieces and pieces without free ends are left as is; the remainder is shown in the report.
- The Length button shows how much neon (or strip, for Neon 2.0) fits in whole cut steps and how much goes to remainders.
- The Fit to cut command in Lines ▾ (NeonCenterline.NeonCut.FitToCut) fits already selected lines to the cut step.
Settings: cutter diameter, whether to ask for it on every run, outline simplify tolerance, interface language (Russian / English / as Windows), daily update check.

Updates
-------

Once a day (after a run) the plugin checks GitHub for a new version. If there is one, it asks, downloads it, verifies the checksum, and the new version is installed automatically when you close CorelDRAW. The previous version is kept as NeonCenterline.gms.bak. Manual check: Settings → "Check for updates".

License
-------

Free for any use, including commercial (sign making, milling for hire). Selling or repackaging the plugin is prohibited. Full text: LICENSE.txt (LICENSE.txt).

The plugin is free. If it helps you, you can support the author: DONATE.md (DONATE.md) (same link in the "About" window).
