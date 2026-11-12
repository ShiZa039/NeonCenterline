NEONCENTERLINE

RU · EN

Плагин для CorelDRAW для рекламного производства. Главное — центральная линия по контурам букв для фрезеровки канала под гибкий неон; вокруг неё — инструменты подготовки к резке: очистка и упрощение контуров, разбивка на буквы, раскладка деталей на листы, группировка одинаковых, лапша, люверсы, длина и кратность реза. Без DLL, чистый VBA.

Что умеет
---------

- Центральные линии — «Моно» для моноширинных (неоновых) шрифтов и «Курсив» (бета) для рукописных. Результат — кривые на слое «Неон - центр».
- Контур — очистить (убрать наложения и петли), упростить (меньше узлов), разбить на буквы (дырки остаются в своих буквах).
- Раскладка — детали на листы по контуру или по габаритам (бета), одинаковые — в группы, лапша (полосы), люверсы по краю баннера.
- Линии — длина контуров и линий, выделить центральные линии, подогнать линии под кратность реза.
- У всех инструментов — окно хода работы (видно, что плагин работает, а не CorelDRAW завис) и отчёт с кнопками «ОК» / «Отмена» (отменить всё сделанное одним шагом). У очистки, упрощения, разбивки и раскладки можно оставить оригинал.
- Интерфейс на русском и английском, автообновление с GitHub.

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
   (нужны права администратора). Или, когда плагин уже загружен, запустите макрос NeonCenterline.NeonAddon.InstallToolbar (Инструменты → Макросы → Запуск макроса) — он сам положит панель в Addons той версии CorelDRAW, из которой запущен.
3. Полностью перезапустите CorelDRAW и включите панель: Окно → Панели инструментов → «Неон (NeonCenterline)».
4. Если панели нет или кнопки пустые (бывает в старых версиях), кнопки можно поставить вручную: Инструменты → Настройка (Ctrl+J) → Команды → категория «Макросы» → перетащить на любую панель команды NeonCenterline.NeonUI.RunMono, RunScript, ShowSettings, CheckUpdates, ShowAbout, NeonCenterline.NeonTools.SimplifyContours, CleanContours, SplitLetters, NestSheets, GroupSame, MakeNoodles, PlaceGrommets, MeasureLength, SelectNeonLines, NeonCenterline.NeonCut.FitToCut. Значки — в папке icons архива (вкладка «Вид» → «Импорт»).
5. Если после обновления плагина на панели старые кнопки — запустите макрос NeonCenterline.NeonUI.ResetToolbar (Инструменты → Макросы → Запуск макроса): он сбросит только панель «Неон» к новому набору кнопок.

Тестовый шрифт font\SABNeonTest.ttf для работы плагина не нужен — он для пробы (моноширинный неоновый шрифт, на нём видно, как должен работать режим «Моно»). Установить: двойной щелчок по файлу → «Установить».

Совместимость: проверено в CorelDRAW 2026 (27). Должно работать в 2017–2025; в X7–X8 расчёт, скорее всего, работает, панель — возможно, только вручную. Работает или нет в вашей версии — напишите в Issues на GitHub.

Панель «Неон»
-------------

Моно, Курсив | Контур ▾, Раскладка ▾, Линии ▾ | Настройки и помощь ▾. Кнопки со стрелкой открывают список:

| Группа | Кнопки |
|---|---|
| Контур ▾ | Упростить контур, Очистить контур, Разбить на буквы |
| Раскладка ▾ | Раскладка на листы (бета), Одинаковые — в группы, Лапша (полосы), Люверсы |
| Линии ▾ | Длина контуров и линий, Выделить центральные линии, Подогнать под кратность реза |
| Настройки и помощь ▾ | Настройки, Проверить обновления, О программе |

После обновления плагина, если поменялся состав кнопок, он один раз предложит обновить файлы панели (права администратора); после перезапуска CorelDRAW панель «Неон» перестроится сама, остальное рабочее пространство не трогается.

Общее для всех инструментов
---------------------------

- Окно хода работы. Пока инструмент работает, открыто окно: этап, полоса, «N из M» и сколько секунд прошло — видно, что плагин работает, а не CorelDRAW завис. Пока оно открыто, документ трогать нельзя; закрывается само.
- Отчёт с кнопками «ОК» и «Отмена». После инструментов, которые меняют документ, — отчёт: ОК — оставить, Отмена — отменить всё сделанное одним шагом (как Ctrl+Z; Esc и крестик — тоже отмена). Если ничего не поменялось — только ОК. После центральных линий — только ОК (отмена — Ctrl+Z, как обычно).
- Оставить оригинал. У раскладки — галочка в её окне (на листы встают копии деталей, детали остаются на месте); у упрощения — в его окне; у очистки, разбивки и упрощения — в Настройках («Оставлять оригинал: очистка, разбивка, упрощение»). Результат — копией поверх оригинала; объекты, которые не поменялись, копий не получают; текст в кривые переводится тоже в копии.
- Очистка и разбивка на буквы переводят текст в кривые сами; упрощение работает только с кривыми. Группы обходятся насквозь — кроме раскладки и «Одинаковых», где группа — одна деталь.

Центральные линии: «Моно» и «Курсив»
------------------------------------

Выделите контуры (кривые, текст, группы) и нажмите Моно или Курсив.

- Моно — для моноширинных (неоновых) шрифтов: прямые штрихи остаются прямыми, дуги — минимумом кривых, восьмёрки замкнуты.
- Курсив ⚠️ бета — для рукописных шрифтов (Segoe Script и т. п.). По умолчанию — шрифт как образец: линия идёт по ходу пера, сквозь развилки и до кончиков штрихов, высота букв как в оригинале. Рукописные шрифты очень разные — на стыках букв возможны изломы и лишние куски, проверяйте линии перед резкой. Если буквы налезают друг на друга или собраны из перекрывающихся частей — сначала «Очистить контур».
- Результат — кривые «Неон: центр N» на слое «Неон - центр» (красная волосяная линия). По этому имени плагин узнаёт свои линии («Длина», «Выделить линии», подгонка).

Окно перед расчётом (можно отключить галочкой «Спрашивать параметры при каждом запуске» — тогда берутся значения из Настроек):
- диаметр фрезы D (по умолчанию 8 мм);
- отступ конца линии от края — свободный конец ставится с отступом от края контура; по умолчанию D/2 (радиус фрезы), можно задать в мм;
- тип неона — светодиодный (лента внутри: кусок неона = длина паза) или неон 2.0 (силиконовый рассеиватель в пазу, лента отдельно);
- «Курсив» — шрифт как образец: нет (фрезеровать сам контур, с учётом фрезы) / да, линия до кончика штриха (по умолчанию) / да, с отступом на полширины штриха;
- кратность реза (10 / 25 / 50 мм или своя) и галочка «Подгонять линии под кратность реза»: после расчёта каждый кусок укорачивается до ближайшей меньшей кратной длины — только за счёт свободных концов, поровну с обоих или только с конца линии (настройка «Укорачивать»). Удлинять нельзя: конец уже стоит с отступом от края. Замкнутые куски и куски без свободных концов не меняются — остаток виден в отчёте.

Контур ▾
--------

Упростить контур (NeonCenterline.NeonTools.SimplifyContours) — меньше узлов у выделенных кривых: кривая не отходит от исходной больше допуска (по умолчанию 0.1 мм), острые углы остаются углами, гладкие места — гладкими. Полезно для трассированных логотипов и контуров из DXF/CAM (тысячи узлов → десятки), шрифты обычно сокращаются на треть. В окне перед запуском — допуск, «Оставить оригинал» и «Показывать окно упрощения при каждом запуске» (то же — в Настройках). Текст и фигуры сначала переведите в кривые (Ctrl+Q).

Очистить контур (NeonCenterline.NeonTools.CleanContours) — убирает наложения и петли: буквы, собранные из перекрывающихся частей (стойка поверх дуг), налезающие друг на друга буквы, контуры, которые перехлёстывают сами себя (цифры, у которых дырка — петля одного контура), «ушки» и ласточкины хвосты у фигур после «Обводка в объект» и сдвига контура. Каждый объект становится одной кривой: внутренние линии уходят, дырки остаются, мелкие петли и лоскуты (меньше 0.3 мм² или тоньше 0.15 мм) — убираются. Форма не меняется: новый контур собирается из исходных кривых, разрезанных точно в точках пересечения. Текст переводится в кривые сам; прямоугольники и эллипсы пропускаются. Отдельные объекты, которые надо слить, сначала объедините (Ctrl+L).

Разбить на буквы (NeonCenterline.NeonTools.SplitLetters) — текст в кривых или объединённые буквы (одна кривая из многих контуров) — по объекту на букву или цифру, дырки остаются в своих буквах: не нужно, как после «Разъединить» (Ctrl+K), вычитать дырку из каждой буквы. Какой контур — дырка, решается вложенностью: контур внутри чётного числа других — наружный контур буквы, внутри нечётного — дырка ближайшего наружного. Точки и акценты остаются со своей буквой (Ё, Й, i, ї, !, ?, ÷). Отдельными остаются куски одного размера (двоеточие, «=», две части «ы»), стоящие рядом с буквой (точка после слова, кольца «%») и детали внутри дырок (буква в «®», точка в нуле моноширинного шрифта). Каждая буква — копия исходной кривой (заливка, обводка, слой те же), буквы выделены — можно сразу «Одинаковые — в группы» или «Раскладка на листы». Если буквы налезают друг на друга — сначала «Очистить контур».

Раскладка ▾
-----------

Раскладка на листы (бета; NeonCenterline.NeonTools.NestSheets) — раскладывает выделенные детали на листы (оргстекло, ПВХ, композит, стол лазера): ни одна деталь не подходит к краю листа ближе поля и к соседней — ближе зазора (для фрезы зазор — не меньше её диаметра). В окне:
- размер листа: готовые размеры в списке (свои запоминаются там же), длина (по горизонтали) и ширина (по вертикали) — любые;
- поле от края и зазор между деталями;
- раскладка: по контуру — плотнее (по умолчанию: детали ставятся по своей форме, входят друг в друга, мелкие — в просветы между крупными), по контуру с поворотом на 180°, по контуру с поворотами по 90° (поворот — только если материалу всё равно: без текстуры, направления волокон или плёнки) или по габаритам — быстрее (прямоугольниками, MaxRects);
- «Детали внутрь других (матрёшкой)» — по контуру: мелкие детали встают и в дырки букв, и внутрь рамок (по умолчанию включено); без галочки дырки деталей считаются занятыми, открытые выемки («П», «Г», «С») остаются;
- заполнять лист: столбцами — сверху вниз, затем вправо (по умолчанию), рядами — слева направо, затем вниз, или как плотнее (авто);
- «Каждый цвет — на свои листы» (цвет детали — заливка, без заливки — обводка);
- «Длина по раскладке»: ширина листа задана, а длина — сколько заняли детали плюс поле (вверх до 10 мм), от заданного минимума до длины листа; что не влезло — на следующий лист;
- «Подписи листов» (лист i из n, размер, цвет, число деталей, заполнение);
- «Оставить оригинал» — на листы встают копии деталей;
- «Группы — целиком» (по умолчанию): каждая группа — одна деталь (таблички, сборки не разбираются); без галочки группы разбираются на элементы — так удобно для импорта SVG/DXF, который приходит одной группой. Группы «Одинаковые» и листы прежней раскладки раскладываются поштучно всегда.

Из нескольких вариантов (порядок деталей, направление заполнения) берётся тот, где меньше листов, а при равенстве — короче последний. Листы рисуются справа от выделения, по пять в ряд, каждый цвет — своими рядами; каждый лист — одна группа на слое деталей (рамка, подпись, детали) — печатается и экспортируется вместе с ними. Повторный запуск по разложенному — новая раскладка, старые рамки и подписи из выделения убираются. В отчёте — листы и заполнение по цветам (по контуру — по площади деталей, по габаритам — по габаритам).
По контуру деталь переводится в сетку клеток (около 0.7 мм на листе 900×600, около 2.5 мм на 3050×2050) с запасом: зазор соблюдается всегда, но из-за клетки детали могут стоять чуть свободнее заданного; на сотнях деталей и больших листах раскладка по контуру идёт дольше. Контур детали должен быть замкнутым (у незамкнутого нет «внутри», туда может встать другая деталь; концы, сошедшиеся в точку, считаются замкнутыми). Проверяйте раскладку перед резкой.

Одинаковые — в группы (NeonCenterline.NeonTools.GroupSame) — находит среди выделенного объекты одной формы и одного цвета и собирает их в группы «Одинаковые: N шт · Ш × В мм». Контур сравнивается с точностью 0.1 мм, без поворота и зеркала; сравнивается сам контур, поэтому число и расстановка узлов, начало и направление обхода не важны — одна и та же цифра из разных переводов текста в кривые собирается вместе. В отчёте — ведомость форм по убыванию числа штук.

Лапша (NeonCenterline.NeonTools.MakeNoodles) — резы для полос на плоттере, лазере, фрезере: длина и ширина лапши, количество полос и хвосты — выход реза за линию другого реза, отдельно у горизонтальных резов (за отрезы) и у вертикальных отрезов (за крайние горизонтальные): лазеру и плоттеру хватает 1–2 мм, фрезе — 2–5 мм запаса. Получается N + 1 горизонтальных резов и два вертикальных отреза на концах. Для фрезы — галочка «Фреза» и её диаметр: фреза режет по центральной линии и съедает по радиусу с каждой стороны, поэтому горизонтальные резы идут через ширину + диаметр, а отрезы стоят на радиус за концами полос — после фрезы полосы выходят ровно заданного размера. Порядок объектов — порядок резки: горизонтальные сверху вниз зигзагом (первый — слева направо, дальше направления чередуются — станок не ходит вхолостую), вертикальные отрезы — в самом конце списка, тоже по ходу. Всё — одной группой в центре страницы.

Люверсы (NeonCenterline.NeonTools.PlaceGrommets) — точки под люверсы по краю выделенного баннера или подложки (прямоугольник, скруглённый прямоугольник, любой замкнутый контур): в каждом углу и равномерно по сторонам, не реже заданного расстояния; у круга и овала — равномерно по контуру. В окне: расстояние от края до центра люверса (по умолчанию 20 мм), расстояние между люверсами — не больше (300 мм), диаметр (5 мм) и цвет точки (из списка или любой через окно выбора цвета CorelDRAW), усиление углов — ещё по точке на обе стороны каждого угла на заданном расстоянии (50 мм), контур под подгиб — линия реза, смещённая наружу от края баннера на выпуск (по умолчанию 50 мм; у прямоугольника — прямоугольник с острыми углами, у кривых — ломаная с шагом 2 мм); люверсы по-прежнему отсчитываются от края баннера. Контур подгиба — на слое баннера, кривой «Подгиб N мм». Люверсы получают только внешние объекты: то, что лежит внутри другого выделенного объекта (рисунок, текст, рамки внутри баннера — в том числе в одной группе с ним), пропускается. Контур, нарисованный линиями без замыкания или собранный из отрезков, считается замкнутым, если концы сходятся. Точки — на слое баннера (галочкой «Люверсы — на отдельный слой» — на слое «Люверсы»), по группе на объект; количество — в отчёте.

Линии ▾
-------

Длина контуров и линий (NeonCenterline.NeonTools.MeasureLength) — длина выделенного: контуры (длина под резку) и центральные линии плагина (длина неона, число кусков, самый длинный кусок) — отдельно; для линий — сколько неона (или ленты для неона 2.0) выходит кратно резу и сколько уйдёт в остатки.

Выделить центральные линии (NeonCenterline.NeonTools.SelectNeonLines) — выделяет все центральные линии плагина на странице (например, чтобы сразу посчитать длину неона или убрать их).

Подогнать под кратность реза (NeonCenterline.NeonCut.FitToCut) — подгоняет под кратность уже готовые выделенные линии (как галочка «Подгонять линии под кратность реза» после расчёта).

Настройки
---------

В окне Настройки (группа «Настройки и помощь ▾»): диаметр фрезы, отступ конца линии от края, тип неона, «Курсив — шрифт как образец», кратность реза, подгонка линий под кратность и с какого конца укорачивать, «Спрашивать параметры при каждом запуске», допуск упрощения контура и показывать ли окно упрощения, «Оставлять оригинал: очистка, разбивка, упрощение», язык интерфейса (русский / English / как в Windows), ежедневная проверка обновлений. Параметры раскладки, лапши и люверсов запоминаются в их собственных окнах.

Обновления
----------

Раз в день (после расчёта центральных линий) плагин проверяет, нет ли новой версии на GitHub. Если есть — спросит, скачает, проверит контрольную сумму, и новая версия установится сама, когда вы закроете CorelDRAW. Старая версия сохраняется рядом как NeonCenterline.gms.bak. Проверить вручную: «Настройки и помощь ▾» → «Проверить обновления». Отключить ежедневную проверку — в Настройках.

Лицензия
--------

Бесплатно для любого использования, в том числе коммерческого (изготовление вывесок, фрезеровка на заказ). Продавать и перепаковывать плагин нельзя. Полный текст — LICENSE.txt (LICENSE.txt).

Плагин бесплатный. Если он вам помогает — можно поблагодарить автора: DONATE.md (DONATE.md) (та же ссылка — в окне «О программе»).

======================================================================

NEONCENTERLINE (ENGLISH)

A CorelDRAW plugin for sign making. Its core is a centerline built from letter outlines for milling flex-neon channels; around it are tools that prepare work for cutting: cleaning and simplifying outlines, splitting into letters, nesting parts on sheets, grouping identical parts, strips, grommets, length and cut step. Pure VBA, no DLL.

Features
--------

- Centerlines — Mono for monoline (neon) fonts and Script (beta) for handwritten ones. Output: curves on the layer "Неон - центр".
- Outline — clean (remove overlaps and loops), simplify (fewer nodes), split into letters (holes stay in their letters).
- Layout — nest parts on sheets by outline or by bounding boxes (beta), group identical, strips, grommets along a banner edge.
- Lines — length of outlines and lines, select centerlines, fit lines to the cut step.
- Every tool has a progress window (you can see the plugin is working and CorelDRAW has not hung) and a report with OK / Cancel (undo all of it in one step). Clean, simplify, split and nesting can keep the original.
- Russian and English interface, automatic updates from GitHub.

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
   (administrator rights required). Or, once the plugin is loaded, run the macro NeonCenterline.NeonAddon.InstallToolbar (Tools → Macros → Run Macro): it puts the toolbar into Addons of the CorelDRAW version it runs in.
3. Fully restart CorelDRAW and turn the toolbar on: Window → Toolbars → "Неон (NeonCenterline)".
4. If there is no toolbar or the buttons are blank (possible in older versions), add the buttons manually: Tools → Customization (Ctrl+J) → Commands → category "Macros" → drag to any toolbar NeonCenterline.NeonUI.RunMono, RunScript, ShowSettings, CheckUpdates, ShowAbout, NeonCenterline.NeonTools.SimplifyContours, CleanContours, SplitLetters, NestSheets, GroupSame, MakeNoodles, PlaceGrommets, MeasureLength, SelectNeonLines, NeonCenterline.NeonCut.FitToCut. Icons are in the icons folder of the archive (Appearance tab → Import).
5. If the toolbar still shows the old buttons after a plugin update, run the macro NeonCenterline.NeonUI.ResetToolbar (Tools → Macros → Run Macro): it resets only the "Неон" toolbar to the new button set.

Test font font\SABNeonTest.ttf is not required by the plugin — it is for trying it out (a monoline neon font that shows how Mono mode is meant to work). To install: double-click the file → Install.

Compatibility: tested in CorelDRAW 2026 (27). Should work in 2017–2025; in X7–X8 the computation most likely works, the toolbar may need manual setup. Please report in GitHub Issues whether it works in your version.

The "Неон" toolbar
------------------

Mono, Script | Outline ▾, Layout ▾, Lines ▾ | Settings and help ▾. Buttons with an arrow open a list:

| Group | Buttons |
|---|---|
| Outline ▾ | Simplify outline, Clean outline, Split into letters |
| Layout ▾ | Nest on sheets (beta), Group identical, Strips, Grommets |
| Lines ▾ | Length of outlines and lines, Select centerlines, Fit to the cut step |
| Settings and help ▾ | Settings, Check for updates, About |

The toolbar captions are in Russian; the windows and reports follow the interface language. After a plugin update that changes the button set, the plugin offers once to update the toolbar files (administrator rights); after a CorelDRAW restart the "Неон" toolbar rebuilds itself, the rest of the workspace is left alone.

Common to all tools
-------------------

- Progress window. While a tool works, a window shows the stage, a bar, "N of M" and the seconds elapsed — you can see the plugin is working and CorelDRAW has not hung. The document cannot be touched while it is open; it closes by itself.
- Report with OK and Cancel. After the tools that change the document a report shows OK to keep the result and Cancel to undo all of it in one step (like Ctrl+Z; Esc and the close button also cancel). If nothing changed, there is only OK. After centerlines there is only OK (undo with Ctrl+Z as usual).
- Keep the original. For nesting — a box in its window (copies of the parts go onto the sheets, the parts stay in place); for simplify — in its window; for clean, split and simplify — in the Settings ("Keep the original: clean, split, simplify"). The result is a copy on top of the original; objects that did not change get no copy; text is converted to curves in the copy too.
- Clean and split into letters convert text to curves themselves; simplify works on curves only. Groups are walked through — except in nesting and group identical, where a group is one part.

Centerlines: Mono and Script
----------------------------

Select outlines (curves, text, groups) and press Mono or Script.

- Mono — for monoline (neon) fonts: straight strokes stay straight, arcs use a minimum of curves, loops are closed.
- Script ⚠️ beta — for handwritten fonts (Segoe Script etc.). By default — font as a reference: the line follows the pen path, through junctions and up to the stroke tips, letter height as in the original. Handwritten fonts differ a lot — kinks and stray pieces may appear at letter joints, check the lines before cutting. If letters overlap or are built from overlapping parts, run Clean outline first.
- Output: curves "Неон: центр N" on the layer "Неон - центр" (red hairline). The plugin recognises its lines by this name (Length, Select lines, fitting).

The window before a run (can be turned off with "Ask for the parameters on every run" — the Settings values are used then):
- cutter diameter D (8 mm by default);
- line end inset from the edge — a free line end is set back from the outline edge; D/2 (the cutter radius) by default, or a value in mm;
- neon type — LED neon (strip inside: neon piece = groove length) or Neon 2.0 (silicone diffuser in a groove, separate strip);
- Script — font as a reference: no (mill the outline itself, with the cutter) / yes, line to the stroke tip (default) / yes, inset by half the stroke width;
- cut step (10 / 25 / 50 mm or custom) and "Fit lines to the cut step": after a run every piece is shortened to the nearest smaller multiple — only at free ends, equally at both or only at the line end ("Shorten" setting). Lines are never extended: a free end is already inset from the edge. Closed pieces and pieces without free ends are left as is; the remainder is shown in the report.

Outline ▾
---------

Simplify outline (NeonCenterline.NeonTools.SimplifyContours) — fewer nodes on the selected curves: the curve deviates from the original by no more than the tolerance (0.1 mm by default), sharp corners stay sharp, smooth parts stay smooth. Useful for traced logos and outlines from DXF/CAM (thousands of nodes become dozens); font outlines usually lose about a third of their nodes. The window before the run: tolerance, "Keep the original" and "Show the simplify window on every run" (also in the Settings). Convert text and shapes to curves first (Ctrl+Q).

Clean outline (NeonCenterline.NeonTools.CleanContours) — removes overlaps and loops: letters built from overlapping parts (a stem lying over arcs), letters running into each other, outlines that loop over themselves (digits whose hole is a loop of one outline), ears and swallowtails of shapes after "Convert outline to object" and offsetting. Each object becomes one curve: inner lines go, holes stay, tiny loops and slivers (under 0.3 mm² or thinner than 0.15 mm) are removed. The shape does not change: the new outline is built from the original curves cut exactly at the crossings. Text is converted to curves automatically; rectangles and ellipses are skipped. Combine separate objects that should merge first (Ctrl+L).

Split into letters (NeonCenterline.NeonTools.SplitLetters) — text converted to curves or combined letters (one curve of many outlines) becomes one object per letter or digit, with the holes kept in their letters: no more Break Apart (Ctrl+K) and then trimming every hole out of its letter. Nesting decides what is a hole: an outline inside an even number of others is a letter's outer outline, inside an odd number — a hole of the nearest outer one. Dots and accents stay with their letter (Ё, Й, i, ї, !, ?, ÷). Pieces of the same size (colon, "=", the two parts of "ы"), pieces next to a letter (a full stop, the rings of "%") and parts inside holes (the letter in ®, the dot in a monospace zero) stay separate. Each letter is a copy of the original curve (same fill, outline and layer), and the letters stay selected, ready for Group identical or Nest on sheets. If letters overlap, run Clean outline first.

Layout ▾
--------

Nest on sheets (beta; NeonCenterline.NeonTools.NestSheets) — lays the selected parts out on sheets (acrylic, PVC, composite, a laser bed): no part comes closer than the margin to the sheet edge or the gap to another part (for a router, use a gap of at least the cutter diameter). The window:
- sheet size: presets in the list (your own sizes are kept there too), length (horizontal) and width (vertical) — any;
- edge margin and gap between parts;
- nesting: by outline — tighter (default: parts are placed by their shape, fit into each other, small ones go into the gaps between big ones), by outline with rotation by 180°, by outline with rotation in 90° steps (rotate only if the material does not care: no texture, grain or film direction) or by bounding boxes — faster (rectangles, MaxRects);
- "Parts inside other parts (nested)" — by outline: small parts also go into letter holes and inside frames (on by default); without it the holes of parts count as taken, open recesses (П, Г, С shapes) still interlock;
- fill the sheet: in columns — top to bottom, then right (default), in rows — left to right, then down, or as tight as possible (auto);
- "Each colour on its own sheets" (a part's colour is its fill, or its outline if it has no fill);
- "Length to fit the parts": the sheet width is fixed and the length is what the parts take plus the margin (rounded up to 10 mm), between the given minimum and the sheet length; what does not fit goes to the next sheet;
- "Sheet labels" (sheet i of n, size, colour, part count, fill);
- "Keep the original" — copies of the parts go onto the sheets;
- "Groups stay whole" (default): each group is one part (plates and assemblies are not split); without it groups are split into their elements — handy for an SVG/DXF import that comes as one group. "Identical" groups and sheets of an earlier layout are always nested one by one.

Of several variants (part order, fill direction) the one with the fewest sheets wins, then the one with the shortest last sheet. Sheets are drawn to the right of the selection, five in a row, each colour in its own rows; each sheet is one group on the parts' layer (frame, label, parts) — printed and exported with them. Running it again on a layout makes a new one and removes the old frames and labels in the selection. The report lists sheets and fill by colour (by part area when nesting by outline, by bounding boxes otherwise).
By outline, a part becomes a grid of cells (about 0.7 mm on a 900×600 sheet, about 2.5 mm on 3050×2050) with a safety allowance: the gap is always kept, but because of the cells parts may sit a little further apart than set; with hundreds of parts on large sheets nesting by outline takes longer. A part's outline must be closed (an open one has no inside, and another part may go there; ends that meet count as closed). Check the layout before cutting.

Group identical (NeonCenterline.NeonTools.GroupSame) — finds selected objects of the same shape and colour and groups them as "Identical: N pcs · W × H mm". Outlines are compared within 0.1 mm, no rotation or mirroring; the outline itself is compared, so the number and placement of nodes, the start point and direction do not matter — the same digit from different text-to-curves conversions is grouped together. The report lists the shapes by count.

Strips (NeonCenterline.NeonTools.MakeNoodles) — cut lines for strips on a plotter, laser or router: strip length and width, the number of strips and the tails — how far a cut runs past the other cut's line, separately for horizontal cuts (past the end cuts) and vertical end cuts (past the outer horizontal ones): a laser or plotter needs 1–2 mm, a router 2–5 mm of margin. You get N + 1 horizontal cuts and two vertical end cuts. For a router tick "Router" and set the cutter diameter: the router cuts along the centre line and eats a radius on each side, so horizontal cuts are spaced by width + diameter and the end cuts sit a radius past the strip ends — after the router the strips come out exactly the size you set. The object order is the cutting order: horizontal cuts top to bottom in a zigzag (the first left to right, then alternating — no idle runs), the vertical end cuts at the very end of the list, also along the way. One group in the middle of the page.

Grommets (NeonCenterline.NeonTools.PlaceGrommets) — points for grommets along the edge of the selected banner or backing (rectangle, rounded rectangle, any closed outline): one in every corner and evenly along the sides, never further apart than the given distance; circles and ovals get them evenly all round. The window: edge to grommet centre (20 mm by default), distance between grommets — maximum (300 mm), point diameter (5 mm) and colour (from the list or any via the CorelDRAW colour dialog), corner reinforcement — one more point on both sides of every corner at the given distance (50 mm), hem outline — a cut line offset outward from the banner edge by the allowance (50 mm by default; a rectangle gets a sharp-cornered rectangle, curves a polyline with a 2 mm step); grommets are still measured from the banner edge. The hem outline goes to the banner layer as a curve "Hem N mm". Only outer objects get grommets: anything lying inside another selected object (artwork, text, frames inside the banner — also in the same group with it) is skipped. An outline drawn with lines and not closed, or assembled from segments, counts as closed if its ends meet. Points go to the banner layer (with "Grommets on a separate layer" — to the "Grommets" layer), one group per object; the count is in the report.

Lines ▾
-------

Length of outlines and lines (NeonCenterline.NeonTools.MeasureLength) — measures the selection: outlines (cutting length) and plugin centerlines (neon length, number of pieces, longest piece) separately; for lines — how much neon (or strip, for Neon 2.0) fits in whole cut steps and how much goes to remainders.

Select centerlines (NeonCenterline.NeonTools.SelectNeonLines) — selects all plugin centerlines on the page (e.g. to measure the neon or remove them).

Fit to the cut step (NeonCenterline.NeonCut.FitToCut) — fits already selected lines to the cut step (like "Fit lines to the cut step" after a run).

Settings
--------

The Settings window (Settings and help ▾): cutter diameter, line end inset, neon type, "Script — font as a reference", cut step, fitting lines to the cut step and which end to shorten, "Ask for the parameters on every run", simplify tolerance and whether to show the simplify window, "Keep the original: clean, split, simplify", interface language (Russian / English / as Windows), daily update check. Nesting, strips and grommets remember their parameters in their own windows.

Updates
-------

Once a day (after a centerline run) the plugin checks GitHub for a new version. If there is one, it asks, downloads it, verifies the checksum, and the new version is installed automatically when you close CorelDRAW. The previous version is kept as NeonCenterline.gms.bak. Manual check: Settings and help ▾ → "Check for updates". The daily check can be turned off in the Settings.

License
-------

Free for any use, including commercial (sign making, milling for hire). Selling or repackaging the plugin is prohibited. Full text: LICENSE.txt (LICENSE.txt).

The plugin is free. If it helps you, you can support the author: DONATE.md (DONATE.md) (same link in the "About" window).
