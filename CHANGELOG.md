# Changelog

## 1.0.2

RU
- Отступ свободного конца линии от края задаётся числом в мм (по умолчанию D/2, как раньше) — в окне перед запуском и в Настройках.
- «Курсив»: режим «шрифт как образец» (по умолчанию) — линия по ходу пера без учёта фрезы, конец у кончика штриха или с отступом полширины; выбор в окне перед запуском и в Настройках.
- «Курсив», конец «до кончика штриха»: концы больше не уходят в угол среза и не загибаются крючком — линия продолжает ход штриха (прямой или дугой у завитков и шариков) до края, высота букв как в оригинале.
- «Курсив», режим образца: перо не рисует «ушки» (короткий выступ стойки над местом, где в неё вливается дуга) — дуга поворачивает прямо в стойку; острые изломы линии скругляются, как поворот пера.
- «Курсив», режим образца: штрих, отходящий от стойки под углом (вторая и третья дуга «т», «м», «п»), идёт прямой диагональю до стойки, без S-изгиба.
- «Курсив», режим образца: две близкие развилки, которые вместе образуют крест (перетяжка «з», петли), сливаются в одну — штрихи проходят крест насквозь, без «паука» из коротких кусков.
- «Курсив», режим образца, острые стыки (ж, т, к): штрихи собираются по продолжению, а не по срединной оси у стыка. Сквозь развилку идёт пара веток, чьи дальние участки ложатся на одну прямую (или дугу), — стойка «ж» теперь сплошная сверху донизу, диагонали упираются в неё; отходящий штрих достраивается по своему ходу (прямой или дугой, при необходимости — отрезком к ближайшей линии), а если его сквозная линия недостижима — к любой соседней линии (чаша «с» в диагональ «ж»).
- «Курсив», режим образца: штрих, подходящий к стойке сбоку по касательной (язычок «з», вход «т», выход «ь», чаша «а»), больше не рвёт стойку — стойка сплошная, штрих примыкает к ней у самого стыка (без «горба» выше стыка, «в»). Крутая гладкая дуга (шпилька «ы», «в») больше не ломается на два куска с зазубриной.
- «Курсив», режим образца: диагональ «3» доходит прямой до кончика перетяжки, нижняя чаша упирается в неё; острый возврат пера («и», «я») — без крючка у основания; стыки «ж» строятся одинаково в плагине и на стенде.
- «Курсив», режим образца: арка, вливающаяся в стойку под углом без продолжения по прямой или дуге (последняя «т» в «текст» мелким кеглем), снова проходит в стойку насквозь — угол в развилке выбирается по касательным у самой развилки.
- «Курсив», режим образца: стойка, через которую проходят другие штрихи («ж», «р», «ф», основание «1»), больше не «виляет» у развилок — штрих идёт сквозь развилку по своему ходу: прямой, если штрих прямой, иначе плавной дугой.
- «Курсив», режим образца: тонкий прямой соединительный штрих перед дугой (диагональ от стойки к арке «т», «м») идёт прямой от стойки и плавно переходит в дугу, как в образце, — раньше он начинался выше и ломался у дуги.
- «Курсив», режим образца: чаша «е» у стыка с перекладиной больше не выгибается наружу — касательные концов развилки берутся по ходу штриха на ширину пера (дугой, если штрих изогнут), а не по ближним точкам утолщения.
- «Курсив», режим образца: штрих, который подходит к соседней линии сбоку (низ чаши «я», «ж», «в», «ь»), достраивается плавной кривой по своему ходу, а не прямым отрезком, срезающим угол развилки у самого края. Прямой отрезок остаётся там, где штрих и так идёт почти прямо (диагональ «т» от арки к стойке).
- «Курсив», режим образца: ноги острого поворота пера прямые до самого поворота — без ступеньки внизу стойки «т» в связном «ст», в низу «и», «л», в перекладине «б».
- «Курсив», режим образца: два близких поворота в одну сторону (крючок входа «т», «к»: верх крючка и сразу поворот в стойку; арки «п», «м») скругляются одной дугой — без бугорка между ними.
- «Курсив», режим образца: острый шпиль принадлежит штриху, который входит в него по прямой, второй штрих примыкает к нему сбоку. Стойка «т» (в «ст»), «и», «р» проходит через шпиль до самого верха, вход упирается в неё; в перетяжке «з», «З» верхняя чаша (диагональ) доходит до кончика, а нижняя примыкает к ней у выреза — без острого возврата пера в кончик.
- «Курсив», режим образца: завиток на конце штриха (верх «с») — линия идёт по изгибу до кончика, а не уходит прямой в сторону.
- «Моно»: линия, идущая прямо через развилку, больше не даёт «волну», когда сразу за развилкой стоит фаска или излом (буква «к» в моноширинных шрифтах).
- Новая кнопка «Неон: упростить контур»: меньше узлов у выделенных кривых в пределах допуска (окно перед запуском, как у центральных линий, и Настройки; по умолчанию 0.1 мм), острые углы сохраняются, гладкие места остаются гладкими; кривые меняются на месте, отмена — Ctrl+Z. Трассированные логотипы и контуры из DXF/CAM — с тысяч узлов до десятков.
- Панель «Неон» обновляется сама: если в новой версии плагина на панели новые кнопки, плагин один раз предложит обновить файлы панели (права администратора), а после перезапуска CorelDRAW (при запуске, при открытии документа или при первой команде плагина) сбросит только панель «Неон» к новому составу и проверит, что новые кнопки действительно на ней (если рабочее пространство при загрузке вернуло старый состав — сброс повторяется); вручную — макрос NeonCenterline.NeonUI.ResetToolbar — остальное рабочее пространство не трогается. Установка панели проходит и при открытом CorelDRAW: у каждого набора кнопок свой файл значков (Resources7.dll и т. д.), поэтому старый, занятый CorelDRAW, не мешает новым значкам; установка проверяет, что файл значков скопирован целиком.
- Новая кнопка «Неон: люверсы» (в группе «Раскладка ▾»): точки под люверсы по краю выделенного баннера или подложки — в каждом углу (скруглённый угол тоже угол) и равномерно по сторонам, не реже заданного расстояния; у круга и овала — равномерно по контуру. Задаются расстояние от края до центра люверса, расстояние между люверсами, диаметр и цвет точки (из списка или любой через окно выбора цвета CorelDRAW), усиление углов (ещё по точке на обе стороны угла, по умолчанию в 50 мм). Точки — на слое «Люверсы», по группе на объект, количество — в отчёте.
- Панель: новая группа «Раскладка ▾» — люверсы, дальше туда же пойдут раскладка модулей и другое для производства.
- Панель «Неон» собрана по инструментам: на панели Моно, Курсив, Упростить контур и две кнопки-списка — «Линии ▾» (длина, выделить линии, подогнать под рез) и «Настройки ▾» (настройки, проверить обновления, о программе); свои значки у «проверить обновления» и «о программе». Панель у пользователей перестроится сама после обновления.
- Кнопка «Поддержать автора» в окне «О программе» (ведёт на страницу DONATE.md).
- README: ручная установка, совместимость, тестовый шрифт; в архиве — папка icons.

EN
- The free line end inset from the edge can be set in mm (default D/2, as before) — in the run window and Settings.
- Script: "font as a reference" mode (default) — a pen-path line ignoring the cutter, ending at the stroke tip or inset by half the stroke width; choice in the run window and Settings.
- Script, "end at the stroke tip": line ends no longer veer into the corner of a cut terminal or hook back — the line follows the stroke (straight, or an arc for curls and ball terminals) to the edge, so letter height matches the original.
- Script, reference mode: the pen no longer draws "ears" (the short stub of a stem above where an arch flows into it) — the arch turns straight into the stem; sharp bends of the line are rounded like a pen turn.
- Script, reference mode: a stroke leaving a stem at an angle (second and third arch of "т", "м", "п") runs as a straight diagonal into the stem, without an S-bend.
- Script, reference mode: two nearby junctions that together form a crossing (waist of "з", loops) are merged into one, so strokes run straight through instead of a "spider" of short pieces.
- Script, reference mode, acute joins (ж, т, к): strokes are assembled by continuation rather than by the medial axis at the join. The pass-through pair at a junction is the one whose far sections lie on a single line (or arc) — the stem of "ж" is now continuous top to bottom with the diagonals ending on it; a departing stroke is extended along its own course (line or arc, or a straight segment to the nearest line), and if its pass-through line is out of reach, to any neighbouring line (the bowl of "с" into the diagonal of "ж").
- Script, reference mode: a stroke joining a stem tangentially from the side (the tongue of "з", the entry of "т", the exit of "ь", the bowl of "а") no longer breaks the stem — the stem stays continuous and the stroke lands right at the join (no "hump" above it, "в"). A tight smooth arc (the hairpin of "ы", "в") is no longer split into two pieces with a notch.
- Script, reference mode: the diagonal of "3" runs straight to the tip of the waist with the lower bowl ending on it; a sharp pen return ("и", "я") has no hook at its base; joins in "ж" are built identically in the plugin and the test bench.
- Script, reference mode: an arch flowing into a stem at an angle with no straight or arc continuation (the last "т" of "текст" at a small size) again runs through into the stem — the corner at the junction is chosen by the tangents right at the junction.
- Script, reference mode: a stem crossed by other strokes ("ж", "р", "ф", the base of "1") no longer wobbles at the junctions — the stroke runs through the junction along its own course: straight if the stroke is straight, otherwise a smooth arc.
- Script, reference mode: a thin straight connecting stroke before an arch (the diagonal from the stem to the arch of "т", "м") now runs straight from the stem and flows smoothly into the arch, as in the reference; it used to start higher and kink at the arch.
- Script, reference mode: the bowl of "е" no longer bulges outwards at the crossbar joint — the end tangents at the junction follow the stroke over one pen width (an arc if the stroke is curved), not the nearby points of the thickening.
- Script, reference mode: a stroke reaching a neighbouring line from the side (bottom of the bowl of "я", "ж", "в", "ь") is completed with a smooth curve along its own course instead of a straight segment cutting the junction corner at the very edge. The straight segment stays where the stroke already runs almost straight (the "т" diagonal from the arch to the stem).
- Script, reference mode: the legs of a sharp pen turn are straight right up to the turn — no step at the bottom of the "т" stem in a connected "ст", at the bottom of "и", "л", in the crossbar of "б".
- Script, reference mode: two close turns in the same direction (entry hook of "т", "к": top of the hook and then the turn into the stem; arches of "п", "м") are rounded with one arc, without a bump between them.
- Script, reference mode: a sharp spike belongs to the stroke that runs straight into it; the other stroke joins it from the side. The stem of "т" (in "ст"), "и", "р" runs through the spike to the very top and the entry stroke meets it; in the waist of "з", "З" the upper bowl (diagonal) reaches the tip and the lower bowl joins it at the notch — no sharp pen return into the tip.
- Script, reference mode: a curl at the end of a stroke (top of "с") — the line follows the curl to its tip instead of leaving it in a straight line.
- Mono: a line running straight through a junction no longer gets a "wave" when a chamfer or bend follows right after the junction (letter "к" in monoline fonts).
- New button "Neon: simplify outline": fewer nodes on the selected curves within a tolerance (a window before the run, like for centerlines, and Settings; 0.1 mm by default); sharp corners are kept and smooth parts stay smooth; curves are changed in place, undo with Ctrl+Z. Traced logos and DXF/CAM outlines go from thousands of nodes to dozens.
- The "Neon" toolbar updates itself: if a new plugin version adds buttons, the plugin offers once to update the toolbar files (administrator rights), and after a CorelDRAW restart (at start, when a document opens or at the first plugin command) resets only the "Neon" toolbar to the new button set and checks that the new buttons are really there (if the workspace brought the old set back while loading, the reset is repeated); manually — macro NeonCenterline.NeonUI.ResetToolbar; the rest of the workspace is untouched. Installing the toolbar also works while CorelDRAW is open: each button set has its own icon file (Resources7.dll etc.), so the old one held by CorelDRAW does not block the new icons; the installer checks that the icon file was copied in full.
- New button "Neon: grommets" (in the "Layout ▾" group): points for grommets along the edge of the selected banner or backing — one in every corner (a rounded corner counts too) and evenly along the sides, never further apart than the given distance; circles and ovals get them evenly all round. Settings: edge to grommet centre, distance between grommets, point diameter and colour (from the list or any via the CorelDRAW colour dialog), corner reinforcement (one more point on both sides of a corner, 50 mm by default). Points go to the "Grommets" layer, one group per object; the count is in the report.
- Toolbar: new "Layout ▾" group — grommets; module layout and other production tools will go there.
- The "Неон" toolbar is grouped by tool: Mono, Script, Simplify outline and two drop-down buttons — "Lines ▾" (length, select lines, fit to cut) and "Settings ▾" (settings, check for updates, about); check for updates and about got their own icons. Users' toolbars rebuild themselves after the update.
- "Support the author" in the About window (opens DONATE.md).
- README: manual installation, compatibility, test font; icons folder in the archive.

## 1.0.1

RU
- Проверка обновлений: запасные способы связи (curl.exe из Windows, прокси из настроек Windows, запасные адреса) — для сетей, где встроенный механизм не достучался до GitHub.
- В архиве релиза — README.txt (простой текст).

EN
- Update check: fallback connection methods (Windows curl.exe, Windows proxy settings, mirror addresses) for networks where the built-in method could not reach GitHub.
- README.txt (plain text) in the release archive.

## 1.0.0 — 2026-10-23

RU
- Плагин вместо отдельного макроса: панель с кнопками «Моно», «Курсив», «Настройки».
- «Моно» — прежняя логика для моноширинных шрифтов (без изменений алгоритма).
- «Курсив» — для рукописных шрифтов: линия не выходит за контур в местах слияния штрихов, канал фрезы не задевает край.
- Окно настроек: диаметр фрезы, спрашивать ли его при запуске, язык интерфейса (RU/EN/авто), проверка обновлений.
- Автообновление с GitHub (установка после закрытия CorelDRAW, предыдущая версия сохраняется как .bak).
- Кнопка «Длина»: длина выделенных контуров (под резку) и центральных линий (неон) с копированием в буфер.
- Кнопка «Выделить линии»: все центральные линии плагина на странице.
- Окно параметров перед расчётом: диаметр, тип неона, кратность реза, подгонка (можно отключить).
- Тип неона (светодиодный / неон 2.0) и кратность реза; подгонка линий под кратность (укорачивание свободных концов), отчёт по кратности в «Длине».

EN
- Plugin instead of a standalone macro: toolbar with Mono, Script and Settings buttons.
- Mono — the previous algorithm for monoline fonts (unchanged).
- Script — for handwritten fonts: the line stays inside the outline where strokes merge, the cutter channel does not touch the edge.
- Settings window: cutter diameter, ask on every run, interface language (RU/EN/auto), update check.
- Auto-update from GitHub (installed after CorelDRAW is closed, the previous version is kept as .bak).
- Length button: length of selected outlines (cutting) and centerlines (neon), copy to clipboard.
- Select lines button: all plugin centerlines on the page.
- Parameters window before a run: cutter, neon type, cut step, fitting (can be turned off).
- Neon type (LED neon / Neon 2.0) and cut step; fitting lines to the cut step (shortening free ends), cut-step summary in Length.
