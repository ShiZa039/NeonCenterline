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
- «Моно»: линия, идущая прямо через развилку, больше не даёт «волну», когда сразу за развилкой стоит фаска или излом (буква «к» в моноширинных шрифтах).
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
- Mono: a line running straight through a junction no longer gets a "wave" when a chamfer or bend follows right after the junction (letter "к" in monoline fonts).
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
