# Changelog

## 1.0.2

RU
- Отступ свободного конца линии от края задаётся числом в мм (по умолчанию D/2, как раньше) — в окне перед запуском и в Настройках.
- «Курсив»: режим «шрифт как образец» (по умолчанию) — линия по ходу пера без учёта фрезы, конец у кончика штриха или с отступом полширины; выбор в окне перед запуском и в Настройках.
- «Курсив», конец «до кончика штриха»: концы больше не уходят в угол среза и не загибаются крючком — линия продолжает ход штриха (прямой или дугой у завитков и шариков) до края, высота букв как в оригинале.
- «Моно»: линия, идущая прямо через развилку, больше не даёт «волну», когда сразу за развилкой стоит фаска или излом (буква «к» в моноширинных шрифтах).
- Кнопка «Поддержать автора» в окне «О программе» (ведёт на страницу DONATE.md).
- README: ручная установка, совместимость, тестовый шрифт; в архиве — папка icons.

EN
- The free line end inset from the edge can be set in mm (default D/2, as before) — in the run window and Settings.
- Script: "font as a reference" mode (default) — a pen-path line ignoring the cutter, ending at the stroke tip or inset by half the stroke width; choice in the run window and Settings.
- Script, "end at the stroke tip": line ends no longer veer into the corner of a cut terminal or hook back — the line follows the stroke (straight, or an arc for curls and ball terminals) to the edge, so letter height matches the original.
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
