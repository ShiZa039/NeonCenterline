# Changelog

## 1.0.0 — 2026-10-23

RU
- Плагин вместо отдельного макроса: панель с кнопками «Моно», «Курсив», «Настройки».
- «Моно» — прежняя логика для моноширинных шрифтов (без изменений алгоритма).
- «Курсив» — для рукописных шрифтов: линия не выходит за контур в местах слияния штрихов, канал фрезы не задевает край.
- Окно настроек: диаметр фрезы, спрашивать ли его при запуске, язык интерфейса (RU/EN/авто), проверка обновлений.
- Автообновление с GitHub (установка после закрытия CorelDRAW, предыдущая версия сохраняется как .bak).
- Кнопка «Длина»: длина выделенных контуров (под резку) и центральных линий (неон) с копированием в буфер.
- Кнопка «Выделить линии»: все центральные линии плагина на странице.
- Тип неона (светодиодный / неон 2.0) и кратность реза; подгонка линий под кратность (укорачивание свободных концов), отчёт по кратности в «Длине».

EN
- Plugin instead of a standalone macro: toolbar with Mono, Script and Settings buttons.
- Mono — the previous algorithm for monoline fonts (unchanged).
- Script — for handwritten fonts: the line stays inside the outline where strokes merge, the cutter channel does not touch the edge.
- Settings window: cutter diameter, ask on every run, interface language (RU/EN/auto), update check.
- Auto-update from GitHub (installed after CorelDRAW is closed, the previous version is kept as .bak).
- Length button: length of selected outlines (cutting) and centerlines (neon), copy to clipboard.
- Select lines button: all plugin centerlines on the page.
- Neon type (LED neon / Neon 2.0) and cut step; fitting lines to the cut step (shortening free ends), cut-step summary in Length.
