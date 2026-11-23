"""Стенд: имя файла по правилам (раздел @@NAMING) в LibreOffice против ожидаемых значений.
Запуск из корня репозитория:  python tools/stands/naming.py  ->  mismatches: 0

Проверяется: разбор файла правил (шаблон, списки, комментарии, метка UTF-8, CRLF), вид частей
(список, текст, размеры, номер), код из строки окна, приведение размеров, запрещённые символы,
сборка имени с пустыми частями, следующий номер по файлам папки.
"""
import os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, 'tools'))
import lo_stand as S

DASH = '—'
RULES = ('﻿# правила\r\n\r\n[Шаблон]\r\n{печать}_{материал}_{размеры}_{клиент}_{номер}\r\n\r\n'
         '[печать]\r\n720 = Интерьерная печать\r\n; комментарий\r\n360=Наружная\r\n\r\n'
         '[материал]\r\nblock = Баннер блэкаут\r\nfrontlit\r\n[лишний]\r\nx = y\r\n')
DEF = '{print}_{material}_{size}_{client}_{number}'


def vb_str(s):
    """Строка Python -> выражение Basic (ChrW для не-ASCII и спецсимволов)."""
    out, cur = [], ''
    for ch in s:
        if 32 <= ord(ch) < 127 and ch != '"':
            cur += ch
        else:
            if cur:
                out.append('"' + cur + '"'); cur = ''
            out.append(f'ChrW({ord(ch)})')
    if cur:
        out.append('"' + cur + '"')
    return ' & '.join(out) if out else '""'


CASES = []   # (имя проверки, выражение Basic, ожидаемое)
main = ['Sub Main', '    Dim g As Integer, i As Long, v(0 To 7) As String, nm(0 To 9) As String', '    g = FreeFile',
        f'    Open "{S.OUT}" For Output As #g']


def check(name, expr, want):
    CASES.append((name, want))
    main.append(f'    Print #g, CStr({expr})')


main.append(f'    NmParse {vb_str(RULES)}, {vb_str(DEF)}')
check('шаблон', 'NmTpl', '{печать}_{материал}_{размеры}_{клиент}_{номер}')
check('частей', 'NmNT', '5')
check('имена частей', 'NmTokName(0) & "|" & NmTokName(1) & "|" & NmTokName(2) & "|" & NmTokName(3) & "|" & NmTokName(4)',
      'печать|материал|размеры|клиент|номер')
check('виды частей', 'NmKind(0) & NmKind(1) & NmKind(2) & NmKind(3) & NmKind(4)', '00213')
check('списков', 'NmNL', '3')
check('пунктов печати', 'NmLsCnt(NmTokList(0))', '2')
check('пункт 720', 'NmItemShow(NmLsFirst(NmTokList(0)))', f'720 {DASH} Интерьерная печать')
check('пункт 360', 'NmItemShow(NmLsFirst(NmTokList(0)) + 1)', f'360 {DASH} Наружная')
check('пункт без расшифровки', 'NmItemShow(NmLsFirst(NmTokList(1)) + 1)', 'frontlit')
check('код из строки', f'NmCodeOf({vb_str("720 " + DASH + " Интерьерная печать")})', '720')
check('код набран руками', 'NmCodeOf("  540 ")', '540')
check('размеры', f'NmSizes({vb_str("3000 х 2500 * 500")})', '3000x2500x500')
check('размеры 2', f'NmSizes({vb_str("x1200×800X ")})', '1200x800')
check('чистка', f'NmClean({vb_str(" Вок/сис: ООО*")})', 'Вок-сис- ООО-')
check('значение номера', 'NmValue(4, "№ 07a")', '07')
check('значение списка', f'NmValue(1, {vb_str("block " + DASH + " Баннер блэкаут")})', 'block')
main += ['    v(0) = NmValue(0, "720" & " " & ChrW(8212) & " x"): v(1) = NmValue(1, "block")',
         f'    v(2) = NmValue(2, {vb_str("3000х2500х500")}): v(3) = NmValue(3, "Voksis"): v(4) = NmValue(4, "1")']
check('имя как в примере', 'NmBuild(v)', '720_block_3000x2500x500_Voksis_1')
main += ['    v(0) = ""']
check('без первой части', 'NmBuild(v)', 'block_3000x2500x500_Voksis_1')
main += ['    v(0) = "720": v(2) = "": v(4) = ""']
check('без середины и номера', 'NmBuild(v)', '720_block_Voksis')
# следующий номер
main += ['    nm(0) = "720_block_3000x2500x500_Voksis_1": nm(1) = "720_block_1x2_Other_12"',
         '    nm(2) = "random file 99": nm(3) = "360_frontlit_10x10_A_3": nm(4) = "720_block_x_Voksis_"']
check('следующий номер', 'NmNextNum(nm, 5)', '13')
check('пустая папка', 'NmNextNum(nm, 0)', '1')
# без раздела шаблона -> шаблон по умолчанию; часть без списка — текст
main.append(f'    NmParse {vb_str("[material]" + chr(10) + "block = Blackout")}, {vb_str(DEF)}')
check('шаблон по умолчанию', 'NmTpl', DEF)
check('виды по умолчанию', 'NmKind(0) & NmKind(1) & NmKind(2) & NmKind(3) & NmKind(4)', '10213')
# шаблон с началом, хвостом и номером не в конце
main.append(f'    NmParse {vb_str("[template]" + chr(10) + "NC-{a}.{number}-{b}!")}, {vb_str(DEF)}')
check('начало и хвост', 'NmPrefix & "|" & NmSuffix & "|" & NmTokSep(1) & "|" & NmTokSep(2)', 'NC-|!|.|-')
main += ['    v(0) = "x": v(1) = "5": v(2) = "y"']
check('сборка с хвостом', 'NmBuild(v)', 'NC-x.5-y!')
check('номер не в конце', 'NmNextNum(nm, 5)', '1')
# образцы файла правил из tools/strings.py (RU и EN) разбираются как надо
sys.path.insert(0, os.path.join(ROOT, 'tools'))
from strings import STRINGS
for lang, sample in zip(('RU', 'EN'), STRINGS['nm_sample']):
    main.append(f'    NmParse {vb_str(sample)}, {vb_str(DEF)}')
    check(f'образец {lang}: виды', 'NmKind(0) & NmKind(1) & NmKind(2) & NmKind(3) & NmKind(4)', '00213')
    main += ['    v(0) = NmValue(0, NmItemShow(NmLsFirst(NmTokList(0)))): v(1) = NmValue(1, NmItemShow(NmLsFirst(NmTokList(1))))',
             f'    v(2) = NmValue(2, {vb_str("3000 х 2500 х 500")}): v(3) = NmValue(3, "Voksis"): v(4) = NmValue(4, "1")']
    check(f'образец {lang}: имя', 'NmBuild(v)', '720_block_3000x2500x500_Voksis_1')
main += ['    Close #g', 'End Sub']

code = S.module(S.section('NAMING_DECL'), S.section('NAMING'), main)
L = S.run(code)
bad = 0
for (name, want), got in zip(CASES, L):
    ok = got == want
    bad += not ok
    print(f'{"OK " if ok else "BAD"} {name:24} {got!r}' + ('' if ok else f'  ждали {want!r}'))
if len(L) != len(CASES):
    print('строк', len(L), 'ждали', len(CASES)); bad += 1
print('mismatches:', bad)
sys.exit(1 if bad else 0)
