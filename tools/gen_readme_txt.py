"""README.txt (простой текст, Windows-1251 не нужен: UTF-8 с BOM, CRLF) из README.md — для архива пользователей."""
import os, re
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
s = open(os.path.join(ROOT, 'README.md'), encoding='utf-8').read()
s = re.sub(r'<a name="[^"]*"></a>\n?', '', s)
s = re.sub(r'\[([^\]]+)\]\(#[^)]*\)', r'\1', s)            # внутренние ссылки
s = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'\1 (\2)', s)      # внешние ссылки -> текст (адрес)
s = re.sub(r'^#{1,6} *(.+)$', lambda m: m.group(1).upper() if m.group(0).startswith('# ') else m.group(1) + '\n' + '-' * len(m.group(1)), s, flags=re.M)
s = s.replace('**', '').replace('`', '')
s = re.sub(r'(?<!\w)\*([^*\n]+)\*(?!\w)', r'\1', s)
s = s.replace('\n---\n', '\n' + '=' * 70 + '\n')
s = re.sub(r'\n{3,}', '\n\n', s).strip() + '\n'
open(os.path.join(ROOT, 'README.txt'), 'w', encoding='utf-8-sig', newline='\r\n').write(s)
print('README.txt', len(s.splitlines()), 'lines')
