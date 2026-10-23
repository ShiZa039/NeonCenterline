"""Копия src/*.bas в кодировке Windows-1251 (CRLF) в папку src_ansi/ — для ручного импорта
в редактор VBA (File > Import читает файлы в кодировке Windows). Сборщик NeonBuild это не нужно."""
import os, glob
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
out = os.path.join(ROOT, 'src_ansi'); os.makedirs(out, exist_ok=True)
for f in glob.glob(os.path.join(ROOT, 'src', '*.bas')):
    s = open(f, encoding='utf-8').read().replace('\r\n', '\n').replace('\n', '\r\n')
    open(os.path.join(out, os.path.basename(f)), 'wb').write(s.encode('cp1251', errors='replace'))
    print(os.path.basename(f))
