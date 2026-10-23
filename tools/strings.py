# Строки интерфейса плагина: ключ -> (русский, английский).
# {0}, {1}... — подстановки (функция LF в NeonUI). \n — перенос строки.
# Новый язык: добавить колонку здесь и ветку в gen_lang.py.
STRINGS = {
    'app_title':     ('Неон: центральная линия', 'Neon: centerline'),
    'mode_mono':     ('Моно', 'Mono'),
    'mode_script':   ('Курсив', 'Script'),
    'err_nodoc':     ('Нет открытого документа.', 'No document is open.'),
    'err_nosel':     ('Выделите контуры букв.', 'Select the letter outlines.'),
    'err_noclosed':  ('В выделении нет замкнутых контуров.', 'The selection contains no closed outlines.'),
    'err_fail':      ('Ошибка: {0}', 'Error: {0}'),
    'err_nomodule':  ('Модуль «{0}» не найден в проекте плагина. Переустановите плагин.',
                      'Module "{0}" was not found in the plugin project. Please reinstall the plugin.'),
    'done':          ('Готово ({0}, фреза {1} мм): объектов {2}, линий создано {3}.\nВремя {4} с (расчёт {5} с).',
                      'Done ({0}, cutter {1} mm): {2} objects, {3} lines created.\nTime {4} s (computing {5} s).'),
    'prompt_tool':   ('Диаметр фрезы / рассеивателя, мм:\n(конец линии ставится на D/2 от края контура)',
                      'Cutter / diffuser diameter, mm:\n(a free line end is placed D/2 from the outline edge)'),
    'bad_tool':      ('Неверный диаметр.', 'Invalid diameter.'),
    'settings_title':('Настройки', 'Settings'),
    'lbl_tool':      ('Диаметр фрезы, мм:', 'Cutter diameter, mm:'),
    'lbl_tool_hint': ('Свободный конец линии ставится на D/2 от края контура.',
                      'A free line end is placed D/2 from the outline edge.'),
    'chk_ask':       ('Спрашивать диаметр при каждом запуске', 'Ask for the diameter on every run'),
    'lbl_lang':      ('Язык интерфейса:', 'Interface language:'),
    'lang_auto':     ('Авто (как в Windows)', 'Auto (as Windows)'),
    'lang_ru':       ('Русский', 'Русский'),
    'lang_en':       ('English', 'English'),
    'chk_upd':       ('Проверять обновления раз в день', 'Check for updates once a day'),
    'lbl_version':   ('Версия {0}', 'Version {0}'),
    'btn_ok':        ('OK', 'OK'),
    'btn_cancel':    ('Отмена', 'Cancel'),
    'btn_update':    ('Проверить обновления', 'Check for updates'),
    'btn_about':     ('О программе', 'About'),
    'lang_changed':  ('Язык будет применён к сообщениям сразу, к окну настроек — при следующем открытии.',
                      'The language applies to messages right away and to this window next time it opens.'),
    'upd_none':      ('У вас последняя версия ({0}).', 'You have the latest version ({0}).'),
    'upd_found':     ('Доступна новая версия {0} (у вас {1}).\n\nЧто нового:\n{2}\n\nСкачать и установить? '
                      'Новая версия заработает после перезапуска CorelDRAW.',
                      'A new version {0} is available (you have {1}).\n\nWhat\'s new:\n{2}\n\nDownload and install? '
                      'The new version will start working after CorelDRAW is restarted.'),
    'upd_ready':     ('Обновление {0} скачано. Оно установится автоматически, когда вы закроете CorelDRAW.',
                      'Update {0} has been downloaded. It will be installed automatically when you close CorelDRAW.'),
    'upd_pending':   ('Обновление {0} уже скачано и установится после закрытия CorelDRAW.',
                      'Update {0} is already downloaded and will be installed when CorelDRAW is closed.'),
    'upd_err':       ('Не удалось проверить обновления: {0}', 'Could not check for updates: {0}'),
    'upd_bad':       ('Скачанный файл обновления повреждён (контрольная сумма не совпала). Обновление отменено.',
                      'The downloaded update is damaged (checksum mismatch). The update was cancelled.'),
    'upd_nourl':     ('Адрес обновлений не настроен.', 'The update address is not configured.'),
    'upd_nogms':     ('Не найден файл плагина (.gms) для замены.', 'The plugin file (.gms) to replace was not found.'),
    'about':         ('NeonCenterline {0}\n\nЦентральная линия по контурам букв для фрезеровки канала под гибкий неон. '
                      'Кнопки: «Моно» — для моноширинных (неоновых) шрифтов, «Курсив» — для рукописных.\n\n'
                      '© 2026 SAB:CORE. Бесплатно для любого использования, в том числе коммерческого. '
                      'Продажа и перепаковка плагина запрещены (см. LICENSE).',
                      'NeonCenterline {0}\n\nCenterline from letter outlines for milling flex-neon channels. '
                      'Buttons: "Mono" for monoline (neon) fonts, "Script" for handwritten fonts.\n\n'
                      '© 2026 SAB:CORE. Free for any use, including commercial. '
                      'Selling or repackaging the plugin is prohibited (see LICENSE).'),
    'about_donate':  ('\n\nПлагин бесплатный. Если он вам помогает — можно поблагодарить автора. Открыть страницу?',
                      '\n\nThe plugin is free. If it helps you, you can support the author. Open the page?'),
    'fallback_lang': ('Язык интерфейса: Да — русский, Нет — English, Отмена — авто.',
                      'Interface language: Yes — Russian, No — English, Cancel — auto.'),
    'fallback_upd':  ('Проверять обновления раз в день?', 'Check for updates once a day?'),
    'fallback_ask':  ('Спрашивать диаметр при каждом запуске?', 'Ask for the diameter on every run?'),
}
