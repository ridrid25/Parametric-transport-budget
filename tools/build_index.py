#!/usr/bin/env python3
"""Собирает standalone index.html (для GitHub Pages) из фрагмента dashboard/dashboard.html.

dashboard.html — фрагмент без <head> (так его ждёт панель артефактов claude.ai,
которая сама оборачивает его в документ с <meta viewport>). Для GitHub Pages
нужен полный HTML-документ с charset и viewport, иначе на телефоне не включится
адаптивная вёрстка. Этот скрипт добавляет корректный <head> и body-обёртку.

Запуск:  python3 tools/build_index.py   (после любых правок dashboard.html)
"""
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, 'dashboard', 'dashboard.html')
OUT = os.path.join(ROOT, 'index.html')

frag = open(SRC, encoding='utf-8').read()

# Вытащить <title> из фрагмента, чтобы поставить его в <head> (а не в <body>).
m = re.match(r'\s*<title>(.*?)</title>\s*', frag, re.S)
title = m.group(1).strip() if m else 'Транспортный бюджет — интерактивный дашборд'
body = frag[m.end():] if m else frag

html = (
    '<!DOCTYPE html>\n'
    '<html lang="ru">\n'
    '<head>\n'
    '<meta charset="utf-8">\n'
    '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
    '<title>' + title + '</title>\n'
    '<style>\n'
    '  *, *::before, *::after { box-sizing: border-box; }\n'
    '  html, body { margin: 0; padding: 0; }\n'
    '  body { background: #f9f9f7; }\n'
    '  @media (prefers-color-scheme: dark) { body { background: #0d0d0d; } }\n'
    '</style>\n'
    '</head>\n'
    '<body>\n'
    + body.strip() + '\n'
    '</body>\n'
    '</html>\n'
)

open(OUT, 'w', encoding='utf-8').write(html)
print('index.html собран (%d байт) из dashboard/dashboard.html' % len(html))
