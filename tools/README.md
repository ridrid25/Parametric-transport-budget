# Инструменты (`tools/`)

Вспомогательные скрипты для проверки модели и генерации образцов данных — **без Google Sheets**. Логика та же, что в `apps-script/Code.gs`: скрипты подставляют заглушку Apps Script API (`SpreadsheetApp`, `Charts`), выполняют `Code.gs` и работают с результатом.

## `harness.js` — самопроверка модели

```bash
node tools/harness.js
```

Прогоняет `apps-script/Code.gs` через заглушку, строит бюджет/сценарии/разбивки и печатает контрольные суммы: годовую выручку и сходимость сумм по дистанции и сегментам с общим бюджетом (расхождение должно быть ~0). Пишет полный дамп в `tools/out.data.json` (в `.gitignore`).

Используйте после правок в `Code.gs`, чтобы убедиться, что расчёты сходятся.

## `export_inputs.js` — выгрузка входных данных

```bash
node tools/export_inputs.js
```

Генерирует входные листы (Vehicles, Seasonality, Params, Customers, Бюджет) и сохраняет их в `tools/out.inputs.json` (в `.gitignore`) — вход для `build_sample_data.py`.

## `build_sample_data.py` — сборка образцов

```bash
node tools/export_inputs.js          # сначала создать out.inputs.json
python3 tools/build_sample_data.py   # затем собрать образцы (нужен openpyxl для xlsx)
```

Пересобирает файлы в `samples/`:
- `transport_budget_data.xlsx` — многовкладочный (Инструкция + входные данные, пример 50 клиентов);
- `transport_budget_data.csv` — те же данные одним листом с секциями, готово для Google Sheets.

## `legacy/`

Ранние версии, оставлены для истории (не соответствуют текущей модели):
- `generate.py` — первоначальный генератор статичного xlsx (до перехода на самогенерацию данных в `Code.gs`);
- `preview.html` — первый статичный снимок дашборда (актуальная интерактивная версия — `dashboard/dashboard.html`).
