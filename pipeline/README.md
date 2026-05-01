# `pipeline/` — каталог скриптов-фиксеров

Каждый скрипт в этой папке решает **одну изолированную проблему** в `.docx`-файле, написанном по ГОСТ. Все скрипты:
- идемпотентны (повторный запуск ничего не ломает),
- работают через `--input X.docx --output Y.docx` (вход не модифицируется),
- не зависят друг от друга (можно запускать в любом порядке, но рекомендуется по номерам),
- работают через `python-docx` + `lxml` напрямую с OOXML, без LibreOffice/MS Office.

## Запуск

### Все шаги сразу

```bash
python pipeline/run_all.py \
    --input  "data/Уч пособ Матлаб ПРАВИЛЬНО3.docx" \
    --output "data/Уч пособ Матлаб ПРАВИЛЬНО3_GOST.docx"
```

Сохранит бэкап оригинала в `data/backups/<имя>_pre_run.docx`, прогонит 12 шагов по порядку, выпишет результат.

### Только определённые шаги

```bash
# Только восстановление номеров страниц
python pipeline/run_all.py --input X.docx --output Y.docx --only 09,10

# Все шаги, кроме объединения разделов
python pipeline/run_all.py --input X.docx --output Y.docx --skip 08
```

### Один шаг отдельно

```bash
python pipeline/01_pagesetup.py --input X.docx --output Y.docx
```

### Список всех шагов

```bash
python pipeline/run_all.py --list
```

## Список шагов

| ID  | Скрипт                            | Что делает                                                                       |
| --- | --------------------------------- | -------------------------------------------------------------------------------- |
| 01  | `01_pagesetup.py`                 | A4, поля по ГОСТ, `<w:titlePg/>`, footer1/footer2 references                    |
| 02  | `02_pagebreaks.py`                | `<w:pageBreakBefore/>` на H1 (каждая глава с новой страницы)                    |
| 03  | `03_subscripts.py`                | Подстрочный текст: `P_a → Pₐ` через `<w:vertAlign w:val="subscript"/>`          |
| 04  | `04_wrap_figures.py`              | Рисунки оборачиваются в 2-ячеечные таблицы (рисунок + подпись)                  |
| 05  | `05_normalize_tables.py`          | Data-таблицы: TNR 12pt, без жирных заголовков, autofit                          |
| 06  | `06_dedup_formulas.py`            | Удалить дубли `(N.M)(N.M)` в подписях формул                                    |
| 07  | `07_renumber_formulas.py`         | Сквозная перенумерация формул по главам: `(глава.порядок_в_главе)`              |
| 08  | `08_merge_sections.py`            | Объединить раздел 9 в раздел 10 (специфика данного проекта)                     |
| 09  | `09_footer_pagenumber.py`         | Канонический PAGE-field в `footer1.xml` (TNR 14pt, `\* MERGEFORMAT`)            |
| 10  | `10_remove_hidewhitespace.py`     | Снять `<w:doNotDisplayPageBoundaries/>` — Word'овский Hide White Space          |
| 11  | `11_turbo_postpass.py`            | Постпроход: пустые pbb-параграфы, TOC stale, «ПРОПАЛА ФОРМУЛА», updateFields    |
| 12  | `12_clear_highlights.py`          | Снять все жёлтые подсветки (применять последним)                                |

Полное описание проблем, которые решает каждый шаг (симптом в Word → причина в OOXML → как проверить) — в [`../PROBLEM_CATALOG.md`](../PROBLEM_CATALOG.md).

## Принципы добавления нового шага

1. Создать `pipeline/NN_<имя>.py` с интерфейсом `--input`/`--output` и функцией `process(input_path, output_path)`.
2. Добавить запись в список `STEPS` в `pipeline/run_all.py`.
3. Добавить строку в таблицу выше и описание в `PROBLEM_CATALOG.md`.
4. Скрипт должен быть **идемпотентным** — повторный запуск не должен ничего ломать.
5. Скрипт **не должен** ничего удалять из `data/`, кроме создания бэкапа в `data/backups/`.

## Что НЕ делать

- **Не редактировать** `.docx` файлы вручную через `cat`/`echo`/`awk` — только через скрипты.
- **Не создавать** разовые диагностические скрипты в корне репо или в `pipeline/`. Если нужна диагностика — пиши inline в `python -c` или ad-hoc, не коммить.
- **Не править** `data/*.docx` напрямую без бэкапа в `data/backups/`.
