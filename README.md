# word

Рабочее пространство для двух Word-документов под ГОСТ 7.32-2017:

- `WORD/data/Исходник/БР.docx` + `БР_титульник.docx` — выпускная работа.
- `WORD/data/Готово/Уч пособ Матлаб ПРАВИЛЬНО3.docx` — учебное пособие по Matlab/Simulink.

## Структура репо

```
.
├── README.md                      ← этот файл
├── GUIDE_FOR_AI.md                ← инструкция для AI-ассистента
├── PROBLEM_CATALOG.md             ← каталог «симптом → причина → скрипт»
├── RULES_GOST.md                  ← выжимка ГОСТ 7.32-2017
├── .gitignore
├── WORD/data/
│   ├── Исходник/                  ← исходные docx (входы пайплайна)
│   ├── Готово/                    ← результаты прогонов
│   ├── Скриншоты/                 ← пользовательские скрины проблем
│   ├── backups/                   ← локальные бэкапы (gitignored)
│   └── задание.md                 ← текущая задача
└── pipeline/
    ├── README.md
    ├── run_all.py                 ← оркестратор 12 шагов
    ├── 01_pagesetup.py
    ├── 02_pagebreaks.py
    ├── 03_subscripts.py
    ├── 04_wrap_figures.py
    ├── 05_normalize_tables.py
    ├── 06_dedup_formulas.py
    ├── 07_renumber_formulas.py
    ├── 08_merge_sections.py
    ├── 09_footer_pagenumber.py
    ├── 10_remove_hidewhitespace.py
    ├── 11_turbo_postpass.py
    ├── 12_clear_highlights.py
    └── md_to_docx/                ← Markdown → DOCX (сборка / append)
        ├── README.md
        ├── build_docx.py
        ├── word_config.py
        ├── word_utils.py
        ├── latex_converter.py
        └── MML2OMML.XSL
```

## Три основных воркфлоу

### 1. Привести существующий DOCX под ГОСТ (12 шагов)

```bash
python pipeline/run_all.py \
    --input  "WORD/data/Исходник/БР.docx" \
    --output "WORD/data/Готово/БР_GOST.docx"
```

### 2. Собрать новый DOCX из Markdown-черновика

```bash
python pipeline/md_to_docx/build_docx.py \
    -i WORD/data/Исходник/draft.md \
    -o WORD/data/Готово/draft.docx \
    --fast
```

### 3. Дописать MD-фрагмент в конец готового DOCX

```bash
python pipeline/md_to_docx/build_docx.py \
    -i WORD/data/Исходник/new_chapter.md \
    -o WORD/data/Готово/БР_GOST.docx \
    --append WORD/data/Готово/БР_GOST.docx \
    --fast
# затем привести к ГОСТ
python pipeline/run_all.py \
    -i WORD/data/Готово/БР_GOST.docx \
    -o WORD/data/Готово/БР_GOST.docx
```

## Окружение

```bash
pip install python-docx lxml mistune pygments
apt-get install -y libreoffice poppler-utils   # для рендера в PDF
```

## Документация

- [`GUIDE_FOR_AI.md`](GUIDE_FOR_AI.md) — как с этим работать AI-ассистенту (что читать, что НЕ делать, как добавить новый шаг).
- [`PROBLEM_CATALOG.md`](PROBLEM_CATALOG.md) — реестр 14 решённых ГОСТ-проблем с симптомом, причиной в OOXML и скриптом-фиксером.
- [`pipeline/README.md`](pipeline/README.md) — описание каждого шага пайплайна.
- [`pipeline/md_to_docx/README.md`](pipeline/md_to_docx/README.md) — как пользоваться MD→DOCX сборщиком.
- [`RULES_GOST.md`](RULES_GOST.md) — выжимка из ГОСТ 7.32-2017 / 2.105-2019.
