# Гайд для AI-ассистента: Word Document Expert

> **Этот файл — инструкция для тебя, AI-ассистента.** Загрузив его, ты получаешь полноценную систему работы с Word-документами по тематике электроснабжения.

---

## Быстрый старт

Ты находишься в воркспейсе с 3-слойной архитектурой:

```
WORD/
├── GEMINI.md                  ← Прочитай ПЕРВЫМ — твои главные инструкции
├── GUIDE_FOR_AI.md            ← Этот файл (ознакомление)
├── .agent/
│   ├── skills/
│   │   └── word-document-expert/
│   │       └── SKILL.md       ← Скилл: ВСЕ правила работы с документами
│   └── workflows/
│       ├── format-docx.md     ← Воркфлоу: форматирование по ГОСТ/эталону
│       └── extract-formulas.md ← Воркфлоу: извлечение формул в LaTeX
├── directives/
│   └── word_processing.md     ← SOP: пошаговая процедура работы
├── execution/                 ← Python-скрипты (НЕ пиши сам — запускай эти)
│   ├── docx_reader.py         ← Чтение .docx
│   ├── docx_writer.py         ← Редактирование .docx
│   ├── docx_formatter.py      ← Форматирование по эталону/ГОСТ
│   ├── latex_converter.py     ← Конвертация формул в LaTeX
│   └── utils/docx_utils.py    ← Утилиты (библиотека)
├── data/                      ← Сюда кладут рабочие и эталонные файлы
└── .tmp/                      ← Временные файлы (можно удалять)
```

---

## Что ты должен сделать в первую очередь

1. **Прочитай `GEMINI.md`** — там описана архитектура и принципы работы
2. **Прочитай `.agent/skills/word-document-expert/SKILL.md`** — все правила, переменные, режимы работы
3. **Прочитай `directives/word_processing.md`** — пошаговый алгоритм выполнения задач

---

## Как работать

### Принцип: ты НЕ пишешь код сам
Ты — оркестратор. Твоя задача — принять запрос пользователя и вызвать нужные скрипты из `execution/`. Не пытайся редактировать .docx вручную или писать код с нуля.

### Python на этой машине
```powershell
& "C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe" скрипт.py аргументы
```
Зависимость `python-docx` уже установлена.

### Типичные сценарии

**Сценарий 1: «Прочитай мой документ»**
```powershell
& "C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe" execution/docx_reader.py --file "data/файл.docx" --output text --all
```

**Сценарий 2: «Отформатируй по ГОСТу»**
```powershell
& "C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe" execution/docx_formatter.py --target "data/файл.docx" --standard gost
```

**Сценарий 3: «Сделай как в эталоне»**
```powershell
& "C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe" execution/docx_formatter.py --source "data/эталон.docx" --target "data/рабочий.docx"
```

**Сценарий 4: «Замени текст»**
```powershell
& "C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe" execution/docx_writer.py --file "data/файл.docx" --replace "старый текст" "новый текст"
```

**Сценарий 5: «Извлеки формулы в LaTeX»**
```powershell
& "C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe" execution/latex_converter.py --scan "data/файл.docx"
```

**Сценарий 6: «Исправь десятичные разделители»**
```powershell
& "C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe" execution/docx_writer.py --file "data/файл.docx" --fix-decimals
```

**Сценарий 7: «Рерайт текста (антиплагиат)»**
1. Извлечь текст из .docx: `docx_reader.py` → `.tmp/текст.txt`
2. Переписать текст в `.tmp/rewritten_*.md` (markdown)
3. Собрать обратно в .docx: `execution/build_docx.py` → `data/файл_рерайт.docx`

**Сценарий 8: «Сборка .docx из markdown по ГОСТ»**
```powershell
& "C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe" execution/build_docx.py
```

> **Важно:** Папка `.tmp/` используется для временных промежуточных файлов (например извлеченный текст). Итоговые документы всегда сохраняются в `data/`. Скрипты запускаются из `execution/`.

---

## Критические правила (ЗАПОМНИ!)

1. **Эталонный файл НЕЛЬЗЯ менять** — если в имени файла слово «эталон», его только ЧИТАЕМ
2. **Формулы не трогай** — исходные формулы 1 в 1, только подставляй значения
3. **Двойной формат формул** — обычная запись + LaTeX-блок кода
4. **Точки → запятые** в числах: `32.09` → `32,09`
5. **Только русский язык** — все обозначения, единицы, текст на русском
6. **Всегда создавай бэкап** перед редактированием (скрипты делают это автоматически)
7. **Формулы бери из PDF**, если есть — LaTeX из .docx ты не видишь
8. **Оборачивать переменные в LaTeX** — любые символы с нижним индексом в тексте (например $T_s$, $U_{\text{н}}$) всегда оборачивай в `$ ... $` для авто-рендеринга в настоящие Word-формулы. В скрипт сборки вшита защита, но лучше сразу писать `$T_s$`.

---

## Обработка ошибок

| Ошибка | Решение |
|---|---|
| `ModuleNotFoundError: docx` | `pip install python-docx` (через Python313) |
| `python` не тот | Используй полный путь: `C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe` |
| Файл не найден | Проверь путь; файлы должны лежать в `data/` |
| Эталон не определён | Ищи файл со словом «эталон» в имени |

---

## Воркфлоу (быстрые сценарии)

Если пользователь просит типовую операцию — смотри файлы в `.agent/workflows/`:
- `format-docx.md` — форматирование документа
- `extract-formulas.md` — извлечение формул

Воркфлоу содержат готовые пошаговые инструкции с командами.

---

## Самоулучшение

Когда обнаруживаешь новые паттерны или ошибки:
1. Исправь скрипт в `execution/`
2. Обнови директиву в `directives/`
3. Система стала сильнее

Удачной работы!
