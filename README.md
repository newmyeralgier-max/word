# word

Рабочее пространство для двух Word-документов под ГОСТ 7.32-2017:

- **`БР.docx`** + **`БР_титульник.docx`** — выпускная работа.
- **`data/Уч пособ Матлаб ПРАВИЛЬНО3.docx`** — учебное пособие по Matlab/Simulink.

## Как пользоваться

```bash
# Прогнать все 12 шагов автоматического форматирования
python pipeline/run_all.py \
    --input  "data/Уч пособ Матлаб ПРАВИЛЬНО3.docx" \
    --output "data/Уч пособ Матлаб ПРАВИЛЬНО3_GOST.docx"
```

Бэкап оригинала автоматически кладётся в `data/backups/<имя>_pre_run.docx`.

## Что где

| Файл/папка                        | Что это                                                                  |
| --------------------------------- | ------------------------------------------------------------------------ |
| [`GUIDE_FOR_AI.md`](GUIDE_FOR_AI.md) | Инструкция для AI-ассистента (как работать с репозиторием).             |
| [`PROBLEM_CATALOG.md`](PROBLEM_CATALOG.md) | Каталог «симптом в Word → причина в OOXML → скрипт-фиксер».      |
| [`RULES_GOST.md`](RULES_GOST.md)  | Выжимка из ГОСТ 7.32-2017 / 2.105-2019.                                  |
| [`pipeline/`](pipeline/)          | 12 скриптов-фиксеров + `run_all.py` + `README.md`.                       |
| `data/`                           | Рабочие .docx, скриншоты, задание, бэкапы.                               |
| `Practice/`                       | Matlab/Simulink-исходники (контент пособия).                             |

## Окружение

```bash
pip install python-docx lxml
apt-get install -y libreoffice poppler-utils  # для рендера в PDF
```

## Как добавить новый шаг (когда находишь новую проблему)

1. Создай `pipeline/NN_<имя>.py` с интерфейсом `--input X.docx --output Y.docx`.
2. Добавь его в список `STEPS` в `pipeline/run_all.py`.
3. Опиши проблему в `PROBLEM_CATALOG.md` и в таблице `pipeline/README.md`.
4. Открой PR.

См. [`GUIDE_FOR_AI.md`](GUIDE_FOR_AI.md) для подробного workflow и принципов.
