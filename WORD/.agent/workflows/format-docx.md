---
description: Форматирование Word-документа по эталону или стандарту ГОСТ
---

# Форматирование документа Word

Этот воркфлоу применяет стиль эталонного документа к рабочему файлу, или форматирует по стандарту ГОСТ.

// turbo-all

1. Определи рабочий файл — спроси у пользователя путь к `.docx` файлу или найди его в `data/`.
2. Определи источник стилей:
   - Есть ли эталонный файл? Проверь `data/` на наличие файла со словом «эталон» в имени.
   - Если эталона нет — используй стандарт ГОСТ.
3. Прочитай текущее состояние рабочего файла:
   ```powershell
   & "C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe" execution/docx_reader.py --file "путь/к/файлу.docx" --output text --styles --page-setup
   ```
4. Создай резервную копию и примени форматирование:
   - С эталоном:
     ```powershell
     & "C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe" execution/docx_formatter.py --source "путь/к/эталону.docx" --target "путь/к/рабочему.docx"
     ```
   - Без эталона (ГОСТ):
     ```powershell
     & "C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe" execution/docx_formatter.py --target "путь/к/рабочему.docx" --standard gost
     ```
5. Исправь десятичные разделители (точки → запятые):
   ```powershell
   & "C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe" execution/docx_writer.py --file "путь/к/рабочему.docx" --fix-decimals --no-backup
   ```
6. Прочитай результат и покажи пользователю изменения:
   ```powershell
   & "C:\Users\New-life\AppData\Local\Programs\Python\Python313\python.exe" execution/docx_reader.py --file "путь/к/рабочему.docx" --output text --styles
   ```
7. Сообщи результат пользователю: что изменилось, путь к файлу и к резервной копии.
