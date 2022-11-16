### Трансформация EXCEL-формы отчета в txt для загрузки в PI или в XML.

Создание инсталлятора
```
pyinstaller  --clean ExcelTransform.spec

Для памяти создание spec-файла:
pyi-makespec --onefile --name ExcelTransform main.py
```
Вызовы программы:
```
ExcelTransform --help

ExcelTransform -xml <имя отчета> 
```
