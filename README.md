# Что это
Скрипт для MS Office, позволяющий преобразовать табличное представление выгрузки отчётов по инцидентам в вид презентации для встречи и обсуждения

# Как использовать
- Открыть файл Excel с таблицей отчёта
- Вызвать окно редактора VB `Alt + F11`
- В меню `Tools` -> `References...` отметить пункт `Microsoft PowerPoint 16.0 Object Library` (или аналог) и нажать ОК
- В меню `File` -> `Import File...` открыть файл из репозитория `Analysis_incidents.bas`
- Найти открывшийся модуль в проводнике проектов слева и запустить единственную процедуру
В результате будет создан экземпляр презентации с упорядоченными данным из строк отчёта
