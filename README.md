# scripts
Общее для версий:

-Запуск:
    -из командной строки -  python3 postanovka_v0.py или python3 postanovka_v1.py
    -из папки - дважды тапаем (при корреткном запуске должен обновиться файл отчета)

-Принимает файл из рабочей папки - убедиться, что файл в папке ОДИН и с расширением .xlsx

-Приводим пронумерованный обычным способом лист

-По умолчанию скрипт читает только первый(0) лист из файла, можно указать внутри скрипта параметром sheet_name=номер 
(если первый лист врачебный, а второй ассистентский). В питоне нумерация с 0, если что.

-Итоговый файл с номерами - постановка_дата_день/ночь.xlsx - появится на рабочем столе

-Автоматически фильтрует котопсов (если на кошку забиты инфекции собаки и наоборот) - высвечивает предупреждение и пишет об этом в файл-отчет.

-Фильтрует двойные ВИК+лейко и одиночные ВИК/лейко.


1)postanovka_v0 
- Итоговый файл - простейшая таблица, где первый столбик - инфекция, последующие столбики - номера.


2)postanovka_v1 
- Итоговый файл представляет собой таблицу, где номера для каждой инфекции сгруппированы по 8. (то есть практически готовая таблица раскапки)

3)postanovka_v2
- то же, что v1, но с приоритетом по инфекциям (сначала комплексы, потом одиночные)
