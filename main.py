import sqlalchemy
from sqlalchemy import create_engine # pip install SQLAlchemy
from sqlalchemy.engine import URL
import pypyodbc # pip install pypyodbc
import pandas as pd # pip install pandas
import matplotlib.pyplot as plt


SERVER_NAME = 'WIN-JGVGGFOF1LR\SQLEXPRESS'
DATABASE_NAME = 'Excel'

excel_file = 'Example.xlsx'

#Строка подключения к бд
connection_string = f"""
    DRIVER={{SQL Server}};
    SERVER={SERVER_NAME};
    DATABASE={DATABASE_NAME};
    Trusted_Connection=yes;
"""

#Подключение к локальной базе данных
connection_url = URL.create('mssql+pyodbc', query={'odbc_connect': connection_string})
enigne = create_engine(connection_url, module=pypyodbc)

#Чтение данных из excel
excel_file = pd.read_excel(excel_file, sheet_name=None)

#Копирование данных из excel в базу данных
d = excel_file['Таблица 1']

d.to_sql('Таблица 1', enigne, if_exists='replace', index=False,
         dtype={'Код города': sqlalchemy.types.Integer(),
                'Название города': sqlalchemy.types.NVARCHAR(length=100),
                'Кол-во': sqlalchemy.types.Integer(),
                'Цена': sqlalchemy.types.Integer(),
                'Реализация': sqlalchemy.types.Integer(),
                'Статус': sqlalchemy.types.NVARCHAR(length=100),
                'Место по объёму реализации': sqlalchemy.types.Integer()})

d = excel_file['Таблица 2']

d.to_sql('Таблица 2', enigne, if_exists='replace', index=False,
         dtype={'Код города': sqlalchemy.types.Integer(),
                'Название города': sqlalchemy.types.NVARCHAR(length=100),
                'Сумма': sqlalchemy.types.NVARCHAR(length=100)})

d = excel_file['Таблица 3']

d.to_sql('Таблица 3', enigne, if_exists='replace', index=False,
         dtype={'Условие': sqlalchemy.types.NVARCHAR(length=100),
                'Статус': sqlalchemy.types.NVARCHAR(length=100)})

#Курсор
connection_to_db = pypyodbc.connect(connection_string)

cursor = connection_to_db.cursor()

#Задание:Посчитайте суммы реализации по каждому городу из таблицы 1 и подставьте значения в таблицу 2
#Реализация SQL запроса
cursor.execute('update [Таблица 1] set [Таблица 1].[Название города] = [Таблица 2].[Название города] from [Таблица 1],[Таблица 2] where [Таблица 1].[Код города] = [Таблица 2].[Код города]')
connection_to_db.commit()

#Задание:Подставьте в таблицу 1 названия городов из таблицы 2 в соответствии с кодом города
#Реализация SQL запроса
cursor.execute('update [Таблица 2] set [Таблица 2].[Сумма] = c.[Сумма реализации] from (Select [Таблица 1].[Код города],SUM([Таблица 1].[Реализация]) as [Сумма реализации] from [Таблица 1] group by [Таблица 1].[Код города]) as c where [Таблица 2].[Код города] = c.[Код города]')
connection_to_db.commit()

#Задание:Проставьте в таблице 1 статус каждой строки из таблицы 3 в зависимости от суммы реализации
#Реализация SQL запроса
cursor.execute("update [Таблица 1] set [Таблица 1].[Статус] = 'А' where [Таблица 1].[Реализация] > 5000000")
connection_to_db.commit()
cursor.execute("update [Таблица 1] set [Таблица 1].[Статус] = 'Б' where [Таблица 1].[Реализация] >= 1000000 and [Таблица 1].[Реализация]  <= 5000000")
connection_to_db.commit()
cursor.execute("update [Таблица 1] set [Таблица 1].[Статус] = 'В' where [Таблица 1].[Реализация] < 1000000")
connection_to_db.commit()

#Задание:Проставьте в таблице 1 статус каждой строки из таблицы 3 в зависимости от суммы реализации
#Реализация SQL запроса
cursor.execute('update [Таблица 1] set [Таблица 1].[Место по объёму реализации] = e.[Место] from (select [Таблица 1].[Код города],[Таблица 1].[Реализация], ROW_NUMBER() over (order by [Таблица 1].[Реализация] DESC) as [Место] from [Таблица 1]) as e where [Таблица 1].[Реализация] = e.[Реализация]')
connection_to_db.commit()

#Вывод полученного результат в excel документ Output1.xlsx
cursor.execute('Select * from [Таблица 1]')

data = cursor.fetchall()

data1 = pd.DataFrame(data, columns=['Код города', 'Название', 'Кол-во', 'Цена', 'Реализация', 'Статус', 'Место по объёму реализации'])

data1.to_excel('Output1.xlsx', index=False)

#Задание:Найти города занявшие первые три места по объёмам реализации(табл.1)
#Реализация SQL запроса
cursor.execute('select Top 3 [Таблица 1].[Название города] from [Таблица 1] order by [Место по объёму реализации] ASC ')
#Вывод результатов в консоль
data = cursor.fetchall()
data1 = pd.DataFrame(data, columns=['Место'])
print('Города занявшие первые три места по объёмам реализации')
print(data1)
print()

#Задание:Найти количество Статусов А для городов Челябинск и Мурманск
#Реализация SQL запроса
cursor.execute("select [Таблица 1].[Название города],count(*) as [Кол-во] from [Таблица 1] where ([Таблица 1].[Название города] = 'Челябинск' or [Таблица 1].[Название города] = 'Мурманск') and [Таблица 1].[Статус] = 'А' group by [Таблица 1].[Название города]")
#Вывод результатов в консоль
data = cursor.fetchall()
data1 = pd.DataFrame(data, columns=['Город', 'Кол-во статусов А'])
print('Количество Статусов А для городов Челябинск и Мурманск')
print(data1)

#Задание:Создать диаграмму (гистаграмма линейчатая) по табл. 2 (название города, сумма), добавить подписи данных.
#Реализация
cursor.execute("select * from [Таблица 2]")

results = cursor.fetchall()

arrays1 = [0] * 10
arrays2 = [0] * 10

for i in range(0, len(results)):
    arrays1[i] = results[i][1]

for i in range(0, len(results)):
    arrays2[i] = int(results[i][2])

x = arrays1
y = arrays2

plt.figure(figsize=(15, 15))
plt.xlabel("Город")
plt.ylabel("Сумма")

plt.bar(x, y)

#Вывод полученного графика
plt.show()

#Чтение данных из excel
excel_file = pd.read_excel('Example2.xlsx', sheet_name=None)

#Копирование данных из excel в базу данных
d = excel_file['Таблица 1']

d.to_sql('Таблица 4', enigne, if_exists='replace', index=False,
         dtype={'Регион': sqlalchemy.types.NVARCHAR(length=100),
                'Клиент': sqlalchemy.types.NVARCHAR(length=100),
                'Адрес': sqlalchemy.types.NVARCHAR(length=200),
                'Сумма продаж': sqlalchemy.types.Integer()})

d = excel_file['Таблица 2']

d.to_sql('Таблица 5', enigne, if_exists='replace', index=False,
         dtype={'Клиент': sqlalchemy.types.NVARCHAR(length=100),
                'Адрес': sqlalchemy.types.NVARCHAR(length=200),
                'Продукт': sqlalchemy.types.NVARCHAR(length=100),
                'Год': sqlalchemy.types.Integer(),
                'Месяц': sqlalchemy.types.NVARCHAR(length=100),
                'Сумма продаж': sqlalchemy.types.Integer()})

#Задание:В табл. 1 в столбец D проставить суммы продаж за Июль 2011 года по рыбе
#Реализация SQL запроса
cursor.execute("update [Таблица 4] set [Таблица 4].[Сумма продаж] = [Таблица 5].[Сумма продаж] from [Таблица 5] where [Таблица 4].[Адрес] = [Таблица 5].[Адрес] and [Таблица 5].[Год] = 2011 and [Таблица 5].[Месяц] = 'Июль' and [Таблица 5].[Продукт] = 'рыба'")
connection_to_db.commit()

#Создать сводную табл. 1 по табл.1 на новой вкладке (клиент, адрес), добавить срез на регион
#Реализация

#Чтение данных из базы данных
cursor.execute('Select * from [Таблица 4]')

data = cursor.fetchall()

#Запись данных в excel документ Output2.xlsx
data1 = pd.DataFrame(data, columns=['Регион', 'Клиент', 'Адрес', 'Сумма продаж'])

data1.to_excel('Output2.xlsx', sheet_name='Таблица 1', index=False)

df = pd.read_excel('Output2.xlsx', sheet_name='Таблица 1')

#Создание сводной таблицы
dh = pd.pivot_table(df, index=["Клиент"], columns=["Адрес"])

with pd.ExcelWriter('Output2.xlsx', mode='a', if_sheet_exists='replace') as writer:
    dh.to_excel(writer, sheet_name='Таблица 2')

#Закрыть подключение к базе данных
connection_to_db.close()
