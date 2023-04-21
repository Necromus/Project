create database Excel

use Excel

create table [Таблица 1](
[Код города] int,
[Название города] nvarchar(100),
[Кол-во] int,
[Цена] int,
[Реализация] int,
[Статус] nvarchar(100),
[Место по объёму реализации] int
)

create table [Таблица 2](
[Код города] int,
[Название города] nvarchar(100),
[Сумма] int
)

create table [Таблица 3](
[Условие] nvarchar(100),
[Статус] nvarchar(100)
)

create table [Таблица 4](
[Регион] nvarchar(100),
[Клиент] nvarchar(100),
[Адрес] nvarchar(200),
[Сумма продаж] float
)

create table [Таблица 5](
[Клиент] nvarchar(100),
[Адрес] nvarchar(200),
[Продукт] nvarchar(100),
[Год] int,
[Месяц] nvarchar(100),
[Сумма продаж] float
)
