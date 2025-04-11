
# Windows Forms Application: Электронные сертификаты

Приложение создано для удобного отображения, добавления и удаления электронных сертификатов компании.

## Технологии
- Lang: C#, .NETFramework (v4.7.2)
- IDE: Visual Studio 2022 Community
- СУБД: MS SQL Server

## Основной функционал
- Просмотр существующих сертификатов
- Добавление сертификатов
- Удаление сертфиикатов
- Экспорт данных в Excel

## Особенности
- Подключение к БД через файл **DBUtils.cs**
- Для второй версии программы с экспортом в Excel необходима лицензионная версия Excel
- Директория **package** содержит пакет Microsoft.Office.Interop.Excel (v15.0+). Его можно подгрузить самостоятельно через встроенный в VS пакетный менеджер NuGet.

---

## EN Version
# Windows Forms Application: Electronic Certificates

This application is designed for convenient viewing, adding, and deleting the company’s electronic certificates.

## Technologies
- Language: C#, .NET Framework (v4.7.2)
- IDE: Visual Studio 2022 Community
- DBMS: MS SQL Server

## Core Features
- View existing certificates
- Add new certificates
- Delete certificates
- Export data to Excel

## Notes
- Database connection is implemented in **DBUtils.cs**
- For the version with Excel export, a licensed version of Excel is required
- The **packages** directory includes the Microsoft.Office.Interop.Excel (v15.0+) package. You can also install it manually via the built-in NuGet Package Manager in Visual Studio.