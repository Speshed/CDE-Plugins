# Project Point Python test application

Реализация сделана по ТЗ `ProjectPoint_LarixManager_TZ_v2.pdf`.

## Что делает

1. Показывает поле `Адрес Project Point` и кнопку `Подключить`.
2. Если схема отсутствует, добавляет `https://`.
3. Открывает адрес во встроенном Qt WebEngine.
4. Наблюдает запросы страницы и ищет `Authorization: Bearer <access_token>`.
5. Если URL запроса содержит `/api/`, определяет API base как часть URL до `/api`.
6. Если API base из URL пока определить нельзя, проверяет:
   - `<server>/ru/api`
   - `<server>/en/api`
   - `<server>/api`
7. Проверяет токен запросом:
   `GET <api-base>/Core/UsersService/GetMyProfile`
8. При HTTP 200 и JSON создаёт HTTP-сессию с:
   - `Authorization: Bearer <access_token>`
   - `Accept: application/json`
9. После подключения можно проверить пример из ТЗ:
   `GET <api-base>/Core/ProjectService/GetAll`

## Запуск без C++

Нужен только Python 3.11/3.12.

1. Установите Python с python.org.
2. При установке включите `Add python.exe to PATH`.
3. Запустите `run.bat`.

Первый запуск установит PySide6 и requests автоматически.

## Сборка EXE

После успешного запуска через `run.bat` можно запустить:

`build_exe.bat`

Готовая сборка будет в:

`dist\ProjectPointTest\ProjectPointTest.exe`

Для Qt WebEngine надёжнее распространять всю папку `ProjectPointTest`, а не один exe.
