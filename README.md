# SOD Manager

Общая оболочка для четырёх приложений СОД:

- Larix Platform
- VitroCAD
- SIGNAL / SGNL
- Project Point

Главное окно работает как лаунчер: по нажатию на карточку запускается соответствующее приложение отдельным процессом.

## Структура

```text
SOD_Manager/
├─ launcher.py
├─ app_registry.py
├─ assets/                     # ЕДИНАЯ папка всех внешних изображений
│  ├─ icon.ico
│  ├─ sun.png
│  ├─ moon.png
│  ├─ ...                      # используемые общие UI-иконки
│  ├─ sod_larix.png            # добавить вручную
│  ├─ sod_vitrocad.png         # добавить вручную
│  ├─ sod_signal.png           # добавить вручную
│  └─ sod_projectpoint.png     # добавить вручную
├─ shared/                     # общий UI-слой всех СОД
│  ├─ theme_core.py            # тема, ThemeToggle, общие цвета и ресурсы
│  └─ ui_components.py         # компактные таблицы и окно «Подробнее»
├─ apps/
│  ├─ larix/
│  ├─ vitrocad/
│  ├─ signal/
│  └─ projectpoint/
├─ requirements.txt
├─ install_dependencies.bat
├─ run.bat
└─ build_all.bat
```

У Larix, VitroCAD, SIGNAL и Project Point больше нет собственных папок с копиями одинаковых UI-иконок. Все основные приложения и лаунчер ищут изображения в корневой папке `assets`.

Larix Comment Importer хранит свои специальные UI-ресурсы внутри Python-кода/EXE, поэтому отдельная папка `apps\larix\coll\icon` ему не требуется.

## Картинки карточек СОД

Положить PNG непосредственно в `assets` под именами:

```text
sod_larix.png
sod_vitrocad.png
sod_signal.png
sod_projectpoint.png
```

Рекомендуется PNG с прозрачным фоном. Лаунчер вписывает картинку в карточку с сохранением пропорций и без обрезки.

## Запуск из исходников

1. Один раз запустить `install_dependencies.bat`.
2. Запустить `run.bat`.

## Сборка общей версии под Windows

Запустить `build_all.bat`.

Результат:

```text
dist\SOD_Manager\
├─ SOD_Manager.exe
├─ assets\
└─ apps\
   ├─ larix\Larix_Platform_Plugin.exe
   ├─ vitrocad\VitroCAD.exe
   ├─ signal\SGNL_Platform\SGNL_Platform.exe
   └─ projectpoint\ProjectPoint.exe
```

Папка `assets` остаётся внешней и общей для всех приложений. Картинки карточек и UI-ресурсы можно заменять без пересборки EXE.

## Единый UI-слой

Общие элементы интерфейса вынесены в `shared`: переключатель темы, синхронизация темы, базовые цвета/иконки и стандарт полного предпросмотра таблиц. Larix остаётся визуальным ориентиром. SIGNAL и VitroCAD используют компактный предпросмотр в основном окне и кнопку «Подробнее» для полной таблицы. Project Point использует тот же базовый стиль таблиц, сохраняя свой Excel-preview с несколькими листами.
