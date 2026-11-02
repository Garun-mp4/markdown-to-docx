---
name: Markdown to DOCX
description: Monospaced document conversion workspace adapted from the user-provided OpenCode design system.
colors:
  primary: "#201d1d"
  ink: "#201d1d"
  ink-deep: "#0f0000"
  charcoal: "#302c2c"
  body: "#424245"
  mute: "#646262"
  ash: "#9a9898"
  canvas: "#fdfcfc"
  surface-soft: "#f8f7f7"
  surface-card: "#f1eeee"
  paper: "#ffffff"
  outline: "rgba(15, 0, 0, 0.12)"
  outline-variant: "#646262"
  tertiary: "#007aff"
  accent-hover: "#0056b3"
  accent-active: "#004085"
  focus: "#0056b3"
  accent-selection: "#dcecff"
  focus-soft: "#f2f7ff"
  ink-muted-inverse: "#d4d0d0"
  danger: "#a50011"
  danger-bg: "#fff2f1"
  danger-dim: "#d99a9e"
  warning: "#714800"
  warning-bg: "#fff8ea"
  warning-dim: "#dfc98f"
  success: "#1f6538"
  success-bg: "#f0f7f1"
  success-dim: "#a9c4ae"
  preview-line: "#dedbd5"
  outline-dim: "#c5c1c1"
  backdrop: "rgba(15, 0, 0, 0.46)"
typography:
  display:
    fontFamily: '"Berkeley Mono", "JetBrains Mono", "IBM Plex Mono", ui-monospace, "Cascadia Code", "Segoe UI Mono", Consolas, monospace'
    fontSize: "36px"
    fontWeight: 700
    lineHeight: 1.2
    letterSpacing: "-0.035em"
  compact-display:
    fontFamily: '"Berkeley Mono", "JetBrains Mono", "IBM Plex Mono", ui-monospace, "Cascadia Code", "Segoe UI Mono", Consolas, monospace'
    fontSize: "31px"
    fontWeight: 700
    lineHeight: 1.2
  mobile-display:
    fontFamily: '"Berkeley Mono", "JetBrains Mono", "IBM Plex Mono", ui-monospace, "Cascadia Code", "Segoe UI Mono", Consolas, monospace'
    fontSize: "27px"
    fontWeight: 700
    lineHeight: 1.2
  heading:
    fontFamily: '"Berkeley Mono", "JetBrains Mono", "IBM Plex Mono", ui-monospace, "Cascadia Code", "Segoe UI Mono", Consolas, monospace'
    fontSize: "16px"
    fontWeight: 700
    lineHeight: 1.45
    letterSpacing: "-0.02em"
  dialog-heading:
    fontFamily: '"Berkeley Mono", "JetBrains Mono", "IBM Plex Mono", ui-monospace, "Cascadia Code", "Segoe UI Mono", Consolas, monospace'
    fontSize: "18px"
    fontWeight: 700
    lineHeight: 1.4
  body:
    fontFamily: '"Berkeley Mono", "JetBrains Mono", "IBM Plex Mono", ui-monospace, "Cascadia Code", "Segoe UI Mono", Consolas, monospace'
    fontSize: "14px"
    fontWeight: 400
    lineHeight: 1.55
  button:
    fontFamily: '"Berkeley Mono", "JetBrains Mono", "IBM Plex Mono", ui-monospace, "Cascadia Code", "Segoe UI Mono", Consolas, monospace'
    fontSize: "13px"
    fontWeight: 500
    lineHeight: 1.35
  label:
    fontFamily: '"Berkeley Mono", "JetBrains Mono", "IBM Plex Mono", ui-monospace, "Cascadia Code", "Segoe UI Mono", Consolas, monospace'
    fontSize: "12px"
    fontWeight: 500
    lineHeight: 1.55
  caption:
    fontFamily: '"Berkeley Mono", "JetBrains Mono", "IBM Plex Mono", ui-monospace, "Cascadia Code", "Segoe UI Mono", Consolas, monospace'
    fontSize: "11px"
    fontWeight: 400
    lineHeight: 1.5
  micro:
    fontFamily: '"Berkeley Mono", "JetBrains Mono", "IBM Plex Mono", ui-monospace, "Cascadia Code", "Segoe UI Mono", Consolas, monospace'
    fontSize: "10px"
    fontWeight: 400
    lineHeight: 1.4
  tiny:
    fontFamily: '"Berkeley Mono", "JetBrains Mono", "IBM Plex Mono", ui-monospace, "Cascadia Code", "Segoe UI Mono", Consolas, monospace'
    fontSize: "9px"
    fontWeight: 400
    lineHeight: 1.4
rounded:
  none: "0px"
  sm: "4px"
spacing:
  xxs: "4px"
  xs: "8px"
  sm: "12px"
  md: "16px"
  lg: "24px"
  xl: "32px"
components:
  button-primary:
    backgroundColor: "{colors.primary}"
    textColor: "{colors.canvas}"
    typography: "{typography.button}"
    rounded: "{rounded.sm}"
    padding: "8px 15px"
    height: "42px"
  button-primary-hover:
    backgroundColor: "{colors.charcoal}"
    textColor: "{colors.canvas}"
    typography: "{typography.button}"
    rounded: "{rounded.sm}"
  button-primary-active:
    backgroundColor: "{colors.ink-deep}"
    textColor: "{colors.canvas}"
    typography: "{typography.button}"
    rounded: "{rounded.sm}"
  button-secondary:
    backgroundColor: "{colors.canvas}"
    textColor: "{colors.ink}"
    typography: "{typography.button}"
    rounded: "{rounded.sm}"
    padding: "8px 15px"
    height: "42px"
  button-disabled:
    backgroundColor: "{colors.surface-card}"
    textColor: "{colors.mute}"
    typography: "{typography.button}"
    rounded: "{rounded.sm}"
  text-input:
    backgroundColor: "{colors.surface-soft}"
    textColor: "{colors.ink}"
    typography: "{typography.label}"
    rounded: "{rounded.sm}"
    padding: "8px 10px"
    height: "42px"
  file-dropzone:
    backgroundColor: "{colors.surface-soft}"
    textColor: "{colors.ink}"
    typography: "{typography.body}"
    rounded: "{rounded.sm}"
    padding: "18px"
  workflow-link:
    backgroundColor: "{colors.canvas}"
    textColor: "{colors.mute}"
    typography: "{typography.label}"
    rounded: "{rounded.none}"
    padding: "8px 16px"
    height: "48px"
  page-description:
    backgroundColor: "{colors.canvas}"
    textColor: "{colors.body}"
    typography: "{typography.body}"
  helper-text:
    backgroundColor: "{colors.canvas}"
    textColor: "{colors.mute}"
    typography: "{typography.label}"
  file-confirmation:
    backgroundColor: "{colors.surface-soft}"
    textColor: "{colors.success}"
    typography: "{typography.caption}"
  page-preview:
    backgroundColor: "{colors.paper}"
    textColor: "{colors.ink}"
    typography: "{typography.body}"
  link:
    backgroundColor: "{colors.canvas}"
    textColor: "{colors.accent-hover}"
    typography: "{typography.label}"
  active-guidance:
    backgroundColor: "{colors.focus-soft}"
    textColor: "{colors.ink}"
    typography: "{typography.label}"
  active-link:
    backgroundColor: "{colors.canvas}"
    textColor: "{colors.accent-active}"
    typography: "{typography.label}"
  focus-ring:
    backgroundColor: "{colors.canvas}"
    textColor: "{colors.focus}"
    typography: "{typography.label}"
  selection:
    backgroundColor: "{colors.accent-selection}"
    textColor: "{colors.ink}"
    typography: "{typography.body}"
  error-state:
    backgroundColor: "{colors.danger-bg}"
    textColor: "{colors.danger}"
    typography: "{typography.label}"
  warning-state:
    backgroundColor: "{colors.warning-bg}"
    textColor: "{colors.warning}"
    typography: "{typography.label}"
  success-state:
    backgroundColor: "{colors.success-bg}"
    textColor: "{colors.success}"
    typography: "{typography.label}"
  preview-rule:
    backgroundColor: "{colors.preview-line}"
  format-chip:
    backgroundColor: "{colors.ink}"
    textColor: "{colors.ink-muted-inverse}"
    typography: "{typography.micro}"
  scrollbar-thumb:
    backgroundColor: "{colors.ash}"
  overlay:
    backgroundColor: "{colors.backdrop}"
    textColor: "{colors.canvas}"
---

## Overview

**Creative North Star: "Текстовая рабочая консоль"**

Интерфейс Markdown → DOCX берёт у пользовательской системы OpenCode её спокойную машинописную типографику, тёплый почти белый фон, тёмные чернила, тонкие разделители и компактные контролы. Визуальная система обслуживает последовательную подготовку документа: сначала источник и ресурсы, затем колонтитулы и профиль, в конце — создание файла.

В приложении вместо маркетингового TUI-макета и брендового ASCII-логотипа используется навигация по реальным этапам конвертации. Маркеры в скобках и моноширинная типографика помогают узнавать структуру Markdown-инструмента, а не превращают интерфейс в декоративный терминал. Акцентный синий обозначает фокус и ссылки; красный, янтарный и зелёный зарезервированы для ошибок, предупреждений и успешных состояний.

**Key Characteristics:**
- Моноширинная типографика во всех текстовых ролях.
- Кремовый фон, почти чёрный текст и тонкие тёплые линии.
- Навигация [01]–[04] повторяет реальный порядок подготовки документа.
- Интерактивные элементы имеют радиус 4 px; панели и разделы остаются квадратными.
- Цвет применяется как сигнал состояния, а не как украшение.

## Colors

Палитра берёт основу OpenCode и добавляет контрастные оттенки состояний, чтобы формы и сообщения оставались читаемыми.

### Primary
- **Чернила** (#201d1d): заголовки, основной текст и главные действия.
- **Глубокие чернила** (#0f0000): нажатое состояние главной кнопки.
- **Графит** (#302c2c): вторичная тёмная поверхность и наведение главного действия.
- **Третичный синий-индикатор** (#007aff): рамка при перетаскивании и активные подсказки.
- **Синий ссылок и фокуса** (#0056b3): контрастный текст ссылок и видимый клавиатурный фокус; главная кнопка остаётся чёрной.

### Neutral
- **Тёплый холст** (#fdfcfc): фон приложения и основная поверхность.
- **Мягкая поверхность** (#f8f7f7): поля, область перетаскивания и вторичные рабочие зоны.
- **Светлая поверхность** (#f1eeee): отключённые элементы и спокойные состояния.
- **Белая бумага** (#ffffff): лист A4 в предпросмотре.
- **Текст** (#424245): вспомогательные описания.
- **Приглушённый текст** (#646262): подписи и метаданные с достаточным контрастом.
- **Линия** (rgba(15, 0, 0, 0.12)): тонкие разделители.
- **Сильная линия** (#646262): обводки полей и интерактивных зон.

### Semantic
- **Ошибка** (#a50011 на #fff2f1): отклонённый ресурс, неверный формат и ошибки конфигурации.
- **Предупреждение** (#714800 на #fff8ea): состояние, требующее внимания, без блокировки.
- **Успех** (#1f6538 на #f0f7f1): принятый файл и готовый DOCX.

## Typography

**Display Font:** Berkeley Mono, затем JetBrains Mono, IBM Plex Mono и системные моноширинные шрифты.
**Body Font:** тот же моноширинный стек для единого ритма.
**Label/Mono Font:** общий моноширинный стек.

**Character:** Типографика напоминает аккуратно набранную документацию, а не рекламную страницу. Русский текст и длинные имена файлов остаются простыми для сканирования; системные и коммерческие шрифты не загружаются снаружи.

### Hierarchy
- **Display** (700, 36 px / 1.2): заголовок рабочей страницы.
- **Compact / mobile display** (700, 31 / 27 px): заголовок на узких экранах.
- **Headline** (700, 16 px / 1.45): названия этапов.
- **Dialog heading** (700, 18 px / 1.4): заголовки диалогов.
- **Title** (700, 14 px / 1.45): результат, блоки и заголовки групп.
- **Body** (400, 14 px / 1.55): пояснения и основной текст интерфейса.
- **Label** (500, 12 px / 1.55): поля, состояния, подписи.
- **Caption / micro / tiny** (400, 11 / 10 / 9 px): вспомогательные подписи и миниатюра листа.
- **Control** (500, 13 px / 1.35): кнопки и короткие действия.

**The One Type Rule.** Не вводить шрифт с засечками или отдельную декоративную гарнитуру в компонентах приложения.

## Layout

Ширина рабочего полотна — до 1080 px с адаптивными боковыми полями. На широком экране название и пояснение этапа занимают левую колонку около 200 px, рабочие контролы — правую. Разделы отделяются одной линией и отступом 24 px; вертикальное пространство лендинга сокращено для сквозной формы.

Навигация этапов расположена под заголовком и горизонтально прокручивается на узком экране. При ширине 700 px и меньше подписи этапов перемещаются над своими контролами, а предпросмотр с конструктором и строка отправки формы складываются в одну колонку. Поля профиля остаются двухколоночными до 460 px; ниже становятся одноколоночными, а действия профиля собираются в две колонки. Заголовок уменьшается с 36 до 31 px, а затем до 27 px на ширине 460 px. Диалоги используют максимум доступной высоты окна и сохраняют доступные действия внизу.

## Elevation & Depth

Система плоская и не использует тени. Иерархию задают цвет поверхности, тонкая линия и инверсия главной кнопки. Сильная обводка обозначает интерактивность, а не декоративный уровень.

**The Flat Surface Rule.** Не добавлять тени, градиенты или декоративную глубину на рабочую форму.

## Shapes

Поля, кнопки, метки формата, зоны загрузки и нажимаемые зоны колонтитулов имеют радиус 4 px. Разделители, рабочие секции, список ресурсов и лист предпросмотра остаются прямоугольными; обводка — 1 px, у зоны перетаскивания она штриховая.

## Components

### Buttons
- **Shape:** радиус 4 px.
- **Primary:** почти чёрная заливка и кремовый текст; стандартная высота не менее 42 px.
- **Hover / Focus / Active:** заливка переходит в графит при наведении; клавиатурный фокус — контрастное синее кольцо; при нажатии цвет становится глубже.
- **Secondary / Quiet:** прозрачная кремовая поверхность с графитовой обводкой; тихие действия оформляются ссылкой с подчёркиванием.
- **Disabled / Loading:** светлая поверхность, затемнённый текст и курсор ожидания без смены геометрии.

### Chips
Метка формата .DOCX — небольшой инвертированный блок с радиусом 4 px. Она сообщает формат результата, но не имитирует навигационную вкладку.

### Cards / Containers
Основные секции не оборачиваются в повторяющиеся карточки. Используются границы-разделители и мягкая поверхность для полей и области загрузки. Отдельная настройка профиля получает одну тонкую рамку без тени.

### Inputs / Fields
- **Style:** фон #f8f7f7, обводка #646262, радиус 4 px. Основные поля имеют высоту 42 px, поля модального редактора — 38 px; на сенсорных устройствах поля и действия увеличиваются до 44 px.
- **Focus:** фон становится кремовым, обводка дополняется кольцом #0056b3.
- **Error / Disabled:** красная текстура ошибки с бледно-красной поверхностью; отключённое состояние остаётся читаемым.

### Navigation
Ссылки этапов показывают маркеры [01]–[04], текст и тонкую нижнюю линию при наведении или фокусе. На мобильном экране они остаются в одну строку с горизонтальной прокруткой.

### File Dropzone
Область принимает .md и .zip, а внутри показывает короткий текстовый маркер [FILE]. При перетаскивании она подсвечивается синим; ошибки получают красную обводку. Успешная загрузка подтверждается отдельной зелёной меткой [OK].

### Signature Component: Document Route
Маршрут [01] Файл → [02] Ресурсы → [03] Колонтитулы → [04] Профиль — это навигация по реальным секциям, а не индикатор фиктивного прогресса. Маркеры [FILE], [IMG] и [OK] переносят квадратную маркерную лексику OpenCode в выбор файлов.

### Dialogs
Нативный dialog удерживает фокус, закрывается Escape и остаётся прокручиваемым на мобильном экране. Фон затемняется; сама панель плоская, с тонкой границей и без тени.

## Do's and Don'ts

### Do
- **Do** использовать общий моноширинный стек во всех компонентах.
- **Do** держать тело интерфейса на #fdfcfc, а поля и зоны загрузки — на #f8f7f7.
- **Do** резервировать синий для фокуса и действий, красный, янтарный и зелёный — для семантических состояний.
- **Do** применять одну геометрию кнопок и полей с радиусом 4 px.
- **Do** оставить разделы квадратными и разделять их линиями толщиной 1 px.
- **Do** сохранять видимую клавиатурную навигацию и отдельный текст ошибки около соответствующего шага.

### Don't
- **Don't** добавлять лэндинговый герой-макет терминала или чужой логотип OpenCode.
- **Don't** использовать serif display, градиенты, тени или карточки из одинаковых плиток.
- **Don't** использовать цвет как декор на неактивных поверхностях.
- **Don't** растягивать интервалы секций до лендинговых 96 px в рабочей форме.
- **Don't** заменять текстовые маркеры [FILE], [IMG] и [OK] разношёрстными иконками.
