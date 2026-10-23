# Листы персонажей в стиле ПЧК

Скилл для ИИ-агентов и браузерных чатов: по описанию персонажа создаёт редактируемый **PPTX**,
**PDF для печати** и JSON-черновик для последующих изменений. Две страницы A5 в альбомной
ориентации: характеристики в щитах, навыки, способности, вещи и заряды; портрет и предыстория.

Это неофициальный инструмент оформления, а не официальный свод правил «Подземелий Чикен Карри».
Механику составляет агент по выбранной системе. По умолчанию используется основа D&D 5e 2014;
можно задать другую редакцию или авторские правила. Для JSON, импортируемого в Long Story Short,
есть отдельный [lss-character](https://github.com/Dezodemius/dnd-lss-character-generator).

## Установка

Готовые файлы лежат в [релизах](https://github.com/Dezodemius/dnd-pchk-character-generator/releases/latest).

| Где пользуетесь | Как подключить |
| --- | --- |
| **claude.ai**, Claude Desktop | Скачать [`pchk-character.zip`](https://github.com/Dezodemius/dnd-pchk-character-generator/releases/latest/download/pchk-character.zip) → Настройки → Возможности → Skills → загрузить. В Desktop можно открыть [`pchk-character.skill`](https://github.com/Dezodemius/dnd-pchk-character-generator/releases/latest/download/pchk-character.skill) двойным кликом |
| **Claude Code** | `/plugin marketplace add Dezodemius/agent-skills`, затем `/plugin install pchk-character@dezodemius-skills` |
| **Codex, Gemini CLI, Cursor, GitHub Copilot, OpenCode** и другие агенты | `npx skills add Dezodemius/dnd-pchk-character-generator` |
| **ChatGPT, Gemini, DeepSeek, Qwen, Grok, Алиса, GigaChat** и другие чаты | Скачать [`pchk-character.prompt.md`](https://github.com/Dezodemius/dnd-pchk-character-generator/releases/latest/download/pchk-character.prompt.md) и приложить в начале чата или вставить в инструкции проекта / GPT / Gem |

Подробно по каждой платформе — в [каталоге](https://github.com/Dezodemius/agent-skills#установка).

Скиллу нужны Python 3.10+ и зависимости из `requirements.txt`: полноценно он работает в агенте или чате,
который умеет запускать код. Однофайловая версия `.prompt.md` в обычном чате поможет составить
JSON-черновик, а собрать PPTX и PDF можно локально (см. «Локальный запуск»).

## Использование в браузерном чате

Если интерфейс поддерживает импорт пользовательских skills, загрузи `.skill` или ZIP в его
раздел импорта. Название раздела и поддерживаемые расширения зависят от сервиса.

В чате, который умеет исполнять Python и выдавать файлы, можно просто прикрепить **ZIP** и написать:

> Распакуй архив, прочитай pchk-character/SKILL.md и создай лист ПЧК: дварф-жрец 5 уровня,
> бывший тюремный повар. Верни PPTX и PDF. Портрет приложен.

Нужны распаковка архивов, Python, возможность установить зависимости (либо готовые библиотеки)
и выдача созданных файлов. Расширение `.skill` само по себе не включает выполнение кода.
Если чат поддерживает только текст, он сможет подготовить JSON-черновик; сборку нужно выполнить
локально или в среде с Python. Доступ к внешним сервисам для самой сборки не требуется.
Генерация нового портрета зависит от доступных инструментов чата; можно передать свою картинку
или явно попросить лист без портрета.

## Локальный запуск

```text
python -m pip install -r skills/pchk-character/requirements.txt
python skills/pchk-character/scripts/pchk.py validate hero.draft.json
python skills/pchk-character/scripts/pchk.py build hero.draft.json --pdf --preview
```

Черновик составляй по [формату](skills/pchk-character/references/draft-format.md) или
[примеру](skills/pchk-character/assets/example.draft.json). Картинка задаётся путём относительно JSON.
Файлы появятся рядом с черновиком; `--output-dir` меняет папку, `--force` разрешает пересборку.

- `--renderer auto` — PowerPoint в Windows, если он доступен; иначе переносимый PDF-экспорт.
- `--renderer portable` — PDF и PNG через Python (ReportLab / PDFium), Windows, Linux и macOS.
- `--renderer powerpoint` — проверка и экспорт установленным Microsoft PowerPoint в Windows.

Портативный PDF использует встроенные **PT Serif и Anton** вместо Cambria и Impact. Возможны
небольшие отличия шрифтов и переносов. Он поддерживает этот шаблон, а не произвольные PPTX.
PPTX остаётся редактируемым; сохранение в PDF другим офисным приложением может изменить переносы.
Если текст не помещается, сборщик сообщает об ошибке, не обрезает описание и не выдаёт старый PDF
за новый. Перед использованием посмотри оба PNG: проверка размеров не заменяет проверку глазами.

## Выпуск релиза

Релиз собирается автоматически по тегу из четырёх чисел на коммите из `main`:

```bash
git tag 1.0.0.0 && git push origin 1.0.0.0
```

Общий CD и проверки живут в [Dezodemius/agent-skills](https://github.com/Dezodemius/agent-skills).

Собственные проверки (валидация данных, сборка PPTX/PDF/PNG, переполнение текста) запускаются
в общем CI на Windows и Linux. Локально:

```text
python -m pip install -r requirements-dev.txt
python -X utf8 tools/check.py
```

## Лицензии и источники

Код и инструкции распространяются по Apache-2.0, см. [LICENSE](LICENSE).
Стороннее оформление шаблона и шрифты не получают эту лицензию автоматически; источники и
ограничения перечислены в [NOTICE](NOTICE). Официальный платный PDF и портрет героя в архив
не входят. Проект не связан с авторами шоу и не утверждает их одобрения.
