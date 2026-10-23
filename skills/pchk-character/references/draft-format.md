# Черновик ПЧК, версия 1

JSON UTF-8. Рабочий пример: `../assets/example.draft.json`. Пути к картинкам разрешаются
относительно черновика. Неизвестные поля отклоняются.

| Поле | Значение |
|---|---|
| `version` | Число `1` |
| `name` | Обязательное имя |
| `role` | Класс, вид, титул или уровень; выводится перед именем |
| `rules` | Система, редакция и допущения; записывается в заметки |
| `ability_modifiers` | Шесть целых модификаторов: `str`, `dex`, `con`, `wis`, `int`, `cha`; от -5 до 15 |
| `armor_class` | Положительное целое до 99 либо строка вида `14/16` |
| `hit_points` | Положительное целое до 9999 |
| `weapon_proficiencies` | Оружие и итоговые бонусы, строка |
| `skills` | Навыки и итоговые бонусы, строка |
| `abilities` | Список записей `name`, необязательные `uses` и `text` |
| `equipment` | Такой же список предметов; в `text` — механика и свойства |
| `special` | `null` либо объект `title`, необязательные `charges`, `recovery`, обязательный список `effects` |
| `backstory` | Предыстория, строка; абзацы через `\n\n` |
| `portrait` | Путь к PNG/JPEG/WebP либо `null` |
| `portrait_focus` | Два числа 0…1; по умолчанию `[0.5, 0.5]`, `[0.5, 0.25]` поднимает фокус |

Обязательны `version`, `name`, `ability_modifiers`, `armor_class`, `hit_points`, `abilities`,
`equipment`, `backstory`. Модификаторы — числа без знака `+`, скрипт добавляет знак сам.
`armor_class` со слешем требует пояснения условий в предметах.

```json
{
  "name": "Второе дыхание",
  "uses": "1 раз; короткий отдых",
  "text": "Бонусным действием восстанови 1d10+3 здоровья."
}
```

Умения и предметы выводятся абзацами: **Название (применения):** *описание*.
Пустой `text` допустим для простого предмета. Не добавляй формулы LSS или вложенный `data`:
это статический лист, числовые значения нужно рассчитать до сборки.

`special` — блок магического предмета с общими зарядами. У `effects` тот же формат, что у умений.
`recovery` описывает восстановление запаса. При `null` блок удаляется, место доступно другим разделам.

Первый лист распределяет место между разделами по длине текста. Основной текст — Cambria 9.4 пт,
предыстория — Cambria Italic 16 пт. Текст не обрезается и не уменьшается до нечитаемого размера.
При переполнении сократи формулировки без потери механики либо согласуй дополнительный лист.
Предварительное измерение приблизительное; проверь PNG выбранного рендерера перед выдачей.

```text
python scripts/pchk.py validate /path/hero.draft.json
python scripts/pchk.py build /path/hero.draft.json --pdf --preview
python scripts/pchk.py build /path/hero.draft.json --output-dir /path/output --pdf --preview --force
python scripts/pchk.py render /path/hero.pchk.pptx --pdf --preview
python scripts/pchk.py build /path/hero.draft.json --pdf --preview --renderer portable
```

Результаты: `<имя>.pchk.pptx`, `<имя>.pchk.pdf`, `<имя>.pchk.preview/` с PNG и служебный
`<имя>.pchk.render-report.json`. Имя файла берётся из имени черновика без `.draft.json`.
При ошибке экспорта сохраняется PPTX, но PDF/PNG не выдаются как готовые.

`--renderer auto` — PowerPoint в Windows, иначе переносимый ReportLab-экспорт со встроенными
PT Serif и Anton. `--renderer powerpoint` требует Windows и Microsoft PowerPoint.
`--renderer portable` не требует офисного приложения и работает с данным шаблоном на Python.
Размеры символов и переносы в PDF и при открытии PPTX другим офисным приложением могут различаться.
