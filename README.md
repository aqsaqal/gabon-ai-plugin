# GABON AI

Плагин Claude Code для ПТО и сметного отдела строительной компании (Казахстан).

## Установка (в Claude Code)

Нужна утилита [uv](https://docs.astral.sh/uv/getting-started/installation/): через неё запускаются скрипты плагина. При первом запуске нужен интернет, чтобы загрузить библиотеки Python.

```
/plugin marketplace add aqsaqal/gabon-ai-plugin
/plugin install gabon-ai@gabon-ai
```

После установки перезапустите Claude Code. Обновление до новой версии: `/plugin marketplace update gabon-ai`.

## Использование

Откройте Claude Code в папке с проектами и напишите обычным языком, например:

> Сделай КП на демонтаж перегородок для ТОО «Заказчик», объёмы в input/ВОР.xlsx

Claude соберёт `kp.json`, создаст PDF, Excel и Word в `output/` и проверит их.

## Структура

```
.claude-plugin/marketplace.json   каталог плагинов (для установки)
plugins/gabon-ai/
  .claude-plugin/plugin.json      описание плагина
  hooks/hooks.json                загрузка общих правил в начале сессии
  rules/RULES.md                  общие правила (ничего не выдумывать, НДС, проверки)
  skills/kp/                      навык «Коммерческое предложение»
    SKILL.md                      инструкция для Claude
    scripts/build_kp.py           сборка Excel + PDF
    scripts/check_kp.py           независимая проверка
    scripts/kp_core.py            общий расчёт
    scripts/make_template.py      пустой Excel-шаблон КП для ручного заполнения
    scripts/xlsx_to_kp.py         заполненный шаблон -> kp.json -> PDF
    references/input-format.md    формат kp.json
    assets/KP_template.xlsx       универсальный шаблон КП (Excel, формулы, печать A4)
    assets/                       примеры kp.json и company.json
  skills/prices/                  навык «Прайс-лист»: шаблон, проверка, поиск, подстановка в КП
  skills/brand/                   навык «Фирменный шаблон»: цвета, логотип, бланк, тексты
  skills/vor/                     навык «ВОР»: объёмы из документов заказчика, переход в КП
  skills/smeta/                   навык «Смета»: ресурсный расчёт, проверка смет подрядчиков
  skills/grafik/                  навык «График»: диаграмма Ганта, критический путь
  skills/akty/                    навык «Акты»: Р-1, накопительная, справка, Word
  skills/proverka/                навык «Проверка»: сверка всех документов проекта
  agents/proverka.md              независимый проверяющий агент
  tests/                          автотесты
projects/test-fundament/          тест: ТЗ -> ВОР -> смета, КП, график, договор, акты; сравнение со сметой подрядчика
projects/demontazh-almaty/        тест: демонтаж 5 000 м³ × 4 500 тг
docs/ROADMAP.md                   следующие этапы
```

## Ручной запуск без Claude

```
uv run plugins/gabon-ai/skills/kp/scripts/build_kp.py   skills/prices/                  навык «Прайс-лист»: шаблон, проверка, поиск, подстановка в КП
  skills/brand/                   навык «Фирменный шаблон»: цвета, логотип, бланк, тексты
  skills/vor/                     навык «ВОР»: объёмы из документов заказчика, переход в КП
  skills/smeta/                   навык «Смета»: ресурсный расчёт, проверка смет подрядчиков
  skills/grafik/                  навык «График»: диаграмма Ганта, критический путь
  skills/akty/                    навык «Акты»: Р-1, накопительная, справка, Word
  skills/proverka/                навык «Проверка»: сверка всех документов проекта
  agents/proverka.md              независимый проверяющий агент
  tests/                          автотесты
projects/test-fundament/input/kp.json
uv run plugins/gabon-ai/skills/kp/scripts/check_kp.py   skills/prices/                  навык «Прайс-лист»: шаблон, проверка, поиск, подстановка в КП
  skills/brand/                   навык «Фирменный шаблон»: цвета, логотип, бланк, тексты
  skills/vor/                     навык «ВОР»: объёмы из документов заказчика, переход в КП
  skills/smeta/                   навык «Смета»: ресурсный расчёт, проверка смет подрядчиков
  skills/grafik/                  навык «График»: диаграмма Ганта, критический путь
  skills/akty/                    навык «Акты»: Р-1, накопительная, справка, Word
  skills/proverka/                навык «Проверка»: сверка всех документов проекта
  agents/proverka.md              независимый проверяющий агент
  tests/                          автотесты
projects/test-fundament/input/kp.json \
    skills/prices/                  навык «Прайс-лист»: шаблон, проверка, поиск, подстановка в КП
  skills/brand/                   навык «Фирменный шаблон»: цвета, логотип, бланк, тексты
  skills/vor/                     навык «ВОР»: объёмы из документов заказчика, переход в КП
  skills/smeta/                   навык «Смета»: ресурсный расчёт, проверка смет подрядчиков
  skills/grafik/                  навык «График»: диаграмма Ганта, критический путь
  skills/akty/                    навык «Акты»: Р-1, накопительная, справка, Word
  skills/proverka/                навык «Проверка»: сверка всех документов проекта
  agents/proverka.md              независимый проверяющий агент
  tests/                          автотесты
projects/test-fundament/output/KP_ТЕСТ-2026-001.xlsx   skills/prices/                  навык «Прайс-лист»: шаблон, проверка, поиск, подстановка в КП
  skills/brand/                   навык «Фирменный шаблон»: цвета, логотип, бланк, тексты
  skills/vor/                     навык «ВОР»: объёмы из документов заказчика, переход в КП
  skills/smeta/                   навык «Смета»: ресурсный расчёт, проверка смет подрядчиков
  skills/grafik/                  навык «График»: диаграмма Ганта, критический путь
  skills/akty/                    навык «Акты»: Р-1, накопительная, справка, Word
  skills/proverka/                навык «Проверка»: сверка всех документов проекта
  agents/proverka.md              независимый проверяющий агент
  tests/                          автотесты
projects/test-fundament/output/KP_ТЕСТ-2026-001.pdf
```

## Тесты

```
cd plugins/gabon-ai
uv run --with pytest --with openpyxl --with reportlab --with pypdf --with pillow --with python-docx python -m pytest -q tests
```

## Лицензия

MIT-0 (MIT No Attribution): код можно использовать, изменять и распространять без условий. См. файл LICENSE.
Данные компаний, прайсы и документы заказчиков в плагин не входят и лицензией не покрываются.
