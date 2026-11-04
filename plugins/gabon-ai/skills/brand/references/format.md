# Формат brand.json

Пути к файлам указываются относительно brand.json. Все поля необязательные:
пустой `{}` даёт стандартный стиль GABON AI.

| Поле | По умолчанию | Описание |
|---|---|---|
| `name` | | Название стиля, для себя |
| `colors.primary` | `#1F3864` | Заголовки, шапка таблицы, линии |
| `colors.header_text` | `#FFFFFF` | Текст в шапке таблицы |
| `colors.section` | `#D9E1F2` | Фон строк разделов |
| `colors.subtotal` | `#F2F2F2` | Фон строк «Итого по разделу» |
| `colors.total` | `#FFF2CC` | Фон строки «ВСЕГО с НДС» |
| `logo` | нет | PNG или JPG. Если нет, берётся `logo` из company.json |
| `logo_position` | `left` | `left` или `right` |
| `logo_width_mm` | `35` | Ширина логотипа в PDF, 10–90 мм |
| `show_requisites` | `true`, а с бланком `false` | Печатать реквизиты из company.json в шапке |
| `letterhead_pdf` | нет | Фирменный бланк A4 в PDF. КП печатается поверх него |
| `letterhead_pages` | `first` | `first` — бланк только на первой странице, `all` — на всех |
| `margins_mm` | `top 15, bottom 18, left 15, right 15` | Поля страницы, 5–120 мм. С бланком увеличьте `top` и `bottom` |
| `fonts.regular`, `fonts.bold` | Arial или DejaVu Sans | Файлы TTF. Проверяется наличие русских и казахских букв |
| `title` | `КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ` | Заголовок документа |
| `footer_text` | нет | Текст в центре подвала каждой страницы |
| `defaults` | нет | Значения для пустых полей kp.json, см. ниже |

## defaults

| Поле | Пример |
|---|---|
| `valid_days` | `14` |
| `numbering` | `continuous` или `section` |
| `amount_decimals` | `0` или `2` |
| `intro` | `"Предлагаем выполнить следующий перечень работ:"` |
| `terms` | список строк или объект `duration`, `payment`, `warranty`, `start` |
| `assumptions`, `exclusions` | списки строк |
| `rates` | `overhead_pct`, `overhead_source`, `profit_pct`, `profit_source`, `vat_pct`, `vat_source`. При `vat_pct` обязателен `vat_source` |

## Что проверяет `check`

- цвета в виде #RRGGBB;
- файлы логотипа, бланка и шрифтов существуют и нужного типа;
- шрифты содержат кириллицу и казахские буквы (ә, ғ, қ, ң, ө, ұ, ү, һ, і);
- бланк читается и имеет формат A4;
- поля страницы и ширина логотипа в допустимых пределах;
- неизвестные поля отмечаются замечанием, чтобы опечатка не прошла незамеченной.
