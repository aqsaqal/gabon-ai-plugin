# GABON AI

Claude Code plugin for the estimating and production-technical (ПТО) office of a construction company in Kazakhstan. It turns client documents into checkable commercial offers, bills of quantities, estimates, work schedules and acts of completed works. Documents and conversation are in Russian.

Плагин для сметного отдела и ПТО строительной компании в Казахстане: КП, ведомости объёмов, сметы, графики работ и акты выполненных работ с проверяемыми расчётами.

## What it does

| Skill | Result |
|---|---|
| `kp` | Commercial offer: PDF, Word and Excel with live formulas |
| `prices` | Company price list with a source and date for every price; flags old prices |
| `brand` | Company logo, colors, letterhead and default wording |
| `vor` | Bill of quantities from client PDF, Word and Excel files, with a formula, source and status on every row |
| `smeta` | Resource-based commercial estimate; comparison with a contractor's estimate |
| `grafik` | Work schedule with a Gantt chart, critical path and working calendar |
| `akty` | Acts of completed works (form Р-1), cumulative progress, advance offset and retention |
| `proverka` | Review of a whole project before documents are issued |

The `proverka` agent is an independent reviewer. It reads finished documents without seeing how they were made and does not edit them.

## Rules it follows

Prices, rates, quantities, standards and signatures are never invented. Every price needs a source and a date, and anything missing is reported or marked as an assumption. Scripts do the arithmetic, and a separate checker recalculates every result. All documents are drafts for a responsible specialist to check and approve.

## Limits

It makes commercial calculations, not official normative estimates with Kazakh rate codes, which require licensed software such as АВС-4. Dimensions read from scanned drawings are always marked as assumptions. Check the Р-1 form and the VAT rate with your accountant before real use.

## What it runs and accesses

- **Where it works:** Claude Code. It runs Python scripts on your computer, and it is tested only in Claude Code.
- **Requirement:** [uv](https://docs.astral.sh/uv/) must be installed. Claude runs the plugin's Python scripts with `uv run`.
- **Downloads:** on first run, uv downloads the Python libraries named at the top of each script from PyPI: openpyxl, reportlab, pypdf, pillow and python-docx. It needs internet access for this.
- **Files:** scripts read and write only the files in the project folder you point them at, plus a system font for PDFs.
- **Data:** the plugin sends nothing to any server, and it has no telemetry, hooks or MCP servers. Claude itself reads your documents under your usual Claude Code terms.

See the [privacy notice](PRIVACY.md).

## Install

```
/plugin marketplace add aqsaqal/gabon-ai-plugin
/plugin install gabon-ai@gabon-ai
```

Then open Claude Code in a folder with a `projects/<object>/input/` subfolder holding the client's documents, and describe the task in plain language.

## Example prompts

1. «Сделай КП на демонтаж здания в Алматы: 5 000 м³ по 4 500 тг/м³ без НДС, НДС 16 % сверху, заказчик пока не указан.»
2. «Прочитай техзадание из input/TZ.pdf и составь ведомость объёмов работ на фундамент. То, чего нет в ТЗ, отметь как допущение или недостающие данные.»
3. «Сравни смету подрядчика input/smeta_podryadchika.xlsx с нашей сметой input/smeta.json и покажи, где он завысил объёмы или цены.»
4. «Проверь все документы проекта projects/sklad и скажи, можно ли отправлять КП заказчику.»

A complete sample project is in the repository under `projects/test-fundament`: a client brief, a bill of quantities, an estimate, an offer, a schedule, a contract and two acts, all built from one another. Every figure in it is marked as test data.

## Troubleshooting

- **`uv: command not found`**: install uv from https://docs.astral.sh/uv/ and restart the terminal.
- **The first run is slow or fails without internet**: uv is downloading the Python libraries. Connect once, and later runs work offline.
- **A build stops with «ОШИБКИ В ИСХОДНЫХ ДАННЫХ»**: the message lists what is missing or wrong in the input file. Nothing is written until it is fixed.
- **A check reports «РЕЗУЛЬТАТ: найдено ошибок»**: the generated files no longer match the input data, usually because they were edited by hand. Rebuild them from the input file.
- **Cyrillic shows as boxes in a PDF**: no font with Cyrillic was found. Set `fonts` in `brand.json` to a TTF font that has it.

## License

MIT-0. See [LICENSE](LICENSE).
