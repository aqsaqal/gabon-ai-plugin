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

## Install

```
/plugin marketplace add aqsaqal/gabon-ai-plugin
/plugin install gabon-ai@gabon-ai
```

Then describe the task in plain language, for example «Сделай КП на демонтаж перегородок, объёмы в input/ВОР.xlsx».

## License

MIT-0. See [LICENSE](LICENSE).
