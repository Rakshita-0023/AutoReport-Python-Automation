# AutoReport | 45-second sales demo

Use the synthetic CSV and the generated workbook. Before recording, open the CSV in a text editor, stage the command in a terminal with the virtual environment active, and close the Excel report. Set Excel to open the generated workbook at Sales Summary. Hide unrelated files and notifications.

| Time | On screen | Say |
| --- | --- | --- |
| 0–10s | Show the raw CSV. Point to two date formats or a blank field. | “This is a typical sales export. Dates arrive in different formats, some rows are repeated, and some prices are missing. Every reporting cycle, someone has to sort that out.” |
| 10–21s | Run the staged command. Pause on the row counts. | “I run AutoReport once. It keeps 942 usable sales, removes 30 repeats, and flags 28 rows for review.” |
| 21–35s | Open Sales Summary. Point to revenue and the monthly chart, then click Product Performance. | “The workbook opens with revenue and the monthly trend. Your team can see which products and regions drive sales using the same reporting rules each time.” |
| 35–45s | Click Data Quality Report. Show the action and details columns. | “Here’s what needs correcting, down to the source row. If you send me a redacted export, I can scope this for your business.” |

Stage this command before you start:

```bash
python main.py --input data/sample_sales.csv --output output/sample_report.xlsx
```
