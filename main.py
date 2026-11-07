"""Command-line entry point for AutoReport."""

import argparse
import logging
from pathlib import Path

from openpyxl.utils.exceptions import IllegalCharacterError

from src.cleaner import ReportError, clean_sales, read_sales
from src.report_generator import write_report


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Turn a sales export into an Excel report.")
    parser.add_argument("--input", type=Path, required=True, help="Source CSV or XLSX")
    parser.add_argument("--output", type=Path, required=True, help="Destination XLSX")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    try:
        if args.output.suffix.lower() != ".xlsx":
            raise ReportError("The output file must have an .xlsx extension.")
        if args.input.resolve() == args.output.resolve():
            raise ReportError("Choose a different output path to protect your source file.")
        sales, audit = clean_sales(read_sales(args.input))
        write_report(sales, audit, args.input.name, args.output)
    except IllegalCharacterError:
        logging.error("A source value contains a control character that Excel cannot store.")
        return 1
    except (ReportError, OSError) as exc:
        logging.error("%s", exc)
        return 1

    logging.info("Report saved to %s", args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
