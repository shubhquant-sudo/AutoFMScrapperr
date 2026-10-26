"""AutoFMScrapper entry point: beginner desktop app or local Excel CLI."""

import argparse
import json
import sys
from pathlib import Path


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Create a three-statement model and DCF with downloaded history and blank forecast inputs."
    )
    parser.add_argument("--ticker")
    parser.add_argument("--import-file", type=Path)
    parser.add_argument("--refresh", action="store_true")
    parser.add_argument(
        "--output",
        type=Path,
        help="A new output directory. Existing directories are preserved.",
    )
    parser.add_argument("--ui", action="store_true")
    args = parser.parse_args(argv)
    if args.ui or not args.ticker:
        from autofmscrapper.gui import main as ui

        ui()
        return 0
    from autofmscrapper.runner import generate

    try:
        result = generate(
            args.ticker,
            import_file=args.import_file,
            refresh=args.refresh,
            output=args.output,
            progress=lambda n, t: print(f"{n}/5: {t}", flush=True),
        )
        print(json.dumps(result, indent=2))
        return 0
    except (ValueError, OSError, TimeoutError) as error:
        print(str(error), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
