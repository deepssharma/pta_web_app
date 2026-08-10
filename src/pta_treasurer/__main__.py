"""
__main__.py
Thin CLI wrapper around pipeline.run_month(), usable before the GUI exists
and as the same code path the GUI will call.

    python -m pta_treasurer generate --month July --year 2025 --data-dir "~/Documents/PTA Treasurer"
"""

import argparse
import sys
from pathlib import Path

from pta_treasurer.config import get_data_dir, load_config
from pta_treasurer.pipeline import run_month


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog='pta_treasurer')
    subparsers = parser.add_subparsers(dest='command', required=True)

    gen = subparsers.add_parser('generate', help='Generate one month\'s treasurer report')
    gen.add_argument('--month', required=True, help='Month name, e.g. July')
    gen.add_argument('--year', required=True, help='Year, e.g. 2025')
    gen.add_argument('--data-dir', help='Data folder (defaults to the saved data folder)')

    args = parser.parse_args(argv)

    data_dir = Path(args.data_dir) if args.data_dir else get_data_dir()
    if data_dir is None:
        print('No data folder set. Pass --data-dir, or run the app once to set one up.',
              file=sys.stderr)
        return 1

    config = load_config(data_dir)
    result = run_month(config, data_dir, args.month, args.year)

    print(f'Saved -> {result.output_path}')
    for w in result.warnings:
        print(f'WARNING: {w}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
