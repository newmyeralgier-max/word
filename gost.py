#!/usr/bin/env python3
"""Entrypoint ГОСТ-форматтера v5.

Usage:
    python3 gost.py [input.docx] [output.docx]

По умолчанию:
    input  = ./БР.docx
    output = ./gost.docx
"""
import argparse
import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / 'WORD' / 'execution'))

from gost.pipeline import run  # noqa: E402


def main(argv=None):
    parser = argparse.ArgumentParser(description='GOST-формтатер v5 (БР.docx → gost.docx)')
    parser.add_argument('input', nargs='?', default=str(HERE / 'БР.docx'),
                        help='исходный DOCX (default: БР.docx)')
    parser.add_argument('output', nargs='?', default=str(HERE / 'gost.docx'),
                        help='целевой DOCX (default: gost.docx)')
    parser.add_argument('-q', '--quiet', action='store_true')
    args = parser.parse_args(argv)

    if not os.path.exists(args.input):
        print(f'ERROR: input not found: {args.input}', file=sys.stderr)
        sys.exit(2)

    print(f'[gost] {args.input} -> {args.output}')
    stats = run(args.input, args.output, verbose=not args.quiet)
    print(f'[gost] done. Stats: {dict(stats)}')


if __name__ == '__main__':
    main()
