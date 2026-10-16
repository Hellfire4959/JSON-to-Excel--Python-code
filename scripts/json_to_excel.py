"""Combine JSON objects or arrays of objects into an Excel worksheet."""
import argparse
import json
from pathlib import Path
import sys
import pandas as pd

def convert_folder(input_dir, output_file):
    input_dir, output_file = Path(input_dir), Path(output_file)
    if not input_dir.is_dir():
        raise ValueError(f'Input directory does not exist: {input_dir}')
    files = sorted(input_dir.glob('*.json'))
    if not files:
        raise ValueError('No .json files found in the input directory.')
    frames = []
    for path in files:
        with path.open(encoding='utf-8-sig') as stream:
            data = json.load(stream)
        if isinstance(data, dict):
            data = [data]
        if not isinstance(data, list) or any(not isinstance(row, dict) for row in data):
            raise ValueError(f'{path.name}: expected an object or an array of objects.')
        frames.append(pd.json_normalize(data))
    result = pd.concat(frames, ignore_index=True)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    result.to_excel(output_file, index=False, engine='openpyxl')
    return len(result)

def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input_dir', type=Path, help='Directory containing .json files (not recursive).')
    parser.add_argument('output_file', type=Path, help='Excel .xlsx file to write.')
    args = parser.parse_args(argv)
    if args.output_file.suffix.lower() != '.xlsx':
        parser.error('output_file must end in .xlsx')
    try:
        rows = convert_folder(args.input_dir, args.output_file)
    except (OSError, ValueError) as error:
        print(f'Error: {error}', file=sys.stderr)
        return 1
    print(f'Wrote {rows} rows to {args.output_file}')
    return 0

if __name__ == '__main__':
    sys.exit(main())
