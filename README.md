# JSON to Excel

Combine JSON files from one directory into a single Excel worksheet. Nested object fields become dotted column names such as `profile.location`.

## Quick start

Requires Python 3.11 or later.

```bash
git clone https://github.com/Hellfire4959/JSON-to-Excel--Python-code.git
cd JSON-to-Excel--Python-code
python -m venv .venv
# macOS/Linux: source .venv/bin/activate
# Windows PowerShell: .venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python scripts/json_to_excel.py examples/json output/people.xlsx
```

The included [examples](examples/json/) produce three rows. To use your own files:

```bash
python scripts/json_to_excel.py path/to/json-folder output/report.xlsx
```

## Input and output

- Each `.json` file must contain an object or an array of objects. UTF-8 files with a byte-order mark are supported.
- Files are processed in filename order; subdirectories are not searched.
- Columns are the union of the input fields. Missing values become empty Excel cells; nested arrays remain cell values rather than becoming additional rows.
- The output is a single `.xlsx` worksheet. An existing file at the output path is replaced.
- Invalid JSON or an unsupported top-level structure stops the conversion before the workbook is written. This utility loads the data into memory; large datasets may exceed memory or Excel limits.

## Repository layout

| Path | Purpose |
| --- | --- |
| [scripts/json_to_excel.py](scripts/json_to_excel.py) | Command-line converter |
| [examples/json/](examples/json/) | Small, synthetic inputs |
| [requirements.txt](requirements.txt) | Python dependencies |
| [tests/](tests/) | Workbook and input-validation tests |

The original extensionless `Python code` file is now `scripts/json_to_excel.py`.

## Development

```bash
python -m unittest discover -s tests -v
```

For changes, include a small input, the expected worksheet layout, and the test result in your pull request.
