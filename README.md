# Excel Cleaner & Summary Tool

A Python script that cleans Excel files and generates a statistical summary sheet using pandas.

## Features

- Removes fully empty rows and columns
- Normalizes column names (lowercase, spaces → underscores)
- Drops duplicate rows
- Strips leading/trailing whitespace from string columns
- Generates a summary table (count, mean, median, std, min, max, missing) for numeric columns
- Saves cleaned data and summary to separate sheets in a new Excel file

## Requirements

```bash
pip install pandas openpyxl
```

## Usage

```bash
python excel_cleaner.py input.xlsx
python excel_cleaner.py input.xlsx --sheet Sheet1 --out output.xlsx
```

### Arguments

| Argument | Default | Description |
|----------|---------|-------------|
| `dosya` | *(required)* | Input `.xlsx` file path |
| `--sheet` | `0` | Sheet name or index |
| `--out` | `cikti.xlsx` | Output file path |

## Output

The output Excel file contains two sheets:

- **Temiz_Veri** — cleaned data
- **Ozet** — statistical summary of numeric columns
