# Unprotect Sheets

This repository provides simple scripts to **remove sheet protection tag** from `.xlsx` (Excel) files. Useful for editing protected sheets or automating access to locked content.

## How to Use

Just run the command bellow:

```bash
python unprotect_sheet.py -i <xlsx_file>
```

A new file will be generated in the format `Unprotect<xlsx_file>`, adding "Unprotect" at the beginning of the original file name.

## How it Works

1. Compress the sheet file `.xlsx` to `.zip` archive.

```python
os.rename('file.xlsx', 'file.zip')
```

2. Remove all `<sheetProtection/>` tags in all `sheetn.xml` files present in `xl/worksheets` folder on zipped file.

```python
pattern = r'<sheetProtection\b[^<>]*?(?:/>|></sheetProtection>)'
cleaned_content = re.sub(pattern, '', original_content)
```

> The files are named `sheet1.xml`, `sheet2.xml`, etc., each corresponding to a worksheet tab.

3. Restore the `.zip` file to a sheet file `.xlsx`.

```python
os.rename('file.zip', 'file.xlsx')
```

## Systems Settings

- Windows 11 24H2
- Python 3.13.3

---

### Author

[Jonas Lima](https://github.com/jonas-lucas)