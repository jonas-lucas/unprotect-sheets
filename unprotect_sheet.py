import os
import re
import shutil
import zipfile
import argparse


def validate_input_file(filepath: str) -> None:
    """Validate that the input file is an .xlsx file."""
    if not filepath.lower().endswith('.xlsx'):
        raise argparse.ArgumentTypeError(f'The file "{filepath}" is not a .xlsx file.')


def remove_sheet_protection_tags(xml_folder: str) -> None:
    """Remove <sheetProtection .../> tags from all .xml files in the given folder."""
    pattern = r'<sheetProtection\b[^<>]*?(?:/>|></sheetProtection>)'

    for filename in os.listdir(xml_folder):
        if filename.endswith('.xml'):
            file_path = os.path.join(xml_folder, filename)

            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

            cleaned_content = re.sub(pattern, '', content)

            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(cleaned_content)


def main():
    parser = argparse.ArgumentParser(description='Remove sheet protection from .xlsx sheets.')
    parser.add_argument('-i', '--input', required=True, help='Input .xlsx file')
    args = parser.parse_args()

    input_file = os.path.basename(args.input)
    validate_input_file(input_file)

    # Output and temp names
    output_file = f'Unprotect{input_file}'
    temp_xlsx = '_temp.xlsx'
    temp_zip = '_temp.zip'
    temp_dir = '_temp'

    # Step 1: Copy and rename input file to .zip
    shutil.copy2(input_file, temp_xlsx)
    os.rename(temp_xlsx, temp_zip)

    # Step 2: Extract .zip content to temporary directory
    with zipfile.ZipFile(temp_zip, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # Step 3: Remove <sheetProtection> tags from sheet XML files
    worksheets_path = os.path.join(temp_dir, 'xl', 'worksheets')
    remove_sheet_protection_tags(worksheets_path)

    # Step 4: Rezip the modified content
    shutil.make_archive('_temp', 'zip', temp_dir)
    os.rename(temp_zip, output_file)

    # Step 5: Clean up
    shutil.rmtree(temp_dir)
    if os.path.exists(temp_zip):
        os.remove(temp_zip)

    print(f'File "{output_file}" created successfully!!')


if __name__ == '__main__':
    main()
