# XLSX Locales Exporter
Simple util that converts xlsx sheet with subsheets containing locales from translators to **JSON** or **XML** translation file.

<img src="https://github.com/user-attachments/assets/96f04915-ff89-476e-af52-5c68532f8235" width="800" height="800" />

## Installation
Install required packages:

`pip install -r requirements.txt`

## Usage
Two locales export formats supported: **JSON** and **XML**.

### JSON

Example usage:

`python lexp.py .\example_input.xlsx json 1`

### XML
Example usage:

`python lexp.py .\example_input.xlsx xml 1`

### Multiple subsheets in one export

You can export translations from more than one subsheet at once. Just write all 1-based index of subsheets that have to be included in export.

Example usage (exports only odd subsheets):

`python lexp.py .\multiple_subsheets.xlsx json 1 3 5 7 9`

## Disclaimer

**The script does not check the consistency of the format on subsheets or the correctness of indexes. Please check this by yourself.**
