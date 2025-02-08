# XLSX Locales Exporter
Simple util that converts xlsx sheet with subsheets containing locales from translators to JSON or XML translation file.

## Installation
Install required packages:

`pip install -r requirements.txt`

## Usage
Two locales export formats supported: JSON and XML.

### JSON

Example usage that creates a JSON file from second subsheet of xlsx file.

`python lexp.py file.xlsx json 2`

### XML
Example usage that creates a XML file from first 4 subsheets of xlsx file.

`python lexp.py file.xlsx xml 1 2 3 4`