# MIT License
# Copyright (c) 2025 OnistDerFalke

# This software is provided "as is", without any warranty. You can use, modify, and share it freely.

import sys
import pandas as pd
import json
import xml.etree.ElementTree as ET
import re

def sanitize_xml_tag(tag):
    """
    Ensures XML tags are valid by replacing invalid characters.
    """
    tag = re.sub(r'[^a-zA-Z0-9_]', '_', tag)
    if tag[0].isdigit():
        tag = "_" + tag
    return tag

def xls_to_json(file_path, sheet_numbers):
    """
    Converts a XLSX locales file into a JSON structure.
    
    :param file_path: Path to the xlsx file.
    :param sheet_numbers: List of sheet indices to process (1-based index).
    :return: Dictionary with language keys mapping.
    """
    result = {}
    
    for sheet_number in sheet_numbers:
        df = pd.read_excel(file_path, sheet_name=sheet_number-1)
        json_col = df.columns[0]
        
        for _, row in df.iterrows():
            key = row[json_col]
            if pd.isna(key):
                continue
            
            for lang in df.columns[1:]:
                if lang not in result:
                    result[lang] = {}
                result[lang][key] = row[lang]
    
    return result


def xls_to_xml(file_path, sheet_numbers):
    """
    Converts a XLSX locales file into a XML structure.
    
    :param file_path: Path to the xlsx file.
    :param sheet_numbers: List of sheet indices to process (1-based index).
    :return: XML tree format.
    """
    root = ET.Element("Root")
    
    for sheet_number in sheet_numbers:
        df = pd.read_excel(file_path, sheet_name=sheet_number-1)
        json_col = df.columns[0]
        
        for _, row in df.iterrows():
            key = row[json_col]
            if pd.isna(key):
                continue
            
            sanitized_key = sanitize_xml_tag(str(key))
            entry = ET.SubElement(root, sanitized_key)
            for lang in df.columns[1:]:
                lang_tag = sanitize_xml_tag(str(lang))
                text_element = ET.SubElement(entry, lang_tag)
                text_element.text = str(row[lang]) if not pd.isna(row[lang]) else ""
    
    return ET.tostring(root, encoding="utf-8", method="xml").decode()


def print_usage_info():
    """
    Prints usage instructions for running the script.
    """
    print(
    "Usage: python lexp.py <xls_file> <format> <sheets>\n"
    "Example: python lexp.py my_locales_sheets.xlsx json 1 2 3 4 5\n"
    "Supported output formats:\n"
    "- JSON\n"
    "- XML\n"
    )


def lexp():
    """
    Parses command-line arguments and converts xlsx file to chosen format.
    """

    if len(sys.argv) < 4:
        print_usage_info()
        return
    
    file_path = sys.argv[1]
    output_format = sys.argv[2]
    sheet_numbers = [int(n) for n in sys.argv[3:]]
    
    if output_format.lower() == "json":

        print(f"Processed file: {file_path}.")
        print(f"Output format: JSON.")
        print(f"Sheets included: {sheet_numbers}.")

        data = xls_to_json(file_path, sheet_numbers)
        output_file = file_path.rsplit('.', 1)[0] + ".json"
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        
        print(f"Locales exported to JSON file and saved as {output_file}.")

    elif output_format.lower() == "xml":
        print(f"Processed file: {file_path}.")
        print(f"Output format: XML.")
        print(f"Sheets included: {sheet_numbers}.")

        xml_data = xls_to_xml(file_path, sheet_numbers)
        output_file = file_path.rsplit('.', 1)[0] + ".xml"
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(xml_data)
        
        print(f"Locales exported to XML file and saved as {output_file}.")

if __name__ == "__main__":
    lexp()
