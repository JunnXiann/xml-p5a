import xml.etree.ElementTree as ET
import pandas as pd
import os
import re
import csv

def process_excel_file(excel_path, xml_base_path, identifier_column='identifier'):
    """Process Excel file and add punctuation information"""
    df = pd.read_excel(excel_path)
    
    for index, row in df.iterrows():
        identifier = row.get(identifier_column, "")
        if not identifier or pd.isna(identifier):
            print(f"Skipping row {index}: identifier is empty or NaN")
            continue

        print(f"Processing identifier: {identifier}")
        xml_path = find_xml_file(identifier, xml_base_path)
        print(f"Found XML files: {xml_path}")

        if xml_path:
            all_punc_values = []
            all_providers = []
            all_paths = []
            
            for path in xml_path:
                punc_value, provider = extract_info(path)
                print(provider, punc_value)
                if punc_value:
                    all_punc_values.append(punc_value)
                if provider:
                    all_providers.append(provider)
                all_paths.append(os.path.basename(path))
            
            # Update the existing columns with extracted information
            df.at[index, '标点类型'] = ';'.join(all_punc_values)
            df.at[index, '标点来源'] = ';'.join(all_providers)
            df.at[index, 'xml文件名'] = ';'.join(all_paths)
    
    # Save the updated DataFrame back to Excel
    output_path = excel_path.replace('.xlsx', '_updated.xlsx')
    df.to_excel(output_path, index=False)
    print(f"Updated Excel file saved as: {output_path}")
    
    return df

def find_xml_file(identifier, base_path):
    dir_name = identifier[0]
    folder_path = os.path.join(base_path, dir_name)

    if not os.path.exists(folder_path):
        return None

    file_name = ''.join(identifier[1:]) + '.xml'

    matched_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.xml') and file_name in file:
                matched_files.append(os.path.join(root, file))

    return matched_files

def extract_info(xml_file):
    with open(xml_file, 'r', encoding='utf-8') as f:
        xml_text = f.read()

    punc_matches = re.findall(r'<punctuation[^>]*>\s*<p>(.*?)</p>', xml_text, re.DOTALL)
    provider_matches = re.findall(r'<projectDesc>.*?<p[^>]*(?:xml:lang="zh-Hant"[^>]*cb:type="ly"|cb:type="ly"[^>]*xml:lang="zh-Hant")[^>]*>(.*?)</p>.*?</projectDesc>', xml_text, re.DOTALL)
    values = [match.strip() for match in punc_matches if match.strip()]

    matched_providers = []
    for provider in provider_matches:
        provider = provider.strip()
        for punc_value in values:
            if punc_value in provider:
                split_items = provider.split('，')
                for item in split_items:
                    if punc_value in item:
                        matched_providers.append(item.strip())
                        break
                break

    if not matched_providers:
        matched_providers = [match.split('，')[-1] for match in provider_matches]

    providers = [match.strip() for match in matched_providers if match.strip()]
    return ','.join(values), ','.join(providers)

# Example usage
if __name__ == "__main__":
    # Set your paths here
    excel_file_path = "径山藏对应CBETA经编码-0906.xlsx" 
    xml_repository_path = "." 

    identifier_column_name = "CBETA经编码" 

    result = process_excel_file(excel_file_path, xml_repository_path, identifier_column_name)