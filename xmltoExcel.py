import xml.etree.ElementTree as ET
import pandas as pd
import os
import sys

def extract_text(element):
    if element is None:
        return "N/A"
    
    texts = []
    for p in element.findall("p"):
        if p.text:
            texts.append(p.text)
        
        for sub_element in p:
            if sub_element.text:
                texts.append(sub_element.text)
            if sub_element.tail:
                texts.append(sub_element.tail)
    
    return " ".join(texts).strip()

def parse_xml_to_excel(xml_file, output_excel):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    
    extracted_data = []
    
    for question in root.findall(".//question"):
        number = question.get("number", "N/A")
        
        headline_text = extract_text(question.find("headline/text"))
        
        extracted_data.append([number, headline_text])
    
    df = pd.DataFrame(extracted_data, columns=["Question Number", "Text"])
    df.to_excel(output_excel, index=False)
    print(f"Data successfully written to {output_excel}")

def process_input(path):
    if os.path.isdir(path):
        for file in os.listdir(path):
            if file.lower().endswith(".xml"):
                input_file = os.path.join(path, file)
                output_file = os.path.join(path, os.path.splitext(file)[0] + ".xlsx")
                parse_xml_to_excel(input_file, output_file)
    elif os.path.isfile(path) and path.lower().endswith(".xml"):
        output_file = os.path.splitext(path)[0] + ".xlsx"
        parse_xml_to_excel(path, output_file)
    else:
        print("Invalid input. Please provide a valid XML file or directory containing XML files.")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py <xml_file_or_directory>")
    else:
        process_input(sys.argv[1])
