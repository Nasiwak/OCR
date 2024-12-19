import os
import csv
import pdfplumber
import openai
import json
import logging
import re
import codecs

logging.basicConfig(filename='extraction_log.txt', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    encoding='utf-8')

openai.api_key = 'you-api-key-here'

def extract_text_from_pdf(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = '\n'.join(page.extract_text() for page in pdf.pages if page.extract_text())
        return full_text
    except Exception as e:
        logging.error(f"Error extracting text from {pdf_path}: {str(e)}")
        return None

def format_floor_number(floors):
    if isinstance(floors, str):
        try:
            floors = json.loads(floors)
        except json.JSONDecodeError:
            floors = [floors]
    
    if isinstance(floors, (list, set)):
        floors = sorted(list(set(floors)))  
    else:
        floors = [floors]
    
    numeric_floors = [str(f) for f in floors if str(f).isdigit()]
    
    if len(numeric_floors) == 1 and numeric_floors[0] == '1':
        return '平屋'
    else:
        return ','.join(numeric_floors) + 'F'

def clean_numeric_value(value):
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        value = re.sub(r'[m²㎡]\s*$', '', value.strip())
        match = re.search(r'\d+\.?\d*', value)
        if match:
            return float(match.group())
    return 0

def query_openai_for_data(text, pdf_filename):
    prompt_text = f"""
    Analyze the text and extract structured information with details for each floor (if multiple floors are mentioned). Give me the summarized values of 'File name', 'Builder name', 'Stud sink direction', 'Wall width', 'Board thickness', 
                      'Floor height', 'Ceiling height', 'Floor number', 'Order number', 'Order name', 'Comment section', 'Floor area 1', 'Floor area 2', 'Floor area 3', 'Loft', 'Penthouse area'. For floor number, return an array of floor numbers mentioned in the document. If there is only one page in the pdf, return ["1"]. 
    The values for floor areas should be mapped as follows:
    - 'Floor area 1': Area of 1st floor (numeric value only, without m² or ㎡)
    - 'Floor area 2': Area of 2nd floor (numeric value only, without m² or ㎡)
    - 'Floor area 3': Area of 3rd floor (numeric value only, without m² or ㎡)
    - 'Loft': Area marked as loft space (ロフト) or attic storage (小屋裏収納) (numeric value only)
    - 'Penthouse area': Area specifically marked as penthouse (numeric value only)
    Only return the numeric part of the floor area (e.g., "75.5" instead of "75.5m²"). The value for the field 'File name' will be '{pdf_filename}'. The output should be in JSON format containing only these fields and their associated values:
    - File name
    - Builder name (ビルダー名)
    - Stud sink direction (スタッド流し⽅向) (Format: '@number')
    - Wall width (壁先⾏) (Format: 'number mm', e.g., '105 mm')
    - Board thickness (壁ボード) (Format: 'number mm', e.g., '12.5 mm')
    - Floor height (階⾼) (Format: 'number mm per floor', e.g., '1階: 2750 mm')
    - Ceiling height (天井⾼) (Format: 'number mm per floor', e.g., '1階: 2200 mm')
    - Floor number (Array of floor numbers, e.g., ["1"] for single floor, ["1", "2"] for two floors)
    - Order number (【案件No】) (Format: '6 digits')
    - Order name (【案件名】) (Format: name)
    - Comment section (【特 記】) (Format: specific comments per floor)
    - Floor area (【⾯積】) (Format: numeric values only, without units)
    - Penthouse area (Format: numeric value only)
    Text provided:
    {text}
    """
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a skilled assistant trained in extracting precise information from the pdf. Always respond with valid JSON, including all specified fields even if the value is empty or not found. For all area measurements, return only numeric values without units (m² or ㎡)."},
                {"role": "user", "content": prompt_text}
            ]
        )
        return response['choices'][0]['message']['content']
    except Exception as e:
        logging.error(f"Error querying OpenAI for {pdf_filename}: {str(e)}")
        return None
    
def clean_json_string(json_string):
    json_string = re.sub(r'^```json\s*|```\s*$', '', json_string, flags=re.MULTILINE)
    return json_string.strip()

pdf_dir = 'file'
csv_file = 'xyz2.csv'
fieldnames = ['File name', 'Builder name', 'Stud sink direction', 'Wall width', 'Board thickness', 
              'Floor height', 'Ceiling height', 'Floor number', 'Order number', 'Order name', 
              'Comment section', 'Floor area 1', 'Floor area 2', 'Floor area 3', 'Loft', 'Penthouse area']

field_mapping = {
    'File name': 'File name',
    'Builder name': 'Builder name',
    'Stud sink direction': 'Stud sink direction',
    'Wall width': 'Wall width',
    'Board thickness': 'Board thickness',
    'Floor height': 'Floor height',
    'Ceiling height': 'Ceiling height',
    'Floor number': 'Floor number',
    'Order number': 'Order number',
    'Order name': 'Order name',
    'Comment section': 'Comment section',
    'Floor area 1': 'Floor area 1',
    'Floor area 2': 'Floor area 2',
    'Floor area 3': 'Floor area 3',
    'Loft': 'Loft',
    'Penthouse area': 'Penthouse area'
}

all_data = []

for filename in os.listdir(pdf_dir):
    if filename.endswith('.pdf'):
        pdf_path = os.path.join(pdf_dir, filename)
        text = extract_text_from_pdf(pdf_path)
        if text is None:
            continue
        
        extracted_data = query_openai_for_data(text, filename)
        if extracted_data is None:
            continue
        
        cleaned_data = clean_json_string(extracted_data)
        
        try:
            data_dict = json.loads(cleaned_data)
        except json.JSONDecodeError as e:
            logging.error(f"Error parsing JSON for file {filename}. Error: {str(e)}")
            logging.error(f"Raw response: {extracted_data}")
            continue
        
        if 'Floor number' in data_dict:
            data_dict['Floor number'] = format_floor_number(data_dict['Floor number'])
        
        mapped_data = {}
        for api_field, csv_field in field_mapping.items():
            value = data_dict.get(api_field, '')
            
            if csv_field in ['Floor area 1', 'Floor area 2', 'Floor area 3', 'Loft', 'Penthouse area']:
                mapped_data[csv_field] = clean_numeric_value(value)
            elif isinstance(value, (list, dict)) and api_field != 'Floor number': 
                mapped_data[csv_field] = json.dumps(value, ensure_ascii=False)
            else:
                mapped_data[csv_field] = value
        
        for field in fieldnames:
            if field not in mapped_data:
                mapped_data[field] = ''
        
        for area_field in ['Floor area 1', 'Floor area 2', 'Floor area 3', 'Loft', 'Penthouse area']:
            if not mapped_data[area_field]:
                mapped_data[area_field] = 0
        
        all_data.append(mapped_data)

with codecs.open(csv_file, mode='w', encoding='utf-8-sig') as file:
    writer = csv.DictWriter(file, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(all_data)

print("Data extraction and storage to CSV completed successfully.")
logging.info("Data extraction and storage to CSV completed successfully.")