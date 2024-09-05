import json
from bicycle_generator.generator import generate_bicycle_json
def clean_bicycle_data(data):
    for item in data:
        for key, value in list(item.items()):
            if value == "=FALSE()":
                item[key] = False
            elif value == "=TRUE()":
                item[key] = True
            elif value == "null":
                item[key] = None
        
        if 'null' in item:
            del item['null']
            
    return data

if __name__ == "__main__":
    file_path = r"C:\Users\swami\OneDrive\Desktop\Bicycle.xlsx"
    
    result = generate_bicycle_json(file_path)
    
    if isinstance(result, str):
        result = json.loads(result)
        result = clean_bicycle_data(result)
    
    output_json_path = r"C:\Users\swami\OneDrive\Desktop\Bicycle.json"
    
    with open(output_json_path, 'w', encoding='utf-8') as json_file:
        json.dump(result, json_file, indent=4, ensure_ascii=False)
    
    print(f"Result saved to {output_json_path}")
