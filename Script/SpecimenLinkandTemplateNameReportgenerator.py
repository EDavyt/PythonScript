#%%
import os
import shutil
import json
from pathlib import Path
import pandas as pd

# This is to delete all the files in the output folder
def empty_output_folder(destination):
    # Check if the destination folder exists, if not create it
    if os.path.exists(destination):
        # Remove all files in the folder
        for filename in os.listdir(destination):
            file_path = os.path.join(destination, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f'Failed to delete {file_path}. Reason: {e}')
    else:
        os.makedirs(destination)

def create_json_file(data, destination, output_filename="Specimen Report.json"):
    # Create output directory if it doesn't exist
    os.makedirs(destination, exist_ok=True)
    
    # Define output path
    output_path = os.path.join(destination, output_filename)
    
    # Write to file
    with open(output_path, 'w') as f:
        json.dump(data, f, indent=4)
    
    return output_path

def create_excel_file(data, destination, output_filename="Specimen Report.xlsx"):
    # Create output directory if it doesn't exist
    os.makedirs(destination, exist_ok=True)
    
    # Define output path
    output_path = os.path.join(destination, output_filename)
    
    # Convert JSON data to DataFrame
    df = pd.DataFrame(data)
    
    # Write to Excel file
    df.to_excel(output_path, index=False)
    
    return output_path

def process_input_directory(input_path, destination, json_filename="Specimen Report.json", excel_filename="Specimen Report.xlsx"):
    json_objects = []
    
    # Get all files in input directory (excluding subdirectories)
    for filename in os.listdir(input_path):
        filepath = os.path.join(input_path, filename)
        
        # Skip directories
        if os.path.isdir(filepath):
            continue
            
        # Get filename without extension
        nameoffile = Path(filename).stem
        
        # Create JSON object
        json_obj = {
            "FileName": nameoffile,
            "SpecimenURL": f"https://hudsonfiles.hudsonportal.com/BL/{nameoffile}.pdf"
        }
        json_objects.append(json_obj)
        print(f"Processed file: {filename}")
    
    # Create single JSON file with all objects
    json_output_path = create_json_file(json_objects, destination, json_filename)
    print(f"\nCreated JSON file at: {json_output_path}")
    
    # Create Excel file with the same data
    excel_output_path = create_excel_file(json_objects, destination, excel_filename)
    print(f"Created Excel file at: {excel_output_path}")
    
    print(f"Total objects processed: {len(json_objects)}")
    
    return json_output_path, excel_output_path

# Example usage
if __name__ == "__main__":
    destination = r"..\Output"
    input_path = r"..\Templates"
    
    # Process all files in input directory and create JSON and Excel files
    empty_output_folder(destination)
    json_file, excel_file = process_input_directory(input_path, destination)

#%%