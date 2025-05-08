#%%
import os
import shutil
import json
from pathlib import Path

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

import json
import os
from pathlib import Path

def create_json_file(data, destination, output_filename="Specimen Report.json"):
    # Create output directory if it doesn't exist
    os.makedirs(destination, exist_ok=True)
    
    # Define output path
    output_path = os.path.join(destination, output_filename)
    
    # Write to file
    with open(output_path, 'w') as f:
        json.dump(data, f, indent=4)
    
    return output_path

def process_input_directory(input_path, destination, output_filename="Specimen Report.json"):

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
    output_path = create_json_file(json_objects, destination, output_filename)
    print(f"\nCreated single JSON file at: {output_path}")
    print(f"Total objects: {len(json_objects)}")
    
    return output_path

# Example usage
if __name__ == "__main__":
    destination = r"C:\Users\EstebanDavyt\Desktop\PythonScript\Output"
    input_path = r"C:\Users\EstebanDavyt\Desktop\PythonScript\Templates"
    
    # Process all files in input directory and create single JSON
    empty_output_folder(destination)
    output_file = process_input_directory(input_path, destination)


# %%