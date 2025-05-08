#%%
import os
import shutil
import json
from pathlib import Path
import aspose.words as aw
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

def add_watermark_to_docx(input_path, output_path):
    """Add diagonal 'SPECIMEN' watermark to a DOCX file using Aspose.Words"""
    try:
        # Load the document
        doc = aw.Document(input_path)
        
        # Create watermark shape
        watermark = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_PLAIN_TEXT)
        
        # Configure watermark text
        watermark.text_path.text = "SPECIMEN"
        watermark.text_path.font_family = "Arial"
        watermark.width = 300  # Width in points
        watermark.height = 70   # Height in points
        watermark.rotation = -45  # Diagonal at 45 degrees
        
        # Set fill color (light gray)
        watermark.fill.color = aw.drawing.Color.from_rgb(200, 200, 200)
        
        # Remove outline
        watermark.stroke_color = aw.drawing.Color.from_rgb(200, 200, 200)
        
        # Make watermark appear behind text
        watermark.behind_text = True
        
        # Center watermark on page
        watermark.relative_horizontal_position = aw.drawing.RelativeHorizontalPosition.PAGE
        watermark.relative_vertical_position = aw.drawing.RelativeVerticalPosition.PAGE
        watermark.left = 0
        watermark.top = 0
        watermark.wrap_type = aw.drawing.WrapType.NONE
        
        # Add watermark to all sections' headers
        for section in doc.sections:
            header = section.headers_footers.get_by_header_footer_type(aw.HeaderFooterType.HEADER_PRIMARY)
            header.append_child(watermark.clone(True))
        
        # Save document
        doc.save(output_path)
        return True
    except Exception as e:
        print(f"Error processing {input_path}: {str(e)}")
        return False

def create_json_file(data, destination, output_filename="output.json"):
    # Create output directory if it doesn't exist
    os.makedirs(destination, exist_ok=True)
    
    # Define output path
    output_path = os.path.join(destination, output_filename)
    
    # Write to file
    with open(output_path, 'w') as f:
        json.dump(data, f, indent=4)
    
    return output_path

def convert_json_to_excel(json_path, excel_path):
    """Convert JSON file to Excel file"""
    with open(json_path) as f:
        data = json.load(f)
    
    df = pd.DataFrame(data)
    df.to_excel(excel_path, index=False)
    return excel_path

def process_input_directory(input_path, destination, output_filename="output"):
    json_objects = []
    
    # Create subdirectories for watermarked files and JSON/Excel
    watermarked_dir = os.path.join(destination, "Watermarked_DOCX")
    os.makedirs(watermarked_dir, exist_ok=True)
    
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
        
        # Process DOCX files (add watermark)
        if filename.lower().endswith('.docx'):
            output_docx_path = os.path.join(watermarked_dir, filename)
            if add_watermark_to_docx(filepath, output_docx_path):
                print(f"Added watermark to: {filename}")
            else:
                print(f"Failed to add watermark to: {filename}")
        else:
            print(f"Skipped non-DOCX file: {filename}")
    
    # Create single JSON file with all objects
    json_path = create_json_file(json_objects, destination, f"{output_filename}.json")
    
    # Convert JSON to Excel
    excel_path = os.path.join(destination, f"{output_filename}.xlsx")
    convert_json_to_excel(json_path, excel_path)
    
    print(f"\nCreated files:")
    print(f"- JSON file at: {json_path}")
    print(f"- Excel file at: {excel_path}")
    print(f"- Watermarked DOCX files in: {watermarked_dir}")
    print(f"Total objects: {len(json_objects)}")
    
    return json_path, excel_path, watermarked_dir

# Example usage
if __name__ == "__main__":
    destination = r"C:\Users\EstebanDavyt\Desktop\PythonScript\Output"
    input_path = r"C:\Users\EstebanDavyt\Desktop\PythonScript\Templates"
    
    # Process all files in input directory
    empty_output_folder(destination)
    json_path, excel_path, watermarked_dir = process_input_directory(input_path, destination)
# %%
