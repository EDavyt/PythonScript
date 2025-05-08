#%%
import pandas as pd
import uuid as guid
import os
import openpyxl
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

def duplicate_fileDocument(fileTemplate, destination, uuid1gen, form_number, template_name):
        # Check if source exists
        if not os.path.exists(fileTemplate):
            print("File does not exist")
            return False
            
        # Check if the destination folder exists, if not create it
        if not os.path.exists(destination):
            os.makedirs(destination)
        
        # Get file extension and base name
        base_name = os.path.basename(fileTemplate)
        name, ext = os.path.splitext(base_name)
        
        # Create new filename with GUID
        new_filename = f"{name}{form_number}_{uuid1gen}{ext}".replace(" ", "_")
        destination_path = os.path.join(destination, new_filename)
        
        shutil.copy2(fileTemplate, destination_path)
        print(f"File duplicated successfully to {destination_path}")
    
        with open(destination_path, 'r+') as file:
            data = json.load(file)
            
            # Element modifications
            data['id'] = str(uuid1gen)
            data['SystemInfo']['DocumentId'] = str(uuid1gen)
            data['SystemInfo']['DocumentVersionId'] = str(uuid1gen)
            data['Content']['FileName'] = template_name + ".docx"
            data['Content']['AzureBlobStorageFileName'] = "hudsoninsgroup0001/01/01/"+uuid1gen+".docx"
            
            # Write back the changes
            file.seek(0)
            json.dump(data, file, indent=4)
            file.truncate()
        
        print(f"File {new_filename} updated successfully to {destination_path}")
        return True

def duplicate_documentDocument(documentTemplatePath, destination, uuid1gen, uuid2gen, form_number, form_description, effective_date, expiration_date, print_order,edit_date,policy_State,OutputTemplateMandatory):
        if not os.path.exists(documentTemplatePath):
            print("File does not exist")
            return False
            
        # Check if the destination folder exists, if not create it
        if not os.path.exists(destination):
            os.makedirs(destination)

        base_name = os.path.basename(documentTemplatePath)
        name, ext = os.path.splitext(base_name)

        # Create new filename with GUID
        new_filename = f"{name}{form_description}_{uuid2gen}{ext}".replace(" ", "_")
        destination_path = os.path.join(destination, new_filename)

        shutil.copy2(documentTemplatePath, destination_path)
        print(f"File duplicated successfully to {destination_path}")

        with open(destination_path, 'r+') as file:
            data = json.load(file)

            # Element modifications
            data['id'] = str(uuid2gen)
            data['SystemInfo']['DocumentId'] = str(uuid2gen)
            data['SystemInfo']['DocumentVersionId'] = str(uuid2gen)
            data['Content']['CommonGUID'] = str(uuid2gen)
            data['Content']['TextPolicyFormNumber'] = form_number
            data['Content']['TextOutputTemplateTitle'] = form_number + " "+edit_date+ " " + form_description
            data['Content']['TemplateFile'] = str(uuid1gen).upper()
            data['Content']['TemplateCriteria'][0]['TemplateStartDate'] = effective_date
            data['Content']['TemplateCriteria'][0]['TemplateEndDate'] = expiration_date
            data['Content']['TemplateCriteria'][0]['PolicyState'] = policy_State
            data['Content']['TemplateCriteria'][0]['OutputTemplateMandatory'] = OutputTemplateMandatory
            data['Content']['TemplateCriteria'][0]['TextOutputTemplateSpecimenUrl'] = "https://hudsonfiles.hudsonportal.com/BL/"+ uuid1gen + ".pdf"
            data['Content']['TemplateCriteria'][0]['OutputTemplatePrintOrder'] = str(print_order)

            # Write back the changes
            file.seek(0)
            json.dump(data, file, indent=4)
            file.truncate()  

def duplicate_TemplatesDocument(folder_path, destination, uuid1gen):
        # Check if source exists
        if not os.path.exists(folder_path):
            print("File does not exist")
            return False
            
        # Check if the destination folder exists, if not create it
        if not os.path.exists(destination):
            os.makedirs(destination)
        
        # Get file extension and base name
        base_name = os.path.basename(folder_path)
        name, ext = os.path.splitext(base_name)
        
        # Create new filename with GUID
        new_filename = f"{uuid1gen}{ext}"
        destination_path = os.path.join(destination, new_filename)
        
        shutil.copy2(folder_path, destination_path)
        print(f"File duplicated successfully to {destination_path}")


# Path for the templates placeholder to use on the generation of the files
fileTemplatePath = r"C:\Users\EstebanDavyt\Desktop\PythonScript\input\File_.json"
documentTemplatePath = r"C:\Users\EstebanDavyt\Desktop\PythonScript\input\Document_.json"
folder_path = r"C:\Users\EstebanDavyt\Desktop\PythonScript\Templates"
destination = r"C:\Users\EstebanDavyt\Desktop\PythonScript\Output"
# This reads the excel file and loads it into a pandas dataframe
rules = pd.read_excel(
    r"C:\Users\EstebanDavyt\Desktop\PythonScript\Rules\rules.xlsx",
    parse_dates=['EffectiveDate', 'ExpirationDate']
)

rules['EffectiveDate'] = pd.to_datetime(rules['EffectiveDate']).dt.strftime('%Y-%m-%d')
rules['ExpirationDate'] = rules['ExpirationDate'].astype(str).str[:10] 

rulesAsJSON = json.loads(rules.to_json(
    orient='records',
    default_handler=str
))

#Show as JSON the rules loaded from the excel file
print(rulesAsJSON)

template_files = [f for f in Path(folder_path).iterdir() if f.is_file()]

#Empty the output folder
empty_output_folder(destination)

for file in template_files:
    try:
        # Generate GUIDs for each file
        template_guid = str(guid.uuid4())
        document_guid = str(guid.uuid4())

        #Extract the file name without extension
        filename = file.stem
        parts = filename.split(' ', 2)
        
        # Get form number and description
        FormNumber = parts[:2]
        
        form_number = ' '.join(parts[:2])
        
        print(f"Form Number: {form_number}")
        print(f"Form Description: {form_description}")

        #Matching rules:

        #Matching Rule based on USStateCode
        # matching_rule = next(
        #     (rule for rule in rulesAsJSON 
        #     if f"{rule['USStateCode']}" == form_state),
        #     None
        # )

        #Matching Rule based on FormNumber
        matching_rule = next(
            (rule for rule in rulesAsJSON 
            if f"{rule['FormNumber']}" == form_number),
            None
        )

        effective_date = matching_rule['EffectiveDate']
        expiration_date = matching_rule['ExpirationDate']
        print_order = matching_rule['DisplaySequence']
        form_description = matching_rule['FormTitle']
        template_name = filename
        form_number_And_Title_FromExcel = matching_rule['FormNumber']
        edit_date = matching_rule['EditionDate']
        policy_State = [matching_rule['USStateCode']]
        OutputTemplateMandatory = ""
        #Required or Optional
        if matching_rule['FormRequired'] == 1:
            OutputTemplateMandatory = "Mandatory"
        else:
            OutputTemplateMandatory = "Optional"


        duplicate_fileDocument(fileTemplatePath, destination, template_guid, form_number, template_name)
        duplicate_documentDocument(documentTemplatePath, destination, template_guid, document_guid, form_number_And_Title_FromExcel, form_description, effective_date, expiration_date, print_order,edit_date,policy_State,OutputTemplateMandatory)
        duplicate_TemplatesDocument(file, destination, template_guid)
    
    except Exception as e:
        print(f"Error: {e}")


#Generate Specimen PDF and Report
    

# %%