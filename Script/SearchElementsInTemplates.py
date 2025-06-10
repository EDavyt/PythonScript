#%%
import os
import json
import docx2txt

def search_word_documents(input_folder, output_json_path):
    search_term = "CGLPremiumWithoutLiquorLiabilityAndTerrorismDelta"
    matching_files = []

    for filename in os.listdir(input_folder):
        if filename.lower().endswith('.docx'):
            file_path = os.path.join(input_folder, filename)
            
            try:
                # Extract ALL text (including dynamic elements)
                text = docx2txt.process(file_path)
                
                if search_term in text:
                    matching_files.append(filename)
            except Exception as e:
                print(f"⚠️ Error reading {filename}: {e}")

    # Save results
    os.makedirs(os.path.dirname(output_json_path), exist_ok=True)
    with open(output_json_path, 'w') as json_file:
        json.dump(matching_files, json_file, indent=4)

    print(f"Searching for '{search_term}' in Word documents in: {input_folder}")
    print(f"✅ Found {len(matching_files)} matches. Results saved to: {output_json_path}")

# Example usage
search_word_documents(r"..\Templates", r"..\Output\matches.json")
# %%
