from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches

import zipfile
import os
import re
import shutil
from pathlib import Path
import glob

from file_renamer import renamer

#**********************Report Generator Function********************************
def render_template(placeholder_dict: dict, doc):
    print("Writing the data.....Please Wait....")
    doc.render(placeholder_dict)
    name_formatted = format_name(placeholder_dict['Client_Name'])
    output_file_name = f"{name_formatted}.docx"
    print("Report name:", output_file_name)
    doc.save(output_file_name)
    print("Report is saved successfully.")

# Report Name Adjuster
def format_name(name: str):
    file_name = name.replace(" ", "_")
    return file_name

# -----------------------------------Creates the dictionary which contains the text placeholder and the images-----------------------------------------
def dict_generator(doc):
    # Inputs for TEXTS
    text_key_list = ['Client_Name', 'Date','no_of_stories', 'average_floor_ht',
                    'bldg_ht', 'grid_pattern', 'bldg_len', 'bldg_width',
                    'bldg_type', 'location', 'footing_type', 'col_size',
                    'beam_size', 'bearing_capacity', 'concrete_grade',
                    'stair_dl', 'stair_ll', 'weight', 'C_ULS', 'C_SLS',
                    'K', 'VB_ULS', 'VB_SLS']
    values_list = []

#****************************************USER INPUT****************************************************

    Client_Name = 'Mr. John Nash'          
    Date = 'JUNE 2024'   
    no_of_stories = 3  
    floor_height_general= 2.870
    bldg_ht = 8.611    
    grid_pattern = 'an Irregular Grid Pattern'  
    bldg_len = 5.715
    bldg_width = 8.433     
    building_type = 'Residential'
    location = 'Kathmandu-10'
    footing_type = 'Strap Footing'
    col_size = '14" x 14"'
    beam_size = '9" x 14"'
    bearing_capacity = 110
    concrete_grade = 'M20'
    stair_dl = 8.49
    stair_ll = 3.23

    # Parameters Related to EQ
    seismic_weight = 1777.605              # kN
    C_ULS = 0.131                           # Base shear coefficient for ULS
    C_SLS = 0.126                           # Base shear coefficient for SLS
    exponent_K = 1.0                        # Exponent related to time period
    VB_ULS = round(C_ULS * seismic_weight,2)         # Base Shear for ULS
    VB_SLS = round(C_SLS * seismic_weight, 2)         # Base Shear for SLS
#***************************************************************************************

    height_common = Inches(4.5)

    hard_values = [Client_Name, Date, no_of_stories, floor_height_general, bldg_ht, grid_pattern, bldg_len, bldg_width, building_type,
                location, footing_type, col_size, beam_size, bearing_capacity, concrete_grade, stair_dl, stair_ll, seismic_weight,
                C_ULS, C_SLS, exponent_K, VB_ULS, VB_SLS]
    # Creating placeholder dictionary
    placeholder_dict = {}
    for i in range(len(text_key_list)):
        placeholder_dict[text_key_list[i]] = hard_values[i]
# *****************************************************************************************************************************
    # Inputs for images dictionary
    image_dict = {'cover_image': InlineImage(doc, 'Snaps/cover_image.png', height= height_common), 
                'beam_layout': InlineImage(doc, 'Snaps/beam_layout.png', height= height_common),
                'shell_ll': InlineImage(doc, 'Snaps/shell_ll.png', height= height_common),
                'shell_ff': InlineImage(doc, 'Snaps/shell_ff.png', height= height_common),
                'wall_load_3d': InlineImage(doc, 'Snaps/wall_load_3d.png', height= height_common),
                'diaphragm_assignment_3d': InlineImage(doc,'Snaps/diaphragm_assignment_3d.png', height= height_common),
                'bmd': InlineImage(doc, 'Snaps/bmd.png', height= height_common),
                'sfd': InlineImage(doc, 'Snaps/sfd.png',height= height_common),
                'afd': InlineImage(doc, 'Snaps/afd.png',height= height_common),
                'modal_participation': InlineImage(doc, 'Snaps/modal_participation.png', width = height_common),
                'base_reactions': InlineImage(doc, 'Snaps/base_reactions.png', height= height_common),
                'beam_rebar': InlineImage(doc, 'Snaps/beam_rebar.png', height= height_common),
                'col_rebar': InlineImage(doc,'Snaps/col_rebar.png', height_common),
                'eq_x_sls_drifts': InlineImage(doc,'Snaps/eq_x_sls_drifts.png', height= height_common),
                'eq_y_sls_drifts': InlineImage(doc,'Snaps/eq_y_sls_drifts.png', height= height_common),
                'eq_x_uls_drifts': InlineImage(doc,'Snaps/eq_x_uls_drifts.png', height= height_common),
                'eq_y_uls_drifts': InlineImage(doc,'Snaps/eq_y_uls_drifts.png', height= height_common),
                'bc_ratio_grid_1': InlineImage(doc,'Snaps/bc_ratio_grid_1.png', height= height_common),
                'bc_ratio_grid_2': InlineImage(doc,'Snaps/bc_ratio_grid_2.png', height= height_common),
                'bc_ratio_grid_3': InlineImage(doc,'Snaps/bc_ratio_grid_3.png', height= height_common),
                'stair_dead': InlineImage(doc, 'Snaps/stair_dead.png', height= height_common),
                'stair_live': InlineImage(doc, 'Snaps/stair_live.png', height= height_common)
                }
    # Merging two dictionaries
    final_dict = {**placeholder_dict, **image_dict} #final dictionary to be used to take text and images to place
    return final_dict

'''Main Function to Replace Texts and Images'''
def report_generator_main(master_dict, doc):
    # Rendering the doc file
    render_template(master_dict, doc)

#--------------Template file path and name--------------
doc = DocxTemplate("Report_template_normal.docx")

#--------------Final Dictionary containing placeholders and images--------------
master_dict = dict_generator(doc)
print(f"{'*'*20}PROGRAM START{'*'*20}")
report_generator_main(master_dict, doc)
print(f"{'*'*20}Report Created Successfully{'*'*20}")


#*************************************************Rename the Excel Files*************************************************************
folder_paths = ['Excel_Sheets_Linked', 'Excel_Sheets_Linked/NBC_Checks']
client_name = format_name(master_dict['Client_Name'])
for paths in folder_paths:  # Replace with the path to your folder
    renamer(paths, client_name)

print("All Workbooks renamed Successfully. \n")

#*********************************************************EXCEL LINK UPDATER************************************************

'''EXCEL LINK UPDATER'''
print(f"{'*'*20}Updating the Links to Excel Sheets{'*'*20}")

# Link matcher function
def link_matcher(parent_link, new_links):
    parent_link_name = parent_link.split('/')[-1].split('\\')[-1]  # Extract the file name
    parent_link_stem = Path(parent_link_name).stem  # Get the stem (name without extension)
    
    for new_link in new_links:
        new_link_name = new_link.split('/')[-1].split('\\')[-1]  # Extract the file name
        if parent_link_stem in new_link_name:
            return parent_link, new_link
    return parent_link, None

def extract_docx(docx_path, extract_to):
    """Extracts the DOCX file to a specified directory."""
    with zipfile.ZipFile(docx_path, 'r') as docx:
        docx.extractall(extract_to)

def find_linked_excel_files(xml_content):
    """Finds linked Excel files in the given XML content."""
    pattern = re.compile(r'Target="([^"]+\.xlsx)"')
    return pattern.findall(xml_content)

def find_linked_excel_files_in_rels(extract_to):
    """Finds linked Excel files in all .rels files within the extracted DOCX content."""
    linked_files = []
    for root, dirs, files in os.walk(extract_to):
        for file in files:
            if file.endswith('.rels'):
                with open(os.path.join(root, file), 'r', encoding='utf-8') as f:
                    content = f.read()
                    linked_files.extend(find_linked_excel_files(content))
    return linked_files

def extract_excel_links(docx_path):
    """Extracts and lists all links to Excel files found in the DOCX file."""
    temp_extract_to = 'temp_extracted_docx'
    os.makedirs(temp_extract_to, exist_ok=True)
    extract_docx(docx_path, temp_extract_to)

    # Read the document.xml file
    document_xml_path = os.path.join(temp_extract_to, 'word/document.xml')
    with open(document_xml_path, 'r', encoding='utf-8') as file:
        xml_content = file.read()

    linked_excel_files = find_linked_excel_files(xml_content)
    linked_excel_files_in_rels = find_linked_excel_files_in_rels(temp_extract_to)
    all_linked_excel_files = list(set(linked_excel_files + linked_excel_files_in_rels))

    # Clean up extracted files
    shutil.rmtree(temp_extract_to)

    return all_linked_excel_files

def update_excel_links_in_rels(extract_to, link_mapping):
    """Updates all found links to Excel files to the new link paths."""
    for root, dirs, files in os.walk(extract_to):
        for file in files:
            if file.endswith('.rels'):
                file_path = os.path.join(root, file)
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                for old_link, new_link in link_mapping.items():
                    if old_link in content:
                        print(f"Updating link in: '{file_path}'")
                        content = content.replace(old_link, new_link)
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)

def repack_docx(extract_to, new_docx_path):
    """Repacks the modified contents into a new DOCX file."""
    with zipfile.ZipFile(new_docx_path, 'w', zipfile.ZIP_DEFLATED) as docx:
        for root, dirs, files in os.walk(extract_to):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, extract_to)
                docx.write(file_path, arcname)

# Get the directory of the current script

current_script_dir = Path(__file__).parent.resolve()
print(f"{'*'*20}Link Updater Running{'*'*20}")
print(f"\nPython script directory: {current_script_dir}\n")

# Extract and print all links
client_name = format_name(master_dict['Client_Name'])
output_file_name = f"{client_name}.docx"
docx_path = output_file_name       # Original Doc file to be updated
all_links = extract_excel_links(docx_path)
print(f"{len(all_links)}\n")
for links in all_links:
    print(links)
print("Existing linked Excel files:")
for index, link in enumerate(all_links):
    print(f"{index + 1}: {link}")

# User-provided list of new links, dynamically using the current script directory
new_links = [
    f'file:///{(current_script_dir / "Excel_Sheets_Linked" / f"Wall_Load_Calculation_{client_name}.xlsx")}',
    f'file:///{(current_script_dir / "Excel_Sheets_Linked" / f"Tank_Load_Calculation_{client_name}.xlsx")}',
    f'file:///{(current_script_dir / "Excel_Sheets_Linked/NBC_Checks" / f"Base_Shear_Distribution_{client_name}.xlsx")}',
    f'file:///{(current_script_dir / "Excel_Sheets_Linked/NBC_Checks" / f"CM_CR_Check_{client_name}.xlsx")}',
    f'file:///{(current_script_dir / "Excel_Sheets_Linked/NBC_Checks" / f"Torsion_Irregularity_Check_{client_name}.xlsx")}',
    f'file:///{(current_script_dir / "Excel_Sheets_Linked/NBC_Checks" / f"Mass_Irregularity_Check_{client_name}.xlsx")}',
    f'file:///{(current_script_dir / "Excel_Sheets_Linked/NBC_Checks" / f"Drift_Check_{client_name}.xlsx")}',
    f'file:///{(current_script_dir / "Excel_Sheets_Linked/NBC_Checks" / f"Soft_Story_Check_{client_name}.xlsx")}',
    f'file:///{(current_script_dir / "Excel_Sheets_Linked" / f"Column_Rebar_{client_name}.xlsx")}',
    f'file:///{(current_script_dir / "Excel_Sheets_Linked" / f"Slab_Design_{client_name}.xlsx")}',
]

# Ensure the number of new links matches the number of old links
if len(new_links) != len(all_links):
    raise ValueError("The number of new links must match the number of existing links.")

# Create the link mapping
link_mapping = {}
for old_link in all_links:
    old_link, new_link = link_matcher(old_link, new_links)
    if new_link:
        link_mapping[old_link] = new_link


# Print the mapping for verification
print("\nLink Mapping:")
for old_link, new_link in link_mapping.items():
    print(f"{old_link} -> {new_link}")

# Proceed with updating the links
extract_to = 'temp_files'  # Directory to extract DOCX contents

# Step 1: Extract the DOCX file
os.makedirs(extract_to, exist_ok=True)
extract_docx(docx_path, extract_to)

# Step 2: Update all found links to new paths
update_excel_links_in_rels(extract_to, link_mapping)

# Step 3: Repack the DOCX file
new_docx_path = f"{client_name}_Report.docx"
repack_docx(extract_to, new_docx_path)

print(f"Document saved as '{new_docx_path}'")
print(f"{'*'*20}All Operations Run Successfully{'*'*20}")
print(f"\n{'*'*20}PROGRAM END{'*'*20}")

