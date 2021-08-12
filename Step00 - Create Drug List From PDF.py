# %%

# Client Name: Corteva
# Client Code/Name: 0091COR32
# Project Code/Name: Audit


# Initial Author: Mason Boles

# Objective: Read provided PDF files for drugs covered/on formulary. Product outputs that can be used in SAS.

# Developer Notes:

# ****** DEFINED VARIABLES *****
# <Please define all quantifiable assumptions here>
# <Include both Numeric and String/List assumptions>


# /**** LIBRARIES, LOCATIONS, LITERALS, ETC. GO ABOVE HERE ****/


# ****** DOCUMENTATION OF INITIAL CHECKING *****
# Checking Approach:

# Program Summary above is complete and useful?
# All variables clearly defined above; no explicit assumptions embedded in code below?
# Code is easy to navigate and understand; useful names are used for all things?
# Assumptions & variables reviewed for reasonableness and/or supporting documentation?
# Results are compared to prior iterations or otherwise tested for reasonableness?

# *ONLY SIGN IF NO ISSUES REMAIN*
# Initial Checker Name:
# Initial Checker Date:

# ****** DOCUMENTATION OF SUBSEQUENT CHANGES *****
# Change Author & Date: Dylan Bakr 8/11/2021
# Description of Change: Making program more robust for handling files generally
# Name of Checker:
# Date of Checking:
# How it was Checked:

# Change Author & Date:
# Description of Change:
# Name of Checker:
# Date of Checking:
# How it was Checked:

# Change Author & Date:
# Description of Change:
# Name of Checker:
# Date of Checking:
# How it was Checked:

# %% Import Libraries

import PyPDF2
import pikepdf as pike
import pandas as pd
import os


# %% Define Functions

# 1) PDF text extraction
def text_extraction(decrypt_file, name):
    # Generate a text string with the entire PDF file inside of it
    pdf = PyPDF2.PdfFileReader(open(decrypt_file, 'rb'))

    # Combine all pages into a single large string
    whole_pdf = ''
    for i in range(pdf.getNumPages()):
        page = pdf.getPage(i)
        page_contents = page.extractText()
        whole_pdf = whole_pdf + page_contents
    # print(whole_pdf)

    # Find the section 'QUICK REFERENCE DRUG LIST' that flags the beginning of the drug list that we want
    # Also find the section 'PREFERRED OPTIONS FOR EXCLUDED SPECIALTY MEDICATIONS' that flags the end of what we want
    begin_strip = whole_pdf.find('QUICK REFERENCE')
    end_strip = whole_pdf.find('PREFERRED OPTIONS')

    # Strip the whole_pdf string into an array that will populate our output
    split_array = whole_pdf[begin_strip:end_strip].split('\n')
    print(split_array)

    # Navigate through the large array using the following principles:
    # 1) The drugs are separated by first-letter, so the first drug will begin after an entry of just 'A'
    # 2) All empty/letter designating (just 1 character)/plain text entries need to be found and removed
    # for i in range(len(split_array)):
    #     # if split_array[i] == 'A':
    #     if split_array[i].find('A ') != -1:
    #         begin = i
    #         break
    # extract_section = split_array[begin:]
    # print(extract_section)

    # Create our first pass array - remove obvious phrases or empty entries
    first_pass = []
    for j in range(len(split_array)):
        # allow the hyphoned names to create a hyphon value in our list
        if split_array[j] == '-':
            first_pass.append(split_array[j])
        # only output entries greater than 1 character long and specifically remove plain text
        elif split_array[j] not in ['', ' ', 'QUICK REFERENCE DRUG', 'LIST', ' LIST',
                                    'Your specific prescription benefit plan design may not cover certain products or categories, regardless of their appearance i',
                                    'n this ', 'document. For specific information, visit ', 'Caremark.com',
                                    'or contact a CVS', 'Caremark Cu', 'stomer Care representative.']:
            first_pass.append(split_array[j])
    print(first_pass)

    # Combine hyphoned named where possible, then remove 1 character long entries
    combine_hyphon = []
    for k in range(len(first_pass)):
        if first_pass[k] != '-':
            combine_hyphon.append(first_pass[k])
        else:
            # Hyphon-named drug
            hyphon_drug = first_pass[k - 1] + first_pass[k] + first_pass[k + 1]
            combine_hyphon.append(hyphon_drug)

    # Remove the 1 digit entries and we are done
    final_drug_list = []
    for z in range(len(combine_hyphon)):
        # Also remove any extra 'rel' products generated during this hyphon business
        if len(combine_hyphon[z]) != 1 and combine_hyphon[z] != 'rel':
            final_drug_list.append(combine_hyphon[z])

    Brand_Generic(final_drug_list)


# 2) Create a corresponding Brand/Generic flag array determined by the case of the drugs
#   - Also create final table layout for export

def Brand_Generic(final_drug_list):
    brand_generic = []
    for i in range(len(final_drug_list)):
        if final_drug_list[i].isupper():
            b_g = 'Brand'
        else:
            b_g = 'Generic'
        brand_generic.append(b_g)

    data = {"Drug": final_drug_list, "Brand_Generic": brand_generic}
    data_frame = pd.DataFrame(data)
    print(data_frame)

    # Export results to excel
    # Base location for results
    python_results = 'C:\\Users\\C-Dylan.Bakr\\Documents\\Top Secret\\Python Results\\'
    output_file = python_results + name.replace('.pdf', '.xlsx')
    data_frame.to_excel(output_file)


# %% Generate a list of all files we need to decrypt and pull drugs from

base_location = 'C:\\Users\\C-Dylan.Bakr\\Documents\\Top Secret\\Formularies\\'
copy_location = 'C:\\Users\\C-Dylan.Bakr\\Documents\\Top Secret\\Formularies Copies\\'
current_files = os.listdir(base_location)

for name in current_files:
    pdf_file = base_location + name
    with pike.open(pdf_file, password='', allow_overwriting_input=True) as pdf:
        print("Working on file: " + str(pdf_file))
        print("Total pages: " + str(len(pdf.pages)))
        pdf.save(pdf_file)
        text_extraction(pdf_file, name)

# %%
