# Program to get required fields from G-star pdf files
# Importing required libraries
import os
import re
import tabula
import pandas as pd
import pdfplumber
import time
# Function getting header information from every pdf
# Input Param: pdf file
# Output format: dictionary containing header information
def getting_header(pdf_file):
    required_dict = {}
    # Appending static data to required dictionary
    required_dict['Inquiry Source'] = 'FABRIC AND GARMENT'
    required_dict['Customer'] = 'G-STAR INTERNATIONAL B.V.'

    # Code to read first line
    with pdfplumber.open(pdf_file) as pdf:
        first_page = pdf.pages[0]
        # Reading only one line of first page
        text = first_page.extract_text().split('\n')
        first_line = text[0:1]
        first_line = [re.split(r"\s+", ele) for ele in first_line]
        required_dict['Buyer Style Ref'] = first_line[0][9] + ' (' + first_line[0][0] + ')'
    return required_dict

# Function for processing page content
# Input Param: pdf file
# Output format: dictionary containing page key and value
def getting_page_content(pdf_file):
    # Read pdf into a list of DataFrame
    df_list = tabula.read_pdf(pdf_file, pages="1", lattice=True, multiple_tables=True)
    # Seperating df based on length
    zero_len_df = [df for df in df_list if len(df) == 0]
    one_len_df = [df for df in df_list if len(df) == 1]
    n_len_df = [df for df in df_list if len(df) > 1]
    # Remove df having unnamed column
    one_len_df = [i for i in one_len_df for col in i.columns if 'Unnamed' not in col]

    # Code for zero length dataframe
    L1 = []
    L1_flat = []
    for i in zero_len_df:
        if len(i.columns) == 2:
            L1.append(i.columns.to_list())
    for sublist in L1:
        for item in sublist:
            L1_flat.append(item)
    def Convert(lst):
        res_dct = {lst[i]: lst[i + 1] for i in range(0, len(lst), 2)}
        return res_dct
    final_zero_dict = Convert(L1_flat)

    # code for n length dataframe
    color = n_len_df[0]['Colorway'].to_list()
    color = [ele for ele in color if str(ele) != 'nan']
    final_n_dict = {'Shade': color}

    # Getting header data
    header_data = getting_header(pdf_file)
    # Combining header data and page content data
    total_data = {**header_data, **final_zero_dict, **final_n_dict}
    return total_data


# Function that will process each pdf file anf returns dictionary object
# Input param: pdf file
# Output format: dict of required fields
def pdf_processing(pdf_file):
    temp_dict = {}
    total_processed_data = getting_page_content(pdf_file)
    # Getting processed data
    # Declaring dataframe with necessary columns
    required_df = pd.DataFrame(
        columns=['Inquiry Source', 'Style Name', 'Shade', 'Default Size', 'Buyer Style Ref', 'Customer'])

    dict_key_list = total_processed_data.keys()
    for i in required_df.columns:
        if i in dict_key_list:
            temp_dict[i] = total_processed_data[i]

    # processing df for proper format
    required_df = pd.DataFrame(temp_dict)
    required_df['Style Description/Shade/Size'] = required_df['Style Name'] + '/' + required_df['Shade'] + '/' + \
                                                  required_df['Default Size']
    required_df.drop(['Style Name', 'Shade', 'Default Size'], inplace=True, axis=1)
    # Converting to dict
    required_dict = required_df.to_dict('records')
    return required_dict

if _name_ == '_main_':
    start_time = time.time()
# List of all pdf to be processed
dir_path = 'D:\Working Directory\TechPack to GRID\G-Star\Fw__Different_tech_pack_for_Gstar'
full_path = []
final_dict_list = []
# Initializing final format
final_format = pd.DataFrame(
    columns=['SO_INDEX', 'DOC_TYPE', 'SALES_ORG', 'DISTR_CHAN', 'DIVISION', 'SALES_GRP', 'SALES_OFF',
             'ORD_REASON Code', 'MATERIAL Code', 'Cotton Type', 'SAL_EMP Code', 'Developer Code',
             'Customer', 'Inq Ref', 'Inq date', 'Delivery Date', 'Inquiry Source', 'Initial Inquiry T',
             'Style Description/Shade/Size', 'Buyer Style Ref', 'Customer',
             'Requested By', 'Developers', 'Cotton Type', 'QTY', 'PLANT', 'Rate', 'Currency', 'SO TYPE'])

for path in os.listdir(dir_path):
    full_path.append(os.path.join(dir_path, path))

for pdf_file in full_path:
    final_dict_list.append(pdf_processing(pdf_file))

for list_ele in final_dict_list:
    for record in list_ele:
        final_format = final_format.append(record, ignore_index=True)

# Converting final dataframe to xlsx format
final_format.to_excel('Inquiry_format.xlsx')
print("%s seconds " % (time.time() - start_time))
