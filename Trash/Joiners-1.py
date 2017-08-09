import yaml
import pandas as pd
import os
import random
import openpyxl, sys
import mergepatch
import insertrow
from openpyxl import *


mergepatch.patch_worksheet()  # Openpyxl (2.4.1) merged cells patch


# ============================== FILE PATHS AND YAML FUNCTION =============================


# FILE READING:

# Data (containing unique ID, application, type, date):
extract_file = r"C:\Users\2022547\Documents\D\CODE\Joiners\Input\Primary input file.xlsx"
# File reading for input extract:
yaml_extraction = r"C:\Users\2022547\Documents\D\CODE\Joiners\Input\Extraction settings.yaml"
# File reading for form filling:
yaml_form = r"C:\Users\2022547\Documents\D\CODE\Joiners\Input\Form settings example.yaml"


# YAML FILE LOADING FUNCTION

def yaml_loader(yaml_file):
    """Loads yaml file"""
    with open(yaml_file, "r") as file_descriptor:
        settings_data = yaml.load(file_descriptor)
    return settings_data





# ============================== EXTRACT SAMPLE RANDOM SELECTION ==============================


# LOADING EXTRACT SETTINGS & SETTING UP COMPONENT VARS

settings_data = yaml_loader(yaml_extraction)

start_date = settings_data[0]
end_date = settings_data[1]
applications = settings_data[2]

apps = applications.get('Applications')
start_date = start_date.get('Start date')
end_date = end_date.get('End date')

mandatory_cols =






# LOOP FOR EACH APPLICATION

for app in apps:
    df = pd.read_excel(extract_file)

    # Filtering applications
    df = df[df['Application'].str.contains(app)]

    # Filtering type
    df = df[df['Type'].str.contains(
        "New]")]  # TODO: find how to have "[New]" #TODO:Consider placing parameter in var/sth else if other engs/apps have dif type indicator. Could have a filter keyword in yaml file

    # Filtering N/A dates
    df = df[df['Date'].notnull()]  # TODO: maybe - create DF containing the rows with nulls

    # Filtering out-of-scope dates
    df = df[df['Date'].isin(pd.date_range(start_date, end_date))]

    # Random selection

    df_length = len(df) - 1
    df_length_range = [i for i in range(df_length)]
    print(df_length_range)

    sample_size = int(df_length / 10)

    print('df_length:', df_length)

    if sample_size > 25:
        sample_size = 25
    elif sample_size < 5:
        sample_size = 5


    cryptogen = random.SystemRandom()
    selected_sample = cryptogen.sample(df_length_range,
                                       sample_size)  # TODO: maybe print the selected sample (and index of DF)

    selected_sample.sort()  # sorting random nos numerically

    selected_sample_df = df.iloc[selected_sample, :]

    print(selected_sample_df)


    writer = pd.ExcelWriter(f'Selected sample - {app}.xlsx')
    selected_sample_df.to_excel(writer)
    writer.save()




# ========================= EXCEL FORM FILLING SECTION =========================

# LOADING FORM SETTINGS & SETTING UP COMPONENT VARS

settings_data = yaml_loader(yaml_form)


# import openpyxl
# read workbook
# read sheet - Put in advanced settings the name of the sheet


wb = openpyxl.load_workbook(r"C:\Users\2022547\Documents\D\CODE\Joiners\Input\310GL-ITGC testing form.xlsx")
sheet = wb.get_sheet_by_name('ITGC Testing')

sheet['C1'].value = 'This the ref'
sheet['F51'].value = 'This is in the top left corner'
sheet['K64'].value = 'This is in the bottom right corner'



##D_wp_ref = settings_data[0]
##D_itgc_ref = settings_data[1]
##D_entity_name = settings_data[2]
##D
#list(settings_data[0].values())[0]

#insertrow.Worksheet.insert_rows = insert_rows
openpyxl.worksheet.Worksheet.insert_rows = insertrow.insert_rows
sheet.insert_rows(75, 10, above=False, copy_style=True, fill_formulae=False)




wb.save(r'C:\Users\2022547\Documents\D\CODE\Joiners\Input\ex_copy-insertrow_75t.xlsx')

