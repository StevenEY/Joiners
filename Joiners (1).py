import yaml
import pandas as pd
import numpy as np
import os
import random
import math
import xlwings as xw
from xlwings.constants import DeleteShiftDirection
from xlwings.constants import InsertShiftDirection





# ============================== FILE PATHS AND YAML FUNCTION =============================


# FILE READING:

# Data (containing unique ID, application, type, date):
extract_file = r"C:\Users\2022547\Documents\D\CODE\Joiners\Input\New primary input file2.xlsx"
# File reading for input extract:
yaml_extraction = r"C:\Users\2022547\Documents\D\CODE\Joiners\Input\New extraction settings2.yaml"
# File reading for form filling:
yaml_form = r"C:\Users\2022547\Documents\D\CODE\Joiners\Input\New form settings example2.yaml"
# 310GL form
excel_form = r"C:\Users\2022547\Documents\D\CODE\Joiners\Input\310GL-ITGC testing form.xlsx"


# YAML FILE LOADING FUNCTION

def yaml_loader(yaml_file):
    """Loads yaml file"""
    with open(yaml_file, "r") as file_descriptor:
        settings_data = yaml.load(file_descriptor)
    return settings_data





# ============================== EXTRACT SAMPLE RANDOM SELECTION ==============================


# LOADING EXTRACT SETTINGS & SETTING UP COMPONENT VARS

settings_data = yaml_loader(yaml_extraction)


#IMPORTANT: if yaml key changes, the below (within []), must be modified accordingly
start_date = settings_data['Start date']
end_date = settings_data['End date']
apps = settings_data['Applications']
mandatory_cols_list = settings_data['Mandatory columns and matching column name in form table']
additional_cols_list = settings_data['Optional columns and matching column name in form table']

# mandatory cols
# IMPORTANT: if yaml dict key changes, the below (within ''), must be modified accordingly
INP_COL_user_id = [list(i.values())[0] for i in mandatory_cols_list if list(i.keys())[0] == 'User ID'][0]
INP_COL_application = [list(i.values())[0] for i in mandatory_cols_list if list(i.keys())[0] == 'Application'][0]
INP_COL_type = [list(i.values())[0] for i in mandatory_cols_list if list(i.keys())[0] == 'Type'][0]
INP_COL_date = [list(i.values())[0] for i in mandatory_cols_list if list(i.keys())[0] == 'Date'][0]




len_apps = len(apps)



df_input_spreadsheet = pd.read_excel(extract_file)  # empty df w/ input_spreadsheet's header
df_random_selection_from = pd.DataFrame(columns=list(df_input_spreadsheet.columns))

# LOOP FOR EACH APPLICATION

for app in apps:
    df = pd.read_excel(extract_file)

    # Filtering applications
    df = df[df[INP_COL_application].str.contains(app)]

    # Filtering type
    df = df[df[INP_COL_type].str.contains("New]")]  # TODO: find how to have "[New]" #TODO:Consider placing parameter in var/sth else if other engs/apps have dif type indicator. Could have a filter keyword in yaml file

    # Filtering N/A dates
    df = df[df[INP_COL_date].notnull()]  # TODO: maybe - create DF containing the rows with nulls

    # Filtering out-of-scope dates
    df = df[df[INP_COL_date].isin(pd.date_range(start_date, end_date))]

    df_random_selection_from = df_random_selection_from.append(df)



# Random selection

df_length = len(df) - 1
df_length_range = [i for i in range(df_length)]
print(df_length_range)

df_length_roundup_ten = int(math.ceil((df_length + 1) / 10.0)) * 10
sample_size = int(df_length_roundup_ten / 10)


print('df_length:', df_length)

if sample_size > 25:
    sample_size = 25
elif sample_size < 5:
    sample_size = 5


cryptogen = random.SystemRandom()
selected_sample = cryptogen.sample(df_length_range, sample_size)  # TODO: maybe print the selected sample (and index of DF)

selected_sample.sort()  # sorting random nos numerically

selected_sample_df = df.iloc[selected_sample, :]

print(selected_sample_df)


writer = pd.ExcelWriter(f'Selected sample - {app}.xlsx')
selected_sample_df.to_excel(writer)
writer.save()



# ========================= EXCEL FORM FILLING SECTION =========================

# LOADING FORM SETTINGS & SETTING UP COMPONENT VARS

settings_data = yaml_loader(yaml_form)



# Assigning values from yaml form
#IMPORTANT: if yaml key changes, the below (within []), must be modified accordingly
V_wp_ref = settings_data['Working paper ref']
V_itgc_ref  = settings_data['ITGC name/ref']
V_entity_name  = settings_data['Entity name']

V_wp_ref_assess = settings_data['Workpaper reference of our assessment of the design of the control']
V_canvas_link = settings_data['Canvas link to our assessment of the design of the control']
V_ctrl_descr = settings_data['Control description']
V_it_process = settings_data['IT process']
V_frequency = settings_data['Frequency']
V_other = settings_data['If other, describe']

V_ind_testing = settings_data['Independent testing']
V_reperformance_ia = settings_data['Reperformance of the work of IA or others']
V_review = settings_data['Review of the work performed by IA or others']

V_inquiry = settings_data['Inquiry regarding performance of the control']
V_observation = settings_data['Observation of the control taking place']
V_inspection = settings_data['Inspection of relevant documentation']
V_reperformance_ctrl = settings_data['Reperformance of the control']

V_population_source = settings_data['Population source']
V_descr_procedures = settings_data['Description of the procedures performed to determine how the population was determined to be complete']

V_procedures_perf = settings_data['Procedures performed']

V_tab_name = settings_data['Name of form tab']


# Assigning Cell Position
#IMPORTANT: if yaml key changes, the below (within []), must be modified accordingly
P_wp_ref = settings_data['Cell position - Working paper ref']
P_itgc_ref  = settings_data['Cell position - ITGC name/ref']
P_entity_name  = settings_data['Cell position - Entity name']
P_audit_date = settings_data['Cell position - Audit date']

P_wp_ref_assess = settings_data['Cell position - Workpaper reference of our assessment of the design of the control']
P_ctrl_descr = settings_data['Cell position - Control description']
P_it_process = settings_data['Cell position - IT process']
P_frequency = settings_data['Cell position - Frequency']
P_other = settings_data['Cell position - If other, describe']

P_ind_testing = settings_data['Cell position - Independent testing']
P_reperformance_ia = settings_data['Cell position - Reperformance of the work of IA or others']
P_review = settings_data['Cell position - Review of the work performed by IA or others']

P_inquiry = settings_data['Cell position - Inquiry regarding performance of the control']
P_observation = settings_data['Cell position - Observation of the control taking place']
P_inspection = settings_data['Cell position - Inspection of relevant documentation']
P_reperformance_ctrl = settings_data['Cell position - Reperformance of the control']

P_from_date = settings_data['Cell position - Period covered by the test From date']
P_to_date = settings_data['Cell position - Period covered by the test To date']

P_population_source = settings_data['Cell position - Population source']
P_population_size = settings_data['Cell position - Population size']
P_sample_size = settings_data['Cell position - Sample size']
P_descr_procedures = settings_data['Cell position - Description of the procedures performed to determine how the population was determined to be complete']
P_sample_method = settings_data['Cell position - Sample selection method']

P_testing_table = settings_data['Cell position - Testing table']
P_procedures_perf = settings_data['Cell position - Procedures performed']



# xlwings sheet set-up
wb = xw.Book(excel_form)
sht = wb.sheets[V_tab_name]  # Excel tab name



sht.range(P_wp_ref).value = V_wp_ref
sht.range(P_itgc_ref).value = V_itgc_ref
sht.range(P_entity_name).value = V_entity_name
sht.range(P_audit_date).value = end_date

sht.range(P_wp_ref_assess).add_hyperlink(V_canvas_link)
sht.range(P_wp_ref_assess).value = V_wp_ref_assess
sht.range(P_ctrl_descr).value = V_ctrl_descr
sht.range(P_it_process).value = V_it_process
sht.range(P_frequency).value = V_frequency
sht.range(P_other).value = V_other

sht.range(P_ind_testing).value = V_ind_testing
sht.range(P_reperformance_ia).value = V_reperformance_ia
sht.range(P_review).value = V_review

sht.range(P_inquiry).value = V_inquiry
sht.range(P_observation).value = V_observation
sht.range(P_inspection).value = V_inspection
sht.range(P_reperformance_ctrl).value = V_reperformance_ctrl

sht.range(P_from_date).value = start_date
sht.range(P_to_date).value = end_date
sht.range(P_population_source).value = V_population_source
sht.range(P_population_size).value = df_length + 1

if df_length + 1 < 50:
    sht.range(P_sample_size).value = f"As per GAM methodology, 10% of the in-scope joiners ({df_length + 1}) with a minimum sample limit (5) applied applied is: {sample_size}"
elif df_length + 1 > 250:
    sht.range(P_sample_size).value = f"As per GAM methodology, 10% of the in-scope joiners ({df_length + 1}) with a maximum sample limit (25) applied applied is: {sample_size}"
else:
    sht.range(P_sample_size).value = f"As per GAM methodology, 10% of the in-scope joiners ({df_length + 1}) rounded up is: {sample_size}"

sht.range(P_descr_procedures).value = V_descr_procedures
sht.range(P_sample_method).value = f"Random:\nThe {sample_size} samples were randomly selected from the new joiner population.\nRandom selection derived from random bytes generated by the Windows cryptographically secure pseudorandom number generator."







def FUNCTION_testing_table():

    # Prep for DF header
    mandatory_cols_header_list = [list(i.values())[0] for i in mandatory_cols_list]
    additional_cols_header_list = [list(i.values())[0] for i in additional_cols_list]
    ###combo_cols_header_list = mandatory_cols_header_list.append(additional_cols_header_list)




    # Inverting dictionary to change col names from old to new


    mandatory_cols_header_list_new_header = [list(i.keys())[0] for i in mandatory_cols_list]
    additional_cols_header_list_new_header = [list(i.keys())[0] for i in additional_cols_list]

    table = selected_sample_df[mandatory_cols_header_list + additional_cols_header_list]
    table.columns = mandatory_cols_header_list_new_header + additional_cols_header_list_new_header





    # Setting table index/Testing #
    table.index = np.arange(1, len(table) + 1)
    table.index.name = 'Testing #'

    # Setting table header
    attribute_letters = [chr(i).upper() for i in range(ord('a'), ord('z')+1)][:len(V_procedures_perf)]
    combined_columns = list(table.columns) + attribute_letters + ['Evidence']
    table = table.reindex(columns=combined_columns)


    # Replacing table
    sht.range(P_testing_table).api.Delete(DeleteShiftDirection.xlShiftUp)
    P_testing_table_first_row = P_testing_table.rsplit(':', 1)[0]
    for i in range(len(table) + 1):
        sht.range(P_testing_table_first_row + ':' + P_testing_table_first_row).api.Insert(InsertShiftDirection.xlShiftToRight)

    # Merging cols D & E -- unused because table would be missing a col
    #limit = len(table)
    #add_rows = 0
    #while limit > 0:
    #    sht.range('D' + str(int(P_testing_table_first_row) + add_rows) + ':' + 'E' + str(int(P_testing_table_first_row) + add_rows)).api.MergeCells = True
    #    limit -= 1
    #    add_rows += 1

    # Removing formatting
    sht.range('B' + str(P_testing_table_first_row) + ':' + [chr(i).upper() for i in range(ord('a'), ord('z')+1)][:len(table.columns) + 2][-1] + str(int(P_testing_table_first_row) + len(table))).api.ClearFormats()

    sht.range('B' + P_testing_table_first_row).value = table

    sht.range('D:D').api.Columns.AutoFit()  # reason: by default, D width very small
    sht.range('L:L').api.Columns.AutoFit()  # reason: by default, L width very large


    # Applying formatting
    sht.range('B' + str(P_testing_table_first_row) + ':' + [chr(i).upper() for i in range(ord('a'), ord('z')+1)][:len(table.columns) + 2][-1] + str(int(P_testing_table_first_row) + len(table))).api.Borders.LineStyle=1
    sht.range('B' + str(P_testing_table_first_row) + ':' + [chr(i).upper() for i in range(ord('a'), ord('z')+1)][:len(table.columns) + 2][-1] + str(int(P_testing_table_first_row) + len(table))).api.Borders.Weight=2

    sht.range('B' + str(P_testing_table_first_row) + ':' + [chr(i).upper() for i in range(ord('a'), ord('z')+1)][:len(table.columns) + 2][-1] + str(P_testing_table_first_row)).api.Interior.ColorIndex = 16
    sht.range('B' + str(P_testing_table_first_row) + ':' + [chr(i).upper() for i in range(ord('a'), ord('z')+1)][:len(table.columns) + 2][-1] + str(P_testing_table_first_row)).api.Font.ColorIndex = 2
    sht.range('B' + str(P_testing_table_first_row) + ':' + [chr(i).upper() for i in range(ord('a'), ord('z')+1)][:len(table.columns) + 2][-1] + str(P_testing_table_first_row)).api.font.bold = True

    sht.range('B' + str(P_testing_table_first_row) + ':' + 'B' + str(int(P_testing_table_first_row) + len(table))).api.Interior.ColorIndex = 16
    sht.range('B' + str(P_testing_table_first_row) + ':' + 'B' + str(int(P_testing_table_first_row) + len(table))).api.Font.ColorIndex = 2
    sht.range('B' + str(P_testing_table_first_row) + ':' + 'B' + str(int(P_testing_table_first_row) + len(table))).api.font.bold = True


def FUNCTION_procedures_perf():
    """
    a.k.a. attributes.
    For loop -formatting and values for each item (attribute) in yaml_extraction
    """

    sht.range(P_procedures_perf).api.Delete(DeleteShiftDirection.xlShiftUp)  # Deleting attribute rows

    P_procedures_perf_first_row = P_procedures_perf.rsplit(':', 1)[0]

    for i in range(len(V_procedures_perf)):
        sht.range(P_procedures_perf_first_row + ':' + P_procedures_perf_first_row).api.Insert(InsertShiftDirection.xlShiftToRight)  # Inserting rows (number of attributes)
        sht.range('C' + P_procedures_perf_first_row + ':K' + P_procedures_perf_first_row).api.MergeCells = True

        sht.range('B' + P_procedures_perf_first_row + ':K' + P_procedures_perf_first_row).api.Borders.LineStyle=1
        sht.range('B' + P_procedures_perf_first_row + ':K' + P_procedures_perf_first_row).api.Borders.Weight=2

        sht.range('B' + P_procedures_perf_first_row).api.Interior.ColorIndex=16
        sht.range('C' + P_procedures_perf_first_row).api.Interior.ColorIndex=2

        sht.range('B' + P_procedures_perf_first_row).api.Font.ColorIndex=2
        sht.range('C' + P_procedures_perf_first_row).api.Font.ColorIndex=1

        sht.range('B' + P_procedures_perf_first_row).api.font.bold = True
        sht.range('C' + P_procedures_perf_first_row).api.font.bold = False

        sht.range('B' + P_procedures_perf_first_row).value = [chr(i).upper() for i in range(ord('a'), ord('z') + 1)][:len(V_procedures_perf)][::-1][i]  # reverse order of attribute letters
        sht.range('C' + P_procedures_perf_first_row).value = V_procedures_perf[::-1][i]  # reverse order of V_procedures_perf



# RUNNING THE ABOVE FUNCTIONS (in order so that the bottom-most set in form is completed first
# in order for the top set not to interfere with the bottom set

P_testing_table_first_row = P_testing_table.rsplit(':', 1)[0]
P_procedures_perf_first_row = P_procedures_perf.rsplit(':', 1)[0]

if P_testing_table_first_row > P_procedures_perf_first_row:
    FUNCTION_testing_table()
    FUNCTION_procedures_perf()

else:
    FUNCTION_procedures_perf()
    FUNCTION_testing_table()












#wb.save(r'C:\Users\2022547\Documents\D\CODE\Joiners\Input\ex_xlw3.xlsx')





