import yaml
import pandas as pd
import os
import random




# INPUT FILES:

# Settings:
yaml_file = r"C:\Users\2022547\Documents\D\CODE\Joiners\Input\Settings.yaml"
# Data (containing unique ID, application, type, date)
data_file = r"C:\Users\2022547\Documents\D\CODE\Joiners\Input\Primary input file.xlsx"



# Settings (from settings (yaml) file:

def yaml_loader(yaml_file):
    """Loads yaml file"""
    with open(yaml_file, "r") as file_descriptor:
        settings_data = yaml.load(file_descriptor)
    return settings_data


# Settings
settings_data = yaml_loader(yaml_file)

start_date = settings_data[0]
end_date = settings_data[1]
applications = settings_data[2]

apps = applications.get('Applications')
start_date = start_date.get('Start date')
end_date = end_date.get('End date')





# Loop for each application
for app in apps:
    df = pd.read_excel(data_file)

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