from flask import Flask
from flask import send_file, render_template, request
import pandas as pd
import numpy as np
import os
import time

"""
Some definitions
"""
combinations = ['111', '112', '113', '121', '122', '123', '131', '132', '133',
                '211', '212', '213', '221', '222', '223', '231', '232', '233']
excel_base_path = '/home/tords/macro_project/excel_files/'
try:
    if os.environ['USERNAME'] == 'Tord':
        excel_base_path = 'excel_files/'
except:
    pass
default_sam_file = excel_base_path + 'sam_long.xlsx'
default_euklems_file = excel_base_path + 'euklems ny.xlsx'
# default_mapping_file = excel_base_path + 'mapping.xlsx'
euklems_codes_column_name_in_excel = "EUKLEMS code"
sam_codes_column_name_in_excel = "SAM code"


def do_calculations(sam_file=default_sam_file, euklems_file=default_euklems_file,
                    sam_sheet_name='Sheet1', euklems_sheet_name='W_shares', mapping_sheet_name='Mapping', year_of_analysis='2017', country_code='AT'):
    start_time = time.time()

    """
    Generating the Pandas Dataframes from the files
    """
    euklems = pd.read_excel(euklems_file, sheet_name=euklems_sheet_name)
    mapping = pd.read_excel(euklems_file, sheet_name=mapping_sheet_name)
    sam_long = pd.read_excel(sam_file, sheet_name=sam_sheet_name,
                             names=['country', 'gets', 'gives', 'amount'])

    sam_for_given_country = sam_long.loc[sam_long['country'] == country_code]
    sam_long_labour_row = sam_for_given_country.loc[sam_for_given_country['gets'] == 'Labour', [
        'amount', 'gives']].set_index('gives')['amount'].to_dict()

    """
    Mapping codes SAM/EUKLEMS
    """
    all_euklems_codes = mapping.get(
        euklems_codes_column_name_in_excel).dropna()
    all_sam_codes = mapping.get(sam_codes_column_name_in_excel).dropna()
    mapped_codes = []
    for i in range(0, len(all_sam_codes)):
        sam_codes = all_sam_codes[i].split(', ')
        euklems_codes = all_euklems_codes[i].split(', ')
        mapped_codes.append((sam_codes, euklems_codes))

    year_of_analysis = int(year_of_analysis)

    data_to_be_added = []
    for code_map in mapped_codes:
        sam_total = 0
        sam_total_per_sector = 0
        codes_from_sam = code_map[0]
        for code_from_sam in codes_from_sam:
            sam_total += sam_long_labour_row[code_from_sam.strip()]
        sam_total_per_sector = sam_total / len(codes_from_sam)

        euklems_relevant_data = euklems[[
            'code', 'gender', 'age', 'edu', year_of_analysis]]
        codes_from_euklems = code_map[1]
        for code_from_euklems in codes_from_euklems:
            for combo in combinations:
                e_df = euklems_relevant_data
                euklems_fraction = e_df.loc[(e_df['gender'] == float(combo[0])) & (e_df['code'] == code_from_euklems.strip()) &
                                            (e_df['age'] == float(combo[1])) & (e_df['edu'] == float(combo[2]))][year_of_analysis]
                new_value_in_sam = float(
                    euklems_fraction/100 * sam_total_per_sector)

                for code_from_sam in codes_from_sam:
                    data_to_be_added.append(('Labour ' + combo,
                                            code_from_sam, new_value_in_sam))

    # Makes the df as small as possible to make it quicker to search
    smaller_index_search_df = sam_long.loc[(sam_long['country'] == country_code) & (
        sam_long['gets'] == 'Labour')]

    # The SAM codes in order
    list_of_sam_codes = list(smaller_index_search_df['gives'])

    testing_dict = {}

    for sam_code in list_of_sam_codes:
        total_sum = float(sam_long.loc[(sam_long['country'] == country_code) & (
            sam_long['gets'] == 'Labour') & (
            sam_long['gives'] == sam_code)].amount)
        testing_dict[sam_code] = {'total_sum': total_sum, 'sum_of_parts': 0}

    """
    Formating the data to be added to a nested dictonary to make it easy to look up when inserting data
    """
    formated_data = {}
    for value in data_to_be_added:
        sam_code, labour_combo, amount = value[1], value[0], value[2]
        try:
            defined = formated_data[sam_code]
        except:
            formated_data[sam_code] = {}
        testing_dict[sam_code]['sum_of_parts'] = testing_dict[sam_code]['sum_of_parts'] + amount
        formated_data[sam_code][labour_combo] = amount

    test_success = True
    for key, value in testing_dict.items():
        if value['total_sum'] != round(value['sum_of_parts'], 2):
            print('ERROR with ', key, ' giving unequal sum ', value)
            test_success = False
    print('Test success: ', test_success)

    # The sceleton to be populated with data
    output_sceleton = sam_long.loc[(sam_long['country'] == country_code)]

    # Gets the index to insert the new rows
    labour_row_index_list = list(
        smaller_index_search_df.loc[sam_long['gets'] == 'Labour'].index)
    labour_row_index = int(labour_row_index_list[0]) - 1

    # Inserts the new rows
    i = len(combinations) - 1
    while (i >= 0):
        for code in list_of_sam_codes:
            combination = 'Labour ' + combinations[i]
            output_sceleton = pd.DataFrame(np.insert(output_sceleton.values, labour_row_index + 1, values=[
                country_code, combination, code, formated_data[code][combination]], axis=0))
        i -= 1

    # Creates the excel file with the populated sceleton
    with pd.ExcelWriter(excel_base_path + "output.xlsx") as writer:
        output_sceleton.to_excel(writer, verbose=False, header=False,
                                 index=False, sheet_name='Results')
    end_time = time.time()
    return round(end_time-start_time, 5)


app = Flask(__name__)


@ app.route('/')
def file_downloads():
    try:
        return render_template('downloads.html')
    except Exception as e:
        return str(e)


@ app.route('/results/', methods=['POST', 'GET'])
def return_files_tut():
    if request.method == 'POST':
        sam = request.files['sam']
        euklems = request.files['euklems']
        sam_sheet_name = request.form['sam_sheet_name']
        euklems_sheet_name = request.form['euklems_sheet_name']
        mapping_sheet_name = request.form['mapping_sheet_name']
        year_of_analysis = request.form['year_of_analysis']
        used_time = do_calculations(
            sam_file=sam, euklems_file=euklems, sam_sheet_name=sam_sheet_name, euklems_sheet_name=euklems_sheet_name,
            mapping_sheet_name=mapping_sheet_name, year_of_analysis=year_of_analysis)
        print('Used time: ', used_time)
        return send_file(excel_base_path + "output.xlsx", attachment_filename='Extented SAM (timestamp ' + str(round(time.time(), 0))[: -2] + ').xlsx')
    try:
        return send_file(excel_base_path + "output.xlsx", attachment_filename='Extented SAM (timestamp ' + str(round(time.time(), 0))[: -2] + '.xlsx')
    except Exception as e:
        return str(e)


if __name__ == "__main__":
    print(do_calculations())
