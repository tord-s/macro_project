from flask import Flask
from flask import send_file, render_template, request
import pandas as pd
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
default_sam_file = excel_base_path + 'sam.xlsx'
default_euklems_file = excel_base_path + 'euklems.xlsx'
default_mapping_file = excel_base_path + 'mapping.xlsx'
euklems_codes_column_name_in_excel = "EUKLEMS code"
sam_codes_column_name_in_excel = "SAM code"


def do_calculations(sam_file=default_sam_file, euklems_file=default_euklems_file, mapping_file=default_mapping_file,
                    sam_sheet_name='AT', euklems_sheet_name='W_shares', mapping_sheet_name='Sheet1', year_of_analysis='2017', country_code='AT'):
    start_time = time.time()

    """
    Generating the Pandas Dataframes from the files
    """
    # sam = pd.read_excel(sam_file, sheet_name=sam_sheet_name,
    #                     index_col=0)
    euklems = pd.read_excel(euklems_file, sheet_name=euklems_sheet_name)
    mapping = pd.read_excel(mapping_file, sheet_name=mapping_sheet_name)

    # Reads the Labour row of the SAM - Generates a dictionary
    # sam_labour_row = sam.loc['Labour'].dropna().to_dict()

    # print(sam_labour_row)
    # print('----------------------------------------')

    # sam_long = pd.read_excel('excel_files/sam_long.xlsx', sheet_name='Sheet1',
    #                          names=['country', 'gets', 'gives', 'amount'])
    sam_long = pd.read_excel(sam_file, sheet_name='Sheet1',
                             names=['country', 'gets', 'gives', 'amount'])

    sam_for_given_country = sam_long.loc[sam_long['country'] == country_code]
    sam_long_labour_row = sam_for_given_country.loc[sam_for_given_country['gets'] == 'Labour', [
        'amount', 'gives']].set_index('gives')['amount'].to_dict()
    # print(sam_for_given_country)
    # print('----------------------------------------')
    # print(sam_long_labour_row)

    sam_labour_row = sam_long_labour_row

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
            sam_total += sam_labour_row[code_from_sam.strip()]
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

    """
    Big Question here - does he want output in 'Long' format as well?
    """                                            
    output_sceleton = sam
    # Gets index of where the Labour column and row is
    labour_column_index = output_sceleton.columns.get_loc("Labour")
    labour_row_index = list(output_sceleton.index).index("Labour")

    # Adds the colums where labour is now
    i = len(combinations) - 1
    while (i > 0):
        output_sceleton.insert(labour_column_index - 1,
                               'Labour ' + combinations[i], None)
        i -= 1

    # Splits data to be able to insert between rows
    first_part = output_sceleton.copy(deep=True).iloc[0:labour_row_index]
    second_part = output_sceleton.copy(deep=True).iloc[labour_row_index:]

    # inserts the data in the right rows and columns
    for row, column, value in data_to_be_added:
        first_part.loc[row, column] = value

    # Merges the data again after the rows have been inserted
    output_sceleton = first_part.append(second_part)

    for combo in combinations:
        combo_values = list(output_sceleton.loc['Labour ' +
                                                combo].dropna().to_dict().values())
        combo_total = 0
        for value in combo_values:
            combo_total += float(value)
        output_sceleton.loc['HOUS', 'Labour ' + combo] = combo_total

    # Removes more general Labour column and row
    output_sceleton.drop(['Labour'], axis=1, inplace=True)
    output_sceleton.drop(['Labour'], axis=0, inplace=True)

    with pd.ExcelWriter(excel_base_path + "output.xlsx") as writer:
        output_sceleton.to_excel(writer, verbose=True,
                                 index=True, sheet_name='Results')
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
        mapping = request.files['mapping']
        sam_sheet_name = request.form['sam_sheet_name']
        euklems_sheet_name = request.form['euklems_sheet_name']
        mapping_sheet_name = request.form['mapping_sheet_name']
        year_of_analysis = request.form['year_of_analysis']
        print(sam_sheet_name)
        used_time = do_calculations(
            sam, euklems, mapping, sam_sheet_name, euklems_sheet_name, mapping_sheet_name, year_of_analysis)
        print('Used time: ', used_time)
        return send_file(excel_base_path + "output.xlsx", attachment_filename='resulting_output_id' + str(round(time.time(), 0))[: -2] + '.xlsx')
    try:
        return send_file(excel_base_path + "output.xlsx", attachment_filename='resulting_output_id' + str(round(time.time(), 0))[: -2] + '.xlsx')
    except Exception as e:
        return str(e)


if __name__ == "__main__":
    print(do_calculations())
