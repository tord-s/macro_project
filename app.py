from flask import Flask
from flask import send_file, render_template, request
import pandas as pd

"""
Some definitions
"""
combinations = ['111', '112', '113', '121', '122', '123', '131', '132', '133',
                '211', '212', '213', '221', '222', '223', '231', '232', '233']
excel_base_path = 'excel_files/'
default_sam_file = excel_base_path + 'sam.xlsx'
default_euklems_file = excel_base_path + 'euklems.xlsx'
default_settings_file = excel_base_path + 'settings.xlsx'
euklems_codes_column_name_in_excel = "EUKLEMS code"
sam_codes_column_name_in_excel = "SAM code"

"""
Generating the Pandas Dataframes from the files
"""


def insert_row(row_number, df, row_name):
    # Slice the upper half of the dataframe
    df1 = df[0:row_number]

    # Store the result of lower half of the dataframe
    df2 = df[row_number:]

    # Insert the row in the upper half dataframe
    print(df1.loc[row_number])
    df1.loc[row_number] = row_name

    # Concat the two dataframes
    df_result = pd.concat([df1, df2])

    # Reassign the index labels
    df_result.index = [*range(df_result.shape[0])]

    # Return the updated dataframe
    return df_result


def do_calcululations(sam_file=default_sam_file, euklems_file=default_euklems_file, settings_file=default_settings_file):
    sam = pd.read_excel(sam_file, sheet_name='AT',
                        index_col=0)
    euklems = pd.read_excel(euklems_file, sheet_name='W_shares')
    settings = pd.read_excel(settings_file, sheet_name='Settings')

    # Reads the Labour row of the SAM - Generates a dictionary
    sam_labour_row = sam.loc['Labour'].dropna().to_dict()

    """
        Mapping codes SAM/EUKLEMS
        """
    all_euklems_codes = settings.get(
        euklems_codes_column_name_in_excel).dropna()
    all_sam_codes = settings.get(sam_codes_column_name_in_excel).dropna()
    mapped_codes = []
    for i in range(0, len(all_sam_codes)):
        sam_codes = all_sam_codes[i].split(', ')
        euklems_codes = all_euklems_codes[i].split(', ')
        mapped_codes.append((sam_codes, euklems_codes))

    year_of_analysis = int(settings.loc[0, 'YoA'])

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

    with pd.ExcelWriter(excel_base_path + "output.xlsx") as writer:

        output_sceleton = sam
        labour_column_index = output_sceleton.columns.get_loc("Labour")
        labour_row_index = list(output_sceleton.index).index("Labour")

        i = len(combinations) - 1
        while (i > 0):
            output_sceleton.insert(labour_column_index - 1,
                                   'Labour ' + combinations[i], None)
            i -= 1

        first_part = output_sceleton.copy(deep=True).iloc[0:labour_row_index]
        second_part = output_sceleton.copy(deep=True).iloc[labour_row_index:]

        for row, column, value in data_to_be_added:
            first_part.loc[row, column] = value

        output_sceleton = first_part.append(second_part)

        output_sceleton.drop(['Labour'], axis=1, inplace=True)
        output_sceleton.drop(['Labour'], axis=0, inplace=True)

        print(output_sceleton)

        output_sceleton.to_excel(writer, verbose=True,
                                 index=True, sheet_name='Results')


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
        settings = request.files['settings']
        do_calcululations(sam, euklems, settings)
        return send_file(excel_base_path + "output.xlsx", attachment_filename='resulting_output.xlsx')
    try:
        return send_file(excel_base_path + "output.xlsx", attachment_filename='resulting_output.xlsx')
    except Exception as e:
        return str(e)


if __name__ == "__main__":
    do_calcululations()
