import pandas as pd

"""
Some definitions
"""
combinations = ['111', '112', '113', '121', '122', '123', '131', '132', '133',
                '211', '212', '213', '221', '222', '223', '231', '232', '233']
excel_base_path = 'excel_files/'
sam_file = excel_base_path + 'sam.xlsx'
euklems_file = excel_base_path + 'euklems.xlsx'
settings_file = excel_base_path + 'settings.xlsx'
euklems_codes_column_name_in_excel = "EUKLEMS code"
sam_codes_column_name_in_excel = "SAM code"

"""
Generating the Pandas Dataframes from the files
"""
sam = pd.read_excel(sam_file, sheet_name='AT',
                    index_col=0)
euklems = pd.read_excel(euklems_file, sheet_name='W_shares')
settings = pd.read_excel(settings_file, sheet_name='Settings')

# Reads the Labour row of the SAM - Generates a dictionary
sam_labour_row = sam.loc['Labour'].dropna().to_dict()

"""
Mapping codes SAM/EUKLEMS
"""
all_euklems_codes = settings.get(euklems_codes_column_name_in_excel).dropna()
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
    columns_to_be_added = {}
    for combo in combinations:
        columns_to_be_added['Labour ' + combo] = {'Labour ' + combo: None}
    columns_df = pd.DataFrame(columns_to_be_added)
    labour_index = len(sam.index)
    output_sceleton = sam.append(columns_df)

    for row, column, value in data_to_be_added:
        output_sceleton.loc[row, column] = value

    output_sceleton.to_excel(writer, verbose=True,
                             index=True, sheet_name='Results')
