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


all_euklems_codes = settings.get(euklems_codes_column_name_in_excel).dropna()
all_sam_codes = settings.get(sam_codes_column_name_in_excel).dropna()

mapped_codes = []
for i in range(0, len(all_sam_codes)):
    sam_codes = all_sam_codes[i].split(', ')
    euklems_codes = all_euklems_codes[i].split(', ')
    mapped_codes.append((sam_codes, euklems_codes))

year_of_analysis = 2017
data_to_be_added = []

for code_map in mapped_codes:
    """
    TODO
    Add support for many to many relation between EUKLEMS industries and SAM industries
    """
    # Get the total value of Labour for the given sectors in the SAM to be distrubuted
    sam_total = 0
    sam_total_per_sector = 0
    # Get the euklems industries with the same code
    codes_from_sam = code_map[0]
    for code_from_sam in codes_from_sam:
        sam_total += sam_labour_row[code_from_sam.strip()]
    # print(sam_total)
    sam_total_per_sector = sam_total / len(codes_from_sam)
    # print(sam_total_per_sector)

    # euklems_rows = euklems.loc[euklems['code'] == code_from_euklems]
    euklems_relevant_data = euklems[[
        'code', 'gender', 'age', 'edu', year_of_analysis]]
    # print(euklems_relevant_data)
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
    """
    TODO
    Add data
    """

    # addiding the requiered columns and rows
    combinations = ['111', '112', '113', '121', '122', '123', '131', '132', '133',
                    '211', '212', '213', '221', '222', '223', '231', '232', '233']
    columns_to_be_added = {}
    for combo in combinations:
        columns_to_be_added['Labour ' + combo] = {'Labour ' + combo: None}
    columns_df = pd.DataFrame(columns_to_be_added)
    labour_index = len(sam.index)
    output_sceleton = sam.append(columns_df)

    for row, column, value in data_to_be_added:
        # print(row, column, value)
        output_sceleton.loc[row, column] = value
        # print(data_to_be_added[datapoint].key)

    output_sceleton.to_excel(writer, verbose=True,
                             index=True, sheet_name='Results')


# xl = pd.ExcelFile(file)

# df = pd.read_excel(xl, sheet_name='AT', index_col=0, names=['Labour'])
# print(df)

# Load spreadsheet

# Print the sheet names
# print(xl.sheet_names)

# Load a sheet into a DataFrame by name: df1
# df1 = xl.parse('excel_files/example.xlsx')

"""

writer = pd.ExcelWriter('excel_files/example_write.xlsx', engine='xlsxwriter')

data = [['A', 'B'], ['C', 'D']]

df = pd.DataFrame(data, columns = ['Product', 'Price'])
# Write your DataFrame to a file
# yourData is a dataframe that you are interested in writing as an excel file
df.to_excel(writer, 'Dummy data')

# Save the result
writer.save()

"""
