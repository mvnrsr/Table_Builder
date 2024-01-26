import pandas as pd

# Load the Excel file
file_path = 'Nets_spec_test.xlsx'  # Adjust the path as needed
df = pd.read_excel(file_path, sheet_name='Nets')

# Extract specific data
list_name = df.at[32, 'C']  # Cell C33
variable_names = df.at[33, 'C'].split(',')  # Cell C34, split by comma
code_label_pairs = df.loc[37:91, ['B', 'C']]  # Range B38:C92

# Creating a structured DataFrame
# Assuming each variable name should be associated with each code-label pair
structured_data = []
for var_name in variable_names:
    for _, row in code_label_pairs.iterrows():
        code, label = row['B'], row['C']
        structured_data.append({'List Name': list_name, 'Variable Name': var_name.strip(), 'Code': code, 'Label': label})

structured_df = pd.DataFrame(structured_data)

print(structured_df.head())  # Display the first few rows of the structured DataFrame
