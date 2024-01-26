#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import os


# Assuming you have an Excel file named 'your_file.xlsx' and the sheet is named 'Tab_Specs'
df = pd.read_excel('SYNTAX_TOOL_old.xlsx', sheet_name='Tab_Specs')

df


def parse_var_name_absolute_pattern(var_name):
    # Initialize empty list to hold loop levels and default actual_question
    loop_levels = []
    actual_question = None

    # Check if var_name is a string
    if isinstance(var_name, str):
        # Splitting by '[..].' to separate loop levels and the actual question
        components = var_name.split('[..].')

        # Extract actual question (last component)
        actual_question = components[-1]

        # Loop through components to find loop levels
        for comp in components[:-1]:
            loop_levels.append(comp)

    # Return tuple of loop levels and actual question
    return loop_levels, actual_question

# Apply the corrected function to the Var_name column
df['Parsed_Info_Absolute'] = df['Var_name'].apply(parse_var_name_absolute_pattern)


# In[ ]:


# Method to determine if a variable is a loop or not based on the parsed information
def is_loop(parsed_info):
    loop_levels, _ = parsed_info
    return True if len(loop_levels) > 0 else False

# Apply the is_loop method to create a new column indicating if a variable is a loop or not
df['Is_Loop'] = df['Parsed_Info_Absolute'].apply(is_loop)


# In[ ]:


# Method to update the 'Grid' column based on the 'Is_Loop' column
def update_grid(row):
    if row['Is_Loop']:
        return 'x'
    else:
        return row['Grid']

# Apply the update_grid method to update the 'Grid' column
df['Grid'] = df.apply(update_grid, axis=1)

# Display the updated DataFrame
# df.head()


# In[ ]:


# Create new columns for each loop level based on the 'Parsed_Info_Absolute' column
def extract_loop_levels(row):
    loop_levels = row['Parsed_Info_Absolute'][0]
    for i, level in enumerate(loop_levels, 1):
        row[f'Loop_Header{i}'] = level
    return row

df = df.apply(extract_loop_levels, axis=1)


# In[ ]:


print(df.columns)


# In[ ]:


# Filter DataFrame for loop variables
filtered_df = df[df['Is_Loop'] == True]

# Select specific columns
filtered_df = filtered_df[['Var_name', 'QN_No', 'Parsed_Info_Absolute', 'Loop_Header1', 'Loop_Header2', 'Loop_Header3','Is_Loop']]
filtered_df.head(15)


# In[ ]:





# In[ ]:


def determine_loop_variables(row):
    parsed_info = row['Parsed_Info_Absolute']
    loop_headers = parsed_info[0]
    loop_var_name = None
    loop_slice_name = parsed_info[1]

    # Determine the main loop_var_name based on the number of elements in loop_headers
    if row['Is_Loop'] == True:  # or use 'True' if it's a string
        if len(loop_headers) == 1:
            loop_var_name = loop_headers[0]
        elif len(loop_headers) >= 2:
            loop_var_name = loop_headers[-1]

    return pd.Series([loop_var_name, loop_slice_name])


# In[ ]:


# Assuming df_loop is the DataFrame containing only loop variables

# df[['Loop_Var_Name', 'Loop_Slice_Name']] = df.apply(determine_loop_variables, axis=1)
# Using apply directly on DataFrame with axis=1
df[['Loop_Var_Name', 'Loop_Slice_Name']] = df.apply(determine_loop_variables, axis=1)


# Code would be applied like this, uncomment to use
# df['Loop_Var_Name'], df['Loop_Slice_Name'] = zip(*df['Parsed_Info_Absolute'].apply(determine_loop_variables))


# In[ ]:


print(df.columns)


# In[ ]:


# Filter DataFrame for loop variables
filtered_df = df[df['Is_Loop'] == True]

# Select specific columns
filtered_df = filtered_df[['Var_name', 'QN_No', 'Parsed_Info_Absolute', 'Loop_Header1', 'Loop_Header2', 'Loop_Header3','Loop_Var_Name','Loop_Slice_Name','Is_Loop']]
filtered_df.head(30)


# In[ ]:


# Filter DataFrame for loop variables
filtered_df = df[df['Is_Loop'] == True]

# Select specific columns
filtered_df = filtered_df[['Var_name', 'QN_No', 'Parsed_Info_Absolute', 'Loop_Header1', 'Loop_Header2', 'Loop_Header3','Loop_Var_Name','Loop_Slice_Name','Is_Loop']]

# Convert the DataFrame to a CSV string for preview
csv_preview = filtered_df.head(15).to_csv(index=False)

print(csv_preview)


# In[ ]:


# Function to generate syntax for non-loop questions
def generate_non_loop_syntax(row):
    if not row['Is_Loop']:  # Check if it's not a loop
        var_name = row['Var_name']
        base_title = row['Base_title']
        table_title = row['Table_Title']
        qn_name = row['QN_No']
        # return f"'--- {var_name}\n\tfnAddTable(TableDoc,\"{var_name}\",banner,\"\",\"{base_title}\")\n\t.Item[.count-1].Rules.Addnew(0,0)"
        header = f"'--- {var_name}"
        fn_add_table = f'\tfnAddTable(TableDoc,"{var_name}",banner,"({qn_name}) {table_title}","{base_title}")'
        additional_line = '\t.Item[.count-1].Rules.Addnew(0,0)'

        return f"{header}\n{fn_add_table}\n{additional_line}"

    return None  # Return None if it's a loop

# Apply the function to generate syntax for non-loop questions
df['Non_Loop_Syntax'] = df.apply(generate_non_loop_syntax, axis=1)

# Filter and display only the non-loop rows to see the generated syntax
# df[df['Is_Loop'] == False][['Var_name', 'Non_Loop_Syntax']].head()


# In[40]:


# Function to generate syntax for non-loop questions
def generate_non_loop_syntax_v2(row):
    if not row['Is_Loop']:  # Check if it's not a loop
        var_name = row['Var_name']
        base_title = row['Base_title']
        table_title = row['Table_Title']
        qn_name = row['QN_No']

        # Initialize each component of the 'fnAddTable' line
        fnAddTable = '\tfnAddTable(TableDoc,"'
        fnAddTable_axis_subtotal = f'{var_name}{{..,sigma \'Total\' subtotal()}}'
        fnAddTable_axis_mean = f'{var_name}{{..,sigma \'Total\' subtotal(),mean \'Mean\' mean()}}'
        fnAddTable_axis_avg_mention = f'{var_name}{{..,sigma \'Total\' subtotal(),mean \'Average Num of Mentions\' mean()}}'
        fnAddTable_banner = '",banner,"'
        fnAddTable_tab_title = f'({qn_name}) {table_title}'
        fnAddTable_base = f'","{base_title}")'

        # Concatenate all components to form the complete 'fnAddTable' line
        fnAddTable_final = fn_add_table_start + var_with_subtotal + fn_add_table_middle + table_details + fn_add_table_end

        # Add a new rule to the last item in the table
        additional_line = '\t.Item[.count-1].Rules.Addnew(0,0)'


        return f"{header}\n{fnAddTable_final}\n{additional_line}"

    return None  # Return None if it's a loop

# Apply the function to generate syntax for non-loop questions
df['Non_Loop_Syntax_v2'] = df.apply(generate_non_loop_syntax_v2, axis=1)

# Filter and display only the non-loop rows to see the generated syntax
# df[df['Is_Loop'] == False][['Var_name', 'Non_Loop_Syntax']].head()


# In[ ]:


# Function to generate table syntax for loop variables
def generate_loop_syntax(row):
    if row['Is_Loop']:  # Check if it's not a loop
        var_name = row['Var_name']
        qn_no = row['QN_No']
        table_title = row['Table_Title']
        base_title = row['Base_title']
        loop_var_name = row['Loop_Var_Name']
        loop_slice_name = row['Loop_Slice_Name']
        loop_header1 = row['Loop_Header1']
        loop_header2 = row['Loop_Header2']
        is_loop = row['Is_Loop']

        syntax = f"'--- {var_name}\n"


        # For 1-level loop
        if loop_header2 is None:
            syntax += f"fnAddGrid(TableDoc,\"{loop_header1}[..].{loop_slice_name}\",\"{loop_header1}\",\"({qn_no}) {table_title}\",\"{base_title}\")\n"
            syntax += f"For each i in MDM.Fields[\"{loop_header1}\"].categories\n"
            syntax += f"\tfnAddTable(TableDoc,\"{loop_header1}[{{\"+i.name+\"}}].{loop_slice_name}\",banner,\"({qn_no}) {table_title} - \" + i.label,\"{base_title}\")\n"
            syntax += f"\t.Item[.count-1].Rules.Addnew(0,0)\n"
            syntax += f"next\n"

        # For 2-level loop
        elif loop_header2 is not None:
            syntax += f"fnAddGrid(TableDoc,\"{loop_header1}[..].{loop_header2}[..].{loop_slice_name}\",\"{loop_header1} > {loop_header1}[..].{loop_header2}\",\"({qn_no}) {table_title}\",\"{base_title}\")\n"
            syntax += f"For each i in MDM.Fields[\"{loop_header1}\"].categories\n"
            syntax += f"    For each j in MDM.Fields[\"{loop_header1}[{{\"+j.name+\"}}].{loop_header2}\"].categories\n"
            syntax += f"\tfnAddTable(TableDoc,\"{loop_header1}[{{\"+i.name+\"}}].{loop_header2}[{{\"+j.name+\"}}].{loop_slice_name}\",banner,\"({qn_no}) {table_title} - \" + i.label + \" - \" + j.label,\"{base_title}\")\n"
            syntax += f"\t.Item[.count-1].Rules.Addnew(0,0)\n"
            syntax += f"    next\n"
            syntax += f"next\n"

        return syntax

# Assuming df_loop is your DataFrame containing only the loop variables, this line applies the function.
# Uncomment to use.
df['Loop_Syntax'] = df.apply(generate_loop_syntax, axis=1)


# In[33]:


# Sample code to fix the issue with NaN loop_header being included in the syntax

def generate_loop_syntax_fixed(row):

    if row['Is_Loop']:  # Check if it's not a loop
        var_name = row['Var_name']
        qn_no = row['QN_No']
        table_title = row['Table_Title']
        base_title = row['Base_title']
        loop_header1 = row.get('Loop_Header1', None)
        loop_header2 = row.get('Loop_Header2', None)
        loop_slice_name = row['Loop_Slice_Name']

        # Initialize loop_syntax string
        loop_syntax = f"'--- {var_name}\n"

        # Single-level loop
        if pd.notna(loop_header1) and pd.isna(loop_header2):
            loop_syntax += f"fnAddGrid(TableDoc,\"{loop_header1}[..].{loop_slice_name}\",\"{loop_header1} '{table_title}'\",\"({qn_no}) {table_title}\",\"{base_title}\")\n"
            loop_syntax += f"For each i in MDM.Fields[\"{loop_header1}\"].categories\n"
            loop_syntax += f"\tfnAddTable(TableDoc,\"{loop_header1}[{{\"+i.name+\"}}].{loop_slice_name}\",banner,\"({qn_no}) {table_title} - \" + i.label,\"{base_title}\")\n"
            loop_syntax += f"\t.Item[.count-1].Rules.Addnew(0,0)\n"
            loop_syntax += "next\n"

        # Two-level loop
        elif pd.notna(loop_header1) and pd.notna(loop_header2):
            loop_syntax += f"fnAddGrid(TableDoc,\"{loop_header1}[..].{loop_header2}[..].{loop_slice_name}\",\"{loop_header1} > {loop_header1}[..].{loop_header2} '{table_title}'\",\"({qn_no}) {table_title}\",\"{base_title}\")\n"
            loop_syntax += f"For each i in MDM.Fields[\"{loop_header1}\"].categories\n"
            loop_syntax += f"\tFor each j in MDM.Fields[\"{loop_header1}[{{\"+i.name+\"}}].{loop_header2}\"].categories\n"
            loop_syntax += f"\t\tfnAddTable(TableDoc,\"{loop_header1}[{{\"+i.name+\"}}].{loop_header2}[{{\"+j.name+\"}}].{loop_slice_name}\",banner,\"({qn_no}) {table_title} - \" + i.label + \" - \" + j.label,\"{base_title}\")\n"
            loop_syntax += f"\t\t.Item[.count-1].Rules.Addnew(0,0)\n"
            loop_syntax += "\tnext\n"
            loop_syntax += "next\n"

        return loop_syntax

# Assuming df_loop is the DataFrame containing only loop variables
# Uncomment the line below to apply the function.
df['Loop_Syntax_Fixed'] = df.apply(generate_loop_syntax_fixed, axis=1)

# Sample output for the first 5 rows
# df.head()



# In[75]:


# Sample code to fix the issue with NaN loop_header being included in the syntax

def generate_loop_syntax_fixed_v2(row):

    if row['Is_Loop']:  # Check if it's not a loop
        var_name = row['Var_name']
        qn_no = row['QN_No']
        table_title = row['Table_Title']
        base_title = row['Base_title']
        loop_header1 = row.get('Loop_Header1', None)
        loop_header2 = row.get('Loop_Header2', None)
        loop_slice_name = row['Loop_Slice_Name']

        # Initialize loop_syntax string
        loop_syntax = f"'--- {var_name}\n"

        # Single-level loop
        if pd.notna(loop_header1) and pd.isna(loop_header2):
            loop_syntax += f"fnAddGrid(TableDoc,\"{loop_header1}[..].{loop_slice_name}\",\"{loop_header1} '{table_title}'\",\"({qn_no}) {table_title} - Summary\",\"{base_title}\")\n"
            loop_syntax += f"For each i in MDM.Fields[\"{loop_header1}\"].categories\n"
            loop_syntax += f"\tfnAddTable(TableDoc,\"{loop_header1}[{{\" + i.name + \"}}].{loop_slice_name}\",banner,\"({qn_no}) {table_title} - \" + i.label,\"{base_title}\")\n"
            loop_syntax += f"\t.Item[.count-1].Rules.Addnew(0,0)\n"
            loop_syntax += "next\n"

        # Two-level loop
        elif pd.notna(loop_header1) and pd.notna(loop_header2):
            loop_syntax += f"For each i in MDM.Fields[\"{loop_header1}\"].categories\n"
            #summary tables
            loop_syntax += f"\tfnAddGrid(TableDoc,\""
            loop_syntax += f"{loop_header1}[{{\" + i.name + \"}}].{loop_header2}[..].{loop_slice_name} {{..,sigma \'Total\' subtotal()}}\","
            #banner
            loop_syntax += f"\"{loop_header1}[{{\" + i.name + \"}}].{loop_header2} '{table_title}'\","
            #table title, base title
            loop_syntax += f"\"({qn_no}) {table_title} - Summary - \" + i.label + \"\",\"{base_title}\")\n"
            loop_syntax += f"\t.Item[.count-1].Rules.Addnew(0,1) 'supress blank column\n"
            loop_syntax += f"\t.Item[.count-1].Rules.Addnew(0,0) 'supress blank row \n"
            # loop_syntax += f"\t"

            #individual tables
            loop_syntax += f"\t'For each j in MDM.Fields[\"{loop_header1}[{{\" + i.name + \"}}].{loop_header2}\"].categories\n"
            # loop_syntax += f"\t"
            loop_syntax += f"\t\t'fnAddTable(TableDoc,\"{loop_header1}[{{\" + i.name + \"}}].{loop_header2}[{{\" + j.name + \"}}].{loop_slice_name} {{..,sigma \'Total\' subtotal()}}\",banner,\"({qn_no}) {table_title} - \" + i.label + \" - \" + j.label,\"{base_title}\")\n"
            loop_syntax += f"\t\t'.Item[.count-1].Rules.Addnew(0,0)\n"
            loop_syntax += "\t'next\n"
            loop_syntax += "next\n"

        return loop_syntax

# Assuming df_loop is the DataFrame containing only loop variables
# Uncomment the line below to apply the function.
df['Loop_Syntax_v2'] = df.apply(generate_loop_syntax_fixed_v2, axis=1)

# Sample output for the first 5 rows
# df.head()



# In[71]:


def generate_manip_syntax_readable(row):
    var_name = row['Var_name']
    qn_no = row['QN_No']
    table_title = row['Table_Title']

    # Common patterns
    sb_title_text = f"sbSetTitleText(MDM,\"{var_name}\",\"Analysis\",LOCALE,\"({qn_no}) {table_title}\")"
    sb_axis_default = f"sbSetAxisExpression(MDM,\"{var_name}\",\"{{..,sigma 'Total' subtotal()}}\")"
    sb_axis_exp_mean = f"sbSetAxisExpression(MDM,\"{var_name}\",\"{{..,sigma 'Total' subtotal(),mean 'Mean' mean() [decimals=2]}}\")"
    sb_avg_mention = f"sbAddAverageEndorsementsToAxis(MDM,\"{var_name}\",\"\",\"Avg. No of Mentions\",false,\"2\")"

    # Initialize the manip_syntax
    manip_syntax = f"'--- {var_name}\n"
    manip_syntax += f"\t{sb_title_text}\n"

    # No elements
    # if row['Mean'] is None and row['Avg_Mention'] is None:
    manip_syntax += f"\t{sb_axis_default}\n"

    # Mean condition
    if row['Mean'] == 'x':
        # Re-initialize the manip_syntax
        manip_syntax = f"'--- {var_name}\n"
        manip_syntax += f"\t{sb_title_text}\n"
        manip_syntax += f"\t{sb_axis_exp_mean}\n"

    # Avg_Mention condition
    if row['Avg_Mention'] == 'x':
        # Re-initialize the manip_syntax
        manip_syntax = f"'--- {var_name}\n"
        manip_syntax += f"\t{sb_title_text}\n"
        manip_syntax += f"\t{sb_avg_mention}\n"

    # Both elements
    if row['Mean'] == 'x' and row['Avg_Mention'] == 'x':
        # Re-initialize the manip_syntax
        manip_syntax = f"'--- {var_name}\n"
        manip_syntax += f"\t{sb_title_text}\n"
        manip_syntax += f"\t{sb_axis_exp_mean}\n"
        manip_syntax += f"\t{sb_avg_mention}\n"

    return manip_syntax

# Uncomment the line below to apply the function.
df['Manip_Syntax_Readable'] = df.apply(generate_manip_syntax_readable, axis=1)


# In[72]:


def generate_manip_syntax_readable(row):
    var_name = row['Var_name']
    qn_no = row['QN_No']
    table_title = row['Table_Title']

    # Common patterns
    sb_title_text = f"sbSetTitleText(MDM,\"{var_name}\",\"Analysis\",LOCALE,\"({qn_no}) {table_title}\")"
    sb_axis_default = f"sbSetAxisExpression(MDM,\"{var_name}\",\"{{..,sigma 'Total' subtotal()}}\")"
    sb_axis_exp_mean = f"sbSetAxisExpression(MDM,\"{var_name}\",\"{{..,sigma 'Total' subtotal(),mean 'Mean' mean() [decimals=2]}}\")"
    sb_avg_mention = f"sbAddAverageEndorsementsToAxis(MDM,\"{var_name}\",\"\",\"Avg. No of Mentions\",false,\"2\")"

    # Initialize the manip_syntax
    manip_syntax = f"'--- {var_name}\n"
    manip_syntax += f"\t{sb_title_text}\n"

    # No elements
    # if row['Mean'] is None and row['Avg_Mention'] is None:
    manip_syntax += f"\t{sb_axis_default}\n"

    # Mean condition
    if row['Mean'] == 'x':
        # Re-initialize the manip_syntax
        manip_syntax = f"'--- {var_name}\n"
        manip_syntax += f"\t{sb_title_text}\n"
        manip_syntax += f"\t{sb_axis_exp_mean}\n"

    # Avg_Mention condition
    if row['Avg_Mention'] == 'x':
        # Re-initialize the manip_syntax
        manip_syntax = f"'--- {var_name}\n"
        manip_syntax += f"\t{sb_title_text}\n"
        manip_syntax += f"\t{sb_avg_mention}\n"

    # Both elements
    if row['Mean'] == 'x' and row['Avg_Mention'] == 'x':
        # Re-initialize the manip_syntax
        manip_syntax = f"'--- {var_name}\n"
        manip_syntax += f"\t{sb_title_text}\n"
        manip_syntax += f"\t{sb_axis_exp_mean}\n"
        manip_syntax += f"\t{sb_avg_mention}\n"

    return manip_syntax

# Uncomment the line below to apply the function.
df['Manip_Syntax_Readable'] = df.apply(generate_manip_syntax_readable, axis=1)


# In[76]:


non_loop_v2_df = df[df['Non_Loop_Syntax_v2'].notna()]

# Write the 'Non_Loop_Syntax' column to a text file
with open('non_loop_syntax_v2.txt', 'w') as f:
    for syntax in non_loop_v2_df['Non_Loop_Syntax_v2']:
        f.write(f"{syntax}\n\n")

# Filter out rows where Non_Loop_Syntax is not None (meaning it has generated syntax)
Loop_v2_df = df[df['Loop_Syntax_v2'].notna()]

with open('Loop_Syntax_v2.txt', 'w') as f:
    for syntax in Loop_v2_df['Loop_Syntax_v2']:
        f.write(f"{syntax}\n\n")


# In[77]:


# Function to combine the syntax
def combine_syntax(row):
    non_loop = row.get('Non_Loop_Syntax', '')
    loop = row.get('Loop_Syntax_v2', '')

    if non_loop and loop:
        return f"{non_loop}\n{loop}"
    elif non_loop:
        return non_loop
    elif loop:
        return loop
    else:
        return None

# Applying the function to combine the syntax
df['Combined_Syntax'] = df.apply(combine_syntax, axis=1)
# df.head(10)


# In[36]:


# Filter out rows where Non_Loop_Syntax is not None (meaning it has generated syntax)
non_loop_df = df[df['Non_Loop_Syntax'].notna()]

# Write the 'Non_Loop_Syntax' column to a text file
with open('non_loop_syntax.txt', 'w') as f:
    for syntax in non_loop_df['Non_Loop_Syntax']:
        f.write(f"{syntax}\n\n")

# Filter out rows where Non_Loop_Syntax is not None (meaning it has generated syntax)
non_loop_manip_df = df[df['Manip_Syntax_Readable'].notna()]

# Write the 'Non_Loop_Syntax' column to a text file
with open('Manip_Syntax.txt', 'w') as f:
    for syntax in non_loop_manip_df['Manip_Syntax_Readable']:
        f.write(f"{syntax}\n\n")

# Filter out rows where Non_Loop_Syntax is not None (meaning it has generated syntax)
Loop_df = df[df['Loop_Syntax_Fixed'].notna()]

with open('Loop_Syntax_Fixed.txt', 'w') as f:
    for syntax in Loop_df['Loop_Syntax_Fixed']:
        f.write(f"{syntax}\n\n")


var_df = df[df['Var_name'].notna()]
# Write the 'Non_Loop_Syntax' column to a text file
with open('Combined_Syntax.txt', 'w') as f:
    for syntax in var_df['Combined_Syntax']:
        f.write(f"{syntax}\n\n")


# In[ ]:


# Filter DataFrame for loop variables
filtered_df = df[df['Is_Loop'] == True]

# Select specific columns
filtered_df = filtered_df[['Var_name', 'QN_No', 'Parsed_Info_Absolute', 'Is_Loop']]

# Convert the DataFrame to a CSV string for preview
csv_preview = filtered_df.head(30).to_csv(index=False)

print(csv_preview)


# In[5]:


# print(filtered_df.head(10))

