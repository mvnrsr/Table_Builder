# Inside survey_data_processor.py
import pandas as pd
import warnings
import re
# from loop_syntax_generator import write_loop_syntax

class SurveyDataProcessor:
     def __init__(self, file_path, sheet_name):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.df = None
        self.load_data()

     def load_data(self):
          """Load data from an Excel file."""
          try:
               #   self.df = pd.read_excel(self.file_path, sheet_name=self.sheet_name)

               warnings.filterwarnings("ignore", message="Data Validation extension is not supported and will be removed")

               # Specify the column indices you want to read
               columns_to_read = 'B,C,D,G,H,M,R,Z'

               # Skip rows 3 and 4 (which are indices 2 and 3 in zero-based indexing)
               skip_rows = [2, 3]

               self.df = pd.read_excel(
                    self.file_path, 
                    sheet_name=self.sheet_name, 
                    usecols=columns_to_read, 
                    skiprows=skip_rows,
                    header=1
               )

               # Ignore specific warnings
          except Exception as e:
               print(f"Error loading data: {e}")

     def parse_var_name_absolute_pattern(self, var_name):
          """Parse variable names to separate loop levels and the actual question."""
          loop_levels = []
          actual_question = None

          if isinstance(var_name, str):
               components = var_name.split('[..].')
               actual_question = components[-1]
               for comp in components[:-1]:
                    loop_levels.append(comp)

          return loop_levels, actual_question

     def display_parsed_info(self, num_rows=5):
          """Display the contents of the 'Parsed_Info_Absolute' column."""
          if self.df is not None and 'Parsed_Info_Absolute' in self.df.columns:
               print(self.df['Parsed_Info_Absolute'].head(num_rows))
          else:
               print("Parsed_Info_Absolute column not found or DataFrame not loaded.")

     def apply_parse_var_name(self):
          """Apply the parse_var_name_absolute_pattern method to the Var_name column."""
          if self.df is not None:
               self.df['Parsed_Info_Absolute'] = self.df['Variable Name'].apply(self.parse_var_name_absolute_pattern)
          else:
               print("DataFrame is not loaded.")

     def is_loop(self, parsed_info):
          """Determine if a variable represents a loop based on the parsed information."""
          loop_levels, _ = parsed_info
          return len(loop_levels) > 0


     def update_grid(self, row):
          """Update the 'Grid' column based on the 'Is_Loop' column."""
          if row['Is_Loop']:
               return 'x'
          else:
               return row['Grid']

     def extract_loop_levels(self, row):
          """Create new columns for each loop level based on the 'Parsed_Info_Absolute' column."""
          loop_levels = row['Parsed_Info_Absolute'][0]
          for i, level in enumerate(loop_levels, 1):
               row[f'Loop_Header{i}'] = level
          return row

     def apply_is_loop(self):
          """Apply the is_loop method to create a new column indicating if a variable is a loop or not."""
          if self.df is not None:
               self.df['Is_Loop'] = self.df['Parsed_Info_Absolute'].apply(self.is_loop)
          else:
               print("DataFrame is not loaded.")

     def apply_update_grid(self):
          """Apply the update_grid method to update the 'Grid' column."""
          if self.df is not None:
               self.df['Grid'] = self.df.apply(self.update_grid, axis=1)
          else:
               print("DataFrame is not loaded.")

     def apply_extract_loop_levels(self):
          """Apply the extract_loop_levels method to create new columns for each loop level."""
          if self.df is not None:
               self.df = self.df.apply(self.extract_loop_levels, axis=1)
          else:
               print("DataFrame is not loaded.")

     def determine_loop_variables(self, row):
          """Determine loop variables and slice names from the parsed information."""
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

     def apply_determine_loop_variables(self):
          self.df[['Loop_Var_Name', 'Loop_Slice_Name']] = self.df.apply(self.determine_loop_variables, axis=1)

     def get_headers(self):
          """Return the column headers of the DataFrame."""
          if self.df is not None:
               return self.df.columns
          else:
               return []

     def generate_non_loop_syntax_v2(self, row):
          """Generate syntax for non-loop questions."""
          if not row['Is_Loop']:  # Check if it's not a loop
               var_name = row['Variable Name']
               base_title = row['Base Title']
               table_title = row['Table Title']
               qn_name = row['Question No.']

               # Syntax generation logic
               fnAddTable = '\tfnAddTable(TableDoc,"'

               # Default axis syntax
               fnAddTable_axis = f'{var_name} {{..,sigma \'Total\' subtotal()}}'

               # Check if 'Mean' is 'x' and adjust the axis syntax
               if row['Mean'] == 'x':
                    fnAddTable_axis = f'{var_name} {{..,sigma \'Total\' subtotal(),mean \'Mean\' mean()}}'

               # Check if another column (e.g., 'Avg_Mention') is 'x' and adjust the axis syntax
               if row['Avg Num of Mentions'] == 'x':
                    fnAddTable_axis = f'{var_name} {{..,sigma \'Total\' subtotal(),mean \'Average Num of Mentions\' mean()}}'

               # fnAddTable_axis_subtotal = f'{var_name} {{..,sigma \'Total\' subtotal()}}'
               # fnAddTable_axis_mean = f'{var_name} {{..,sigma \'Total\' subtotal(),mean \'Mean\' mean()}}'
               # fnAddTable_axis_avg_mention = f'{var_name} {{..,sigma \'Total\' subtotal(),mean \'Average Num of Mentions\' mean()}}'

               fnAddTable_banner = '",banner,"'
               #   fnAddTable_tab_title = f'({qn_name}) {table_title}'

               fnAddTable_tab_title = f"({var_name if pd.isna(qn_name) else qn_name}) {table_title}"
               fnAddTable_base = f'","{base_title}")'

               # Concatenate all components to form the complete 'fnAddTable' line
               fnAddTable_final = fnAddTable + fnAddTable_axis + fnAddTable_banner + fnAddTable_tab_title + fnAddTable_base

               # Add a new rule to the last item in the table
               additional_line = '\t.Item[.count-1].Rules.Addnew(0,0)'

               return f"{fnAddTable_final}\n'{additional_line}"

          return None  # Return None if it's a loop

     def apply_generate_non_loop_syntax(self):
        """Apply the generate_non_loop_syntax_v2 method to the DataFrame."""
        if self.df is not None:
            self.df['Non_Loop_Syntax'] = self.df.apply(self.generate_non_loop_syntax_v2, axis=1)
        else:
            print("DataFrame is not loaded.")

     def write_non_loop_syntax_to_file(self, filename='non_loop_syntax_v2.txt'):
          """Write the non-loop syntax to a text file."""
          if self.df is not None:
               non_loop_df = self.df[self.df['Non_Loop_Syntax'].notna()]
               with open(filename, 'w') as f:
                    for syntax in non_loop_df['Non_Loop_Syntax']:
                         f.write(f"{syntax}\n\n")
          else:
               print("DataFrame is not loaded.")

     def log_base_title_warnings(self):
          """Log warnings for null values in 'Base Title'."""
          with open('base_title_warnings.log', 'w') as log_file:
               for index, row in self.df.iterrows():
                    if pd.isnull(row['Base Title']):
                         log_file.write(f"Base Title Warning: Table Spec - Row {index + 5} - {row['Variable Name']} has no base title.\n")

     def log_data_warnings(self):
          """Log warnings for specific data issues."""
          with open('data_warnings.log', 'w') as log_file:
               for index, row in self.df.iterrows():
                    # Check Base Title if null
                    if pd.isnull(row['Base Title']):
                         log_file.write(f"Warning: Base Title is Empty. Table Spec - Row {index + 5} - {row['Variable Name']}\n")

                    # Check Table Title if null
                    if pd.isnull(row['Table Title']):
                         log_file.write(f"Warning: Table Title is Empty. Table Spec - Row {index + 5} - {row['Variable Name']}\n")

                    table_title = str(row['Table Title']) if pd.notnull(row['Table Title']) else ''
                    # Check for HTML content or &apos; entity
                    if self.contains_html_or_entity(table_title):
                         log_file.write(f"Warning: Table Title HTML Tags or &apos; entity: Row {index + 5} - {row['Variable Name']}\n")

                    # Check if both 'Avg Num of Mentions' and 'Mean' have values
                    if pd.notnull(row['Avg Num of Mentions']) and pd.notnull(row['Mean']):
                         log_file.write(f"Warning: Redundant Element (Mean and Avg Mention) - Table Spec - Row {index + 5} - {row['Variable Name']}\n")

     def contains_html_or_entity(self, text):
          """Check if a string contains HTML tags or specific HTML entities."""
          html_tag_detected = bool(re.search(r'<[^>]+>', str(text)))  # Check for HTML tags
          html_entity_detected = "&apos;" in text  # Check for specific HTML entity
          return html_tag_detected or html_entity_detected

     def generate_loop_syntax_fixed_v2(self, row):
          if row['Is_Loop']:  # Check if it's not a loop
               var_name = row['Variable Name']
               qn_no = row['Question No.']
               table_title = row['Table Title']
               base_title = row['Base Title']
               loop_header1 = row.get('Loop_Header1', None)
               # loop_header2 = row.get('Loop_Header2', None)
               loop_slice_name = row['Loop_Slice_Name']

               # Initialize loop_syntax string
               loop_syntax = f"'--- {var_name}\n"

               # Single-level loop
               # if pd.notna(loop_header1) and pd.isna(loop_header2):
               if pd.notna(loop_header1):
                    loop_syntax += f"fnAddGrid(TableDoc,\"{loop_header1}[..].{loop_slice_name}\",\"{loop_header1} '{table_title}'\",\"({loop_header1 if pd.isna(qn_no) else qn_no}) {table_title} - Summary\",\"{base_title}\")\n"
                    loop_syntax += f"For each i in MDM.Fields[\"{loop_header1}\"].categories\n"
                    loop_syntax += f"\tfnAddTable(TableDoc,\"{loop_header1}[{{\" + i.name + \"}}].{loop_slice_name}\",banner,\"({loop_header1 if pd.isna(qn_no) else qn_no}) {table_title} - \" + i.label,\"{base_title}\")\n"
                    loop_syntax += f"\t''.Item[.count-1].Rules.Addnew(0,0)\n"
                    loop_syntax += "next\n"

               # # Two-level loop
               # elif pd.notna(loop_header1) and pd.notna(loop_header2):
               #      loop_syntax += f"For each i in MDM.Fields[\"{loop_header1}\"].categories\n"
               #      #summary tables
               #      loop_syntax += f"\tfnAddGrid(TableDoc,\""
               #      loop_syntax += f"{loop_header1}[{{\" + i.name + \"}}].{loop_header2}[..].{loop_slice_name} {{..,sigma \'Total\' subtotal()}}\","
               #      #banner
               #      loop_syntax += f"\"{loop_header1}[{{\" + i.name + \"}}].{loop_header2} '{table_title}'\","
               #      #table title, base title
               #      loop_syntax += f"\"({qn_no}) {table_title} - Summary - \" + i.label + \"\",\"{base_title}\")\n"
               #      loop_syntax += f"\t.Item[.count-1].Rules.Addnew(0,1) 'supress blank column\n"
               #      loop_syntax += f"\t.Item[.count-1].Rules.Addnew(0,0) 'supress blank row \n"
               #      # loop_syntax += f"\t"
                    
               #      #individual tables
               #      loop_syntax += f"\t'For each j in MDM.Fields[\"{loop_header1}[{{\" + i.name + \"}}].{loop_header2}\"].categories\n"
               #      # loop_syntax += f"\t"
               #      loop_syntax += f"\t\t'fnAddTable(TableDoc,\"{loop_header1}[{{\" + i.name + \"}}].{loop_header2}[{{\" + j.name + \"}}].{loop_slice_name} {{..,sigma \'Total\' subtotal()}}\",banner,\"({qn_no}) {table_title} - \" + i.label + \" - \" + j.label,\"{base_title}\")\n"
               #      loop_syntax += f"\t\t'.Item[.count-1].Rules.Addnew(0,0)\n"
               #      loop_syntax += "\t'next\n"
               #      loop_syntax += "next\n"

               return loop_syntax

     def apply_generate_loop_syntax(self):
          """Apply the generate_non_loop_syntax_v2 method to the DataFrame."""
          if self.df is not None:
               self.df['loop_syntax'] = self.df.apply(self.generate_loop_syntax_fixed_v2, axis=1)
          else:
               print("DataFrame is not loaded.")

     def write_loop_syntax_to_file(self, filename='loop_syntax.txt'):
          """Write the loop syntax to a text file."""
          if self.df is not None:
               loop_df = self.df[self.df['loop_syntax'].notna()]
               with open(filename, 'w') as f:
                    for syntax in loop_df['loop_syntax']:
                         f.write(f"{syntax}\n\n")
          else:
               print("DataFrame is not loaded.")

     # def apply_generate_loop_syntax(self):
     #      if self.df is not None:
     #           self.df['Loop_Syntax'] = self.df.apply(self.generate_loop_syntax_fixed_v2, axis=1)

     # def combine_syntax(self, row):
     #      non_loop = row.get('Non_Loop_Syntax', '')
     #      loop = row.get('Loop_Syntax', '')

     #      if non_loop and loop:
     #           return f"{non_loop}\n{loop}"
     #      elif non_loop:
     #           return non_loop
     #      elif loop:
     #           return loop
     #      else:
     #           return None

     def combine_syntax(self, row):
          non_loop = row['Non_Loop_Syntax'] if pd.notna(row['Non_Loop_Syntax']) else ''
          loop = row['loop_syntax'] if pd.notna(row['loop_syntax']) else ''

          if non_loop or loop:
               return f"{non_loop}\n{loop}".strip()  # Remove leading/trailing newline if one is empty
          else:
               return None

     def apply_combine_syntax(self):
          self.df['Combined_Syntax'] = self.df.apply(self.combine_syntax, axis=1)

     def write_combined_syntax_to_file(self, filename='Combined_Syntax.txt'):
          """Write the combined syntax to a text file."""
          var_df = self.df[self.df['Variable Name'].notna()]
          with open(filename, 'w') as f:
               for syntax in var_df['Combined_Syntax']:
                    if pd.notna(syntax):  # Ensure the syntax is not null
                         f.write(f"{syntax}\n\n")

     def generate_manip_syntax_readable(self, row):
          var_name = row['Variable Name']
          qn_no = row['Question No.']
          table_title = row['Table Title']
     
          # Common patterns
          sb_title_text = f"sbSetTitleText(MDM,\"{var_name}\",\"Analysis\",LOCALE,\"({var_name if pd.isna(qn_no) else qn_no}) {table_title}\")"

          sb_axis_default = f"sbSetAxisExpression(MDM,\"{var_name}\",\"{{..,sigma 'Total' subtotal()}}\")"
          sb_axis_exp_mean = f"sbSetAxisExpression(MDM,\"{var_name}\",\"{{..,sigma 'Total' subtotal(),mean 'Mean' mean(??) [decimals=2]}}\")"
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
          if row['Avg Num of Mentions'] == 'x':
               # Re-initialize the manip_syntax
               manip_syntax = f"'--- {var_name}\n"
               manip_syntax += f"\t{sb_title_text}\n"
               manip_syntax += f"\t{sb_avg_mention}\n"

          # Both elements
          if row['Mean'] == 'x' and row['Avg Num of Mentions'] == 'x':
               # Re-initialize the manip_syntax
               manip_syntax = f"'--- {var_name}\n"
               manip_syntax += f"\t{sb_title_text}\n"
               manip_syntax += f"\t{sb_axis_exp_mean}\n"
               manip_syntax += f"\t{sb_avg_mention}\n"
               
          return manip_syntax

     def apply_generate_manip_syntax_readable(self):
          self.df['Manip_Syntax_Readable'] = self.df.apply(self.generate_manip_syntax_readable, axis=1)

     def write_manip_syntax_to_file(self, filename='Manip_Syntax.txt'):
          """Write the Manip_Syntax to a text file."""
          var_df = self.df[self.df['Variable Name'].notna()]
          with open(filename, 'w') as f:
               for syntax in var_df['Manip_Syntax_Readable']:
                    if pd.notna(syntax):  # Ensure the syntax is not null
                         f.write(f"{syntax}\n\n")