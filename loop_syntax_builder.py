import pandas as pd

def generate_loop_syntax_fixed_v2(row):
    # ... your function code ...
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