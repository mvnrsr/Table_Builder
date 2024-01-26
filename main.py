from survey_data_processor import SurveyDataProcessor
from FileChecker import FileChecker

def main():
    # Call the select_file static method to get the file path
    file_path = FileChecker.select_file()

    # If a file was selected, proceed with the operation
    if file_path:
        print(f"Selected file: {file_path}")
    else:
        # If the user canceled the file selection, handle the cancellation
        print("File selection was canceled.")

    # Check if a file was selected
    if file_path:
        processor = SurveyDataProcessor(file_path, 'Table Spec')

        processor.apply_parse_var_name()
        processor.apply_is_loop()
        processor.apply_update_grid()
        processor.apply_extract_loop_levels()
        processor.apply_determine_loop_variables()
        processor.apply_generate_non_loop_syntax()
        processor.apply_generate_loop_syntax()
        processor.apply_combine_syntax()
        processor.apply_generate_manip_syntax_readable()

        # write syntax to txt
        processor.write_non_loop_syntax_to_file()
        processor.write_loop_syntax_to_file()
        processor.write_combined_syntax_to_file()
        processor.write_manip_syntax_to_file()

        #write combined syntax to txt

        # processor.log_base_title_warnings()

        #write log file
        processor.log_data_warnings()

        # Test headers
        headers = processor.get_headers()
        print("Column Headers:", headers)
    else:
        print("No file selected. Exiting program.")

if __name__ == "__main__":
    main()
