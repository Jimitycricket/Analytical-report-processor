import os
import xlrd
from xlwt import easyxf, Workbook, Font, XFStyle, Alignment
from xlutils.copy import copy
import datetime
import shutil


# Edit these file paths to your local settings
masterlist= r'C:/Users/17361/OneDrive - Region of Peel/Udrive/Desktop/Automation/masterlist.xls'
source_folder= r'C:/Users/17361/OneDrive - Region of Peel/Udrive/Desktop/Automation/Unprocessed'
destination_folder = r'C:/Users/17361/OneDrive - Region of Peel/Udrive/Desktop/Automation/Processed'

date_style = XFStyle()
date_style.num_format_str = 'yyyy/mm/dd'
date_style.alignment = Alignment()
date_style.alignment.horz = Alignment.HORZ_CENTER
style = easyxf('borders: bottom thick;')
bold_style = easyxf('font: bold 1')

def find_master_data(masterlist, lwm_number):
    workbook = xlrd.open_workbook(masterlist, lwm_number)
    sheet = workbook.sheet_by_index(0)
    
    
    for row_index in range(sheet.nrows):
        # Assuming LWM # is in the first column
        if sheet.cell(row_index, 0).value == lwm_number:
            # Extract necessary data (adjust columns as needed)
            #Columns in masterlist
            contact_name = sheet.cell(row_index, 7).value
            company_name = sheet.cell(row_index, 1).value
            billing_address = sheet.cell(row_index, 2).value
            city_postal_code = sheet.cell(row_index, 3).value
            inspector_name = sheet.cell(row_index, 5).value
            inspector_initials = sheet.cell(row_index, 6).value
            site_address= sheet.cell(row_index, 4).value
            inspector_email= sheet.cell(row_index, 8).value
            return contact_name, company_name, billing_address, city_postal_code, inspector_name, inspector_initials, site_address, inspector_email
    return None, None, None, None, None, None, None  # Return None if LWM # is not found






def read_rewrite_workbook(source_folder, destination_folder):
    for filename in os.listdir(source_folder):
        if filename.lower().endswith('.xls'):
            source_file_path = os.path.join(source_folder, filename)
            workbook = xlrd.open_workbook(source_file_path, formatting_info=True)
            read_sheet = workbook.sheet_by_index(0)  # Assuming you want the first sheet
            writable_workbook = copy(workbook)
            sheet = writable_workbook.get_sheet(0)
            
            # Iterate through each row starting from row 12 (index 11)
            for row_index in range(11, read_sheet.nrows):
                cell_value = read_sheet.cell_value(row_index, 0)  # Column A values
                
                # Check if the cell is empty
                if not cell_value:  # This checks for an empty string which indicates an empty cell
                    reference_point_for_comments = row_index
                    # You can perform further actions here based on finding the empty cell
                    break   # Exit the loop if you only care about the first empty cell found
                # If you need to perform actions for every empty cell, remove the break statement
                
            # Initialize a variable to track if the previous cell was empty
            previous_cell_empty = False
            
            # Iterate through each row starting from row 12 (index 11)
            for row_index in range(11, read_sheet.nrows):
                cell_value = read_sheet.cell_value(row_index, 2)  # Column C values
                
                # Check if the current cell is empty
                if not cell_value:  # This checks for an empty string which indicates an empty cell
                    if previous_cell_empty:  # Check if the previous cell was also empty
                        deletable_rows_for_A_start = row_index - 1  # Mark the start row for deletion
                        break  # Exit the loop after finding two consecutive empty cells
                    previous_cell_empty = True  # Mark this cell as empty for the next iteration
                else:
                    previous_cell_empty = False  # Reset if the current cell is not empty
                
                # Reset the flag if no consecutive empty cells are found in this iteration
                if row_index == read_sheet.nrows - 1 and not deletable_rows_for_A_start:
                    print(f"No consecutive empty cells found in Column C starting from row 12 in {filename}")



            #Writing empty cells for useless data
            #Writing empty cells for cells in column D from row 15 to deletable_rows_for_A_start
            rdl_empty_start = 13
            rdl_empty_end = deletable_rows_for_A_start
            for row_index in range(rdl_empty_start, rdl_empty_end):  # Adjust for zero-based indexing
                sheet.write(row_index, 3, '')  # Write an empty string to column D (index 3)
            
            style_with_border = easyxf('borders: bottom thin')
                #styling previous RDL column with border
            for row_index in range(rdl_empty_start,rdl_empty_end):  # Column D
                sheet.write(row_index, 3, '', style_with_border)
                
                #Deleting cells with no data, below the parameter table
                for i in range(8):  # Loop from 0 to 7
                    row = deletable_rows_for_A_start + i
                    # Default text is an empty string
                    

                    # Write to the sheet for each column
                    sheet.write(row, 0, '')
                    sheet.write(row, 1, '')
                    sheet.write(row, 2, '')
                    sheet.write(row, 3, '')
                    sheet.write(row, 4, '')     
                    sheet.write(deletable_rows_for_A_start + 6, 3, 'Date:',easyxf('align: horiz right'))     
                    sheet.write(deletable_rows_for_A_start + 3, 0, 'Inspectors Comments:', easyxf('align: horiz right'))
                    sheet.write(deletable_rows_for_A_start + 6, 0, 'Reviewed by:', easyxf('align: horiz right')) 
                sheet.write(5, 3, "")

                #sheet.write(inspectors_comments, 0, "Inspectors Comments:")
                #sheet.write(reviewed_by, 0, "Reviewed By:")
            #Writing data into new sheet
    
            # Use the value from cell E7 for LWM # to match with master list            
            read_sheet = workbook.sheet_by_index(0)
            lwm_number = read_sheet.cell_value(13, 2)  # E7 (LWM #)
            date_cell = read_sheet.cell(11, 2)
            if date_cell.ctype == xlrd.XL_CELL_DATE:
                date_tuple = xlrd.xldate_as_tuple(date_cell.value, workbook.datemode)
                report_date = datetime.datetime(*date_tuple)
                report_date_str = report_date.strftime('%Y%m%d')  # Format date as you like
            else:
                report_date_str = 'Unknown-Date'
            if date_cell.ctype == xlrd.XL_CELL_DATE:
                date_value = datetime.datetime(*xlrd.xldate_as_tuple(date_cell.value, workbook.datemode))
                sheet.write(11, 2, date_value, date_style)
            
            # Define custom styles
            red_style = XFStyle()
            blue_underline_style = XFStyle()

            # Set font color to red
            red_font = Font()
            red_font.colour_index = 2  # Excel's default color index for red
            red_font.bold = True
            red_style.font = red_font

            # Set font color to blue and underline
            blue_underline_font = Font()
            blue_underline_font.colour_index = 4  # Excel's default color index for blue
            blue_underline_font.underline = True  # Apply underline
            blue_underline_font.bold = True
            blue_underline_style.font = blue_underline_font
            
            
            #Column A new values (values to index from masterlist)
            contact_name, company_name, billing_address, city_postal_code, inspector_name, inspector_initials, site_address, inspector_email = find_master_data(masterlist, lwm_number)
            # If data is found, write it to the sheet
            if contact_name:
                sheet.write(4, 0, contact_name, bold_style)
                sheet.write(5, 0, company_name, bold_style)
                sheet.write(6, 0, billing_address, bold_style)
                sheet.write(7, 0, city_postal_code, bold_style)
                sheet.write(8, 3, (f'SITE: {site_address}'), bold_style)
                sheet.write(1, 0, "REGION OF PEEL", bold_style)
                sheet.write(deletable_rows_for_A_start+7, 1, inspector_name)
                #Red_text
                email_message = 'If you have any questions or concerns regarding the report, please email ' + inspector_name + " at "
                sheet.write(deletable_rows_for_A_start+10, 0, email_message, red_style)
                #Blue + underline
                sheet.write(deletable_rows_for_A_start+11, 0, inspector_email, blue_underline_style)


            if isinstance(lwm_number, str) and lwm_number.endswith('SX'):
                # Perform actions for numbers ending with 'SX'
                lwm_number = str(lwm_number[:-3]) + ' - SX'
                sheet.write(deletable_rows_for_A_start+13, 0, 'SURCHARGE:', bold_style)
            elif isinstance(lwm_number, (float, int)):
                # Directly convert to int if it's a number; ensure it's an integer
                lwm_number_int = int(lwm_number)
                # Formatting the number with ' - MX' suffix
                lwm_number = str(lwm_number_int) + ' - MX'
                lwm_in_sheet = 'Site Location: ' + lwm_number
                sheet.write(6, 3, lwm_in_sheet, bold_style)
            else:
                # Fallback for any other unexpected types, treating it as a string
                lwm_number = str(lwm_number) + ' - MX'
                lwm_in_sheet = 'Site Location: ' + lwm_number
                sheet.write(6, 3, lwm_in_sheet, bold_style)


            #finding inspector comments row#
            for row_index in range(11, read_sheet.nrows):
                cell_value = read_sheet.cell_value(row_index, 0)  # Corrected to Column A
                if cell_value == "Inspector's Comments:":
                    inspector_comments_row = row_index
                    break  # Exit the loop once found

            if deletable_rows_for_A_start:
                sheet.write(deletable_rows_for_A_start+6, 1, inspector_name, bold_style)
                style_with_border = easyxf('borders: bottom thin')
                #styling date column with border
                for col_index in range(4, 5):  # Columns D and E
                    sheet.write(deletable_rows_for_A_start + 6, col_index, "", style_with_border)
                #styling inspector comment column w/ border
                for col_index in range(1, 5):  # Columns B to E
                    sheet.write(deletable_rows_for_A_start+ 3, col_index, "", style_with_border)
                #styling reviewed by column w/ border
                for col_index in range(1, 3):
                    sheet.write(deletable_rows_for_A_start+6, col_index, "", style_with_border)

        lwm_number_int = int(lwm_number) if isinstance(lwm_number, (float, int)) else None

        # Use lwm_number_int when constructing the filename
        if lwm_number_int is not None:
            new_file_name = f"{inspector_initials} - Report- {report_date_str} - {company_name}- LWM {lwm_number_int} - MX.xls"
            new_file_path = os.path.join(destination_folder, new_file_name)
            writable_workbook.save(new_file_path)
            # Save the Excel workbook
            writable_workbook.save(new_file_path)
            print(f"Saved {new_file_path}")




        else:
            new_file_path = os.path.join(destination_folder, f"{inspector_initials} - Report - {report_date_str} - {company_name} - LWM {lwm_number}.xls")
            writable_workbook.save(new_file_path)
            print(f"Saved {new_file_path} from {filename}")
        




# Process workbooks and convert
read_rewrite_workbook(source_folder, destination_folder)


def organize_files_by_inspector_companytype(destination_folder):
    # List all files in the source directory
    files = [f for f in os.listdir(destination_folder) if os.path.isfile(os.path.join(destination_folder, f))]
    
    for file in files:
        # Get the first two letters of the file name
        prefix = file[:2]
        
        # Create a directory for this prefix if it doesn't exist
        directory = os.path.join(destination_folder, prefix)
        if not os.path.exists(directory):
            os.makedirs(directory)
        
        # Move the file into the created directory
        shutil.move(os.path.join(destination_folder, file), os.path.join(directory, file))
        
    for prefix in os.listdir(destination_folder):
        prefix_directory = os.path.join(destination_folder, prefix)
        if os.path.isdir(prefix_directory):  # Ensure it's a directory
            prefix_files = [f for f in os.listdir(prefix_directory) if os.path.isfile(os.path.join(prefix_directory, f))]
            for file in prefix_files:
                if len(file) > 1:  # Ensure the filename is at least 2 characters long
                    suffix = file[-6:-4]
                    suffix_directory = os.path.join(prefix_directory, suffix)
                    if not os.path.exists(suffix_directory):
                        os.makedirs(suffix_directory)
                    shutil.move(os.path.join(prefix_directory, file), os.path.join(suffix_directory, file))
    print("Files have been organized.")


# Usage example (Replace 'your_directory_path' with the actual path where your files are stored)

organize_files_by_inspector_companytype(destination_folder)
