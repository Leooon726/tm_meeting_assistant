import openpyxl
from openpyxl import load_workbook
from openpyxl.worksheet.page import PrintPageSetup
from openpyxl.worksheet.page import PageMargins
from openpyxl.styles import Font

def set_print_setting(excel_file_path='/home/lighthouse/output_files/adhoc/jabil_test.xlsx',sheet_name='Sheet',print_range=None):
    # Load the workbook
    workbook = load_workbook(excel_file_path)

    # Choose the worksheet
    worksheet = workbook[sheet_name]

    # Specify the print area range
    print_area_range = print_range if print_range is not None else worksheet.dimensions

    # Set the print area
    worksheet.print_area = print_area_range
    
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.print_options.horizontalCentered = True
    worksheet.print_options.verticalCentered = True

    margins = PageMargins(
        left=0.1,   # Adjust the left margin value (in inches)
        right=0.1,  # Adjust the right margin value (in inches)
        top=0.1,    # Adjust the top margin value (in inches)
        bottom=0.1  # Adjust the bottom margin value (in inches)
    )

    # Assign the margins to the worksheet
    worksheet.page_margins = margins

    # page_setup = worksheet.page_setup
    # page_setup.horizontalDpi = 300  # Set the desired horizontal DPI (e.g., 300 DPI)
    # page_setup.verticalDpi = 300    # Set the desired vertical DPI (e.g., 300 DPI)

    # Save the changes
    workbook.save(excel_file_path)


if __name__ == '__main__':
    set_print_setting()
