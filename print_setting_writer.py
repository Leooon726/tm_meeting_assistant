import openpyxl
from openpyxl import load_workbook
from openpyxl.worksheet.page import PrintPageSetup
from openpyxl.worksheet.page import PageMargins
from openpyxl.styles import Font

def _cm_to_inch(cm_value):
    return cm_value / 2.54

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
        left=_cm_to_inch(0.64),
        right=_cm_to_inch(0.64),
        top=_cm_to_inch(0.91),
        bottom=_cm_to_inch(0.91)
    )

    # Assign the margins to the worksheet
    worksheet.page_margins = margins

    # page_setup = worksheet.page_setup
    # page_setup.horizontalDpi = 300  # Set the desired horizontal DPI (e.g., 300 DPI)
    # page_setup.verticalDpi = 300    # Set the desired vertical DPI (e.g., 300 DPI)

    # Save the changes
    workbook.save(excel_file_path)


if __name__ == '__main__':
    set_print_setting('/home/lighthouse/output_files/adhoc/rise_test.xlsx','第一页','/home/lighthouse/output_files/adhoc/rise_print_test.xlsx')
    set_print_setting('/home/lighthouse/output_files/adhoc/rise_test.xlsx','第二页','/home/lighthouse/output_files/adhoc/rise_print_test.xlsx')
