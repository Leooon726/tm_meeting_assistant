import openpyxl
import pyvips

# Load the Excel workbook and sheet
wb = openpyxl.load_workbook('/home/lighthouse/tm_meeting_assistant/tm_303_calendar.xlsx')
sheet = wb['template']  # Replace with your sheet name

# Create an image from the sheet
image = pyvips.Image.new_from_array(sheet.values)

# Save the image as a JPEG
image.jpegsave('sheet_image.jpg', Q=90)  # Adjust quality if needed