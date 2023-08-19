from openpyxl import load_workbook

from excel_utils import coordinate_string_to_index,add_coordinates,get_left_top_coordinate,get_sheet_dimensions,get_right_bottom_coordinate

class XlsxTemplateReader():
    def __init__(self,template_xlsx_file_path,template_sheet_name,position_sheet_name):
        self.template_workbook = load_workbook(template_xlsx_file_path)
        self.template_sheet = self.template_workbook[template_sheet_name]
        self.template_position_sheet = self.template_workbook[position_sheet_name]
        self.template_block_position_dict = {}

    def __exit__(self):
        self.template_workbook.close()

    @staticmethod
    def _extract_field(input_string):
        start = input_string.find('{')
        end = input_string.find('}')
        
        if start != -1 and end != -1 and start < end:
            extracted_text = input_string[start : end + 1]
            return extracted_text
        else:
            return None

    def get_field_list(self,template_block_name):
        position = self.template_block_position_dict[template_block_name]
        if position is None:
            start_coord = get_left_top_coordinate(self.template_sheet)
            end_coord = get_right_bottom_coordinate(self.template_sheet)
        else:
            start_coord = position['start_coord']
            end_coord = position['end_coord']

        start_col, start_row = coordinate_string_to_index(start_coord)
        end_col, end_row = coordinate_string_to_index(end_coord)
        field_list = []
        for row_num in range(start_row, end_row + 1):
            for col_num in range(start_col, end_col + 1):
                cell = self.template_sheet.cell(row=row_num, column=col_num)
                if cell.value is not None and '{' in cell.value:
                    field_list.append(self._extract_field(cell.value))
        return field_list

    def _read_sheet_as_dict_list(self):
        headers = [cell.value for cell in self.template_position_sheet[1]]

        column_as_key_list = []
        for row in self.template_position_sheet.iter_rows(min_row=2, values_only=True):
            row_data = {}
            for header, cell_value in zip(headers, row):
                row_data[header] = cell_value
            column_as_key_list.append(row_data)
        return column_as_key_list

    def get_template_positions(self):
        column_as_key_list = self._read_sheet_as_dict_list()
        data = {}
        for item in column_as_key_list:
            data[item['block_name']] = {'start_coord': item['start_coord'], 'end_coord': item['end_coord']}
        self.template_block_position_dict = data
        return self.template_block_position_dict

if __name__=='__main__':
    template_file = '/home/didi/myproject/tmma/tm_303_calendar.xlsx'
    template_sheet_name = 'template'
    template_position_sheet_name = 'template_position'

    tr = XlsxTemplateReader(template_file,template_sheet_name,template_position_sheet_name)
    position = {'start_coord':'B2','end_coord':'O8'}
    field_list = tr.get_field_list(position)
    print(field_list)
    print(tr.get_template_positions())
