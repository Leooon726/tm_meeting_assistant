from openpyxl.cell import Cell
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Border, PatternFill
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook, load_workbook
from copy import copy, deepcopy
from openpyxl.utils.cell import coordinate_from_string, get_column_letter,column_index_from_string,range_boundaries
from openpyxl.drawing.image import Image

from excel_utils import coordinate_string_to_index,add_coordinates,get_left_top_coordinate,get_sheet_dimensions,get_right_bottom_coordinate


class XlsxWriter():
    def __init__(self,template_xlsx_file_path,target_xlsx_file_path,target_sheet_name):
        self.template_workbook = load_workbook(template_xlsx_file_path)
        self.target_workbook = Workbook()  # Create a new workbook if the target doesn't exist
        self.target_xlsx_file_path = target_xlsx_file_path
        # Create or get the sheet with the desired sheet name
        if target_sheet_name in self.target_workbook.sheetnames:
            self.target_sheet = self.target_workbook[target_sheet_name]
        else:
            self.target_sheet = self.target_workbook.create_sheet(target_sheet_name)
    
    def save(self):
        self.target_workbook.save(self.target_xlsx_file_path)  # 保存
        self.target_workbook.close()  # 关闭文件
        self.template_workbook.close()

    def write_sheet(self,source_sheet_name,source_start_coord=None,source_end_coord=None,target_start_coord='A1',data=None):
        src_file_sheet = self._copy_sheet_impl(self.template_workbook[source_sheet_name],source_start_coord,source_end_coord)
        if data is not None:
            src_file_sheet = self._modify_sheet(src_file_sheet,data)
        self._paste_sheet_impl(self.target_sheet,target_start_coord,src_file_sheet,source_start_coord,source_end_coord)
        
    def add_image(self,image_path, target_cell_coord,image_width=100,image_height=100):
        # Load the image
        image = Image(image_path)

        # col_idx,row_idx = coordinate_string_to_index(targesource_sheet_name,source_start_coordt_cell_coord)
        # col_letter = get_column_letter(col_idx)
        # cell_width = sheet.column_dimensions[col_letter].width
        # cell_height = sheet.row_dimensions[row_idx].height
        image.width = image_width
        image.height = image_height

        # Add the image to the worksheet
        self.target_sheet.add_image(image, target_cell_coord)
    
    def _shift_coordinates(self,coord, offset):
        coord_col_idx,coord_row_idx = coordinate_string_to_index(coord)
        result_tuple = (coord_col_idx + offset[0], coord_row_idx + offset[1])  # Add the corresponding row and column values
        result_coord = get_column_letter(result_tuple[0]) + str(result_tuple[1])  # Convert the result back to a coordinate string
        return result_coord
    
    def _shift_range(self,range_string,offset):
        min_col, min_row, max_col, max_row = range_boundaries(range_string)
        result_tuple1 = (min_col + offset[0], min_row + offset[1])  # Add the corresponding row and column values
        result_coord1 = get_column_letter(result_tuple1[0]) + str(result_tuple1[1])  # Convert the result back to a coordinate string
        result_tuple2 = (max_col + offset[0], max_row + offset[1])  # Add the corresponding row and column values
        result_coord2 = get_column_letter(result_tuple2[0]) + str(result_tuple2[1])  # Convert the result back to a coordinate string
        return '{}:{}'.format(result_coord1,result_coord2)
    
    def _is_cell_in_block(self,src_start_coord,src_end_coord,coord_to_check):
        min_col_idx,min_row_idx = coordinate_string_to_index(src_start_coord)
        max_col_idx,max_row_idx = coordinate_string_to_index(src_end_coord)
        cell_col_idx,cell_row_idx = coordinate_string_to_index(coord_to_check)
        return min_col_idx<=cell_col_idx<=max_col_idx and min_row_idx<=cell_row_idx<=max_row_idx

    def _subtract_coordinates(self,coord1, coord2):
        coord1_col_idx,coord1_row_idx = coordinate_string_to_index(coord1)
        coord2_col_idx,coord2_row_idx = coordinate_string_to_index(coord2)
        return (coord1_col_idx-coord2_col_idx,coord1_row_idx-coord2_row_idx)

    def _copy_sheet_impl(self,original_sheet,start_coord=None,end_coord=None):
        if start_coord is None:
            start_coord = get_left_top_coordinate(original_sheet)
        if end_coord is None:
            end_coord = get_right_bottom_coordinate(original_sheet)

        new_workbook = Workbook()
        new_workbook.remove(new_workbook.active)  # Remove the default sheet
    
        new_sheet = new_workbook.create_sheet(original_sheet.title)
        start_col, start_row = coordinate_string_to_index(start_coord)
        end_col, end_row = coordinate_string_to_index(end_coord)
        for row in original_sheet.iter_rows(min_row=start_row, min_col=start_col, max_row=end_row, max_col=end_col):
            # 遍历源xlsx文件制定sheet中的所有单元格
            for cell in row:  # 复制数据
                target_cell_coord = cell.coordinate
                new_sheet[target_cell_coord].value = cell.value
                if cell.has_style:  # 复制样式
                    new_sheet[target_cell_coord].font = copy(cell.font)
                    new_sheet[target_cell_coord].border = copy(cell.border)
                    new_sheet[target_cell_coord].fill = copy(cell.fill)
                    new_sheet[target_cell_coord].number_format = copy(
                        cell.number_format
                    )
                    new_sheet[target_cell_coord].protection = copy(cell.protection)
                    new_sheet[target_cell_coord].alignment = copy(cell.alignment)
    
        wm = list(zip(original_sheet.merged_cells))  # 开始处理合并单元格
        if len(wm) > 0:  # 检测源xlsx中合并的单元格
            for i in range(0, len(wm)):
                cell2 = (
                    str(wm[i]).replace("(<MergedCellRange ", "").replace(">,)", "")
                )  # 获取合并单元格的范围
                shifted_cell2 = cell2
                new_sheet.merge_cells(shifted_cell2)  # 合并单元格
        # 开始处理行高列宽
        for i in range(1, original_sheet.max_row + 1):
            new_sheet.row_dimensions[i].height = original_sheet.row_dimensions[
                i
            ].height
        for i in range(1, original_sheet.max_column + 1):
            new_sheet.column_dimensions[
                get_column_letter(i)
            ].width = original_sheet.column_dimensions[get_column_letter(i)].width

        for img in new_sheet._images:
            img_copy = Image(img._data)
            new_sheet.add_image(img_copy, img.anchor)
    
        return new_sheet

    def _paste_sheet_impl(self,target_file_sheet,target_start_coord,src_file_sheet,src_start_coord=None,src_end_coord=None):
        if src_start_coord is None:
            src_start_coord = get_left_top_coordinate(src_file_sheet)
        if src_end_coord is None:
            src_end_coord = get_right_bottom_coordinate(src_file_sheet)

        shift_offset = self._subtract_coordinates(target_start_coord,src_start_coord)
        start_col, start_row = coordinate_string_to_index(src_start_coord)
        end_col, end_row = coordinate_string_to_index(src_end_coord)
        for row in src_file_sheet.iter_rows(min_row=start_row, min_col=start_col, max_row=end_row, max_col=end_col):
            # 遍历源xlsx文件制定sheet中的所有单元格
            for cell in row:  # 复制数据
                target_cell_coord = self._shift_coordinates(cell.coordinate,shift_offset)
                target_file_sheet[target_cell_coord].value = cell.value
                if cell.has_style:  # 复制样式
                    target_file_sheet[target_cell_coord].font = copy(cell.font)
                    target_file_sheet[target_cell_coord].border = copy(cell.border)
                    target_file_sheet[target_cell_coord].fill = copy(cell.fill)
                    target_file_sheet[target_cell_coord].number_format = copy(
                        cell.number_format
                    )
                    target_file_sheet[target_cell_coord].protection = copy(cell.protection)
                    target_file_sheet[target_cell_coord].alignment = copy(cell.alignment)
    
        wm = list(zip(src_file_sheet.merged_cells))  # 开始处理合并单元格
        if len(wm) > 0:  # 检测源xlsx中合并的单元格
            for i in range(0, len(wm)):
                cell_range = (
                    str(wm[i]).replace("(<MergedCellRange ", "").replace(">,)", "")
                )  # 获取合并单元格的范围
                begin_coord,end_coord = cell_range.split(':')
                if not self._is_cell_in_block(src_start_coord,src_end_coord,begin_coord) and not self._is_cell_in_block(src_start_coord,src_end_coord,end_coord):
                    continue
                shifted_cell_range = self._shift_range(cell_range,shift_offset)
                target_file_sheet.merge_cells(shifted_cell_range)  # 合并单元格
        # 开始处理行高列宽
        col_shift_offset, row_shift_offset = shift_offset
        for i in range(1, src_file_sheet.max_row + 1):
            target_file_sheet.row_dimensions[i+row_shift_offset].height = src_file_sheet.row_dimensions[
                i
            ].height
        for i in range(1, src_file_sheet.max_column + 1):
            target_file_sheet.column_dimensions[
                get_column_letter(i+col_shift_offset)
            ].width = src_file_sheet.column_dimensions[get_column_letter(i)].width

    def _modify_sheet(self, sheet, data):
        max_row = sheet.max_row
        max_column = sheet.max_column
    
        for key, value in data.items():
            for row_num in range(1, max_row + 1):
                for col_num in range(1, max_column + 1):
                    cell = sheet.cell(row=row_num, column=col_num)
                    cell_value = cell.value
                    if cell_value == key:
                        cell.value = value
        return sheet


if __name__=='__main__':
    # tests
    # res = add_coordinates('B12',(0,1))

    template_file = '/home/didi/myproject/tmma/tm_303_calendar.xlsx'
    target_file = '/home/didi/myproject/tmma/generated_calendar.xlsx'
    target_sheet_name = 'Sheet'
    image_path = '/home/didi/myproject/tmma/tm_logo.jpg'

    title_block = 'title_block'
    theme_block = 'theme_block'
    parent_event = 'parent_event'
    child_event = 'child_event'
    notice_block = 'notice_block'
    rule_block = 'rule_block'
    information_block = 'information_block'

    template_block_config = {'title_block': {'start_coord':'B2','end_coord':'O8'},'theme_block':{'start_coord':'B9','end_coord':'K11'}}

    xlsx_writer = XlsxWriter(template_file,target_file,target_sheet_name)
    xlsx_writer.write_sheet('calendar',source_start_coord=template_block_config[title_block]['start_coord'],source_end_coord=template_block_config[title_block]['end_coord'],target_start_coord='B2')
    xlsx_writer.write_sheet('calendar',source_start_coord=template_block_config[theme_block]['start_coord'],source_end_coord=template_block_config[theme_block]['end_coord'],target_start_coord='B9',data={'{theme_name}':'主题：父亲节','{time}':'15:00','{organizer_name}':'林长虹','{SAA_name}':'Leon Lin'})
    # xlsx_writer.write_sheet('theme_block',target_start_coord='B9',data={'{theme_name}':'主题：父亲节','{time}':'15:00','{organizer_name}':'林长虹','{SAA_name}':'Leon Lin'})
    # cur_start_coord = 'B12'
    # for i in range(3):
    #     cur_start_coord = add_coordinates(cur_start_coord,(0,1))
    #     xlsx_writer.write_sheet(parent_event,target_start_coord=cur_start_coord,data={'{start_time}':'12:01','{event_name}':'开场','{duration}':'8'})
    #     cur_start_coord = add_coordinates(cur_start_coord,(0,1))
    #     xlsx_writer.write_sheet(child_event,target_start_coord=cur_start_coord,data={'{event_name}':'开场白','{end_time}':'15:01','{duration}':'8','{host_name}':'林长宏'})
    # xlsx_writer.write_sheet('rule_block',target_start_coord='L9')
    # xlsx_writer.write_sheet('information_block',target_start_coord='L16')
    xlsx_writer.add_image(image_path, target_cell_coord='B2')
    xlsx_writer.save()