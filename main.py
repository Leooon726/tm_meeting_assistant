'''
Example:
python3 main.py -i /home/lighthouse/tm_meeting_assistant/user_input.txt -o /home/lighthouse/tm_meeting_assistant/generated_calendar.xlsx -c /home/lighthouse/tm_meeting_assistant/engine_config.yaml
'''
import argparse,os

from input_parser import InputTxtParser
from json_input_parser import InputJsonParser
from text_blocks_json_input_parser import TextBlocksJsonInputParser
from agenda_event import ParentEvent, NoticeEvent
from config_reader import ConfigReader
from block_position_csp_solver import BlockPositionCSPSolverAdaptor
from block_position_calculator import PositionCalculator
import xlsx_writer as xw
from template_reader import XlsxTemplateReader
from excel_utils import add_coordinates
from print_setting_writer import set_print_setting
from excel_border_writer import set_outer_borders


def _get_all_block_names_by_block_start_coord_dict(block_start_coord_dict):
    # image block is not mentioned in block_position_config_path, so we need to add it manually.
    return list(block_start_coord_dict.keys())


class ExcelAgendaEngine():
    '''
    This class is used to create a excel agenda with multiple sheets.
    '''
    def __init__(self,user_input_txt_file_path,target_file_path,config_path):
        self.user_input_txt_file_path = user_input_txt_file_path
        self.target_file_path = target_file_path
        # Check if the target file exists
        if os.path.exists(self.target_file_path):
            os.remove(self.target_file_path)

        cr = ConfigReader(config_path)
        self.config = cr.get_config()

    def write(self):
        for target_sheet_config_dict in self.config['target_sheets']:
            # sheet_dict = target_sheet_config_dict
            target_sheet_config_dict['template_excel'] = self.config['template_excel']
            sheet_engine = ExcelAgendaSheetEngine(self.user_input_txt_file_path,self.target_file_path,target_sheet_config_dict)
            sheet_engine.write()


class ExcelAgendaSheetEngine():
    '''
    This class is used to create a excel sheet of agenda.
    '''
    def __init__(self,user_input_file_path,target_file_path,config):
        block_position_config_path = config['block_position_config']
        template_file = config['template_excel']
        template_sheet_name = config['template_sheet_name']
        template_position_sheet_name = config['template_position_sheet_name']
        self.target_sheet_name = config['target_sheet_name']
        self.target_file_path = target_file_path

        # Parse all the user input.
        self.user_input_parser = self._select_input_parser(user_input_file_path)
        self.user_input_parser.parse_file(user_input_file_path)
        self.event_list = self.user_input_parser.event_list
        role_dict = self.user_input_parser.role_dict
        meeting_info_dict = self.user_input_parser.meeting_info_dict
        project_info = self.user_input_parser.project_info
        self.data_dict_list = [role_dict, meeting_info_dict, project_info]

        # time calcuation.
        cur_time = self.user_input_parser.get_meeting_start_time()
        for parent_event in self.event_list:
            parent_event.calculate_time(cur_time)
            cur_time = parent_event.get_end_time()

        # Create template reader.
        self.template_reader = XlsxTemplateReader(template_file, template_sheet_name,
                                template_position_sheet_name)
        self.template_position_config = self.template_reader.get_template_positions()
        template_block_size_dict = self.template_reader.get_template_block_sizes()

        # Calculate the block start coord.
        block_position_calculator = self._select_block_calculator(block_position_config_path)
        # Most of the block size to be written have the same size as template blocks, so we can get the block sizes by template_block_size_dict
        block_position_calculator.set_block_size_base_on_template_block(template_block_size_dict)
        block_position_calculator.set_schedule_block_height(self.user_input_parser.get_total_event_num())
        self.block_start_coord_dict = block_position_calculator.get_start_coords()

        # Create excel writer.
        self.xlsx_writer = xw.XlsxWriter(template_file, template_sheet_name,
                                self.template_position_config, self.target_file_path,
                                self.target_sheet_name)
        
        self.block_names_to_be_written = _get_all_block_names_by_block_start_coord_dict(self.block_start_coord_dict)

    @staticmethod
    def _select_block_calculator(file_path):
        if file_path.endswith('csp_config.yaml'):
            return BlockPositionCSPSolverAdaptor(file_path)
        else:
            return PositionCalculator(file_path)

    @staticmethod
    def _select_input_parser(user_input_file_path):
        if user_input_file_path.endswith('txt'):
          return InputTxtParser()
        elif user_input_file_path.endswith('text_blocks.json'):
          return TextBlocksJsonInputParser()
        elif user_input_file_path.endswith('json'):
          return InputJsonParser()

    @staticmethod
    def _find_data_for_template_fields(field_list, data_dict_list):
        '''
        field_list: [{'field_name': '会议标题'}, {'field_name': '会议经理'}]
        '''
        def _find_data_for_field_key(field_key):
            for data_dict in data_dict_list:
                if field_key in data_dict:
                    return data_dict[field_key]
            raise AttributeError('Cannot find data for the field to be filled {}.'.format(field_key))
    
        res_dict = {}
        for field_dict in field_list:
            # assert field_name.startswith('{') and field_name.endswith('}')
            # field_key = field_name[1:-1]
            field_name = field_dict['field_name']
            res_dict[field_name] = _find_data_for_field_key(field_name)
        return res_dict

    def _write_fixed_block(self,block_name):
        self.xlsx_writer.write_sheet(
            source_template_block_name=block_name,
            target_start_coord=self.block_start_coord_dict[block_name])

    def _write_variable_block(self,block_name):
        fields_to_write = self.template_reader.get_field_list(block_name)
        block_data = self._find_data_for_template_fields(
            fields_to_write, self.data_dict_list)
        self.xlsx_writer.write_sheet(source_template_block_name=block_name,
                                target_start_coord=self.block_start_coord_dict[block_name],
                                data=block_data)
        
    def _write_schedule_block(self):
        cur_start_coord = self.block_start_coord_dict['schedule_block']
        for event in self.event_list:
            if isinstance(event, NoticeEvent):
                self.xlsx_writer.write_sheet(source_template_block_name='notice_block',
                                        target_start_coord=cur_start_coord,
                                        data=event.get_event())
                cur_start_coord = add_coordinates(cur_start_coord, (0, 1))
            elif isinstance(event, ParentEvent):
                self.xlsx_writer.write_sheet(source_template_block_name='parent_block',
                                        target_start_coord=cur_start_coord,
                                        data=event.get_parent_event())
                cur_start_coord = add_coordinates(cur_start_coord, (0, 1))
                for child_event in event.get_child_events():
                    self.xlsx_writer.write_sheet(
                        source_template_block_name='child_block',
                        target_start_coord=cur_start_coord,
                        data=child_event)
                    cur_start_coord = add_coordinates(cur_start_coord, (0, 1))

    def _write_images(self):
        image_coords = self.template_reader.get_image_coords()
        for image_coord in image_coords:
            self.xlsx_writer.add_image(source_cell_coord=image_coord,
                                        target_cell_coord=image_coord)

    def _is_fixed_block(self,block_name):
        '''
        If there is no field to be filled, then it is a fixed block.
        '''
        return self.template_reader.is_pure_text_block(block_name)

    def write(self):
        # TODO: merge cells if the template size is smaller than actual block size.
        # xlsx_writer.merge_cells(start_coord='K25', end_coord='N32')
        for block_name in self.block_names_to_be_written:
            if block_name == 'schedule_block':
                self._write_schedule_block()
            elif block_name == 'images':
                pass
                # self._write_images()
            elif self._is_fixed_block(block_name):
                self._write_fixed_block(block_name)
            else:
                self._write_variable_block(block_name)
        self.xlsx_writer.save()
        set_outer_borders(self.target_file_path,self.target_sheet_name,self.target_file_path)
        set_print_setting(self.target_file_path,self.target_sheet_name)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Generate Agenda.')
    parser.add_argument('-i','--input', dest='user_input_file_path', type=str, required=True,
                        help='Path to the input txt file')
    parser.add_argument('-o','--output', dest='output_excel', type=str, required=True,
                        help='Path to the output excel file')
    parser.add_argument('-c','--config', dest='config_path', type=str, required=True,
                        help='Path to the configuration file')
    args = parser.parse_args()

    engine = ExcelAgendaEngine(args.user_input_file_path,args.output_excel,args.config_path)
    engine.write()
