from input_parser import MeetingParser, ParentEvent, NoticeEvent
from config_reader import PositionCalculator
import xlsx_writer as xw
from template_reader import XlsxTemplateReader
from excel_utils import add_coordinates


def _find_data_for_template_fields(field_list, data_dict_list):
    def _find_data_for_field_key(field_key):
        for data_dict in data_dict_list:
            if field_key in data_dict:
                return data_dict[field_key]
        raise AttributeError('key {} not found.'.format(field_key))

    res_dict = {}
    for field_name in field_list:
        assert field_name.startswith('{') and field_name.endswith('}')
        field_key = field_name[1:-1]
        res_dict[field_name] = _find_data_for_field_key(field_key)
    return res_dict

class ExcelAgendaEngine():
    def __init__(self,user_input_txt_file_path="/home/didi/myproject/tmma/user_input.txt",target_file_path='/home/didi/myproject/tmma/generated_calendar.xlsx'):
        output_excel_config='/home/didi/myproject/tmma/config.yaml'
        template_file = '/home/didi/myproject/tmma/tm_303_calendar.xlsx'
        template_sheet_name = 'template'
        template_position_sheet_name = 'template_position'
        target_sheet_name = 'Sheet'


        # Parse all the user input.
        self.user_input_parser = MeetingParser()
        self.user_input_parser.parse_file(user_input_txt_file_path)
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

        # Calculate the block start coord.
        block_position_calculator = PositionCalculator(output_excel_config)
        block_position_calculator.set_schedule_block_height(self.user_input_parser.get_total_event_num())
        self.block_start_coord_dict = block_position_calculator.get_start_coords()
        
        # Create template reader.
        self.template_reader = XlsxTemplateReader(template_file, template_sheet_name,
                                template_position_sheet_name)
        self.template_position_config = self.template_reader.get_template_positions()

        # Create excel writer.
        self.xlsx_writer = xw.XlsxWriter(template_file, template_sheet_name,
                                self.template_position_config, target_file_path,
                                target_sheet_name)
        
    def _write_fixed_block(self,block_name):
        self.xlsx_writer.write_sheet(
            source_template_block_name=block_name,
            target_start_coord=self.block_start_coord_dict[block_name])

    def _write_variable_block(self,block_name):
        fields_to_write = self.template_reader.get_field_list(block_name)
        block_data = _find_data_for_template_fields(
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
                                        target_cell_coord=image_coord,
                                        image_width=150)

    def write(self):
        variable_block_list = ['title_block','theme_block','project_block']
        for block_name in variable_block_list:
            self._write_variable_block(block_name)

        self._write_schedule_block()

        fixed_block_list = ['contact_block','rule_block','information_block']
        for block_name in fixed_block_list:
            self._write_fixed_block(block_name)

        self._write_images()
        self.xlsx_writer.save()

if __name__ == '__main__':
    # TODO: fix project info writing.
    import sys
    engine = ExcelAgendaEngine()
    engine.write()
    sys.exit()

    parser = MeetingParser()
    parser.parse_file("/home/didi/myproject/tmma/user_input.txt")
    event_list = parser.event_list
    role_dict = parser.role_dict
    meeting_info_dict = parser.meeting_info_dict

    cur_time = meeting_info_dict['开始时间']
    for parent_event in event_list:
        parent_event.calculate_time(cur_time)
        cur_time = parent_event.get_end_time()

    pc = PositionCalculator('/home/didi/myproject/tmma/config.yaml')
    pc.set_schedule_block_height(parser.get_total_event_num())
    start_coord_dict = pc.get_start_coords()
    # print(start_coord_dict)
    # print(pc.calculate_sizes())

    template_file = '/home/didi/myproject/tmma/tm_303_calendar.xlsx'
    template_sheet_name = 'template'
    template_position_sheet_name = 'template_position'
    target_file = '/home/didi/myproject/tmma/generated_calendar.xlsx'
    target_sheet_name = 'Sheet'
    image_path = '/home/didi/myproject/tmma/tm_logo.jpg'

    theme_block = 'theme_block'
    parent_event_template = 'parent_event'
    child_event_template = 'child_event'
    notice_event_template = 'notice_block'
    rule_block = 'rule_block'
    information_block = 'information_block'

    tr = XlsxTemplateReader(template_file, template_sheet_name,
                            template_position_sheet_name)

    template_position_config = tr.get_template_positions()

    xlsx_writer = xw.XlsxWriter(template_file, template_sheet_name,
                                template_position_config, target_file,
                                target_sheet_name)
    field_list = tr.get_field_list('title_block')
    title_block_data = _find_data_for_template_fields(
        field_list, [role_dict, meeting_info_dict])
    xlsx_writer.write_sheet(source_template_block_name='title_block',
                            target_start_coord=start_coord_dict['title_block'],
                            data=title_block_data)

    field_list = tr.get_field_list('theme_block')
    theme_block_data = _find_data_for_template_fields(
        field_list, [role_dict, meeting_info_dict])
    xlsx_writer.write_sheet(source_template_block_name='theme_block',
                            target_start_coord=start_coord_dict['theme_block'],
                            data=theme_block_data)

    cur_start_coord = start_coord_dict['schedule_block']
    for event in event_list:
        if isinstance(event, NoticeEvent):
            xlsx_writer.write_sheet(source_template_block_name='notice_block',
                                    target_start_coord=cur_start_coord,
                                    data=event.get_event())
            cur_start_coord = add_coordinates(cur_start_coord, (0, 1))
        elif isinstance(event, ParentEvent):
            xlsx_writer.write_sheet(source_template_block_name='parent_block',
                                    target_start_coord=cur_start_coord,
                                    data=event.get_parent_event())
            cur_start_coord = add_coordinates(cur_start_coord, (0, 1))
            for child_event in event.get_child_events():
                xlsx_writer.write_sheet(
                    source_template_block_name='child_block',
                    target_start_coord=cur_start_coord,
                    data=child_event)
                cur_start_coord = add_coordinates(cur_start_coord, (0, 1))
    xlsx_writer.write_sheet(
        source_template_block_name='contact_block',
        target_start_coord=start_coord_dict['contact_block'])
    # TODO: merge cells if the template size is smaller than actual block size.
    xlsx_writer.write_sheet(
        source_template_block_name='project_block',
        target_start_coord=start_coord_dict['project_block'],
        data={'{project_info}': parser.project_info})
    xlsx_writer.merge_cells(start_coord='K25', end_coord='N32')
    xlsx_writer.write_sheet(source_template_block_name='rule_block',
                            target_start_coord=start_coord_dict['rule_block'])
    xlsx_writer.write_sheet(
        source_template_block_name='information_block',
        target_start_coord=start_coord_dict['information_block'])
    # TODO: make image coord configable.
    xlsx_writer.add_image(source_cell_coord='A1',
                          target_cell_coord='A1',
                          image_width=150)
    xlsx_writer.add_image(source_cell_coord='M1',
                          target_cell_coord='M1',
                          image_width=150)
    xlsx_writer.save()
