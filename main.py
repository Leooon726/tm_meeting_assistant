from input_parser import MeetingParser, ParentEvent, NoticeEvent
from config_reader import PositionCalculator
import xlsx_writer as xw
from excel_utils import add_coordinates

if __name__=='__main__':
    parser = MeetingParser()
    parser.parse_file("/home/didi/myproject/tmma/user_input.txt")
    event_list = parser.event_list

    # TODO: Move cur_time to input.txt
    cur_time = '15:00'
    for parent_event in event_list:
        parent_event.calculate_time(cur_time)
        cur_time = parent_event.get_end_time()

    pc = PositionCalculator('/home/didi/myproject/tmma/config.yaml')
    pc.set_schedule_block_height(43)
    start_coord_dict = pc.get_start_coords()
    # print(start_coord_dict)
    # print(pc.calculate_sizes())

    template_file = '/home/didi/myproject/tmma/tm_303_calendar.xlsx'
    template_sheet_name = 'calendar'
    target_file = '/home/didi/myproject/tmma/generated_calendar.xlsx'
    target_sheet_name = 'Sheet'
    image_path = '/home/didi/myproject/tmma/tm_logo.jpg'

    theme_block = 'theme_block'
    parent_event_template = 'parent_event'
    child_event_template = 'child_event'
    notice_event_template = 'notice_block'
    rule_block = 'rule_block'
    information_block = 'information_block'

    # TODO: Read template position config from config.yaml.
    template_position_config = {'title_block': {'start_coord':'B2','end_coord':'O8'},'theme_block':{'start_coord':'B9','end_coord':'K11'},'parent_block':{'start_coord':'B12','end_coord':'K12'},'child_block':{'start_coord':'B13','end_coord':'K13'},'notice_block':{'start_coord':'B49','end_coord':'K49'},'rule_block':{'start_coord':'L9','end_coord':'O15'},'information_block':{'start_coord':'L33','end_coord':'O54'},'contact_block':{'start_coord':'L16','end_coord':'O24'},'project_block':{'start_coord':'L25','end_coord':'O32'}}

    xlsx_writer = xw.XlsxWriter(template_file,template_sheet_name,target_file,target_sheet_name)
    # TODO: save templae position config as private member of writer.
    xlsx_writer.write_sheet(source_position=template_position_config['title_block'],target_start_coord=start_coord_dict['title_block'])
    # TODO：Parse theme from input.txt
    xlsx_writer.write_sheet(source_position=template_position_config['theme_block'],target_start_coord=start_coord_dict['theme_block'],data={'{theme_name}':'主题：父亲节','{time}':'15:00','{organizer_name}':'林长虹','{SAA_name}':'Leon Lin'})
    cur_start_coord = start_coord_dict['schedule_block']
    for event in event_list:
        if isinstance(event,NoticeEvent):
            xlsx_writer.write_sheet(source_position=template_position_config['notice_block'],target_start_coord=cur_start_coord,data=event.get_event())
            cur_start_coord = add_coordinates(cur_start_coord,(0,1))
        elif isinstance(event,ParentEvent):
            xlsx_writer.write_sheet(source_position=template_position_config['parent_block'],target_start_coord=cur_start_coord,data=event.get_parent_event())
            cur_start_coord = add_coordinates(cur_start_coord,(0,1))
            for child_event in event.get_child_events():
                xlsx_writer.write_sheet(source_position=template_position_config['child_block'],target_start_coord=cur_start_coord,data=child_event)
                cur_start_coord = add_coordinates(cur_start_coord,(0,1))
    xlsx_writer.write_sheet(source_position=template_position_config['contact_block'],target_start_coord=start_coord_dict['contact_block'])
    xlsx_writer.write_sheet(source_position=template_position_config['project_block'],target_start_coord=start_coord_dict['project_block'])
    xlsx_writer.write_sheet(source_position=template_position_config['rule_block'],target_start_coord=start_coord_dict['rule_block'])
    xlsx_writer.write_sheet(source_position=template_position_config['information_block'],target_start_coord=start_coord_dict['information_block'])
    # TODO: make image coord configable.
    xlsx_writer.add_image(image_path, target_cell_coord='B2')
    xlsx_writer.save()