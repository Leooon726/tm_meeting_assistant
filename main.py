from input_parser import MeetingParser, ParentEvent, NoticeEvent
import xlsx_writer as xw

class Block():
    def __init__(sheet_template,start_coord):
        self.sheet_template = sheet_template
        self.

    def 

if __name__=='__main__':
    parser = MeetingParser()
    parser.parse_file("/home/didi/myproject/tmma/user_input.txt")
    event_list = parser.event_list

    cur_time = '15:00'
    for parent_event in event_list:
        parent_event.calculate_time(cur_time)
        cur_time = parent_event.get_end_time()


    template_file = '/home/didi/myproject/tmma/tm_303_calendar.xlsx'
    target_file = '/home/didi/myproject/tmma/generated_calendar.xlsx'
    target_sheet_name = 'Sheet'
    image_path = '/home/didi/myproject/tmma/tm_logo.jpg'

    title_block = 'title_block'
    theme_block = 'theme_block'
    parent_event_template = 'parent_event'
    child_event_template = 'child_event'
    notice_event_template = 'notice_block'
    rule_block = 'rule_block'
    information_block = 'information_block'

    xlsx_writer = xw.XlsxWriter(template_file,target_file,target_sheet_name)
    xlsx_writer.write_sheet('title_block')
    xlsx_writer.write_sheet('theme_block',start_coord='B9',data={'{theme_name}':'主题：父亲节','{time}':'15:00','{organizer_name}':'林长虹','{SAA_name}':'Leon Lin'})
    cur_start_coord = 'B12'
    for event in event_list:
        cur_start_coord = xw._add_coordinates(cur_start_coord,(0,1))
        if isinstance(event,NoticeEvent):
            xlsx_writer.write_sheet(notice_event_template,start_coord=cur_start_coord,data=event.get_event())
        elif isinstance(event,ParentEvent):
            xlsx_writer.write_sheet(parent_event_template,start_coord=cur_start_coord,data=event.get_parent_event())
            for child_event in event.get_child_events():
                cur_start_coord = xw._add_coordinates(cur_start_coord,(0,1))
                xlsx_writer.write_sheet(child_event_template,start_coord=cur_start_coord,data=child_event)
    xlsx_writer.write_sheet('rule_block',start_coord='L9')
    xlsx_writer.write_sheet('information_block',start_coord='L16')
    xlsx_writer.add_image(image_path, target_cell_coord='B2')
    xlsx_writer.save()