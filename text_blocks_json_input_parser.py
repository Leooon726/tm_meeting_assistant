import json

from input_parser import InputTxtParser
from agenda_event import ParentEvent, NoticeEvent


class TextBlocksJsonInputParser(InputTxtParser):
    def __init__(self):
        super(TextBlocksJsonInputParser, self).__init__() 

    def parse_file(self, json_path):
        with open(json_path, "r") as file:
            json_data = file.read()
        # Parse the .json data into a dictionary
        data_dict = json.loads(json_data)
        self._parse_roles(data_dict['role_name_list'])
        self._parse_meeting_info(data_dict['meeting_info'])
        self._parse_agenda(data_dict['agenda_content'])

if __name__ == '__main__':
    # tests.
    parser = TextBlocksJsonInputParser()
    parser.parse_file("/home/lighthouse/tm_meeting_assistant/example/output_files/user_input_text_blocks.json")

    role_dict = parser.role_dict
    event_list = parser.event_list
    meeting_info_dict = parser.meeting_info_dict
    # print(meeting_info_dict)
    # print(parser.get_total_event_num())

    # print("Role Dictionary:")
    # print(role_dict)
    # print("\nAgenda List:")
    cur_time = '15:00'
    for event in event_list:
        event.calculate_time(cur_time)
        cur_time = event.get_end_time()
        if isinstance(event, ParentEvent):
            print('Parent: ', event.get_parent_event())
            print('Child', event.get_child_events())
        elif isinstance(event, NoticeEvent):
            print('notice: ', event.get_event())