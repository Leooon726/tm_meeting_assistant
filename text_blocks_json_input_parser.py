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

    def _parse_meeting_info(self,meeting_info):
        '''
        meeting_info: a list of dict, every dict contains keys: field_name, content
        '''
        for meeting_info_item in meeting_info:
            name = meeting_info_item['field_name']
            info = meeting_info_item['content']
            self.meeting_info_dict[name] = info

    def _parse_agenda(self,parent_event_dict_list):
        def _is_notice_block(parent_event_dict):
            return len(parent_event_dict) == 1 and ('event_name' in parent_event_dict)
            # return 'child_events' not in parent_event_dict or len(parent_event_dict['child_events'])==0

        def _create_parent_event(parent_event_dict):
            parent_event = ParentEvent(parent_event_dict['event_name'])
            if 'duration' in parent_event_dict:
                parent_event.set_parent_event_duration(
                    parent_event_dict['duration'])
            if 'child_events' in parent_event_dict:
                for child_event in parent_event_dict['child_events']:
                    child_event_name = child_event['event_name']
                    child_event_duration = child_event['duration']
                    person_name = self._get_person_name_from_role(child_event['role'])
                    parent_event.add_child_event(child_event_name, child_event_duration, person_name)
            return parent_event

        for parent_event_dict in parent_event_dict_list:
            if _is_notice_block(parent_event_dict):
                event = NoticeEvent(parent_event_dict['event_name'])
            else:
                event = _create_parent_event(parent_event_dict)
            self.event_list.append(event)
        
if __name__ == '__main__':
    # tests.
    parser = TextBlocksJsonInputParser()
    parser.parse_file("/home/lighthouse/tm_meeting_assistant/example/output_files/user_input_text_blocks.json")

    role_dict = parser.role_dict
    event_list = parser.event_list
    meeting_info_dict = parser.meeting_info_dict
    print('total_event_num: ',parser.get_total_event_num())

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
            