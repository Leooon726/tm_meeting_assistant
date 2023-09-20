import re
from datetime import datetime, timedelta
import json

from agenda_event import ParentEvent, NoticeEvent


class InputJsonParser:
    def __init__(self):
        self.role_dict = {}
        self.meeting_info_dict = {}
        self.event_list = []
        self.project_info = {}

    def parse_file(self, json_path):
        with open(json_path, "r") as file:
            json_data = file.read()
        # Parse the .json data into a dictionary
        data_dict = json.loads(json_data)
        self._parse_roles(data_dict['role_name_list'])
        self._parse_meeting_info(data_dict['meeting_info'])
        self._parse_agenda(data_dict['schedule_events'])

    def get_total_event_num(self):
        cnt = len(self.event_list)
        for event in self.event_list:
            if isinstance(event, ParentEvent):
                cnt += len(event.get_child_events())
        return cnt

    def get_meeting_start_time(self):
        return self.meeting_info_dict['开始时间']

    def _parse_roles(self, role_name_list):
        '''
        input: like [{'role': '破冰师', 'name': '海龙（宾客）'}, {'role': '即兴主持人', 'name': '杨然'}]

        '''
        for role_name_dict in role_name_list:
            role = role_name_dict['role']
            name = role_name_dict['name']
            self.role_dict[role] = name

    def _parse_meeting_info(self,meeting_info_list):
        for info in meeting_info_list:
            field_name = info['field_name']
            content = info['content']
            self.meeting_info_dict[field_name] = content

    @staticmethod
    def _is_parent_event(event_dict):
        if 'is_parent' in event_dict and event_dict['is_parent'] is True:
            return True
        return False

    @staticmethod
    def _create_parent_event(event_dict):
        parent_event = ParentEvent(event_dict['event_name'])
        if 'duration' in event_dict:
            parent_event.set_parent_event_duration(event_dict['duration'])
        return parent_event

    @staticmethod
    def _is_notice_event(preprocessed_event):
        return preprocessed_event.parent_event['duration'] is None and (not preprocessed_event.has_child_event())

    def _get_person_name_from_role(self, role):
        if role not in self.role_dict:
            return role
        return self.role_dict[role]

    def _parse_agenda(self,event_list):
        preprocessed_event_list = []
        for event_dict in event_list:
            if self._is_parent_event(event_dict):
                event = self._create_parent_event(event_dict)
                preprocessed_event_list.append(event)
            else:
                last_parent_event = preprocessed_event_list[-1]
                person_name = self._get_person_name_from_role(event_dict['role'])
                last_parent_event.add_child_event(event_dict['event_name'],event_dict['duration'],person_name)

        # Postprocessing: if a parent event doesn't have child event and duration field, change it into notice event.
        for preprocessed_event in preprocessed_event_list:
            event = preprocessed_event
            if self._is_notice_event(preprocessed_event):
                event = NoticeEvent(preprocessed_event.parent_event['event_name'])
            self.event_list.append(event)

if __name__ == '__main__':
    # tests.
    parser = InputJsonParser()
    parser.parse_file("/home/lighthouse/tm_meeting_assistant/example/output_files/user_input.json")

    role_dict = parser.role_dict
    event_list = parser.event_list
    meeting_info_dict = parser.meeting_info_dict
    print(meeting_info_dict)
    print(parser.get_total_event_num())

    print("Role Dictionary:")
    print(role_dict)
    print("\nAgenda List:")
    cur_time = '15:00'
    for event in event_list:
        event.calculate_time(cur_time)
        cur_time = event.get_end_time()
        if isinstance(event, ParentEvent):
            print('111 ', event.get_parent_event())
            print('2222', event.get_child_events())
        elif isinstance(event, NoticeEvent):
            print('notice: ', event.get_event())
