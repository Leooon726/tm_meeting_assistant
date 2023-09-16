import re
from datetime import datetime, timedelta
import json

class NoticeEvent():

    def __init__(self, notice):
        self.notice_event = {'time': None, 'notice': notice}

    def calculate_time(self, start_time):
        self.notice_event['time'] = start_time

    def get_end_time(self):
        return self.notice_event['time']

    def get_event(self):
        return self.notice_event


class ParentEvent():

    def __init__(self, parent_event_name):
        '''
        parent: {'start_time':'12:01','event_name':'开场','duration':'8'}
        child: {'event_name':'开场白','end_time':'15:01','duration':'8','host_name':'林长宏'}
        only event_start_time is given, other time and duration are derived.
        '''
        self.parent_event = {
            'start_time': None,
            'event_name': parent_event_name,
            'duration': None
        }
        self.child_events = []

    def set_parent_event_duration(self, duration):
        self.parent_event['duration'] = int(duration)

    def add_child_event(self, child_event_name, duration, host_name):
        '''
        Adds to self.child_events with order.
        '''
        self.calculated = False
        data = {
            'event_name': child_event_name,
            'end_time': None,
            'duration': int(duration),
            'host_name': host_name
        }
        self.child_events.append(data)

    def _convert_time_string_to_datetime(self, time_string):
        try:
            # Assuming the time_string is in the format "15:00"
            time_format = "%H:%M"
            return datetime.strptime(time_string, time_format)
        except ValueError:
            # If the time_string is not in the correct format, handle the error here
            raise ValueError("Invalid time format. Expected format: '15:00'")

    def calculate_time(self, start_time_str):
        '''
        Input: start_time is a string like 15:00
        Calcuates the parent duration, and all the child event end time.
        '''
        start_time = self._convert_time_string_to_datetime(start_time_str)
        self.parent_event['start_time'] = start_time
        cur_time = start_time
        for child_event in self.child_events:
            child_event['end_time'] = cur_time + timedelta(
                minutes=child_event['duration'])
            cur_time = child_event['end_time']

        # If parent event duration already set, no need to derive it.
        if self.parent_event['duration'] is not None:
            return
        duration_sum = 0
        for child_event in self.child_events:
            duration_sum += child_event['duration']
        self.parent_event['duration'] = duration_sum

    @staticmethod
    def _convert_datetime_values_to_string(input_dict):
        converted_dict = {}
        for key, value in input_dict.items():
            if isinstance(value, datetime):
                converted_dict[key] = value.strftime("%H:%M")
            else:
                converted_dict[key] = value
        return converted_dict

    def get_parent_event(self):
        return self._convert_datetime_values_to_string(self.parent_event)

    def get_child_events(self):
        return [
            self._convert_datetime_values_to_string(child_event)
            for child_event in self.child_events
        ]

    def has_child_event(self):
        return len(self.child_events) >0

    def get_end_time(self):
        end_time = self.parent_event['start_time'] + timedelta(
            minutes=self.parent_event['duration'])
        end_time_str = end_time.strftime("%H:%M")
        return end_time_str

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
