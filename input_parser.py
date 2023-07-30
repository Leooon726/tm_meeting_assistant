import re
from datetime import datetime, timedelta

class NoticeEvent():
    def __init__(self,notice):
        self.notice_event = {'{time}':None,'{notice}':notice}

    def calculate_time(self,start_time):
        self.notice_event['{time}'] = start_time

    def get_end_time(self):
        return self.notice_event['{time}']

    def get_event(self):
        return self.notice_event


class ParentEvent():
    def __init__(self,parent_event_name):
        '''
        parent: {'{start_time}':'12:01','{event_name}':'开场','{duration}':'8'}
        child: {'{event_name}':'开场白','{end_time}':'15:01','{duration}':'8','{host_name}':'林长宏'}
        only event_start_time is given, other time and duration are derived.
        '''
        self.parent_event = {'{start_time}':None,'{event_name}':parent_event_name,'{duration}':None}
        self.child_events = []

    def set_parent_event_duration(self,duration):
        self.parent_event['{duration}'] = int(duration)

    def add_child_event(self,child_event_name,duration,host_name):
        '''
        Adds to self.child_events with order.
        '''
        self.calculated = False
        data = {'{event_name}':child_event_name,'{end_time}':None,'{duration}':int(duration),'{host_name}':host_name}
        self.child_events.append(data)

    def _convert_time_string_to_datetime(self,time_string):
        try:
            # Assuming the time_string is in the format "15:00"
            time_format = "%H:%M"
            return datetime.strptime(time_string, time_format)
        except ValueError:
            # If the time_string is not in the correct format, handle the error here
            raise ValueError("Invalid time format. Expected format: '15:00'")

    def calculate_time(self,start_time_str):
        '''
        Input: start_time is a string like 15:00
        Calcuates the parent duration, and all the child event end time.
        '''
        start_time = self._convert_time_string_to_datetime(start_time_str)
        self.parent_event['{start_time}'] = start_time
        cur_time = start_time
        for child_event in self.child_events:
            child_event['{end_time}'] = cur_time+timedelta(minutes=child_event['{duration}'])
            cur_time = child_event['{end_time}']

        # If parent event duration already set, no need to derive it.
        if self.parent_event['{duration}'] is not None:
            return
        duration_sum = 0
        for child_event in self.child_events:
            duration_sum+=child_event['{duration}']
        self.parent_event['{duration}'] = duration_sum

    def _convert_datetime_values_to_string(self,input_dict):
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
        return [self._convert_datetime_values_to_string(child_event) for child_event in self.child_events]

    def get_end_time(self):
        end_time = self.parent_event['{start_time}']+timedelta(minutes=self.parent_event['{duration}'])
        end_time_str = end_time.strftime("%H:%M")
        return end_time_str

class MeetingParser:
    def __init__(self):
        self.role_dict = {}
        self.event_list = []

    def parse_roles(self, content):
        lines = content.strip().split('\n')
        for line in lines:
            if '：' in line:
                role, name = line.split('：', 1)
                self.role_dict[role] = name

    def _split_with_hash(self,document):
        event_list = re.split(r"(?=#)", document)
        # Remove leading and trailing whitespaces from each split string
        event_list = [event.strip() for event in event_list if event.strip()]
        return event_list

    def _parser_parent_events_string(self,parent_event_string):
        parts = parent_event_string.split()
        if len(parts) == 2 and parts[1].isdigit():
            return {'event_name':parts[0],'duration':int(parts[1])}
        return {'event_name': parent_event_string}

    def _get_person_name_from_role(self,role):
        if role not in self.role_dict:
            return role
        return self.role_dict[role]

    def _parse_event_string(self,event_string):
        lines = event_string.strip().split('\n')
        parent_event_string, *lines = map(str.strip, lines)

        # parse parent event.
        parent_event_string = parent_event_string.strip("# ").strip()
        parent_event_dict = self._parser_parent_events_string(parent_event_string)
        parent_event = ParentEvent(parent_event_dict['event_name'])
        if 'duration' in parent_event_dict:
            parent_event.set_parent_event_duration(parent_event_dict['duration'])

        # Parse child events.
        for line in lines:
            parts = line.split()
            role = parts[2]
            person_name = self._get_person_name_from_role(role)
            parent_event.add_child_event(parts[0],int(parts[1]),person_name)
        
        return parent_event

    def _is_notice_block(self,event_string):
        '''
        A notice block is like "# 会议结束"
        '''
        if event_string.count('\n')==0:
            if len(event_string.split())==2:
                return True
        return False

    def _parse_notice_string(self,event_string):
        notice = event_string.split()[-1]
        return NoticeEvent(notice)

    def parse_agenda(self, content):
        block_list = self._split_with_hash(content)
        for block_string in block_list:
            if self._is_notice_block(block_string):
                event = self._parse_notice_string(block_string)
            else:
                event = self._parse_event_string(block_string)
            self.event_list.append(event)

    def parse_file(self, filename):
        with open(filename, 'r', encoding='utf-8') as file:
            content = file.read()

        role_section = content.split("【角色表】", 1)[1]
        agenda_section = content.split("【议程表】", 1)[1]

        self.parse_roles(role_section)
        self.parse_agenda(agenda_section)


if __name__=='__main__':
    # tests.
    parser = MeetingParser()
    parser.parse_file("/home/didi/myproject/tmma/user_input.txt")

    role_dict = parser.role_dict
    event_list = parser.event_list

    print("Role Dictionary:")
    print(role_dict)
    print("\nAgenda List:")
    cur_time = '15:00'
    for event in event_list:
        event.calculate_time(cur_time)
        cur_time = event.get_end_time()
        if isinstance(event, ParentEvent):
            print('111 ',event.get_parent_event())
            print('2222',event.get_child_events())
        elif isinstance(event, NoticeEvent):
            print('notice: ',event.get_event())
