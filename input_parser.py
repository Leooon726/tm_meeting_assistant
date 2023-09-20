import re
from datetime import datetime, timedelta
from agenda_event import ParentEvent, NoticeEvent


class InputTxtParser:
    def __init__(self):
        self.role_dict = {}
        self.meeting_info_dict = {}
        self.event_list = []
        self.project_info = {}

    def _parse_roles(self, content):
        lines = content.strip().split('\n')
        for line in lines:
            if '【' in line:
                continue
            if '：' in line:
                role, name = line.split('：', 1)
                role = role.strip()
                name = name.strip()
                self.role_dict[role] = name
            elif ':' in line:
                role, name = line.split(':', 1)
                role = role.strip()
                name = name.strip()
                self.role_dict[role] = name

    def _split_with_hash(self, document):
        event_list = re.split(r"(?=#)", document)
        # Remove leading and trailing whitespaces from each split string
        event_list = [event.strip() for event in event_list if event.strip()]
        return event_list

    def _parser_parent_events_string(self, parent_event_string):
        parts = parent_event_string.split()
        if len(parts) == 2 and parts[1].isdigit():
            return {'event_name': parts[0], 'duration': int(parts[1])}
        return {'event_name': parent_event_string}

    def _get_person_name_from_role(self, role):
        if role not in self.role_dict:
            return role
        return self.role_dict[role]

    def _parse_event_string(self, event_string):
        lines = event_string.strip().split('\n')
        parent_event_string, *lines = map(str.strip, lines)

        # parse parent event.
        parent_event_string = parent_event_string.strip("# ").strip()
        parent_event_dict = self._parser_parent_events_string(
            parent_event_string)
        parent_event = ParentEvent(parent_event_dict['event_name'])
        if 'duration' in parent_event_dict:
            parent_event.set_parent_event_duration(
                parent_event_dict['duration'])

        # Parse child events.
        for line in lines:
            parts = line.split()
            role = parts[2]
            person_name = self._get_person_name_from_role(role)
            parent_event.add_child_event(parts[0], int(parts[1]), person_name)

        return parent_event

    def _is_notice_block(self, event_string):
        '''
        A notice block is like "# 会议结束"
        '''
        if event_string.count('\n') == 0:
            if len(event_string.split()) == 2:
                return True
        return False

    def _parse_notice_string(self, event_string):
        notice = event_string.split()[-1]
        return NoticeEvent(notice)

    def _parse_agenda(self, content):
        if content.startswith('【'):
            content = content.split('\n', 1)[1]
        block_list = self._split_with_hash(content)
        for block_string in block_list:
            if self._is_notice_block(block_string):
                event = self._parse_notice_string(block_string)
            else:
                event = self._parse_event_string(block_string)
            self.event_list.append(event)

    def _parse_meeting_info(self, content):
        lines = content.strip().split('\n')
        for line in lines:
            if '：' in line:
                name, info = line.split('：', 1)
                name = name.strip()
                info = info.strip()
                self.meeting_info_dict[name] = info
            elif ':' in line:
                name, info = line.split(':', 1)
                name = name.strip()
                info = info.strip()
                self.meeting_info_dict[name] = info

    def parse_project_info(self, content):
        '''
        Removes the first line and keep the following content.
        '''
        self.project_info = {'备稿演讲项目简介':content.split('\n', 1)[1]}

    def parse_file(self, filename):
        with open(filename, 'r', encoding='utf-8') as file:
            content = file.read()

        sections = content.split('\n\n')

        for section in sections:
            if '【角色表】' in section:
                role_section = section
            elif '【会议信息】' in section:
                meeting_info_section = section
            elif '【议程表】' in section:
                agenda_section = section
            elif '【备稿演讲项目简介】' in section:
                project_section = section

        self._parse_roles(role_section)

        self._parse_meeting_info(meeting_info_section)

        self._parse_agenda(agenda_section)

        self.parse_project_info(project_section)

    def get_total_event_num(self):
        cnt = len(self.event_list)
        for event in self.event_list:
            if isinstance(event, ParentEvent):
                cnt += len(event.get_child_events())
        return cnt
    
    def get_meeting_start_time(self):
        return self.meeting_info_dict['开始时间']


if __name__ == '__main__':
    # tests.
    parser = InputTxtParser()
    parser.parse_file("/home/lighthouse/tm_meeting_assistant/example/output_files/user_input.txt")

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
