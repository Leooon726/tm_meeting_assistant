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
        child: {'start_time':'12:01','event_name':'开场白','end_time':'15:01','duration':'8','host_name':'林长宏', 'duration_buffer': 0}
        only event_start_time is given, other time and duration are derived.
        '''
        self.parent_event = {
            'start_time': None,
            'event_name': parent_event_name,
            'duration': None
        }
        self.child_events = []

    def set_parent_event_duration(self, duration):
        self.parent_event['duration'] = float(duration)

    def add_child_event(self, child_event_name, duration, host_name, comment='', duration_buffer=0):
        '''
        Adds to self.child_events with order.
        '''
        self.calculated = False
        data = {
            'event_name': child_event_name,
            'end_time': None,
            'duration': float(duration),
            'host_name': host_name,
            'comment': comment,
            'duration_buffer': float(duration_buffer)
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
            child_event['start_time'] = cur_time
            child_event['end_time'] = cur_time + timedelta(
                minutes=(child_event['duration']+child_event['duration_buffer']))
            cur_time = child_event['end_time']

        # If parent event duration already set, no need to derive it.
        if self.parent_event['duration'] is not None:
            return
        duration_sum = 0
        for child_event in self.child_events:
            duration_sum += child_event['duration']
        self.parent_event['duration'] = duration_sum

    @staticmethod
    def format_float(value):
        """
        if value is 1.5, print 1.5,
        if value is 1.0, print 1.
        """
        if value % 1 == 0:
            return str(int(value))
        else:
            return str(value)

    @staticmethod
    def _convert_datetime_values_to_string(input_dict):
        converted_dict = {}
        for key, value in input_dict.items():
            if isinstance(value, datetime):
                converted_dict[key] = value.strftime("%H:%M")
                continue
            if key == 'duration':
                converted_dict[key] = ParentEvent.format_float(value)
                continue
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