from excel_utils import coordinate_string_to_index, index_to_coordinate_string, add_coordinates
from config_reader import ConfigReader

class PositionCalculator:

    def __init__(self, config_path):
        cr = ConfigReader(config_path)
        self.config = cr.get_config()
        self.start_coord_idx_dict = {}

    def set_block_size_base_on_template_block(self,template_block_size_dict):
        def set_width_if_not_exist(template_block_name,template_block_size):
            if 'width' not in self.config:
                self.config[template_block_name]['width'] = template_block_size['width']

        def set_height_if_not_exist(template_block_name,template_block_size):
            if 'height' not in self.config:
                self.config[template_block_name]['height'] = template_block_size['height']

        for template_block_name, template_block_size in template_block_size_dict.items():
            # schedule_block is an exception, it only exists in block to be written, and stacks vertially with parent template block and child template block. So they have the same width.
            if template_block_name == 'parent_block' and 'schedule_block' in self.config:
                set_width_if_not_exist('schedule_block',template_block_size)
                continue
            if template_block_name not in self.config:
                continue
            set_width_if_not_exist(template_block_name,template_block_size)
            set_height_if_not_exist(template_block_name,template_block_size)

    def set_schedule_block_height(self, height):
        '''
        THe schedule block height depends on how many events in the meeting, it should be derived from user input.
        '''
        self.config['schedule_block']['height'] = height

    def _check_if_all_block_size_exist(self):
        for block_name, block_data in self.config.items():
            assert 'height' in block_data, f"Height field is not in {block_name}"
            assert 'width' in block_data, f"Width field is not in {block_name}"

    def get_start_coords(self):
        self._check_if_all_block_size_exist()
        self._calcuate_start_coords()
        start_coord_dict = {}
        for block_name, start_coord_idx in self.start_coord_idx_dict.items():
            start_coord_dict[block_name] = index_to_coordinate_string(
                start_coord_idx[0], start_coord_idx[1])
        return start_coord_dict

    def _calcuate_start_coords(self):
        for key, value in self.config.items():
            if 'coordinate' in value:
                self.start_coord_idx_dict[key] = coordinate_string_to_index(
                    value['coordinate'])
            else:
                start_coord = self._calculate_start_coord(
                    self.start_coord_idx_dict, key, value)
                self.start_coord_idx_dict[key] = start_coord

    def calculate_sizes(self):
        size_dict = {}
        for block_name, constraint in self.config.items():
            size_dict[block_name] = dict(width=-1, height=-1)
            if 'width' in constraint and isinstance(constraint['width'], int):
                size_dict[block_name]['width'] = constraint['width']
            else:
                size_dict[block_name]['width'] = self.calculate_block_width()

            if 'height' in constraint and isinstance(constraint['height'],
                                                     int):
                size_dict[block_name]['height'] = constraint['height']
            else:
                subject_start_row = self.start_coord_idx_dict[block_name][1]
                constraint_object_start_row = self.start_coord_idx_dict[
                    constraint['height']['object']][1]
                constraint_object_height = size_dict[constraint['height']
                                                     ['object']]['height']
                size_dict[block_name]['height'] = self.calculate_block_height(
                    subject_start_row, constraint['height']['type'],
                    constraint['height']['orientation'],
                    constraint_object_start_row, constraint_object_height)
        return size_dict

    def _calculate_start_coord(self, start_coord_dict, subject_name,
                              coord_constraint):
        column_ref_object = coord_constraint['column_position']['object']
        column_ref_object_start_col_idx, _ = start_coord_dict[
            column_ref_object]
        column_ref_object_end_col_idx = column_ref_object_start_col_idx + self.config[
            column_ref_object]['width'] - 1
        column_constraint_type = coord_constraint['column_position']['type']
        column_constraint_orientation = coord_constraint['column_position'][
            'orientation']
        subject_width = int(coord_constraint['width'])
        col_idx = self.calculate_column_index(column_ref_object_start_col_idx,
                                              column_ref_object_end_col_idx,
                                              subject_width,
                                              column_constraint_type,
                                              column_constraint_orientation)

        row_ref_object = coord_constraint['row_position']['object']
        _, row_ref_object_start_row_idx = start_coord_dict[row_ref_object]
        row_ref_object_end_row_idx = row_ref_object_start_row_idx + self.config[
            row_ref_object]['height'] - 1
        row_constraint_type = coord_constraint['row_position']['type']
        row_constraint_orientation = coord_constraint['row_position'][
            'orientation']
        subject_height = coord_constraint['height']
        row_idx = self.calculate_row_index(row_ref_object_start_row_idx,
                                           row_ref_object_end_row_idx,
                                           subject_height, row_constraint_type,
                                           row_constraint_orientation)

        return (col_idx, row_idx)

    @staticmethod
    def calculate_block_height(subject_start_row, constraint_type,
                               constraint_orientation,
                               constraint_object_start_row,
                               constraint_object_height):
        height = None
        if constraint_type == 'adjacent':
            if constraint_orientation == 'top':
                height = constraint_object_start_row - subject_start_row
            elif constraint_orientation == 'bottom':
                assert False
        elif constraint_type == 'aligned':
            if constraint_orientation == 'bottom':
                constraint_object_end_row = constraint_object_start_row + constraint_object_height - 1
                height = constraint_object_end_row - subject_start_row + 1
            else:
                assert False
        return height

    @staticmethod
    def calculate_block_width():
        return -1

    @staticmethod
    def calculate_column_index(column_ref_object_start_idx,
                               column_ref_object_end_idx, subject_column_width,
                               column_constraint_type,
                               column_constraint_orientation):
        col_start_idx = None
        if column_constraint_type == 'adjacent':
            if column_constraint_orientation == 'left':
                col_start_idx = column_ref_object_start_idx - subject_column_width
            elif column_constraint_orientation == 'right':
                col_start_idx = column_ref_object_end_idx + 1
        elif column_constraint_type == 'aligned':
            if column_constraint_orientation == 'left':
                col_start_idx = column_ref_object_start_idx
            elif column_constraint_orientation == 'right':
                col_start_idx = column_ref_object_end_idx - subject_column_width + 1
        return col_start_idx

    @staticmethod
    def calculate_row_index(row_ref_object_start_row_idx,
                            row_ref_object_end_row_idx, subject_height,
                            row_constraint_type, row_constraint_orientation):
        row_start_idx = None
        if row_constraint_type == 'adjacent':
            if row_constraint_orientation == 'top':
                assert isinstance(subject_height, int)
                row_start_idx = row_ref_object_start_row_idx - subject_height
            elif row_constraint_orientation == 'bottom':
                row_start_idx = row_ref_object_end_row_idx + 1
        elif row_constraint_type == 'aligned':
            if row_constraint_orientation == 'top':
                row_start_idx = row_ref_object_start_row_idx
            elif row_constraint_orientation == 'bottom':
                assert isinstance(subject_height, int)
                row_start_idx = row_ref_object_end_row_idx - subject_height + 1
        return row_start_idx


if __name__ == '__main__':
    pc = PositionCalculator('/home/lighthouse/tm_meeting_assistant/example/jabil_jouse_template_for_print/block_position_config.yaml')
    simulated_template_block_size_dict = {'title_block': {'width': 14, 'height': 7}, 'theme_block': {'width': 10, 'height': 3}, 'parent_block': {'width': 10, 'height': 1}, 'child_block': {'width': 10, 'height': 1}, 'notice_block': {'width': 10, 'height': 1}, 'rule_block': {'width': 4, 'height': 7}, 'information_block': {'width': 4, 'height': 22}, 'contact_block': {'width': 4, 'height': 9}, 'project_block': {'width': 4, 'height': 3}}
    pc.set_block_size_base_on_template_block(simulated_template_block_size_dict)
    pc.set_schedule_block_height(43)
    start_coord_dict = pc.get_start_coords()
    print(start_coord_dict)
    print(pc.calculate_sizes())