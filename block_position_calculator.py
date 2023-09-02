from excel_utils import coordinate_string_to_index, index_to_coordinate_string, add_coordinates
from config_reader import ConfigReader

class PositionCalculator:

    def __init__(self, config_path):
        cr = ConfigReader(config_path)
        self.config = cr.get_config()
        self.start_coord_idx_dict = {}

    def set_schedule_block_height(self, height):
        self.config['schedule_block']['height'] = height

    def get_start_coords(self):
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
                start_coord = self.calculate_start_coord(
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

    def calculate_start_coord(self, start_coord_dict, subject_name,
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
    # col_size, row_size
    # block_size_dict = {'title_block':(2,1),'theme_block':(2,2),'schedule_block':(2,10)}

    pc = PositionCalculator('/home/didi/myproject/tmma/block_position_config.yaml')
    pc.set_schedule_block_height(43)
    start_coord_dict = pc.get_start_coords()
    print(start_coord_dict)
    print(pc.calculate_sizes())