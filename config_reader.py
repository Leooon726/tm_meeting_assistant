import yaml

from excel_utils import coordinate_string_to_index,index_to_coordinate_string,add_coordinates

class ConfigReader:
    def __init__(self, file_path):
        self.file_path = file_path
        self.config = self._load_config()

    def _load_config(self):
        with open(self.file_path, "r") as file:
            config = yaml.safe_load(file)
        return config

    def get_config(self):
        return self.config

def calculate_column_index(column_ref_object_start_idx,column_ref_object_end_idx,subject_column_width,column_constraint_type,column_constraint_orientation):
    col_start_idx = None
    if column_constraint_type == 'adjacent':
        if column_constraint_orientation == 'left':
            col_start_idx = column_ref_object_start_idx-subject_column_width
        elif column_constraint_orientation == 'right':
            col_start_idx = column_ref_object_end_idx+1
    elif column_constraint_type == 'aligned':
        if column_constraint_orientation == 'left':
            col_start_idx = column_ref_object_start_idx
        elif column_constraint_orientation == 'right':
            col_start_idx = column_ref_object_end_idx-subject_column_width+1
    return col_start_idx

def calculate_row_index(row_ref_object_start_row_idx,row_ref_object_end_row_idx,subject_height,row_constraint_type,row_constraint_orientation):
    row_start_idx = None
    if row_constraint_type == 'adjacent':
        if row_constraint_orientation == 'top':
            row_start_idx = row_ref_object_start_row_idx-subject_height
        elif row_constraint_orientation == 'bottom':
            row_start_idx = row_ref_object_end_row_idx +1
    elif row_constraint_type =='aligned':
        if row_constraint_orientation == 'top':
            row_start_idx = row_ref_object_start_row_idx
        elif row_constraint_orientation == 'bottom':
            row_start_idx = row_ref_object_end_row_idx - subject_height+1
    return row_start_idx

def calculate_start_coord(start_coord_dict,block_size_dict,subject_name,coord_constraint):
    column_ref_object = coord_constraint['column_position']['object']
    column_ref_object_start_col_idx,_ = coordinate_string_to_index(start_coord_dict[column_ref_object])
    column_ref_object_end_col_idx = column_ref_object_start_col_idx+block_size_dict[column_ref_object][0]-1
    column_constraint_type = coord_constraint['column_position']['type']
    column_constraint_orientation = coord_constraint['column_position']['orientation']
    subject_width = block_size_dict[subject_name][0]
    col_idx = calculate_column_index(column_ref_object_start_col_idx,column_ref_object_end_col_idx,subject_width,column_constraint_type,column_constraint_orientation)
    
    row_ref_object = coord_constraint['row_position']['object']
    _,row_ref_object_start_row_idx = coordinate_string_to_index(start_coord_dict[row_ref_object])
    row_ref_object_end_row_idx = row_ref_object_start_row_idx+block_size_dict[row_ref_object][1]-1
    row_constraint_type = coord_constraint['row_position']['type']
    row_constraint_orientation = coord_constraint['row_position']['orientation']
    subject_height = block_size_dict[subject_name][1]
    row_idx = calculate_row_index(row_ref_object_start_row_idx,row_ref_object_end_row_idx,subject_height,row_constraint_type,row_constraint_orientation)

    return index_to_coordinate_string(col_idx,row_idx)

if __name__=='__main__':
    # col_size, row_size
    block_size_dict = {'title_block':(2,1),'theme_block':(2,2),'schedule_block':(2,10)}

    cr = ConfigReader('/home/didi/myproject/tmma/config.yaml')
    config = cr.get_config()

    start_coord_dict = {}
    for key,value in config.items():
        if 'coordinate' in value:
            start_coord_dict[key]=value['coordinate']
        else:
            start_coord = calculate_start_coord(start_coord_dict,block_size_dict,key,value)
            start_coord_dict[key] = start_coord

    print(start_coord_dict)
