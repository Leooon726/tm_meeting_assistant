from openpyxl.utils.cell import coordinate_from_string, get_column_letter,column_index_from_string

def coordinate_string_to_index(coord):
    coord_tuple = coordinate_from_string(coord)  # Parse the first coordinate string
    coord_col_idx = column_index_from_string(coord_tuple[0])
    return coord_col_idx,coord_tuple[1]

def index_to_coordinate_string(col_idx,row_idx):
    col_letter = get_column_letter(col_idx)
    return col_letter+str(row_idx)

def add_coordinates(coord, offsets):
    '''
    coord can be 'A1', offsets can be (1, 2)
    '''
    coord_col_idx,coord_row_idx = coordinate_string_to_index(coord)
    res_coord_col_idx = coord_col_idx+offsets[0]
    res_coord_row_idx = coord_row_idx+offsets[1]
    return index_to_coordinate_string(res_coord_col_idx,res_coord_row_idx)