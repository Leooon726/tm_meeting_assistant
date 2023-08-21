from openpyxl.utils.cell import coordinate_from_string, get_column_letter, column_index_from_string


def coordinate_string_to_index(coord):
    coord_tuple = coordinate_from_string(
        coord)  # Parse the first coordinate string
    coord_col_idx = column_index_from_string(coord_tuple[0])
    return coord_col_idx, coord_tuple[1]


def index_to_coordinate_string(col_idx, row_idx):
    col_letter = get_column_letter(col_idx)
    return col_letter + str(row_idx)


def add_coordinates(coord, offsets):
    '''
    coord can be 'A1', offsets can be (1, 2)
    '''
    coord_col_idx, coord_row_idx = coordinate_string_to_index(coord)
    res_coord_col_idx = coord_col_idx + offsets[0]
    res_coord_row_idx = coord_row_idx + offsets[1]
    return index_to_coordinate_string(res_coord_col_idx, res_coord_row_idx)


def get_left_top_coordinate(sheet):
    min_row = sheet.min_row
    min_col = sheet.min_column
    left_top_coordinate = f"{get_column_letter(min_col)}{min_row}"
    return left_top_coordinate


def get_right_bottom_coordinate(sheet):
    max_row = sheet.max_row
    max_col = sheet.max_column
    right_bottom_coordinate = f"{get_column_letter(max_col)}{max_row}"
    return right_bottom_coordinate


def get_sheet_dimensions(sheet):
    width = sheet.max_column - sheet.min_column + 1
    height = sheet.max_row - sheet.min_row + 1
    return width, height
