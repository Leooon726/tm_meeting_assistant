import uuid
import os

from openpyxl.drawing.image import Image

from excel_utils import coordinate_string_to_index,get_merged_cell_size,is_cell_in_block,subtract_coordinates,add_coordinates
from sheet_image_loader import SheetImageLoader

class XlsxBlockImageWriter:
    def __init__(self,template_sheet,source_block_position,target_block_start_coord,temp_dir):
        self.template_sheet = template_sheet
        self.temp_dir = temp_dir
        self.source_block_start_coord = source_block_position['start_coord']
        self.source_block_end_coord = source_block_position['end_coord']
        self.target_block_start_coord = target_block_start_coord
        self.image_loader = SheetImageLoader(self.template_sheet)

        self.temp_image_paths = []

    def write(self,target_sheet):
        # step1: get image source coords from template sheet block.
        image_source_coords = self._get_image_source_coords()
        for image_source_coord in image_source_coords:
            image_coord_offset = self._get_image_coord_offset(image_source_coord)
            image_tartget_coord = self._get_image_target_coords(image_coord_offset)
            self._add_image(target_sheet,image_source_coord,image_tartget_coord)
        return target_sheet

    def _get_image_source_coords(self):
        image_source_coords = []
        start_col, start_row = coordinate_string_to_index(self.source_block_start_coord)
        end_col, end_row = coordinate_string_to_index(self.source_block_end_coord)
        for row in self.template_sheet.iter_rows(min_row=start_row,
                                            min_col=start_col,
                                            max_row=end_row,
                                            max_col=end_col):
            for cell in row:
                if not self.image_loader.image_in(cell.coordinate):
                    continue
                image_source_coords.append(cell.coordinate)
        return image_source_coords

    def _get_image_coord_offset(self,image_source_coord):
        image_offset = subtract_coordinates(image_source_coord,self.source_block_start_coord)
        return image_offset

    def _get_image_target_coords(self,image_offset):
        target_coord = add_coordinates(self.target_block_start_coord,image_offset)
        return target_coord

    def _add_image(self,target_sheet,
                  source_cell_coord,
                  target_cell_coord):
        if not self.image_loader.image_in(source_cell_coord):
            print(source_cell_coord,"does not contain an image")
            return

        pil_image = self.image_loader.get(source_cell_coord)
        # Save the PIL image as a temporary PNG image
        temp_image_path = f"{self.temp_dir}/{str(uuid.uuid4())}_temp_image.png"
        pil_image.save(temp_image_path, format="PNG")

        # Add the PNG image to the worksheet
        image = Image(temp_image_path)
        self._resize_image_base_on_cell_height(image,source_cell_coord)

        # Add the image to the worksheet
        target_sheet.add_image(image, target_cell_coord)
        self.temp_image_paths.append(temp_image_path)
        return target_sheet

    @staticmethod
    def _points_to_pixels(points, conversion_factor=1.333):
        pixels = points * conversion_factor
        return pixels

    @staticmethod
    def _characters_to_pixels(character_width, font_size):
        character_width_pixels = 9.3
        # Convert character width to pixels
        pixels = character_width * character_width_pixels
        return pixels

    def _resize_image(self, image, desired_width=-1, desired_height=-1):
        original_width, original_height = image.width, image.height

        if desired_width == -1 and desired_height == -1:
            # Return the original image if no resizing is desired
            return image

        if desired_width == -1:
            # Calculate new width while preserving aspect ratio
            new_width = int(desired_height *
                            (original_width / original_height))
            new_height = desired_height
        elif desired_height == -1:
            # Calculate new height while preserving aspect ratio
            new_width = desired_width
            new_height = int(desired_width *
                             (original_height / original_width))
        else:
            # Resize with desired width and height
            new_width = desired_width
            new_height = desired_height

        image.height = new_height
        image.width = new_width

    def _resize_image_base_on_cell_height(self,image,source_cell_coord):
        cell_width_in_characters,cell_height_in_points = get_merged_cell_size(self.template_sheet,source_cell_coord)
        cell_height_in_pixels = self._points_to_pixels(cell_height_in_points)
        
        cell_font_size = self.template_sheet[source_cell_coord].font.size
        cell_width_in_pixels = self._characters_to_pixels(cell_width_in_characters,cell_font_size)

        cell_aspect_ratio = cell_width_in_pixels/cell_height_in_pixels
        image_aspect_ratio = image.width/image.height
        if cell_aspect_ratio>image_aspect_ratio:
            self._resize_image(image,
                            desired_width=-1,
                            desired_height=cell_height_in_pixels)
        else:
            self._resize_image(image,
                            desired_width=cell_width_in_pixels,
                            desired_height=-1)

    def delete_temp_images():
        for temp_jpeg_path in self.temp_image_paths:
            os.remove(temp_jpeg_path)

if __name__ == '__main__':
    from openpyxl import Workbook, load_workbook
    template_workbook = load_workbook('/home/lighthouse/agenda_template_zoo/huangpu_rise_template_for_print/huangpu_rise_template_for_print.xlsx')
    template_sheet = template_workbook['page2']
    source_block_position = {'start_coord':'A1','end_coord':'L38'}
    target_coord = 'A1'
    writer = XlsxBlockImageWriter(template_sheet,source_block_position,target_coord)
    print(writer._get_image_source_coords())
    print(writer._get_image_coord_offset('B23'))
    print(writer._get_image_target_coords((11,0)))
