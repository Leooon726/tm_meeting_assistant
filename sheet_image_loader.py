"""
Contains a SheetImageLoader class that allow you to loadimages from a sheet.
A modify version of https://github.com/ultr4nerd/openpyxl-image-loader
"""

import io,uuid
import string

from PIL import Image


class SheetImageLoader:
    """Loads all images in a sheet"""
    def __init__(self, sheet,temp_dir):
        """Loads all sheet images"""
        self.temp_dir = temp_dir
        self._images = {}
        sheet_images = sheet._images
        for image in sheet_images:
            row = image.anchor._from.row + 1
            col = string.ascii_uppercase[image.anchor._from.col]
            image_dict = {'image_path':self._save_image(image._data),'col_offset':image.anchor._from.colOff,'row_offset':image.anchor._from.rowOff}
            coord = f'{col}{row}'
            if coord not in self._images:
                self._images[coord] = [image_dict]
            else:
                self._images[coord].append(image_dict)

    def image_in(self, cell):
        """Checks if there's an image in specified cell"""
        return cell in self._images

    def get_images_from_cell(self, cell_coord):
        """Retrieves image data from a cell"""
        if cell_coord not in self._images:
            raise ValueError("Cell {} doesn't contain an image".format(cell_coord))
        else:
            return self._images[cell_coord]
            
    def _save_image(self,byte_data):
        pil_image = self._parse_image(byte_data)
        # Save the PIL image as a temporary PNG image
        temp_image_path = f"{self.temp_dir}/{str(uuid.uuid4())}_temp_image.png"
        pil_image.save(temp_image_path, format="PNG")
        return temp_image_path

    @staticmethod
    def _parse_image(byte_data):
        image = io.BytesIO(byte_data())
        return Image.open(image)

if __name__ == '__main__':
    from openpyxl import Workbook, load_workbook
    template_workbook = load_workbook('/home/lighthouse/agenda_template_zoo/huangpu_rise_template_for_print/huangpu_rise_template_for_print.xlsx')
    template_sheet = template_workbook['page1']
    sil = SheetImageLoader(template_sheet,'/home/lighthouse/output_files/adhoc')