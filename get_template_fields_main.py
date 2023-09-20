import argparse
import json

from config_reader import ConfigReader
from template_reader import XlsxTemplateReader


def _remove_duplicated_fields(input_fields):
    output_fields = []
    seen_field_names = set()
    for field_item in input_fields:
        field_name = field_item['field_name']
        if field_name in seen_field_names:
            continue
        output_fields.append(field_item)
        seen_field_names.add(field_name)
    return output_fields

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Get fields of an agenda that need to be filled by users.')
    parser.add_argument('-c','--config', dest='config_path', type=str, required=True,
                        help='Path to the configuration file')
    parser.add_argument('-o','--output_json_path', dest='output_json_path', type=str, required=True,
                        help='Path to the output json file path')
    args = parser.parse_args()

    cr = ConfigReader(args.config_path)
    config = cr.get_config()
    fields = []
    for target_sheet_config_dict in config['target_sheets']:
        template_sheet_name = target_sheet_config_dict['template_sheet_name']
        template_position_sheet_name = target_sheet_config_dict['template_position_sheet_name']
        template_file = config['template_excel']
        tr = XlsxTemplateReader(template_file, template_sheet_name,
                            template_position_sheet_name)
        fields += tr.get_user_filled_fields_from_sheet()

    # Remove Duplicated fields.
    filtered_fields = _remove_duplicated_fields(fields)

    with open(args.output_json_path, "w") as file:
        json.dump(filtered_fields, file, ensure_ascii=False)
        