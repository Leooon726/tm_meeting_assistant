# tm_meeting_assistant

Run a example:
```
python3 main.py -i /home/lighthouse/tm_meeting_assistant/example/output_files/user_input.txt -o /home/lighthouse/tm_meeting_assistant/example/output_files/319meeting.xlsx -c /home/lighthouse/tm_meeting_assistant/example/jabil_jouse_template_for_print/engine_config.yaml
```

Run a example of get_template_fields_main.py:
```
python3 get_template_fields_main.py -c /home/lighthouse/tm_meeting_assistant/example/jabil_jouse_template_for_print/engine_config.yaml -o /home/lighthouse/tm_meeting_assistant/example/output_files/fields.json
```

## About template preparation
1. The fields of parent event block should be:
{start_time}
{event_name}
{duration}

2. The fields of and child event block should be:
{event_name}
{duration}
{end_time}
{host_name}

3. The fields of notice event block should be:
{time}
{notice}

## Print to pdf
```
unoconv -f pdf /path/to/excel.xlsx
```

## Convert pdf to image
```
convert -density 300 -quality 100 -units PixelsPerInch  /path/to/pdf.pdf path/to/jpg.jpg
```

## Issue Shooting
### Add Chinese fonts:
```
sudo yum install -y fontconfig mkfontscale
sudo yum install ttf-wqy-microhei
sudo yum install ImageMagick
```