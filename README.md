# tm_meeting_assistant

Run a example:
```
python3 main.py -i /home/lighthouse/tm_meeting_assistant/example/output_files/user_input.txt -o /home/lighthouse/tm_meeting_assistant/example/output_files/319meeting.xlsx -c /home/lighthouse/tm_meeting_assistant/example/jabil_jouse_template_for_print/engine_config.yaml
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