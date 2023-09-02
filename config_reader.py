import yaml

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

if __name__ == '__main__':
    cr = ConfigReader('/home/lighthouse/agenda_template_zoo/jabil_jouse_template_for_print/engine_config.yaml')
    config = cr.get_config()
    print(config)
