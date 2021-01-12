import os
import json


class Configuration:
    def __init__(self, config_path, DEFAULT_CONFIG):
        self.config_path = config_path
        if os.path.exists(config_path):
            with open(config_path, 'r') as jsonfile:
                conf = json.load(jsonfile)
                # print(conf)
                self._config = conf
        else:
            with open(config_path, 'w') as jsonfile:
                json.dump(DEFAULT_CONFIG, jsonfile, indent=4, ensure_ascii=False)
                print('Program init')
                self._config = DEFAULT_CONFIG

    def __setitem__(self, key, value):
        self._config[key] = value
        with open(self.config_path, 'w') as jsonfile:
            json.dump(self._config, jsonfile, indent=4, ensure_ascii=False)

    def __getitem__(self, key):
        return self._config[key] if key in self._config else None

    def __contains__(self, key):
        return key in self._config
