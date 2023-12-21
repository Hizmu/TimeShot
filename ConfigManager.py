import os
import configparser
import sys

class ConfigManager:
    CONFIG_FILE = os.path.abspath(os.path.join(sys.path[0], '..', '..', 'config.ini'))
    PATH = os.path.split(CONFIG_FILE)[0]

    DEFAULT_CONFIG = {
        'DATA': {
            'AutoStart': False,
            'Delete': False,
            'Arhive': False,
        },
        'IntervalBetweenScreenshot': {
            'seconds': 5,
            'minutes': 0,
            'hours': 0,
        },
        'TimeToDelete': {
            'days': 2,
            'months': 0,
        },
        'Dirs': {
            'arhive': 'C:/screenshotArchives/',
            'screenshot': 'C:/screenshots/',
            'exe': F'{PATH}',
        },
    }

    def __init__(self):
        self.config = configparser.ConfigParser()
        self.read()

    def read(self):
        if not os.path.exists(self.CONFIG_FILE):
            self.__create_default_config()
        else:
            self.__read_config()

    def __read_config(self):
        with open(self.CONFIG_FILE, 'r') as configfile:
            self.config.read_file(configfile)

    def save(self):
        with open(self.CONFIG_FILE, 'w') as configfile:
            self.config.write(configfile)

    def reset(self):
        self.__create_default_config()

    def __create_default_config(self):
        self.config.read_dict(self.DEFAULT_CONFIG)
        self.save()

    def __getitem__(self, arg):
        return self.config[arg]

    def set(self, section, option, value):
        self.config.set(section, option, str(value))
        self.save()
