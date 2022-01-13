from configparser import ConfigParser
import logging, os

class SettingsMeta(type):
    def __init__(self, name, bases, dic):
        self.__instance = None
        super().__init__(name, bases, dic)

    def __call__(cls, *args, **kwargs):
        if cls.__instance:
            return cls.__instance
        obj = cls.__new__(cls)
        obj.__init__(*args, **kwargs)
        cls.__instance = obj
        return obj

class Settings(metaclass=SettingsMeta):
    def __init__(self):
        logging.info('Create settings manager')
        self._settings_file = "settings.ini"
        self._config_parser = ConfigParser()
        if self.check_settings():
            logging.info('Reading settings.ini')
            self._config_parser.read(self._settings_file)
        else:
            logging.error('settings.ini not found')
            logging.info('Create new settings.ini')
            self.create_new_settings_file()
            logging.info('Reading settings.ini')
            self._config_parser.read(self._settings_file)
        
            
        # global settings
        self.device = self._config_parser.get("global_settings", "device")
        self.connection_speed = self._config_parser.get("global_settings", "connection_speed")
        self.data_bits = self._config_parser.get("global_settings", "data_bits")
        self.parity_check = self._config_parser.get("global_settings", "parity_check")
        self.stop_bit = self._config_parser.get("global_settings", "stop_bit")
        self.number_of_weighings = self._config_parser.get("global_settings", "number_of_weighings")
        self.weighing_frequency = self._config_parser.get("global_settings", "weighing_frequency")


    def get(self):
        settings = {"device": self.device, "connection_speed": self.connection_speed,
                    "data_bits": self.data_bits, "parity_check": self.parity_check,
                    "stop_bit": self.stop_bit, "number_of_weighings": self.number_of_weighings, 
                    "weighing_frequency": self.weighing_frequency}
        return settings

    def update(self, config_values):
        logging.info('Update config parametrs')
        for k, v in config_values.items():
            if k == "device":
                self.device = v
            if k == "connection_speed":
                self.connection_speed = v
            if k == "data_bits":
                self.data_bits = v
            if k == "parity_check":
                self.parity_check = v
            if k == "stop_bit":
                self.stop_bit = v
            if k == "number_of_weighings":
                self.number_of_weighings = v
            if k == "weighing_frequency":
                self.weighing_frequency = v

    def save(self):
        logging.info('Save config parametrs')
        # global settings
        self._config_parser.set("global_settings", "device", self.device)
        self._config_parser.set("global_settings", "connection_speed", self.connection_speed)
        self._config_parser.set("global_settings", "data_bits", self.data_bits)
        self._config_parser.set("global_settings", "parity_check", self.parity_check)
        self._config_parser.set("global_settings", "stop_bit", self.stop_bit)
        self._config_parser.set("global_settings", "number_of_weighings", self.number_of_weighings)
        self._config_parser.set("global_settings", "weighing_frequency", self.weighing_frequency)

        with open(self._settings_file, "w") as config:
            self._config_parser.write(config)


    def check_settings(self):
        return os.path.isfile(self._settings_file)

    def create_new_settings_file(self):
        settings = ["[global_settings]", "device = COM7", "connection_speed = 9600", "data_bits = 8", 
                    "parity_check = ODD", "stop_bit = 1", "number_of_weighings = 10", "weighing_frequency = 3"]

        with open(self._settings_file, "w") as file:
            for i in settings:
                file.write(i + "\n")