from configparser import ConfigParser
import logging
import os


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
        self.mode = self._config_parser.get("global_settings", "mode")

        # weight sum
        self.date_sum = self._config_parser.get("weight_sum", "date")
        self.weight_sum = self._config_parser.get("weight_sum", "weight")

        # weight kort
        self.date_kort = self._config_parser.get("weight_kort", "date")
        self.weight_kort = self._config_parser.get("weight_kort", "weight")

    def get(self):
        settings = {"device": self.device, "connection_speed": self.connection_speed,
                    "data_bits": self.data_bits, "parity_check": self.parity_check,
                    "stop_bit": self.stop_bit, "date_sum": self.date_sum,
                    "weight_sum": self.weight_sum, "date_kort": self.date_kort,
                    "weight_kort": self.weight_kort, "mode": self.mode}
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
            if k == "mode":
                self.mode = v
            if k == "date_sum":
                self.date_sum = v
            if k == "weight_sum":
                self.weight_sum = v
            if k == "date_kort":
                self.date_kort = v
            if k == "weight_kort":
                self.weight_kort = v

    def save(self):
        logging.info('Save config parametrs')
        # global settings
        self._config_parser.set("global_settings", "device", self.device)
        self._config_parser.set("global_settings", "connection_speed", self.connection_speed)
        self._config_parser.set("global_settings", "data_bits", self.data_bits)
        self._config_parser.set("global_settings", "parity_check", self.parity_check)
        self._config_parser.set("global_settings", "stop_bit", self.stop_bit)
        self._config_parser.set("global_settings", "mode", self.mode)

        # weight sum
        self._config_parser.set("weight_sum", "date", self.date_sum)
        self._config_parser.set("weight_sum", "weight", self.weight_sum)

        # weight kort
        self._config_parser.set("weight_kort", "date", self.date_kort)
        self._config_parser.set("weight_kort", "weight", self.weight_kort)

        with open(self._settings_file, "w") as config:
            self._config_parser.write(config)

    def check_settings(self):
        return os.path.isfile(self._settings_file)

    def create_new_settings_file(self):
        settings = ("[global_settings]", "device = COM7", "connection_speed = 9600", "data_bits = 8",
                    "parity_check = ODD", "stop_bit = 1", "mode = 0\n", "[weight_sum]", "date = A",
                    "weight = F\n", "[weight_kort]", "date = A", "weight = G")

        with open(self._settings_file, "w") as file:
            for i in settings:
                file.write(i + "\n")
