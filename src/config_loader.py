import json
import os


class ConfigLoader:
    def __init__(self, json_path="config/matrices_margen.json"):
        self.json_path = json_path
        self.config_data = {}

    def load_config(self):
        if not os.path.exists(self.json_path):
            raise FileNotFoundError(
                f"No se encontró el archivo de configuración en: {self.json_path}")

        with open(self.json_path, 'r', encoding='utf-8') as file:
            self.config_data = json.load(file)
        return self.config_data
