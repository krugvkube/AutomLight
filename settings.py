import json
import os

class Settings:
    def __init__(self):
        self.config_file = "config.json"
        self.default_settings = {
            "save_path": os.path.expanduser("~/Desktop"),
            "column_to_keep": [1, 2, 3, 4, 5, 6, 7]
        }
        self.settings = self.load_settings()

    def load_settings(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r") as f:
                    return json.load(f)
            except:
                return self.default_settings
        return self.default_settings
    
    def save_settings(self, save_path, column_to_keep):
        self.settings = {
            "save_path": save_path,
            "column_to_keep": column_to_keep
        }
        with open(self.config_file, "w") as f:
            json.dump(self.settings, f)
    
    def get_save_path(self):
        return self.settings.get("save_path", self.default_settings["save_path"])
    
    def get_column_to_keep(self):
        return self.settings.get("column_to_keep", self.default_settings["column_to_keep"]) 
