import json
import os
from pathlib import Path


default_config = {
  "font_list": [
    "Arial",
    "Arial Black",
    "Comic Sans MS",
    "Courier New",
    "Georgia",
    "Impact",
    "Times New Roman",
    "Trebuchet MS",
    "Verdana"
  ],
  "hyperlink": {
    "color": "#0563c1",
    "font_name": "Times New Roman",
    "font_size": 12
  },
  "path": {
    "base_registry_path": "",
    "base_scan_path": ""
  }
}


class Config:
    def __init__(self):
        self.local_dir_config_path = f'{str(Path.home())}\\HyperlinkCreator'
        self.local_file_config_path = f'{str(Path.home())}\\HyperlinkCreator\\config.json'
        self.default_config = default_config
        self.create_local_config()

    def create_local_config(self):
        is_dir_exist = os.path.exists(self.local_dir_config_path)
        if not is_dir_exist:
            os.mkdir(self.local_dir_config_path)
        is_dir_exist = os.path.exists(self.local_dir_config_path)
        is_file_exist = os.path.exists(self.local_file_config_path)
        if is_dir_exist and not is_file_exist:
            with open(self.local_file_config_path, 'w', encoding='cp1251') as file:
                json.dump(self.default_config, file, sort_keys=True, indent=2)

    def save(self, settings):
        is_file_exist = os.path.exists(self.local_file_config_path)
        config = self.load()
        if is_file_exist:
            with open(self.local_file_config_path, 'w', encoding='cp1251') as file:
                for key in list(settings):
                    config[key].update(settings[key])
                json.dump(config, file, sort_keys=True, indent=2)
        else:
            self.default_config = config

    def load(self):
        is_file_exist = os.path.exists(self.local_file_config_path)
        if is_file_exist:
            with open(self.local_file_config_path, 'r', encoding='cp1251') as file:
                local_config = json.load(file)
            return local_config
        else:
            return self.default_config

    def reset_config(self):
        is_file_exist = os.path.exists(self.local_file_config_path)
        if is_file_exist:
            os.remove(self.local_file_config_path)
        self.create_local_config()


if __name__ == '__main__':
    pass
    # local_dir_config_path = f'{str(Path.home())}\\HyperlinkCreator'
    # local_file_config_path = f'{str(Path.home())}\\HyperlinkCreator\\config.json'
    # is_dir_exist = os.path.exists(local_dir_config_path)
    # is_file_exist = os.path.exists(local_file_config_path)
    # print(is_dir_exist and not is_file_exist)
