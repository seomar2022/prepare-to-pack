import os
import json

APP_NAME = "PrepareToPack"


def get_config_path():
    appdata = os.getenv("APPDATA")
    config_dir = os.path.join(appdata, APP_NAME)
    os.makedirs(config_dir, exist_ok=True)
    return os.path.join(config_dir, "config.json")


def load_download_path():
    config_path = get_config_path()
    if not os.path.exists(config_path):
        return os.path.expanduser("~/Downloads")
    with open(config_path, "r") as f:
        config = json.load(f)
    return os.path.expanduser(config.get("download_path", "~/Downloads"))


def save_download_path(path):
    config_path = get_config_path()
    with open(config_path, "w") as f:
        json.dump({"download_path": path}, f, indent=2)


def get_instruction_folder():
    appdata = os.getenv("APPDATA")
    instruction_folder = os.path.join(appdata, APP_NAME, "product_instruction")
    os.makedirs(instruction_folder, exist_ok=True)
    return instruction_folder


def get_product_code_mapping():
    appdata = os.getenv("APPDATA")
    product_code_mapping = os.path.join(appdata, APP_NAME, "product_code_mapping.xlsx")
    return product_code_mapping
