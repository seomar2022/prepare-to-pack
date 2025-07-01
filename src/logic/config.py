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
    with open(config_path, 'r') as f:
        config = json.load(f)
    return os.path.expanduser(config.get("download_path", "~/Downloads"))

def save_download_path(path):
    config_path = get_config_path()
    with open(config_path, 'w') as f:
        json.dump({"download_path": path}, f, indent=2)