import json
import os

USER_SETTINGS_PATH = "data/user_data.json"


# Save the data
def save_user_data(data):
    with open(USER_SETTINGS_PATH, 'w') as f:
        json.dump(data, f)


# Load the data
def load_user_data():
    if not os.path.exists(USER_SETTINGS_PATH):
        # Create json if it doesnt exists
        data = {
            "file_path": None,
            "sheet_name": None
        }
        with open(USER_SETTINGS_PATH, 'w') as json_file:
            json.dump(data, json_file, indent=4)
        return data
    else:
        with open(USER_SETTINGS_PATH, 'r') as f:
            return json.load(f)
