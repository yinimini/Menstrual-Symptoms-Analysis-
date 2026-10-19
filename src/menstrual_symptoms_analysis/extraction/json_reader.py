import pandas as pd
import json
from pathlib import Path
from typing import Any

def read_json() -> Any:
    """
    Load and return data from the user_data.json file.

    The file path is constructed relative to this Python file,
    so the function does not depend on the current working directory.

    Returns:
        Any: The data parsed from the JSON file. 

    """

    project_root = Path(__file__).resolve().parents[3]
    data_path = project_root / "raw_data" / "user_data.json"

    with data_path.open("r", encoding="utf-8") as file:
        data = json.load(file)
    return data


if __name__ == "__main__":
    data = read_json()
    print(data)
