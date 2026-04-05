import json
# Load and print formatted JSON
with open('raw_data/user_data.json') as f:
    data = json.load(f)
    print(json.dumps(data, indent=4))