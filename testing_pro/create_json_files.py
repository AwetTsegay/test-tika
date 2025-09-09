# Create JSON files
import json
import os

# Create a directory to store the JSON files
if not os.path.exists("json_files"):
    os.makedirs("json_files")

for i in range(1, 6):
    data = {"file_id": i, "content": f"This is the content for JSON file {i}."}
    with open(f"json_files/data_{i}.json", "w") as f:
        json.dump(data, f, indent=4)

