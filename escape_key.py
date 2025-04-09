import json

# Load your local keys.json file
with open("keys.json", "r") as f:
    creds = json.load(f)

# Print out a properly escaped version
print(json.dumps(creds))
