import json

with open('config.json','r') as f:
    config = json.load(f)

print(config['DEFAULT']['key_decisions']) 
