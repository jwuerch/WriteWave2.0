import textrazor
from dotenv import load_dotenv
import os

# Load the .env file
load_dotenv()

textrazor.api_key = os.getenv('TEXTRAZOR_API_KEY')

client = textrazor.TextRazor(extractors=["entities", "topics"])

# Replace this with your text to analyze
text_to_analyze = """
What To Feed a Guinea Pig | RSPCA.
"""

response = client.analyze(text_to_analyze)

# Create a dictionary to store entities and their details
entity_dict = {}

for entity in response.entities():
    # Print original details
    # If entity is already in dictionary, increment count
    if entity.id in entity_dict:
        entity_dict[entity.id][2] += 1
    # If entity is not in dictionary, add it with count 1
    else:
        entity_dict[entity.id] = [entity.relevance_score, entity.confidence_score, 1]

# Sort entities by relevance score and print them
for entity_id, details in sorted(entity_dict.items(), key=lambda item: item[1][0], reverse=True):
    print(entity_id, details[0], details[1], details[2])
