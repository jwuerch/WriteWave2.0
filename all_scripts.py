import subprocess
import time

# scripts = [
#     'create_sheet_serp_data.py',
#     'fetch_serp_data.py',
#     'update_datasheet_page_structure.py',
#     'scrape_page_structure.py',
#     'update_datasheet_keyword_variations.py'
# ]

scripts = ['create_sheet_serp_data.py',
           'create_sheet_page_structure.py',
           'create_sheet_keyword_variations.py'
]

for script in scripts:
    subprocess.call(["python", script])
    time.sleep(10)  # wait for 10 seconds