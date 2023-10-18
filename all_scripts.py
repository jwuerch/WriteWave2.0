import subprocess
import time

# scripts = [
#     'create_sheet_serp_data.py',
#     'fetch_serp_data.py',
#     'update_datasheet_page_structure.py',
#     'scrape_page_structure.py',
#     'copy_serp_data.py'
# ]

# scripts = ['create_sheet_serp_data.py',
#            'create_sheet_page_structure.py',
#            'create_sheet_keyword_variations.py',
#            'create_sheet_entities.py',
#            'fetch_serp_data.py',
#            'copy_serp_data.py'
# ]

scripts = ['create_sheet_serp_data.py',
           'create_sheet_page_structure.py',
           'create_sheet_keyword_variations.py'
]

for script in scripts:
    subprocess.call(["python", script])
    time.sleep(5)  # wait for 5 seconds