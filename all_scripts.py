import subprocess
import time

scripts = ['create_datasheet.py', 'fetch_serp_data.py', 'update_datasheet_page_structure.py', 'scrape_page_structure.py']

for script in scripts:
    subprocess.call(["python", script])
    time.sleep(30)  # wait for 30 seconds