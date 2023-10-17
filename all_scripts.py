import subprocess
import time

scripts = ['create_datasheet.py', 'fetch_serp_data.py']

for script in scripts:
    subprocess.call(["python", script])
    time.sleep(5)  # wait for 5 seconds