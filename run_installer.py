import subprocess
import sys
import psutil
import time

installer_path = sys.argv[1]
pid = int(sys.argv[2])

while psutil.pid_exists(pid):
    print("yee")
    time.sleep(1)

subprocess.call([sys.argv[1]])
sys.exit(0)
