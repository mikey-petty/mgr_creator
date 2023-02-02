import subprocess
import os
import sys
import updater


def main():
    installer_path = updater.check_and_run_updater("gui.exe")

    if installer_path == None:
        subprocess.run("gui.exe")

    else:
        subprocess.Popen([installer_path, "/SILENT"])
        sys.exit(0)


if __name__ == "__main__":
    main()
