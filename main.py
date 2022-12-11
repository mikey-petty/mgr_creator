import subprocess

def main():

    print("Running Updater")

    subprocess.run("updater.exe")

    print("Now Running MGR")

    subprocess.run("gui.exe")

if __name__ == "__main__":
    main()