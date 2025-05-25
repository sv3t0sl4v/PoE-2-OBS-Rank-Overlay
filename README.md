# PoE-2-OBS-Rank-Overlay

How to Install & Run
1. Install Python

Make sure Python 3.x is installed on your Windows machine. You can download it from python.org.

Make sure to add Python to PATH during installation.
2. Install Dependencies

Open PowerShell or CMD as Administrator and run:

pip install pywin32 psutil selenium webdriver-manager

    pywin32 — needed for Windows service and event logging.

    psutil — to check if OBS is running.

    selenium and webdriver-manager — to automate Chrome browser.

3. Chrome Browser & Driver

Make sure Google Chrome is installed on your system.

webdriver-manager will automatically download the compatible ChromeDriver binary.
4. Save Script

Save the above script as, for example, poe2_rank_service.py somewhere accessible.
5. Register the Service

Open Command Prompt as Administrator, then:

python poe2_rank_service.py install
python poe2_rank_service.py start

    install registers the script as a Windows service.

    start starts the service immediately.

To stop or remove the service:

python poe2_rank_service.py stop
python poe2_rank_service.py remove

6. Logs & Output

    The service writes the rank output to the file you set in OUTPUT_FILE (change this path if needed).

    Service logs (success/errors) can be found in Windows Event Viewer under Application logs, source: PoE2RankChecker.

7. Running the Service

    The service checks if OBS is running.

    If OBS is running, it checks the Path of Exile 2 ladder rank and updates the rank file.

    It repeats this check every 1–5 minutes randomly.

    If OBS is not running, it sleeps and waits 30 seconds before checking again.

Notes

    Running headless Chrome requires a modern Chrome installed.

    Make sure the script has write permissions to the output directory.

    Running as a service means it will start on Windows boot (if set to Automatic).
