# PoE-2-OBS-Rank-Overlay

![Service Demo](demo.png)

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

Save the above script as, for example, poe2_rank_service.py somewhere accessible. IMPORTANT: Edit the config part to your need. You may want to change "Global", if you want it to say something else. Modify to your liking. I decided not to use var, because league title may be too long.

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

    You can set up the service to start automatically when the system starts, if you like.

8. You need to add text item to your scene and adjust as you like. Point OBS to read from the file where the script stores your rank. GL HF SS!
 
Notes
    
    IMPORTANT: If you make changes to the script, e.g. update character name and/or ladder link, you have to also restart the PoE2 Ladder Rank Checker Service! If you don't know how to, just restart your computer. That will do. :)
    
    Running headless Chrome requires a modern Chrome installed.

    Make sure the script has write permissions to the output directory.

    Running as a service means it will start on Windows boot (if set to Automatic).
    
    More about me: https://www.youtube.com/@vest0levs
