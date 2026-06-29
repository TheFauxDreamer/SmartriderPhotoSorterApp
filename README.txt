SCHOOL PHOTO ORGANISER
======================

This sorts the school photos in a folder into "Staff" and "Left", and creates
a "Missing-Photos.txt" list of active students who have no photo. It checks the
GridExport.CSV first and warns you if the CSV looks out of date.


WHAT'S IN THIS FOLDER
---------------------
  Organise-Photos-GUI.ps1        The actual app.
  Launch-Photo-Organiser.vbs     Double-click this to open the app (no black window).
  Build-PhotoOrganiser-EXE.ps1   Optional: turn the app into a proper .exe (see below).
  README-Photo-Organiser.txt     This file.


FOR STAFF - HOW TO USE IT (the easy bit)
----------------------------------------
  1. Make sure GridExport.CSV is in the photos folder, exported from
     People Management in Compass.
  2. Open the Photo Organiser (double-click the launcher, or the .exe if set up).
  3. Check the "Photos folder" shown near the top is the right one.
     If not, click "Change..." and pick the correct folder.
  4. Click "Organise Photos".
  5. If a pop-up says the CSV looks old, follow the advice - usually you'll want
     to export a fresh one from Compass first.
  6. When it finishes you'll see a summary. Done.

No typing, no commands - just buttons.


FOR WHOEVER SETS IT UP - MAKING IT A REAL .EXE (recommended)
------------------------------------------------------------
Turning it into a single .exe means staff just double-click one icon, with no
launcher file and no PowerShell involved at all.

  1. Put all the files above into one folder on a Windows PC with internet.
  2. Right-click  Build-PhotoOrganiser-EXE.ps1  and choose "Run with PowerShell".
  3. It installs a small build tool (first time only) and creates
     "Photo Organiser.exe" in the same folder.
  4. Copy "Photo Organiser.exe" wherever you like - it's self-contained.

Optional icon: drop a file named  icon.ico  into the folder before building and
the .exe will use it.


A NOTE ABOUT WINDOWS SECURITY WARNINGS
--------------------------------------
The .exe is not code-signed, so the very first time it runs on a PC Windows may
show a blue "Windows protected your PC" box. Click "More info" then "Run anyway".
This only happens once per PC. If you'd like to avoid it entirely, your IT team
can either approve the file or sign it with a code-signing certificate.


CHANGING THE "OLD CSV" THRESHOLD
--------------------------------
The warning triggers when the CSV is 3 or more days old. To change that, open
Organise-Photos-GUI.ps1 in Notepad and edit this line near the top:

    $MaxCsvAgeDays = 3

(If you build the .exe, re-run the build after changing it.)
