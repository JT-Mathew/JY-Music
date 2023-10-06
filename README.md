# JY-Music
Jesus Youth Music Slide Maker

How to create the executables


Mac:
1. Open terminal and navigate to the folder and run: 

	python3 -m PyInstaller --onefile main.py

2. After code run, copy the resources folder into the dist folder
3. Create a start.txt file and add the below into it:

	#!/bin/sh

	cd "$(dirname "$0")"
	./main

4. Save the start.txt file and rename it as start.command
5. In terminal enter command:

	 chmod +x start.command

6. Copy the dist folder out to wherever, name it whatever, zip it up, should work on any mac.



Windows:
1. Open cmd and navigate to the folder (using dir and cd) and run:

	PyInstaller --onefile main.py

2. Copy the resources folder into the dist/main
3. Double click on main.exe
4. Copy the main folder out to wherever, name it whatever, zip it up, should work on any windows. 