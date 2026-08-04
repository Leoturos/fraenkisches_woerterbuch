Installation:
You need to have a working python instalation on your system to use this, get the latest version here: https://www.python.org/downloads/ (the tool is tested under python 3.10)
If you are using a windows system, just run the setup.bat.
On a Unix Os run the command '''pip install -r requirements.txt''' 
How to use:
Run script.py
This will open a window, you will first need to press on "Eingabedatenbank" to input your database you want to check, this database has to be csv readable(like a .tab or .xls file) it needs to have following columns for the program to work:
"Lemma","Grundform" and	"Grammatik". It can have any number of additional columns.
After that choose what function you want to run (available are: "Ableitungen","Hapax Legomena", "Diminuitiva","Komposita","Letzte Lemma").
Choose a fitting file to store the results in (this will be an xlsx file).(make sure that you are not overwriting a file you currently have open, otherwise the program will fail.)
Wait a moment for the program to run, you will get a confirmation window once it is finished and it will appear unresponsive while it is working.
You can now close the program with "beenden" or choose another function.