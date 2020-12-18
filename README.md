# About
Send Outlook email w/covid-19 XLSX attachments, using New York Times as data source.
* Configurable spreadsheets complete with graphs 
    * Mix/match states & counties via json file
    * US summary by state
    * State summary by county 

My goal is to facilitate human analysis - so while emails are created automatically, 
they are not sent; the idea is the sender reviews covid-19 data and makes some sort of conclusion.

Otherwise...why generate spreadsheets at all?  If we aren't looking at data, thinking about what we see,
sharing our thoughts with others, it becomes a mindless data generation task; there are too many
spreadsheet generators, dumping data in a directory, never to be looked at, "just in case".
 
# Dependencies
* Git - https://git-scm.com/downloads  (for New York Times sync) 
* Python 3 -  https://www.python.org/downloads/ or https://www.microsoft.com/en-us/p/python-38/

# Options
* Microsoft Visual Studio Code - Simplest way to run ( https://code.visualstudio.com/Download )
* Microsoft Outlook - Optional, needed to send emails
* Microsoft Excel - Optional, for spreadsheets

# Windows Install & Execution (VS Code)Execution
1. Download & Install git 
2. Download & Install Python 3 from Windows App Store
3. Start Microsoft Visual Studio Code, install git / python extensions -- https://code.visualstudio.com/docs/python/python-tutorial#_prerequisites
4. Clone repo https://github.com/const-void/covid-19-spreadsheet: -- https://code.visualstudio.com/docs/editor/github#_setting-up-a-repository
5. Menu: Terminal, New Terminal --  https://code.visualstudio.com/docs/editor/integrated-terminal
6. In terminal, `~\AppData\Local\Microsoft\WindowsApps\pip.exe install -r requirements.txt`
7. File, Open `covid19_data_gather.py`
8. Menu: Run, Run without Debugging
9. File, Open `covid19_data_gather_conf.json`
10. Edit! Go to step 8. 

# Windows Install & Execution (Powershell)
1. Download & Install git 
2. Download & Install Python 3 from Windows App Store
3. Start Powershell
```
PS > cd Documents  (or any location)
PS > git clone https://github.com/const-void/covid-19-spreadsheet
PS > cd covid-19-spreadsheet
PS > ~\AppData\Local\Microsoft\WindowsApps\pip.exe install -r requirements.txt
PS > python .\covid19_data_gather.py
PS > python .\covid19_data_gather.py <path\to\conf.json>
PS > python .\covid19_data_gather.py sample_coivd19_data_gether_conf.json
```

# OSX/linux Install & execution
```
$ git clone https://github.com/const-void/covid-19-spreadsheet
$ cd covid19-19-spreadsheet
$ pip install -r requirements
$ python ./covid19_data_gather.py
$ python ./covid19_data_gather.py </path/to/conf.json>
$ python ./covid19_data_gather.py sample_covid19_data_gather_conf.json
```

# Input Data Definition

## all-geocodes-v2018.csv
To do

## co-est2019-annres.csv
To do

## nst-est2019-01.csv
To do

## state-geocodes-v2018.csv
To do

# JSON Configuration
To do

# Spreadsheet Notes
To do

# Email Notes
Email is a bit limited.  Currently - no osx/linux email, and Windows email is via Outlook.

Outlook is standard in the enterprise world, but not at home.  I haven't figured out Win10 mail yet, and I don't have a Mac/Linux box to use.  I am unwilling to smtp - so clients TBR.

# TO DO 
US Territories (Guam, Puerto Rico, etc)
Win10 Mail
Apple Mail
