# About
Send Outlook email w/covid-19 XLSX attachments, using New York Times as data source.
* Configurable spreadsheets complete with graphs 
    * Mix/match states & counties via json file
    * US summary by state
    * State summary by county 

# Dependencies
* Git - https://git-scm.com/downloads  (for New York Times sync) 
* Python 3 -  https://www.python.org/downloads/ or https://www.microsoft.com/en-us/p/python-38/

# Options
* Microsoft Visual Studio Code - Simplest way to run
* Microsoft Outlook - Optional, needed to send emails
* Microsoft Excel - Optional, for spreadsheets

# VS Code Install & Execution

# Windows Install & Execution (Powershell)
1. Download & Install git & Python 3
2. Start Powershell
```
PS > cd Documents  (or any location)
PS > git clone https://github.com/const-void/covid-19-spreadsheet
PS > cd 
PS > ~\AppData\Local\Microsoft\WindowsApps\pip.exe install -r requirements.txt
PS > python .\covid19_data_gather.py
PS > python .\covid19_data_gather.py <path\to\conf.json>
PS > python .\covid19_data_gather.py sample_coivd19_data_gether_conf.json
```

# OSX/linux Install & execution
```
git clone
cd 
python ./covid19_data_gather.py
python ./covid19_data_gather.py </path/to/conf.json>
python ./covid19_data_gather.py sample_covid19_data_gather_conf.json
pip install -r requirements
```

# JSON Configuration


# Spreadsheet Notes

# Email Notes
Email is a bit limited.  Currently - no osx/linux email, and Windows email is via Outlook.

Outlook is standard in the enterprise world, but not at home.  I haven't figured out Win10 mail yet, and I don't have a Mac/Linux box to use. 

# TO DO 
US Territories (Guam, Puerto Rico, etc)
Win10 Mail
Apple Mail
