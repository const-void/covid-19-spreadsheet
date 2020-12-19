# About
Send Outlook email w/covid-19 XLSX attachments, using New York Times as data source.
* Configurable spreadsheets complete with graphs 
    * Mix/match states & counties via json file
    * US summary by state
    * State summary by county 

# Quick install & execution
```
$ git clone https://github.com/const-void/covid-19-spreadsheet
$ cd covid19-19-spreadsheet
$ pip install -r requirements
$ python ./covid19_data_gather.py
$ python ./covid19_data_gather.py </path/to/conf.json>
$ python ./covid19_data_gather.py sample_covid19_data_gather_conf.json
```

# Goal
I was pretty frusterated by how hard it was to answer simple questions--how many people HAVE covid, right now?  How does where I live, compare to where others live?  Are we doing ok, as a county, state, country? When I found out about the NYT data store, I jumped on it IMMEDIATELY.

I also wanted to facilitate  analysis - so while emails are created automatically, they are not sent; the idea is, I, as a sender review, covid-19 data and makes some sort of conclusion.

Otherwise...why generate spreadsheets at all?  If we aren't looking at data, thinking about what we see, sharing our thoughts with others, it becomes a mindless data generation task; there are too many spreadsheet generators, dumping data in a directory, never to be looked at, "just in case".
 
# Dependencies
* Git - https://git-scm.com/downloads  (for New York Times sync) 
* Python 3 -  https://www.python.org/downloads/ or https://www.microsoft.com/en-us/p/python-38/
* pip c/o [requirement.txt](https://pip.pypa.io/en/stable/user_guide/#requirements-files):
    * openpyxl - native spreadsheet generation
    * gitpython - nyt coid-19-data git repo clone / pull
    * fastjsonschema - json configuration file  
    * pywin32 - Outlook email

# Optional Packages
* Microsoft Visual Studio Code - IDE ( https://code.visualstudio.com/Download )
* Microsoft Outlook - Send emails
* Microsoft Excel - Viewing generated spreadsheets

# Windows Install & Execution (VS Code)Execution
1. Download & Install git 
2. Download & Install Python 3 from Windows App Store
3. Start Microsoft Visual Studio Code, [install git / python extensions](https://code.visualstudio.com/docs/python/python-tutorial#_prerequisites)
4. [Clone repo](https://code.visualstudio.com/docs/editor/github#_setting-up-a-repository) https://github.com/const-void/covid-19-spreadsheet: -- 
5. Menu: Terminal, [New Terminal](https://code.visualstudio.com/docs/editor/integrated-terminal)
6. In terminal, `~\AppData\Local\Microsoft\WindowsApps\pip.exe install -r requirements.txt`
7. File, Open `covid19_data_gather.py`
8. Menu: Run, Run without Debugging
9. File, Open `covid19_data_gather_conf.json`
10. Edit! Go to step 8. 

# Windows Install & Execution (Powershell)
1. Download & Install git 
2. Download & Install Python 3 from Windows App Store
3. Start Powershell
*One time install*
```
PS > cd Documents  (or any location)
PS > git clone https://github.com/const-void/covid-19-spreadsheet
PS > cd covid-19-spreadsheet
PS > ~\AppData\Local\Microsoft\WindowsApps\pip.exe install -r requirements.txt
```

*Execution*
```
PS > python .\covid19_data_gather.py
PS > python .\covid19_data_gather.py <path\to\conf.json>
PS > python .\covid19_data_gather.py sample_coivd19_data_gether_conf.json
```

# JSON Configuration
Case sensitive!

`cov19_data_gather.py` without a command line parameter will look for a `covid19_data_gather_conf.json` file. 

If one isn't found, `sample_covid19_data_gather_conf.json` is used as starting place. 

Note that `covid19_data_gather_conf.json` is in `.gitignore` -- this way you can have your own setup without having to worry about git collisions etc.  Alternatively, a given `/path/to/conf.json` can specified on the command line.

All json configurations are validated -- once at a schema level via `covid19_data_gather_conf.schema.json`, and again, to make sure the intended geographies are accurate - counties have to match both US & NYT data *exactly*, including case sensitivity.  

`los angeles, ca` will *fail* validation.  `Los Angeles County, CA` will pass.

There are two blocks, both required:
* `"spreadsheets": {}`
* `"settings": {}`

## `"spreadsheets": {}`
This block controls spreadsheet generation. Fiddle with these settings to hone in on geographies of interest. If an email will be sent, each spreadsheet generated is attached.

For a while, I was generating spreadsheets and not sending...but...what's the point?  The data isn't going anywhere.  When we want to look at a region -- add it in! When we tire of it...take it out!

key | type | required | desc | example
--- | ---- | -------- | ---- | -------
us  | Boolean | Yes | Controls generation of summary (US) level spreadsheet, consisting of all 50 states! | `"us": true`
state-detail | Array of states | Yes | Controls generation of summary (state) level spreadsheets, consisting of each county in a given state.  One spreadsheet per state is generated. | `"state-detail": ['ND', 'SD']`
custom | List | Yes | List of spreadsheets to generate. each property is a  spreadsheet name; property value is a state or a county, state.  One spreadsheet per property is generated. | `"custom": { "north_dakota": [ "Burleigh County, ND", "ND" ] }`

## "settings": {}
This configuration block controls the script itself.  Fiddle with these settings to change the data we see.  Don't like my take of a 28 day average case...lower it.  Or, raise it.  what happens?   Want to change the comparative per scales -- is 100k too big? Too small?  Or do we want to exclude vast swathes of the country?  The below settings allow you to do *just* that--manipulate the data as *you* see fit.

*Data Controls*
key | type | required | desc | example
--- | ---- | -------- | ---- | -------
case-min-benchmark | Number | Yes | Minimum number of cases; acts a a reporting gate. If we want to eliminate low caseload geographies, we set this property to filter to just the caseloads we are interested in--say those at 100,000 or more, or even--minimum of 10, 1000.  | `"case-min-benchmark": 1`
case-days-duration | Number | Yes | Average case duration - used in active vs recovered calculations. | `"case-days-duration": 28`
geography-per-county | Number | Yes | Scaling factor for counties. Per capita is a value of one; cdc uses 100k. | `"geography-per-county": 100000`
geography-per-state | Number | Yes |  Scaling factor for states. Per capita is a value of one; cdc uses 100k. | `"geography-per-state": 100000`

*Email Settings*
key | type | required | desc | example
--- | ---- | -------- | ---- | -------
send-email | Boolean | Yes | Send email if true | `"send-email": true`
send-email-client | Enumeration | No | One of `Outlook` or `N/A`, sadly. | `"send-email-client": "Outlook"`
send-email-to | Array of emails | No | List of email addresses to send to. | `"send-email-to": [ "a@bc.com" ]`
send-email-style | String | No | HTML styling for a swank email. | `"send-email-style": "font-family: Trebuchet MS; color:#25253b; font-size:14pt"`
send-email-greeting | String | No | HTML email greeting | `"send-email-greeting": "Hello!<br>"`
send-email-signature | String | No | HTML signaure | `"send-email-signature": "xoxo<br>Yours Truly!"`

# Spreadsheet Notes
To do

# Covid-19 Data
c/o The New York Times. (2020). Coronavirus (Covid-19) Data in the United States. Retrieved from https://github.com/nytimes/covid-19-data.

*Usage*
* The function `update_data` sync's Covid-19 via `git pull` into a sibiling directory c/o https://github.com/nytimes/covid-19-data :
```
covid-19-data/
  us-counties.csv
covid-19-spreadsheet/
  xlsx/
     <generated ouutputs>
```
* If the sibiling directory/repository is missing, it is created via a `git clone`
* The function `set_county_covid19_cases` loads covid-19 data into Counties, joining via `fips`. 

Col #  | Field Name  | Desc | Sample 
------ | ----------- | ---- | ------
1 | Date | | `2020-01-21`
2 | County | | `Snohomish`
3 | State | | `Washington`
4 | Fips | State FIPS + County FIPS | `53061`
5 | Cases | | `1`
6 | Deaths | | `0`

# Geography Data
When NYT started reporting on covid-19, each day would introduce a slew of new US geographies.  Initially, I thought maybe I could pull in geography data as it was found--but, as data came in, the performance impact just got to be too great. It is important to cache geograpy in adavnce - so that as new geography data comes pouring in, it has a place to go.

From [Census.gov](https://www.census.gov/geographies/reference-files/2018/demo/popest/2018-fips.html):
## [all-geocodes-v2018.csv](https://github.com/const-void/covid-19-spreadsheet/blob/master/all-geocodes-v2018.csv)
[spreadsheet](https://www2.census.gov/programs-surveys/popest/geographies/2018/all-geocodes-v2018.xlsx) => csv.
**Country Data**
* Estimates Geography File: Vintage 2018
* Source: U.S. Census Bureau, Population Division
* Internet Release Date: May 2019

*Usage*
* Loaded by `Counties` constructor to create individual `County` objects.
* Used to join County to State **( State Code (FIPS) )**
* Used to join NYT Covid-19 data to County **( NYT FIPS Code = State Code (FIPS) + County Code (FIPS) )**

Col #  | Field Name  | Desc | Sample 
------ | ----------- | ---- | ------
1 | Summary Level | | `050`
2 | State Code (FIPS) | | `01`
3 | County Code (FIPS) | | `001`
4 | County Subdivsion Code (FIPS) | | `00000`
5 | Place Code (FIPS) | | `00000`
6 | Consolidated City Code (FIPS) | | `00000`
7 | Area Name | | `Autauga County`

## [state-geocodes-v2018.csv](https://github.com/const-void/covid-19-spreadsheet/blob/master/state-geocodes-v2018.csv)
**State Data**
[spreadsheet](https://www2.census.gov/programs-surveys/popest/geographies/2018/state-geocodes-v2018.xlsx) => csv
* Source: U.S. Census Bureau, Population Division
* Internet Release Date: May 2019

*Usage*
* Loaded by `States` constructor to create individual `State` objects.
* Used to join State to County  **( State Code (FIPS) )**

Col # | Field Name | Desc | Sample
----- | ---------- | ---- | ------
1 | Region  | | `1`
2 | Division | | `1`
3 | State (FIPS ) | | `09`
4 | Name | | `Connecticut`

# Population
Simply knowing geography and case wasn't enough--I wanted to know the sense of scale, and not 
just within a geography...but also *across* geographies.  How do counties compare to other counties? And states?

## co-est2019-annres.csv
**County Population Estimates**
(src)[https://www.census.gov/newsroom/press-kits/2020/pop-estimates-county-metro.html] [spreadsheet](https://www2.census.gov/programs-surveys/popest/tables/2010-2019/counties/totals/co-est2019-annres.xlsx) => csv
* Annual Estimates of the Resident Population for Counties in the United States: April 1, 2010 to July 1, 2019 (CO-EST2019-ANNRES)
* Source: U.S. Census Bureau, Population Division	
* Release Date: March 2020	

*Usage*
* Loaded by `set_county_population` function
* Joins to county & state via name
							
Col # | Field Name | Desc | Sample
----- | ---------- | ---- | ------
1 | Geographic Area | | `".Autauga County, Alabama"`
2 | Census | | `54571`
3 | Estimates Base | | `54597`
4 | 2010 | | `54773`
5 | 2011 | | `55227`
6 | 2012 | | `54954`
7 | 2013 | | `54727`
8 | 2014 | | `54893`
9 | 2015 | | `54864`
10 | 2016 | | `55243`
11 | 2017 | | `55390`
12 | 2018 | | `55533`
13 | 2019 | | `55869`

## nst-est2019-01.csv
**State Population Estimates**
[src](https://www.census.gov/data/tables/time-series/demo/popest/2010s-national-total.html
) [spreadsheet](https://www2.census.gov/programs-surveys/popest/tables/2010-2019/state/totals/nst-est2019-01.xlsx) => csv
* Table 1. Annual Estimates of the Resident Population for the United States, Regions, States, and Puerto Rico: April 1, 2010 to July 1, 2019 (NST-EST2019-01)												
* Source: U.S. Census Bureau, Population Division												
* Release Date: December 2019	

*Usage* 					
* Loaded by `States` contstructor to assign population to each U.S state
* State name joins population to state

Col # | Field Name | Desc | Sample
----- | ---------- | ---- | ------
1 | Geographic Area | | `Alabama`
2 | Census | | `4779736`
3 | Estimates Base | | `4780125`
4 | 2010 | | `4785437`
5 | 2011 | | `4799069`
6 | 2012 | | `4815588`
7 | 2013 | | `4830081`
8 | 2014 | | `4841799`
9 | 2015 | | `4852347`
10 | 2016 | | `4863525`
11 | 2017 | | `48744861`
12 | 2018 | | `4887681`
13 | 2019 | | `4903185`

# Email Notes
Email is a bit limited.  Currently - no osx/linux email, and Windows email is via Outlook.

Outlook is standard in the enterprise world, but not at home.  I haven't figured out Win10 mail yet, and I don't have a Mac/Linux box to use.  I am unwilling to smtp - so clients TBR.

# TO DO 
* US Territories (Guam, Puerto Rico, etc)
* Win10 Mail
* Apple Mail
