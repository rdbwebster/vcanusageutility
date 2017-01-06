Overview
--------

This is a Excel automation program and opens an Excel Workbook with an existing sheet and
then merges in data from a SQL Server Database SQL Query 
and then does the following

- Adds the data to a new Sheet
- Adds the corresponding entry for each provider row to the original Sheet that was opened in a column named rebate
- Creates a number of Sums and Pivot table using Automation.


How to setup development environment
------------------------------------

Install Excel 2016

Install Visual Studio 2015
  https://www.microsoft.com/en-us/download/details.aspx?id=48146
  Custom install, default settings

Install Git Plugin for Visual Studio
  https://visualstudio.github.com/

Install Git command line
Team->settings->install 3rd party tools -> Web Platform Installer (WebPI
  for some reason goes to...
  https://git-scm.com/download/win
  exe installer installs git 2.11.0


Checkout code from this repo
  In visual studio Team-> Settings Git Global Settings
  make sure info is correct.
  
  Note: To avoid Visual Studio clone issues.  Remote already exists error.
  Verify contents of .gitconfig file (specificaylly that no value is set for [remote])
  .gitconfig located in windows based on  env variables 
  HOMEDRIVE is set to H:
  HOMEPATH is set to /

  In visual Studio Team 
  Login to Git using rdbwebster and creds + two form auth code.
  (be sure not to click on local git repo section)

  Then navigate to repo and clone.
  Then open solution.



Create odbc connection for C# connection to DB.
  Windows Control Panel - ODBC
  ODBC Connection Info - User DSN
                     - Add - SQL Server
                     - Name - store-dbrepo
                     - Description - 'blank'
                     - Server store-dbrepo1
                     - Check - With Windows NT Authentication uing the network login ID
                     - Check Connect to SWK Server to obtain default settings for the additional configuration Optinos
                     - Login ID - bwebster
                     - Password - 'blank'

                     - check - Use ANSI nuls, paddings and warnings.
                     - check - perform translatino for character data

Install Heidi
   http://www.heidisql.com/

Create Heidi connection
  New Connection
  New SQL SERVER tcp/ip connection   (not mysql)
  Server - store-dbrepo1
  Check - Use windows Authentication
  User - bwebster
  Port - 1433
  Database - Bookings 



Build in visio studio and run.

