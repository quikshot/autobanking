# autobanking

Automatic banking to keep track of your expenses

## Objectives

* Small project to keep online banking expenses under control.

* Automatizes the access to online banking and keeps your expenses data locally in a database.

* Requests the status of all your online bank accounts and saves it in an excel automatically.

* This Excel will be parsed and categorised and inserted into a SQLite DB.

* Weekly reports sent to e-mail and specific analysis available.

## Status
### What is currently done:
* Access Spanish Kutxabank online bank: enter ID and password from config file. 
* Select desired bank account (pending to use XPATH to parse Account number)
* Download data between two dates into /tmp

### TO-DO list:
* Refactor and organise code.
* Use Python-Excel library to read excel data.
* Insert this data into SQLite db
* Reparse SQLite data to categorize it into topics: household, transport, food, ...
* Automatic categorization by requesting the user
* Create expenses reports:
   * Current status
   * Weekly expenses, comparing to previous weeks
   * Monthly expenses, comparing to previous months. show trend
   * Evolution of categories in time
* Add ING bank

### In the future:
* Create all above for other bank provider and join data in one place
* Multiplatform packaging (windows, linux, mac) with Python for easy install and config.
* Docker server running service
* Small app with frontend info and charts


## Getting started

### Prerequisites

Yo need Python 2.7 and the following packages:
```
sudo apt-get install python
sudo apt-get install pip
sudo apt-get install xvfb
#Browser automation
pip install selenium
#Python image processing:
pip install Pillow
#SQL Lite database:
sudo apt-get install sqlite3
#SQLite bindings for python
#Python-excel library
```

download geckodriver and put it in your path

### Installation

clone or download the repository

Copy configexample.ini to config.ini and put your ID/code to access the Spanish ING personal account.

### Usage

To test it, run:

```
python ingdirect.es.py
```

This will open firefox and you will see the execution.
If something fails you will see it in the exceptions.


### Disclamer:
This is a work in progress, use it at your own risk.


### Author:
quikshot

### Acknowledgements
Roger Barnes for inspiration on its Australian ING online access:
https://gist.github.com/mindsocket/3015852


