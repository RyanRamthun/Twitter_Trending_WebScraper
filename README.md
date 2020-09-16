# Python Twitter Web Scraper

A Python web scraper using Selenium and Openpyxl that takes the top 10 current tweets and adds them to a new Excel spreadsheet.
<br />
## Functionality:

* Automatically scrapes the top 10 current trending tweets 
* Organizes raw data
* Exports organized data to an excel spreadsheet using openpyxl


# Building
## Requirements.txt
```
selenium==3.141.0
openpyxl==3.0.5
```
## Installation
```python
$ pip install selenium
$ pip install openpyxl
```
# Usage
Running Selenium will open the chrome web driver. The program will then wait for the correct HTML element to be loaded. When the element is located, the data will be scraped into an array and exported  to a new/existing excel file in the project directory named twitter_trends using openpyxl.
