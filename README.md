# 365ExchangeSecurityDashboard

A simple Excel Dashboard for Exchange Security to give you a quick view into the main data points from Exchange Security. These data points are all in the exchange portal, but they are all seperated out. Furthermore, between 90-180 days depending on the data point, the data is removed so there is no long term reporting.

## Current Status

Cleaning up some of the code for the main report. Created the functions, but need to add help to them. 

## Next Steps

* [X] Finish importing last bits of data

  * [X] Top Malware
* [X] Create Functions v1
* [X] Finalize v1 Design
* [X] Create Design
* [X] Each month it will add a new sheet to existing file
* [ ] Create Module

## Future Features

* [ ] More Data Points
* [ ] V2 Design with charts
* [ ] Run as a schedule task on the first of the month it will pull last month's data
* [ ] Pull Data into SQL for long term reporting and/or for large organizations
* [ ] Add Quarterly Summary
* [ ] Outputs

  * [X] Only Excel Dashboard
  * [X] Excel Dashboard and seperate raw file
  * [ ] Function to only write to SQL
  * [ ] Function to only export PDF File
  * [ ] Function to Email PDF

## Change Log

Dec 11, 2022

* Uploaded v1 Design with and without separate raw file
* Added functionality so that when you run in it no matter what day of the month, it will pull the previous month information and have everything labeled correctly
* Added help to functions. Although those functions aren't required to run the script, I did use them when making the script originally. Allows you to customize it how you want too.
