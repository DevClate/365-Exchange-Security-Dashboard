# 365-Exchange-Security-Dashboard

A simple Excel Dashboard for Exchange Security to give you a quick view into the main data points from Exchange Security. These data points are all in the exchange portal, but they are all seperated out. Furthermore, between 90-180 days depending on the data point, the data is removed so there is no long term reporting.

## Current Status

I have created the functions for 365 Security Dashboard v1 and working on the design of the dashboard.

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
* [ ] Email PDF version
* [ ] Outputs

  * [X] Only Excel Dashboard
  * [X] Excel Dashboard and seperate raw file
  * [ ] Only write to SQL
  * [ ] PDF File
  * [ ] Email PDF

## Change Log

Dec 11, 2022

* Uploaded v1 Design with and without separate raw file
* Added functionality so that when you run in it no matter what day of the month, it will pull the previous month information and have everything labeled correctly
