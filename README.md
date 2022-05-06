# Daily-Sales

Is a program I made that crawls the archived CRM of the company I am working at.

It was created to solve an issue the CRM that was not able to clearly show the daily sales of the company.

In order to solve that I created this program that gets an .xlsx file as an input downloaded from XERO and after a few transformations on
the .xlsx file creates a list of all the new sales that happened the previous day.

It then takes the list and crawls the CRM by going to all the client from the previous day and extracts cetrain details in order to create the report.

The time complexity is On^2 it is a very slow program to complete, but at the time the company was using an external IT company that wouldnt share the BD loggings or fix the bugs.
