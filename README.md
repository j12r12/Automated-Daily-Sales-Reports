# Automated Daily Sales Report

This project allowed me to work on my object-oriented programming with Python while delivering a useful report for daily use by my Senior Management Team and saving over 10 hours of work per week.

I have used three different bespoke classes:

Daily_Data reads the data from the .csv attachment sent over email and uses simple pivot tables to prepare the data to be graphed.

Graph takes this information and saves three figures for use in the Excel output file.

Excel brings both the data and graphs together to display the information in the format requested by my organisation. I used the very good xlsxwriter library to put this together.

This simple project solved a number of problems for my employers, in that it measured the daily sales completed by each salesperson, displayed this information in an easy-to-understand format and automated the whole process.

# Running My Code

To run this project, download all four python files to your working directory. The main.py file is the only one that needs to be run.

This project is quite specific to the needs I had at the time, however if required it is easy to remove the sections of code regarding receiving and sending emails. 

The data you want to review will have to be in a .csv format with headings including Salesperson, Opportunity Type, Stage and Close Date. However these can obviously be easily changed in the code to suit your needs. 
