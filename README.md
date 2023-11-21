<h1>Creating a Change Of Address Report by Identify Matching Ownership Records within CAMA Database</h1>

<h2>Tools Used</h2>

- <b>SQL</b>
- <b>Python</b>
- <b>PowerShell</b>
- <b>SSMS (Scheduled Job)</b>
- <b>Windows Task Manager</b>

<h2>Description</h2>

<b> Problem: </b> A Change of Address Report is recieved from the Clerk of Court every month listing jurors who have updated their mailing addresses. Every record on the report is manually searched within the CAMA application one at a time to see if the person listed owns property or if the address listed has any exemptions within the CAMA database. This search process is tedious and time consuming.
<br><br>
 <b> Solution: </b> Automate the search process by comparing the data on the original report to data within the CAMA database and produce any matching results in a new report.
 <br><br>
<b> Quick Overview:  </b>
 
  - <b>Step 1:</b> The Change of Address Report is recieved via email as an Excel file and placed into a network directory.
  - <b>Step 2:</b> A scheduled task runs biweekly on a VM executing a PowerShell script that checks if the Excel file is present in the network directory. If the Excel file is present, the PowerShell script kicks off the Python script.
  - <b>Step 3:</b> The Python script reads the data in the Excel file to a Pandas dataframe & filters out any unwanted columns. Then the data is further extracted and manipulated so that the formatted names and addresses are broken down into additional fields that are more ideal for searching. For Example The "FullName" field is broken out into First, Middle, & Last Name fields. After the data has been extracted the script exports the dataframe as a csv file and deletes the original Excel file.
  - <b>Step 4:</b> A scheduled job runs biweekly via SSMS that executes a pair of stored procedures. The first imports the data from the csv file into a table. The second processes the data in the table by looping through one record at a time finding any matching records in the CAMA database and producing a report with the results. The final report is then exported as a .xlsx file and saved to a network directory and an email notification is sent out stating the report for that month has been processed and is ready for review.

<p align="center">
<img src="https://i.imgur.com/faFb5zY.png" height="75%" width="75%" alt="CoA Process Flowchart"/>
</p>

<h2>Screenshots</h2>
*** For the sake of security, any email addresses, network paths, and anything deemed potentially sensitive will be removed from production code & screenshots *** .
<br />

<h3>Original Excel File from Clerk of Court</h3>
<p align="center">
<img src="https://i.imgur.com/Loycsjm.png" height="95%" width="95%" alt="CoC CoA Excel File"/>
</p>

<h3>Data Imported from CSV into Table</h3>
<p align="center">
<img src="https://i.imgur.com/nwde5Hj.png" height="85%" width="85%" alt="Extracted Data in Table"/>
</p>

<h3>Final Excel Report</h3>
<p align="center">
<img src="https://i.imgur.com/cp8fxkt.png" height="85%" width="85%" alt="Final Excel Report"/>
</p>

<h3>Email Notification</h3>
<p align="center">
<img src="https://i.imgur.com/qfY9Qny.png" height="85%" width="85%" alt="Email Notification"/>
</p>

<h3>Report Statistics</h3>
<p align="center">
<img src="https://i.imgur.com/aYF2Zk2.png" height="85%" width="85%" alt="Report Stats"/>
</p>

<h2>The Good Stuff</h2>

The following items are present in the python code involved:

- Pandas
- Logging
- Try-Except Error Handling
- If / Else Logic

The following items are present in the SQL stored procedure involved:

- Dynamic SQL
- While Loop
- Try-Catch Error Handling
- If / Else Logic
- Update / Insert
- Case Statements
- #Temp Tables
- Window Functions
- Pivot
- Pat Index
- Table Variable

Links to SQL scripts involved in this process:
- [Exception Handling Table & Stored Procedure](https://github.com/Deltron2020/ExceptionHandling)
- [Does File Exist Function](https://github.com/Deltron2020/doesFileExist)
- [Export Data to CSV](https://github.com/Deltron2020/ExportDataToCsv)
- [CSV to Excel File wTable](https://github.com/Deltron2020/CSVtoXLSXwTable)

<!--
 ```diff
- text in red
+ text in green
! text in orange
# text in gray
@@ text in purple (and bold)@@
```
--!>
