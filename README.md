<h1>Creating a Change Of Address Report by Identify Matching Ownership Records within CAMA Database</h1>

<h2>Tools Used</h2>

- <b>SQL</b>
- <b>Python</b>
- <b>PowerShell</b>
- <b>SSMS (Scheduled Job)</b>
- <b>Windows Task Manager</b>

<h2>Description</h2>

<b> Problem: </b> A Change of Address Report is recieved from the Clerk of Court every month listing people who have updated their mailing addresses. Every record on the report is manually searched within the CAMA application one at a time to see if the person listed owns property or if the address listed has any exemptions within the CAMA database. This search process is tedious and time consuming.
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
<img src="https://i.imgur.com/zN8izXm.png" height="95%" width="95%" alt="XY Excel File"/>
</p>

<h3>Windows Task</h3>
<p align="center">
<img src="https://i.imgur.com/X0g4X4p.png" height="85%" width="85%" alt="XY Excel File"/>
</p>

<h3>SSMS Job</h3>
<p align="center">
<img src="https://i.imgur.com/Nw1ISAk.png" height="85%" width="85%" alt="XY Excel File"/>
</p>

<h3>CSV to CAMA</h3>
<p align="center">
<img src="https://i.imgur.com/zPx0t5i.png" height="85%" width="85%" alt="XY Excel File"/>
</p>

<h3>Email Notification</h3>
<p align="center">
<img src="https://i.imgur.com/xm0u7dn.png" height="85%" width="85%" alt="XY Excel File"/>
</p>


<h2>The Good SQL Stuff</h2>

The following items are present in the stored procedure [sp_ImportXYCoordinates](https://github.com/Deltron2020/XYCoordinateImport/blob/main/sp_ImportXYCoordinates.sql):
- Dynamic SQL
- Try-Catch Error Handling
- If / Else Logic
- Update / Insert
- #Temp Tables
- Commit / Rollback Transactions
- Data Validation
- Send DB Mail

Links to SQL scripts involved in this process:
- [Exception Handling Table & Stored Procedure](https://github.com/Deltron2020/ExceptionHandling)
- [Does File Exist Function](https://github.com/Deltron2020/doesFileExist)

<!--
 ```diff
- text in red
+ text in green
! text in orange
# text in gray
@@ text in purple (and bold)@@
```
--!>
