# BruteForceEmailer
Emailer for TFX request authorization

This Powershell script uses winscp script to download files from an ftp server, using files called Get-Data.cmd and get-data.txt (txt file has not connection chain per security),
but format can be found on winscp support web page, after that clear content on files used along this process, from line 7 to 98 it loads data from csv files on arrays, 
next version will use a funtion to reduce lines of code, this example loads data from 3 files so regarding performace it is ok

From line 99 to 111 prepare objects that will no change and are necessary for send email as sender, smtp server and credentials to log in powershell to office 365

next part is explained on script it self and it format data and send emails to approvers and errors to service desk 

List of files to run this script


123.txt
body.txt
crqsnotfound.txt
emailerlog.txt
error.txt
get_data.txt
get-data.cmd
requester.txt
