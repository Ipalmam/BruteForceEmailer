C:\CRQ_Reminder\Get-Data.cmd
Clear-content -Path "123.txt" -Force
Clear-content -Path "Error.txt" -Force
Clear-content -Path "Requester.txt" -Force
Clear-content -Path "EmailerLog.txt" -Force
Clear-content -Path "CRQsNotFound.txt" -Force
$FilePath = "P1.xlsx"
$SheetName1 = "P1"
$objExcel1 = New-Object -ComObject Excel.Application
$WorkBook1 = $objExcel1.Workbooks.Open($FilePath) <#Loads excel file in a object#>
$WorkSheet1 = $WorkBook1.sheets.item($SheetName1)
$xlCellTypeLastCell = 11
$endRow1 = $WorkSheet1.UsedRange.SpecialCells($xlCellTypeLastCell).Row
$startRow1 = 5
$col1 = 2
$IncidentIDs = @()
for ($j = 0; $j -le $endRow1; $j++)    <#Saves Incidents col in an array#>
{
  	$IncidentIDs += $WorkSheet1.Cells.Item($startRow1 + $j , $col1).Value2
}


$IncidentIDs = $IncidentIDs | Where-Object {-not [string]::IsNullOrWhiteSpace($_)}    ##Clean data, delete empty spaces 
$startRow2 = 5
$col2 = 5
$Requesters = @()
for ($j = 0; $j -le $endRow1; $j++) <#Saves Task Name col in an array#>
{
   	$Requesters += $WorkSheet1.Cells.Item($startRow2 + $j , $col2).Value2
}


$Requesters = $Requesters | Where-Object {-not [string]::IsNullOrWhiteSpace($_)}
$startRow3 = 5
$col3 = 7
$Requests1 = @()
for ($j = 0; $j -le $endRow1; $j++)      <#Saves Summary col in an array#>
{
   	$Requests1 += $WorkSheet1.Cells.Item($startRow3 + $j , $col3).Value2
}


$Requests1 = $Requests1 | Where-Object {-not [string]::IsNullOrWhiteSpace($_)}
$objExcel1.Workbooks.Close()                                                                                                        <#Close excel document#>
$FilePath2 = "P2.xlsx"
$SheetName2 = "P2"
$objExcel2 = New-Object -ComObject Excel.Application
$WorkBook2 = $objExcel2.Workbooks.Open($FilePath2)     <#Loads excel file in a object#>
$WorkSheet2 = $WorkBook2.sheets.item($SheetName2)
$endRow2 = $WorkSheet2.UsedRange.SpecialCells($xlCellTypeLastCell).Row
$startRow4 = 5
$col4 = 3
$Summaries = @()
for ($j = 0; $j -le $endRow2; $j++)   <#Saves Summary col in an array#>
{
   	$Summaries += $WorkSheet2.Cells.Item($startRow4 + $j , $col4).Value2
}


$Summaries = $Summaries | Where-Object {-not [string]::IsNullOrWhiteSpace($_)}
$startRow5 = 5
$col5 = 1
$Requests2 = @()
for ($j = 0; $j -le $endRow2; $j++)   <#Saves Summary col in an array#>
{
   	$Requests2 += $WorkSheet2.Cells.Item($startRow5 + $j , $col5).Value2
}


$Requests2 = $Requests2 | Where-Object {-not [string]::IsNullOrWhiteSpace($_)}
$objExcel2.Workbooks.Close()                                                                                                            <#Close excel document#>
$FilePath3 = "P3.xlsx"
$SheetName3 = "P3"
$objExcel3 = New-Object -ComObject Excel.Application
$WorkBook3 = $objExcel3.Workbooks.Open($FilePath3)     <#Loads excel file in a object#>
$WorkSheet3 = $WorkBook3.sheets.item($SheetName3)
$endRow2 = $WorkSheet3.UsedRange.SpecialCells($xlCellTypeLastCell).Row
$startRow6 = 5
$col6 = 1
$Requests3 = @()
for ($j = 0; $j -le $endRow2; $j++)   <#Saves Summary col in an array#>
{
   	$Requests3 += $WorkSheet3.Cells.Item($startRow6 + $j , $col6).Value2
}


$Requests3 = $Requests3 | Where-Object {-not [string]::IsNullOrWhiteSpace($_)}
$startRow7 = 5
$col7 = 2
$Approvers = @()
for ($j = 0; $j -le $endRow2; $j++)    <#Saves Summary col in an array#>
{
   	$Approvers += $WorkSheet3.Cells.Item($startRow7 + $j , $col7).Value2
}


$Approvers = $Approvers | Where-Object {-not [string]::IsNullOrWhiteSpace($_)}
$objExcel3.Workbooks.Close()       <#Close excel document#>
$From = Get-content "from.txt"
$Pw = Get-content "PASSWORD.txt"
$PWord = ConvertTo-SecureString -String $Pw -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $From,$PWord
$SMTPServer = Get-content "smtpserver.txt"
$AppoverEmail = $null
$to = "123"
$RequesterEmailAddress = "123"
$UPNs = $null
$Cc = $null
$Bcc = "bcc.txt"
$EmailSent = "789"
$FoundCRQ = "false"
for($i = 0; $i -lt $Requests1.Length; $i++){
    for($j = 0; $j -lt $Requests2.Length; $j++){
        if($Requests1[$i] -contains $Requests2[$j]){
            for($k = 0; $k -lt $Requests3.Length; $k++){
                if($Requests1[$i] -contains $Requests3[$k]){
                    $Body = get-content "Body.txt"                        ##here we start to personalize each body email
                    $Body = $Body -replace "aaaaaaaaaa",$Requests1[$i]    ##add requester's name
                    $array = $Requesters[$i].split(" ")					  ##starting to format requester name 
                    for($l = 0; $l -lt $array.length; $l++){
                            $array[$l] = $array[$l].substring(0,1).toupper()+$array[$l].substring(1).tolower()
                        }
                    $RequesterName = "123"
                    for($l = 0; $l -lt $array.length; $l++){
                        if($RequesterName -contains "123"){
                            $RequesterName = $array[$l] + " "
                        }
                        else {
                            $RequesterName = $RequesterName + $array[$l]
                        }
                    }
                    $Body = $Body -replace "bbbbbbbbbb",$RequesterName
                    $Body = $Body -replace "cccccccccc",$Summaries[$j]
                    $RequesterEmail = $Requesters[$i].split(' ')
                    [array]::Reverse($RequesterEmail)
                    for($l = 0; $l -lt $RequesterEmail.Length; $l++){
						if($RequesterEmailAddress -contains "123")
						{
							$RequesterEmailAddress = $RequesterEmail[$l] + ", "
						}
						else
						{
							$RequesterEmailAddress = $RequesterEmailAddress + $RequesterEmail[$l]
						}
                    }
                    try{$Cc = get-aduser -filter 'Name -like $RequesterEmailAddress' | Select-Object UserPrincipalName} Catch{$_ | Out-File 123.txt -Append} ##searching aduser by email
					$RequesterEmailAddress = "123"
                    if($null -eq $Cc){                                  ##add erroe message to log in case user was not found on AD
						Clear-content -Path "123.txt"
						$message = $Requesters[$i] + " User email address don't found"
						Add-content  -Path "Error.txt" $message
					}
					else{												##modifying requester names as i didn't use json but txt files and CC them in email
						Add-content -Path "Requester.txt" $Cc
						$Cc = Get-content -Path "Requester.txt"
						Clear-Content -Path "Requester.txt" -Force
						$array = $Cc.split("=")
						$value = $array.length
						$value = $value -1
						$UPNs = $array[$value]
						$UPNs = $UPNs.Substring(0,$UPNs.Length-1)
                    }
                    $Cc = $UPNs																																								##obtener correo de requester
					$Subject = $Requests1[$i] + " is pending for your approval"   ##personalizing subject for each email
					$ApproversList = $Approvers[$k].split(';')
                    for($l = 0; $l -lt $ApproversList.Length; $l++){
						try{$AppoverEmail = get-ADUser -identity $ApproversList[$l] | select-object UserPrincipalName} Catch{$_ | Out-File 123.txt -Append}##valitaind email addresses on AD
						if($null -eq $AppoverEmail)
						{
							Clear-content -Path "123.txt"
							$message = $Requests1[$i] + $ApproversList[$j] +" Approver email is not correct, check active directory"   ##saing error message on error log file
							Add-content  -Path "Error.txt" $message
						}
						else
						{
							Add-content -Path "Requester.txt" $AppoverEmail
							$AppoverEmail = Get-content -Path "Requester.txt"
							Clear-Content -Path "Requester.txt" -Force
							$array = $AppoverEmail.split("=")
							$value = $array.length
							$value = $value -1
							$UPNs = $array[$value]
							$UPNs = $UPNs.Substring(0,$UPNs.Length-1)
							$AppoverEmail = $UPNs
							if($to -contains "123")		##validating appprover email 
							{
								$To = $AppoverEmail
							}
							else
							{
								$To = $To + ";" + $AppoverEmail
							}
						}	
					}			##from here we start to create a log from emails sent
					Add-content -Path "EmailerLog.txt" -value "`n`n###############################################################`n`nSending email with next information...`n`n"
					$arrayTo = $to.split(";")
					$To = "123"
					Add-content -Path "EmailerLog.txt" -value "List of approvers..."        ##Starts to save log
					Add-content -Path "EmailerLog.txt" $arrayTo
					Add-content -Path "EmailerLog.txt" -value "`n`nEmail subject		", $Subject
                    Add-content -Path "EmailerLog.txt" -value "`n`nRequester email		", $Cc, "`n`n"
					Add-content -Path "EmailerLog.txt" $Body
					if($null -ne $Cc)      ##if all data required to send email is correct from here email is sent
					{
						try{Send-MailMessage -From $From -To $arrayTo -Cc ($Cc | Out-String) -Bcc $Bcc -Credential $Credential -SmtpServer $SMTPServer -UseSsl -Port 587 -Subject $Subject -Priority High -Attachments C:\CRQ_Reminder\Instructions.pdf -Body ($Body | Out-String)} Catch{$_ | Out-File C:\CRQ_Reminder\Errors.txt -Append}
						Clear-content -Path "Errors.txt"
						$EmailSent = "123"
					}
					else     ##if all data required to send email is correct except requester email from here email is sent
					{
						try{Send-MailMessage -From $From -To $arrayTo -Credential $Credential -Bcc $Bcc -SmtpServer $SMTPServer -UseSsl -Port 587 -Subject $Subject -Priority High -Attachments C:\CRQ_Reminder\Instructions.pdf -Body ($Body | Out-String)} Catch{$_ | Out-File C:\CRQ_Reminder\Errors.txt -Append}
						Clear-content -Path "Errors.txt"
						$EmailSent = "123"
					}
					$arrayTo = @()
					$FoundCRQ = "true"  										##creating a flag for request found   
					$Requests3 = $Requests3 | Where-Object { $_ -ne $k }		##remove this item from array to improve performace
					break
                }
                else {
					$EmailSent = "456"
					$FoundCRQ = "false"
	            }
			}				
			if ($EmailSent -eq "456") {										##validating data on all files, if a request is not found on one saves an error on error log
				$message = $Requests1[$i] + " CRQ not listed in P3 report"
				Add-content  -Path "CRQsNotFound.txt" $message
			}
        }
        else {
			$FoundCRQ = "false"
		}
		if ($FoundCRQ -eq "true") {											##validating flag status for request found to avoid search a request when a related email has been sent
			$Requests2 = $Requests2 | Where-Object { $_ -ne $j }			##remove this item from array to improve performace
			break
		}
	}
	if ($FoundCRQ -eq "false") {											##validating data on all files, if a request is not found on one saves an error on error log
		$message = $Requests1[$i] + " CRQ not listed in P2 report"
		Add-content  -Path "CRQsNotFound.txt" $message
	}else {
		
	}
}
$message = Get-content -Path "CRQsNotFound.txt"					##generating error email for requests not found
if($null -eq $message){
	
}else{
	$ErrorForSD = "CRQs not found on automate reminder reports, please contact CIM leads"
	Send-MailMessage -From $From -To $From -Credential $Credential -SmtpServer $SMTPServer -UseSsl -Port 587 -Subject $ErrorForSD -Priority High -Body ($message | Out-String)
}


Remove-Item -Path * -Include *.xlsx			##removing currend data as reports are updated and dowloaded every day
