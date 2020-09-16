Clear-content -Path "C:\CRQ_Reminder\123.txt"
Clear-content -Path "C:\CRQ_Reminder\Error.txt"
Clear-content -Path "C:\CRQ_Reminder\Requester.txt"
$FilePath = "C:\CRQ_Reminder\P1.xlsx"
$SheetName1 = "P1 - Reminders Report"
$objExcel1 = New-Object -ComObject Excel.Application
$WorkBook1 = $objExcel1.Workbooks.Open($FilePath) <#Loads excel file in a object#>
$WorkSheet1 = $WorkBook1.sheets.item($SheetName1)   ########
$xlCellTypeLastCell = 11
$endRow1 = $WorkSheet1.UsedRange.SpecialCells($xlCellTypeLastCell).Row
$startRow1 = 5
$col1 = 2
$IncidentIDs = @()
for ($j = 0; $j -le $endRow1; $j++)<#Saves Incidents col in an array#>
{
  	$IncidentIDs += $WorkSheet1.Cells.Item($startRow1 + $j , $col1).Value2
}
$IncidentIDs = $IncidentIDs | Where-Object {-not [string]::IsNullOrWhiteSpace($_)}
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
for ($j = 0; $j -le $endRow1; $j++)<#Saves Summary col in an array#>
{
   	$Requests1 += $WorkSheet1.Cells.Item($startRow3 + $j , $col3).Value2
}
$Requests1 = $Requests1 | Where-Object {-not [string]::IsNullOrWhiteSpace($_)}
$objExcel1.Workbooks.Close()                                                                                                        <#Close excel document#>
$FilePath2 = "C:\CRQ_Reminder\P2.xlsx"
$SheetName2 = "P2 - CRQs In Request For Auth"
$objExcel2 = New-Object -ComObject Excel.Application
$WorkBook2 = $objExcel2.Workbooks.Open($FilePath2) <#Loads excel file in a object#>
$WorkSheet2 = $WorkBook2.sheets.item($SheetName2)
$endRow2 = $WorkSheet2.UsedRange.SpecialCells($xlCellTypeLastCell).Row
$startRow4 = 5
$col4 = 3
$Summaries = @()
for ($j = 0; $j -le $endRow2; $j++)<#Saves Summary col in an array#>
{
   	$Summaries += $WorkSheet2.Cells.Item($startRow4 + $j , $col4).Value2
}
$Summaries = $Summaries | Where-Object {-not [string]::IsNullOrWhiteSpace($_)}
$startRow5 = 5
$col5 = 1
$Requests2 = @()
for ($j = 0; $j -le $endRow2; $j++)<#Saves Summary col in an array#>
{
   	$Requests2 += $WorkSheet2.Cells.Item($startRow5 + $j , $col5).Value2
}
$Requests2 = $Requests2 | Where-Object {-not [string]::IsNullOrWhiteSpace($_)}
$objExcel2.Workbooks.Close()                                                                                                            <#Close excel document#>
$FilePath3 = "C:\CRQ_Reminder\P3.xlsx"
$SheetName3 = "P3 - CRQs Pending to be Approve"
$objExcel3 = New-Object -ComObject Excel.Application
$WorkBook3 = $objExcel3.Workbooks.Open($FilePath3) <#Loads excel file in a object#>
$WorkSheet3 = $WorkBook3.sheets.item($SheetName3)
$endRow2 = $WorkSheet3.UsedRange.SpecialCells($xlCellTypeLastCell).Row
$startRow6 = 5
$col6 = 1
$Requests3 = @()
for ($j = 0; $j -le $endRow2; $j++)<#Saves Summary col in an array#>
{
   	$Requests3 += $WorkSheet3.Cells.Item($startRow6 + $j , $col6).Value2
}
$Requests3 = $Requests3 | Where-Object {-not [string]::IsNullOrWhiteSpace($_)}
$startRow7 = 5
$col7 = 2
$Approvers = @()
for ($j = 0; $j -le $endRow2; $j++)<#Saves Summary col in an array#>
{
   	$Approvers += $WorkSheet3.Cells.Item($startRow7 + $j , $col7).Value2
}
$Approvers = $Approvers | Where-Object {-not [string]::IsNullOrWhiteSpace($_)}
$objExcel3.Workbooks.Close()<#Close excel document#>
$From = "ivan.palma@teleflex.com"##cambiar por servicedesk@teleflex.com
$PWord = ConvertTo-SecureString –String "Urticaria_149" –AsPlainText -Force ##cambiar por password de servicedesk@teleflex.com
$Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $From, $PWord
$SMTPServer = "smtp.office365.com"
$AppoverEmail = $null
$to = "123"
$RequesterEmailAddress = "123"
$UPNs = $null
$Cc = $null
$rest = @()
$Bcc = "ivan.palma@teleflex.com"##cambiar por servicedesk@teleflex.com
for($i = 0; $i -lt $Requests1.Length; $i++){
    for($j = 0; $j -lt $Requests2.Length; $j++){
        if($Requests1[$i] -contains $Requests2[$j]){
            for($k = 0; $k -lt $Requests3.Length; $k++){
                if($Requests1[$i] -contains $Requests3[$k]){
                    $Body = get-content "C:\CRQ_Reminder\Body.txt"
                    $Body = $Body -replace "aaaaaaaaaa",$Requests1[$i]
                    $array = $Requesters[$i].split(" ")	
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
                    try{$Cc = get-aduser -filter 'Name -like $RequesterEmailAddress' | Select-Object UserPrincipalName} Catch{$_ | Out-File C:\CRQ_Reminder\123.txt -Append}
					$RequesterEmailAddress = "123"
                    if($null -eq $Cc){
						Clear-content -Path "C:\CRQ_Reminder\123.txt"
						$message = $Requesters[$i] + " User email address don't found"
						Add-content  -Path "C:\CRQ_Reminder\Error.txt" $message
					}
					else{
						Add-content -Path "C:\CRQ_Reminder\Requester.txt" $Cc
						$Cc = Get-content -Path "C:\CRQ_Reminder\Requester.txt"
						Clear-Content -Path "C:\CRQ_Reminder\Requester.txt" -Force
						$array = $Cc.split("=")
						$value = $array.length
						$value = $value -1
						$UPNs = $array[$value]
						$UPNs = $UPNs.Substring(0,$UPNs.Length-1)
                    }
                    $Cc = $UPNs																																								##obtener correo de requester
					$Subject = $Requests1[$i] + " is pending for your approval"
                    $ApproversList = $Approvers[$j].split(';')
                    for($l = 0; $l -lt $ApproversList.Length; $l++){
						try{$AppoverEmail = get-ADUser -identity $ApproversList[$l] | select-object UserPrincipalName} Catch{$_ | Out-File C:\CRQ_Reminder\123.txt -Append}
						if($null -eq $AppoverEmail)
						{
							Clear-content -Path "C:\CRQ_Reminder\123.txt"
							$message = $Requests1[$i] + $ApproversList[$j] +" Approver email is not correct, check active directory"
							Add-content  -Path "C:\CRQ_Reminder\Error.txt" $message
						}
						else
						{
							Add-content -Path "C:\CRQ_Reminder\Requester.txt" $AppoverEmail
							$AppoverEmail = Get-content -Path "C:\CRQ_Reminder\Requester.txt"
							Clear-Content -Path "C:\CRQ_Reminder\Requester.txt" -Force
							$array = $AppoverEmail.split("=")
							$value = $array.length
							$value = $value -1
							$UPNs = $array[$value]
							$UPNs = $UPNs.Substring(0,$UPNs.Length-1)
							$AppoverEmail = $UPNs
							if($to -contains "123")
							{
								$To = $AppoverEmail
							}
							else
							{
								$To = $To + ";" + $AppoverEmail
							}
						}	
					}
					Write-Output "`n`n###############################################################`n`nSending email with next information...`n`n"
					$arrayTo = $to.split(";")
					Write-Output "List of approvers..."
					$arrayTo
					Write-Output "`n`nEmail subject		" $Subject
                    Write-Output "`n`nRequester email		" $Cc "`n`n"
					$Body
					if($null -ne $Cc)
					{
						try{Send-MailMessage -From $From -To $arrayTo -Cc ($Cc | Out-String) -Bcc $Bcc -Credential $Credential -SmtpServer $SMTPServer -UseSsl -Port 587 -Subject $Subject -Priority High -Attachments C:\CRQ_Reminder\Instructions.docx -Body ($Body | Out-String)} Catch{$_ | Out-File C:\CRQ_Reminder\Errors.txt -Append}
						Clear-content -Path "C:\CRQ_Reminder\Errors.txt"
						$message = $Requests1[$i] + $ApproversList[$j] + $Requesters[$i] + " Approver or requester email is not correct, check mailbox status"
                        Add-content  -Path "C:\CRQ_Reminder\Error.txt" $message
					}
					else
					{
						try{Send-MailMessage -From $From -To $arrayTo -Credential $Credential -Bcc $Bcc -SmtpServer $SMTPServer -UseSsl -Port 587 -Subject $Subject -Priority High -Attachments C:\CRQ_Reminder\Instructions.docx -Body ($Body | Out-String)} Catch{$_ | Out-File C:\CRQ_Reminder\Errors.txt -Append}
						Clear-content -Path "C:\CRQ_Reminder\Errors.txt"
						$message = $Requests1[$i] + $ApproversList[$j] + + $Requesters[$i] + " Approver or requester email is not correct, check mailbox status"
                        Add-content  -Path "C:\CRQ_Reminder\Error.txt" $message
					}
                }
                else {
                    $message = $Requests1[$i] + " CRQ not listed in all reports"
					Add-content  -Path "C:\CRQ_Reminder\Error.txt" $message
                    ##Write-Output "lalalalala"
                }
            }
        }
        else {
            ##Write-Output "srsrrsrsrsrsrsr"
        }
    }
}
