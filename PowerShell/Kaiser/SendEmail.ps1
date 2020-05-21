
$OL = New-Object -ComObject outlook.application
Start-Sleep 5
<#

olAppointmentItem
olContactItem
olDistributionListItem
olJournalItem
olMailItem
olNoteItem
olPostItem
olTaskItem
#>

#Create Item
$mItem = $OL.CreateItem("olMailItem")
$mItem.To = "kuchuk.sv@gmail.com"
$mItem.Subject = "PowerMail"
$mItem.Body = "SENT FROM POWERSHELL2"

$mItem.Send()