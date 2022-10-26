#Only can go back 90 days per Microsoft as of 6/11/2022
#Only 30 days for Top Phish and Malware this looks like it has been updated to 90 days as of 10/26/2022
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]
    $UserPrincipalName,
    [Parameter(Mandatory,HelpMessage='Enter start time, ex 04/01/2022')]
    [datetime]
    $StartDate,
    [Parameter(Mandatory,HelpMessage='Enter end date, ex 04/30/2022')]
    [datetime]
    $EndDate,
    [Parameter(Mandatory)]
    [string]
    $Path = 'C:\Scripts\RawMonthlySecurityReport.xlsx'

)

Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName

#All
Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate | Export-Excel -Path $Path -workSheetName All

#Inbound Edge Block Spam
Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate -Direction Inbound -EventType EdgeBlockSpam  | Export-Excel -Path $Path -WorkSheetName EdgeBlockSpam-Inbound

#Outbound Edge Block Spam
Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate -Direction Outbound -EventType EdgeBlockSpam  | Export-Excel -Path $Path -WorkSheetName EdgeBlockSpam-Outbound

#Inbound Email Malware
Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate -Direction Inbound -EventType EmailMalware  | Export-Excel -Path $Path -workSheetName Malware-Inbound

#Outbound Email Malware
Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate -Direction Outbound -EventType EmailMalware  | Export-Excel -Path $Path -workSheetName Malware-Outbound

#Inbound Email Phish
Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate -Direction Inbound -EventType EmailPhish  | Export-Excel -Path $Path -workSheetName Phish-Inbound

#Outbound Email Phish
Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate -Direction Outbound -EventType EmailPhish  | Export-Excel -Path $Path -workSheetName Phish-Outbound

#Inbound Good Mail
Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate -Direction Inbound -EventType GoodMail  | Export-Excel -Path $Path -workSheetName GoodMail-Inbound

#Outbound Good Mail
Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate -Direction Outbound -EventType GoodMail  | Export-Excel -Path $Path -workSheetName GoodMail-Outbound

#Inbound Spam Detections
Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate -Direction Inbound -EventType SpamDetections  | Export-Excel -Path $Path -workSheetName SpamDetections-Inbound

#Outbound Spam Detections
Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate -Direction Outbound -EventType SpamDetections  | Export-Excel -Path $Path -workSheetName SpamDetections-Outbound

#Top Malware
Get-MailTrafficSummaryReport -Category TopMalware –StartDate $startdate -EndDate $enddate | Select-Object C1,C2 | Export-Excel $Path -workSheetName topmalware

#Top Malware Recipients
Get-MailTrafficSummaryReport -Category TopMalwareRecipient –StartDate $startdate -EndDate $enddate | Select-Object C1,C2 | Export-Excel $Path -workSheetName topmalwarerecip

#Top Phish Recipients
Get-MailTrafficSummaryReport -Category TopphishRecipient –StartDate $startdate -EndDate $enddate | Select-Object C1,C2 | Export-Excel $Path -workSheetName topphish

#Top Phish Recipients 10
Get-MailTrafficSummaryReport -Category TopphishRecipient –StartDate $startdate -EndDate $enddate | Select-Object C1,C2 -First 10 | Export-Excel $Path -workSheetName topphish10

#Top Malware Recipients 10
Get-MailTrafficSummaryReport -Category TopMalwareRecipient –StartDate $startdate -EndDate $enddate | Select-Object C1,C2 -First 10 | Export-Excel $Path -workSheetName topmalwarerecip10
