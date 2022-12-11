# Set timeframe and get our data
$StartDate = (Get-Date -Day 1).AddMonths(-1).ToString("yyyy-MM-dd")
$EndDate = (Get-Date -Day 1).AddDays(-1).ToString("yyyy-MM-dd") 
$MonthlyMailFlow = Get-MailFlowStatusReport -StartDate $StartDate -EndDate $EndDate
$Path = "C:\Scripts\Reports\365ExchangeMonthlySecurity.xlsx"
$PathRaw = "C:\Scripts\Reports\365ExchangeMonthlySecurityraw.xlsx"
$MonthTitle = (Get-Date -Day 1).AddMonths(-1).ToString("y")

# Get available directions and event types
$Directions = $MonthlyMailFlow.Direction | Sort-Object -Unique
$EventTypes = $MonthlyMailFlow.EventType | Sort-Object -Unique
$MessageOutput = New-Object System.Collections.ArrayList

# Loop through event types for each direction, put data into our output
foreach ($d in $Directions) {
    foreach ($e in $EventTypes) {
        $Messages = $MonthlyMailFlow | Where-Object { ($_.Direction -eq $d) -and ($_.EventType -eq $e) }
        $Sum = ($Messages | Measure-Object -Property MessageCount -Sum).Sum
        if ($null -eq $Sum) {
            $Sum = 0
        }
        $Totals = [PSCustomObject]@{
            Direction = $d
            Type      = $e
            Count     = $Sum
            Percentage = $null
        }
        $null = $MessageOutput.Add($Totals)
    }
}

# Use $MessageOutput to output to host or build report file
$MessageOutput | Export-Excel $Path -WorksheetName "$Monthtitle Summary" -TableStyle Medium16 -Title "$monthtitle Exchange Mailflow" -StartRow 1 -AutoSize

#Top 10 Summary Sheet
"$Monthtitle Exchange Top 10" | Export-Excel $Path -workSheetName "$Monthtitle Top" -StartRow 1 -AutoSize

#Top 10 Malware Recipient
Get-MailTrafficSummaryReport -Category TopMalwareRecipient –StartDate $startdate -EndDate $enddate | Select-Object @{N='Malware Recipients';E={$_.C1}},@{N='Count';E={$_.C2}} -First 10  | Export-Excel $Path -workSheetName "$Monthtitle Top" -TableStyle Medium16 -StartRow 2 -AutoSize

#Top 10 Phish Recipient
Get-MailTrafficSummaryReport -Category TopphishRecipient –StartDate $startdate -EndDate $enddate | Select-Object @{N='Phishing Recipients';E={$_.C1}},@{N='Count';E={$_.C2}} -First 10 | Export-Excel $Path -workSheetName "$Monthtitle Top" -TableStyle Medium16 -StartRow 2 -StartColumn 4 -AutoSize

#Top Malware
Get-MailTrafficSummaryReport -Category TopMalware –StartDate $startdate -EndDate $enddate | Select-Object @{N='Malware Type';E={$_.C1}},@{N='Count';E={$_.C2}} -First 10 | Export-Excel $Path -workSheetName "$Monthtitle Top" -TableStyle Medium16 -StartRow 2 -StartColumn 7 -AutoSize

$excel = Open-ExcelPackage $Path
$SummaryPage = $excel.Workbook.Worksheets["$Monthtitle Summary"]
$SummaryPage.View.ShowGridLines = $false
$BorderBottom = "Thin"
$BorderColor = "Black"
Set-ExcelRange -Address $SummaryPage.Cells["B:B"] -AutoFit
Set-ExcelRange -Address $SummaryPage.Cells["D3:D21"] -NumberFormat 'Percentage'
Set-ExcelRange -Address $SummaryPage.Cells["D3"] -Formula "=C3/SUM(C3:C8)"
Set-ExcelRange -Address $SummaryPage.Cells["D4"] -Formula "=C4/SUM(C3:C8)"
Set-ExcelRange -Address $SummaryPage.Cells["D5"] -Formula "=C5/SUM(C3:C8)"
Set-ExcelRange -Address $SummaryPage.Cells["D6"] -Formula "=C6/SUM(C3:C8)"
Set-ExcelRange -Address $SummaryPage.Cells["D7"] -Formula "=C7/SUM(C3:C8)"
Set-ExcelRange -Address $SummaryPage.Cells["D8"] -Formula "=C8/SUM(C3:C8)"
Set-ExcelRange -Address $SummaryPage.Cells["A8:D8"] -BorderBottom $BorderBottom -BorderColor $BorderColor

Set-ExcelRange -Address $SummaryPage.Cells["D9"] -Formula "=C9/SUM(C9:C14)"
Set-ExcelRange -Address $SummaryPage.Cells["D10"] -Formula "=C10/SUM(C9:C14)"
Set-ExcelRange -Address $SummaryPage.Cells["D11"] -Formula "=C11/SUM(C9:C14)"
Set-ExcelRange -Address $SummaryPage.Cells["D12"] -Formula "=C12/SUM(C9:C14)"
Set-ExcelRange -Address $SummaryPage.Cells["D13"] -Formula "=C13/SUM(C9:C14)"
Set-ExcelRange -Address $SummaryPage.Cells["D14"] -Formula "=C14/SUM(C9:C14)"
Set-ExcelRange -Address $SummaryPage.Cells["A14:D14"] -BorderBottom $BorderBottom -BorderColor $BorderColor

Set-ExcelRange -Address $SummaryPage.Cells["D15"] -Formula "=C15/SUM(C15:C20)"
Set-ExcelRange -Address $SummaryPage.Cells["D16"] -Formula "=C16/SUM(C15:C20)"
Set-ExcelRange -Address $SummaryPage.Cells["D17"] -Formula "=C17/SUM(C15:C20)"
Set-ExcelRange -Address $SummaryPage.Cells["D18"] -Formula "=C18/SUM(C15:C20)"
Set-ExcelRange -Address $SummaryPage.Cells["D19"] -Formula "=C19/SUM(C15:C20)"
Set-ExcelRange -Address $SummaryPage.Cells["D20"] -Formula "=C20/SUM(C15:C20)"
Set-ExcelRange -Address $SummaryPage.Cells["A20:D20"] -BorderBottom $BorderBottom -BorderColor $BorderColor

$TopPage = $excel.Workbook.Worksheets["$Monthtitle Top"]
$TopPage.View.ShowGridLines = $false
Set-ExcelRange -Address $TopPage.Cells["A:C"] -AutoFit
Set-ExcelRange -Address $TopPage.Cells["A1"] -FontSize 22

Close-ExcelPackage $excel

#Create Raw File for month in seperate raw file
Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate | Export-Excel -Path $PathRaw -workSheetName "$MonthTitle Raw" -AutoSize

#Create Raw file for Malware Recipients in seperate raw file
Get-MailTrafficSummaryReport -Category TopMalwareRecipient –StartDate $startdate -EndDate $enddate | Select-Object @{N='Malware Recipients';E={$_.C1}},@{N='Count';E={$_.C2}} | Export-Excel -Path $PathRaw -workSheetName "$MonthTitle Mal Recip Raw" -AutoSize

#Create Raw file for Phish Recipients in seperate raw file
Get-MailTrafficSummaryReport -Category TopphishRecipient –StartDate $startdate -EndDate $enddate | Select-Object @{N='Phishing Recipients';E={$_.C1}},@{N='Count';E={$_.C2}} | Export-Excel -Path $PathRaw -workSheetName "$MonthTitle Phish Recip Raw" -AutoSize

#Create Raw file for Malware in seperate raw file
Get-MailTrafficSummaryReport -Category TopMalware –StartDate $startdate -EndDate $enddate | Select-Object @{N='Malware Type';E={$_.C1}},@{N='Count';E={$_.C2}} | Export-Excel -Path $PathRaw -workSheetName "$Monthtitle Mal Raw" -AutoSize
