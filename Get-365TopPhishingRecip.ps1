function Get-365TopPhishingRecip {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,HelpMessage='Enter start time, ex 04/01/2022')]
        [datetime]
        $StartDate,
        [Parameter(Mandatory,HelpMessage='Enter end date, ex 04/30/2022')]
        [datetime]
        $EndDate,
        [Parameter()]
        [string]
        $Path = 'C:\Scripts\TopPishingRecipReport.xlsx'
    )
    
    begin {
    
    }
    
    process {
        #Top Phishing Recipients
        Get-MailTrafficSummaryReport -Category TopphishRecipient –StartDate $startdate -EndDate $enddate | Select-Object C1,C2
    }
    
    end {
        
    }
}