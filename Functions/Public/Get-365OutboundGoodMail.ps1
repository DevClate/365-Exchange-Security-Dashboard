function Get-365OutboundGoodMail {
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
        $Path = 'C:\Scripts\OutboundGoodMail.xlsx'
    )
    
    begin {
        
    }
    
    process {
         #Outbound Good Mail
        Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate -Direction Outbound -EventType GoodMail

    }
    
    end {
        
    }
}