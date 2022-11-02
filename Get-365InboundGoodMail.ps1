function Get-365InboundGoodMail {
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
        $Path = 'C:\Scripts\InboundGoodMail.xlsx'
    )
    
    begin {
        
    }
    
    process {
           #Inbound Good Mail
           Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate -Direction Inbound -EventType GoodMail

    }
    
    end {
        
    }
}