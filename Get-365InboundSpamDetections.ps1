function Get-365InboundSpamDetections {
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
        [Parameter()]
        [string]
        $Path = 'C:\Scripts\InboundSpamDetections.xlsx'
    )
    
    begin {
        
    }
    
    process {
         #Inbound Spam Detections
         Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate -Direction Inbound -EventType SpamDetections

    }
    
    end {
        
    }
}