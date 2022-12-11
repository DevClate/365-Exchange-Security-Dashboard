function Get-365OutboundSpamDetections {
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
        $Path = 'C:\Scripts\OutboundSpamDetections.xlsx'
    )
    
    begin {
        
    }
    
    process {
        #Outbound Spam Detections
        Get-MailFlowStatusReport -StartDate $startdate -EndDate $enddate -Direction Outbound -EventType SpamDetections

    }
    
    end {
        
    }
}