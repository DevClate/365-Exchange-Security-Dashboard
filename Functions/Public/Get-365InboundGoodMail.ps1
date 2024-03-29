function Get-365InboundGoodMail {
            <#
        .SYNOPSIS
            PowerShell script to pull all Inbound Good Mail from Exchange Mailflow data
        .DESCRIPTION
            PowerShell script to pull all Inbound Good Mail from Exchange Mailflow data with a certain date range.
            Microsoft doesn't allow past 180 days for Mailflow and only 90 days for Top information. 
        .PARAMETER
            -StartDate [datetime]
            Specifies the starting date that you want to search

            -EndDate [datetime]
            Specifies the ending date that you want to search

            -Path [string]
            Specifies the location you want to save the information
        .NOTES
            Version:        0.1.0
            Author:         Clayton Tyger
            Twitter:        @clatent
            Github:         DevClate
           
        .LINK
            https://github.com/DevClate/365-Exchange-Security-Dashboard/
        #>

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