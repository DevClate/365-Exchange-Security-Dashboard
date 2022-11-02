# Disconnecting from 365 Online
function Disconnect-CTExchangeOnline {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $UserPrincipalName
        
    )
    
    begin {
        
    }
    
    process {
       Disconnect-ExchangeOnline -UserPrincipalName $UserPrincipalName
    }
    
    end {
        
    }
}