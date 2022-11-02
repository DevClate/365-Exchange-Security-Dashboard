# Connecting to 365 Online
function Connect-CTExchangeOnline {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $UserPrincipalName
        
    )
    
    begin {
        
    }
    
    process {
       Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName
    }
    
    end {
        
    }
}