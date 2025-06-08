Function Test-IsRegistryPOLGood
    {
        [cmdletbinding()]
        Param
            (
                [Parameter(Mandatory=$false)]
                    [string[]]$PathToRegistryPOLFile = $(Join-Path $env:windir 'System32\GroupPolicy\Machine\Registry.pol')
            )
 
        if(!(Test-Path -Path $PathToRegistryPOLFile -PathType Leaf)) { return $null }
 
        [Byte[]]$FileHeader = Get-Content -Encoding Byte -Path $PathToRegistryPOLFile -TotalCount 4
 
        if(($FileHeader -join '') -eq '8082101103') { return $true } else { return $false }
    }
    
    Test-IsRegistryPOLGood