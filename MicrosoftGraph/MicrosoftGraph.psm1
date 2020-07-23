[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Function Get-MSGraphAuthToken{
[cmdletbinding()]
Param(
    [parameter(Mandatory=$true)]
    [pscredential]$credential,
    [parameter(Mandatory=$true)]
    [string]$tenantID
    )

    #Get token
    $AuthUri = "https://login.microsoftonline.com/$TenantID/oauth2/token"
    $Resource = 'graph.microsoft.com'
    $AuthBody = "grant_type=client_credentials&client_id=$($credential.UserName)&client_secret=$($credential.GetNetworkCredential().Password)&resource=https%3A%2F%2F$Resource%2F"

    $Response = Invoke-RestMethod -Method Post -Uri $AuthUri -Body $AuthBody
    If($Response.access_token){
        return $Response.access_token
    }
    Else{
        Throw "Authentication failed"
    }
}

Function Invoke-MSGraphQuery{
[CmdletBinding(DefaultParametersetname="Default")]
Param(
    [Parameter(Mandatory=$true,ParameterSetName='Default')]
    [Parameter(Mandatory=$true,ParameterSetName='Refresh')]
    [string]$URI,

    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [Parameter(Mandatory=$false,ParameterSetName='Refresh')]
    [string]$Body,

    [Parameter(Mandatory=$true,ParameterSetName='Default')]
    [Parameter(Mandatory=$true,ParameterSetName='Refresh')]
    [string]$token,

    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [Parameter(Mandatory=$false,ParameterSetName='Refresh')]
    [ValidateSet('GET','POST','PUT','PATCH','DELETE')]
    [string]$method = "GET",
    
    [Parameter(Mandatory=$false,ParameterSetName='Default')]
    [Parameter(Mandatory=$false,ParameterSetName='Refresh')]
    [switch]$recursive,
    
    [Parameter(Mandatory=$true,ParameterSetName='Refresh')]
    [switch]$tokenrefresh,
    
    [Parameter(Mandatory=$true,ParameterSetName='Refresh')]
    [pscredential]$credential,
    
    [Parameter(Mandatory=$true,ParameterSetName='Refresh')]
    [string]$tenantID
)
    $authHeader = @{
        'Accept'= 'application/json'
        'Content-Type'= 'application/json'
        'Authorization'= $Token
    }
    [array]$returnvalue = $()
    Try{
        If($body){
            $Response = Invoke-RestMethod -Uri $URI -Headers $authHeader -Body $Body -Method $method -ErrorAction Stop
        }
        Else{
            $Response = Invoke-RestMethod -Uri $URI -Headers $authHeader -Method $method -ErrorAction Stop
        }
    }
    Catch{
        If(($Error[0].ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue).error.Message -eq 'Access token has expired.' -and $tokenrefresh){
            $token =  Get-MSGraphAuthToken -credential $credential -tenantID $TenantID

            $authHeader = @{
                'Content-Type'='application/json'
                'Authorization'=$Token
            }
            $returnvalue = $()
            If($body){
                $Response = Invoke-RestMethod -Uri $URI -Headers $authHeader -Body $Body -Method $method -ErrorAction Stop
            }
            Else{
                $Response = Invoke-RestMethod -Uri $uri -Headers $authHeader -Method $method
            }
        }
        Else{
            Throw $_
        }
    }

    $returnvalue += $Response
    If(-not $recursive -and $Response.'@odata.nextLink'){
        Write-Warning "Query contains more data, use recursive to get all!"
        Start-Sleep 1
    }
    ElseIf($recursive -and $Response.'@odata.nextLink'){
        If($PSCmdlet.ParameterSetName -eq 'default'){
            If($body){
                $returnvalue += Invoke-MSGraphQuery -URI $Response.'@odata.nextLink' -token $token -body $body -method $method -recursive -ErrorAction SilentlyContinue
            }
            Else{
                $returnvalue += Invoke-MSGraphQuery -URI $Response.'@odata.nextLink' -token $token -method $method -recursive -ErrorAction SilentlyContinue
            }
        }
        Else{
            If($body){
                $returnvalue += Invoke-MSGraphQuery -URI $Response.'@odata.nextLink' -token $token -body $body -method $method -recursive -tokenrefresh -credential $credential -tenantID $TenantID -ErrorAction SilentlyContinue
            }
            Else{
                $returnvalue += Invoke-MSGraphQuery -URI $Response.'@odata.nextLink' -token $token -method $method -recursive -tokenrefresh -credential $credential -tenantID $TenantID -ErrorAction SilentlyContinue
            }
        }
    }
    Return $returnvalue
}

Function Get-OneDriveItems {
Param(
    [parameter(Mandatory=$true)]
    [string]$UserId,
    [parameter(Mandatory=$true)]
    [string]$token,
    [parameter(Mandatory=$true)]
    [PSCredential]$credential,
    [parameter(Mandatory=$true)]
    [string]$TenantID,
    [parameter(Mandatory=$false)]
    [string]$itemId
)
    $AllItems = @()

    If($ItemId){
        $AllItems += Invoke-MSGraphQuery -URI "https://graph.microsoft.com/v1.0/users/$UserId/drive/items/$itemId/children" -token $token -tokenrefresh -recursive -credential $credential -tenantID $TenantID | Select -ExpandProperty value
    }
    else{
        $AllItems += Invoke-MSGraphQuery -URI "https://graph.microsoft.com/v1.0/users/$UserId/drive/root/children" -token $token -tokenrefresh -recursive -credential $credential -tenantID $TenantID | Select -ExpandProperty value
    }

    $AllItems | Where-Object{$_.folder.childCount -ge 1} | ForEach-Object{
        $AllItems += Get-OneDriveItems -UserId $UserId -token $token -credential $credential -TenantID $TenantID -itemId $_.Id
    }

    Write-Verbose "Found $($AllItems.Count) files in folder $itemId"

    Return $AllItems
}
