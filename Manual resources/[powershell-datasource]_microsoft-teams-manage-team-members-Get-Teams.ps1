# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Set from Global variables
# $AADTenantID = ""
# $AADAppId = ""
# $AADAppSecret = ""

# Get Microsoft Teams
try {
    Write-Information "Generating Microsoft Graph API Access Token"

    $baseUri = "https://login.microsoftonline.com/"
    $authUri = $baseUri + "$AADTenantID/oauth2/token"

    $body = @{
        grant_type    = "client_credentials"
        client_id     = "$AADAppId"
        client_secret = "$AADAppSecret"
        resource      = "https://graph.microsoft.com"
    }
 
    $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
    $accessToken = $Response.access_token;
         
    Write-Verbose "Querying Microsoft Teams"

    #Add the authorization header to the request
    $authorization = @{
        Authorization = "Bearer $accesstoken";
        'Content-Type' = "application/json";
        Accept = "application/json";
    }

    $baseSearchUri = "https://graph.microsoft.com/"
    $searchUri = $baseSearchUri + 'v1.0/groups?$filter=resourceProvisioningOptions' + "/Any(x:x eq  'Team')"

    $getMicrosoftTeamsResponse = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
    $microsoftTeams = $getMicrosoftTeamsResponse.value
    while (![string]::IsNullOrEmpty($getMicrosoftTeamsResponse.'@odata.nextLink')) {
        $getMicrosoftTeamsResponse = Invoke-RestMethod -Uri $getMicrosoftTeamsResponse.'@odata.nextLink' -Method Get -Headers $authorization -Verbose:$false
        $microsoftTeams += $getMicrosoftTeamsResponse.value
    }    

    Write-Information "Successfully queried Microsoft Teams. Result count: $($microsoftTeams.id.Count)"

    # Sort by DisplayName
    $objects = $microsoftTeams | Sort-Object -Property DisplayName

    # Output result to HelloID
    foreach ($object in $objects) {
        $displayValue = "$($object.displayName) ($($object.id))"
        $returnObject = @{
            displayValue            = "$($displayValue)"
            displayName             = "$($object.displayName)"
            description             = "$($object.description)"
            mail                    = "$($object.mail)"
            mailEnabled             = "$($object.mailEnabled)"
            mailNickName            = "$($object.mailNickName)"
            proxyAddresses          = "$($object.proxyAddresses)"
            id                      = "$($object.id)"
            securityIdentifier      = "$($object.securityIdentifier)"    
        }
        
        Write-output $returnObject
    }
}
catch {
    throw "Could not query Microsoft Teams. Error: $($_.Exception.Message)"
}
