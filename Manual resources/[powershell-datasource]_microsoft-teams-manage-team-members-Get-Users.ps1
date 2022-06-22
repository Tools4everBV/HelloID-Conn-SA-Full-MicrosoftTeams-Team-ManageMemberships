# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Set from Global variables
# $AADTenantID = ""
# $AADAppId = ""
# $AADAppSecret = ""

# Get users
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
         
    Write-Verbose "Querying Azure AD users"

    #Add the authorization header to the request
    $authorization = @{
        Authorization = "Bearer $accesstoken";
        'Content-Type' = "application/json";
        Accept = "application/json";
    }

    $baseSearchUri = "https://graph.microsoft.com/"
    $searchUri = $baseSearchUri + 'v1.0/users?$orderby=displayName'

    $azureADUsersResponse = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
    $azureADUsers = $azureADUsersResponse.value
    while (![string]::IsNullOrEmpty($azureADUsersResponse.'@odata.nextLink')) {
        $azureADUsersResponse = Invoke-RestMethod -Uri $azureADUsersResponse.'@odata.nextLink' -Method Get -Headers $authorization -Verbose:$false
        $azureADUsers += $azureADUsersResponse.value
    }    

    Write-Information "Successfully queried Azure AD users. Result count: $($azureADUsers.id.Count)"

    # Sort by DisplayName
    $objects = $azureADUsers | Sort-Object -Property DisplayName

    # Output result to HelloID
    foreach ($object in $objects) {
        $displayValue = "$($object.displayName) ($($object.id))"
        $returnObject = @{
            displayValue      = "$($displayValue)"
            displayName       = "$($object.displayName)"
            mail              = "$($object.mail)"
            givenName         = "$($object.givenName)"
            surname           = "$($object.surname)"
            userPrincipalName = "$($object.userPrincipalName)"
            userId            = "$($object.id)"
            membershipId      = $null # Empty value but needed since it is provided from the pre-filled options
        }
        
        Write-output $returnObject
    }
}
catch {
    throw "Could not query Azure AD users. Error: $($_.Exception.Message)"
}
