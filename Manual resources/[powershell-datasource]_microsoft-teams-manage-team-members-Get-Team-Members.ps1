# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Set from Global variables
# $AADTenantID = ""
# $AADAppId = ""
# $AADAppSecret = ""

# Form input
$team = $datasource.team

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
         
    Write-Verbose "Querying Members of Microsoft Team $($team.displayName) ($($team.id))"

    #Add the authorization header to the request
    $authorization = @{
        Authorization  = "Bearer $accesstoken";
        'Content-Type' = "application/json";
        Accept         = "application/json";
    }

    $baseSearchUri = "https://graph.microsoft.com/"
    $searchUri = $baseSearchUri + "v1.0/teams/$($team.id)/members"

    $getMicrosoftTeamMembersResponse = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
    $microsoftTeamMembers = $getMicrosoftTeamMembersResponse.value
    while (![string]::IsNullOrEmpty($getMicrosoftTeamMembersResponse.'@odata.nextLink')) {
        $getMicrosoftTeamMembersResponse = Invoke-RestMethod -Uri $getMicrosoftTeamMembersResponse.'@odata.nextLink' -Method Get -Headers $authorization -Verbose:$false
        $microsoftTeamMembers += $getMicrosoftTeamMembersResponse.value
    }    

    Write-Information "Successfully queried Members of Microsoft Team $($team.displayName) ($($team.id)). Result count: $($microsoftTeamMembers.id.Count)"

    # Sort by DisplayName
    $objects = $microsoftTeamMembers | Sort-Object -Property DisplayName

    # Output result to HelloID
    foreach ($object in $objects) {
        $displayValue = "$($object.displayName) ($($object.userId))"
        $returnObject = @{
            displayValue      = "$($displayValue)"
            displayName       = "$($object.displayName)"
            mail              = "$($object.email)"
            userId            = "$($object.userId)"
            membershipId      = "$($object.id)"
        }
        
        Write-output $returnObject
    }
}
catch {
    throw "Could not query Members of Microsoft Team $($team.displayName) ($($team.id)). Error: $($_.Exception.Message)"
}
