# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Set from Global variables
# $AADTenantID = ""
# $AADAppId = ""
# $AADAppSecret = ""

# variables configured in form
$team = $form.dropDownTeam
$usersToAdd = $form.dualListMembers.leftToRight
$usersToRemove = $form.dualListMembers.rightToLeft

# Create authorization token and add to headers
try{
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

    #Add the authorization header to the request
    $authorization = @{
        Authorization  = "Bearer $accesstoken";
        'Content-Type' = "application/json";
        Accept         = "application/json";
    }
}
catch{
    throw "Could not generate Microsoft Graph API Access Token. Error: $($_.Exception.Message)"
}

# Get current members because we need the membership id to remove it
try{
    Write-Verbose "Querying Members of Microsoft Team $($team.displayName) ($($team.id))"

    $baseSearchUri = "https://graph.microsoft.com/"
    $searchUri = $baseSearchUri + "v1.0/teams/$($team.id)/members"

    $getMicrosoftTeamMembersResponse = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
    $microsoftTeamMembers = $getMicrosoftTeamMembersResponse.value
    while (![string]::IsNullOrEmpty($getMicrosoftTeamMembersResponse.'@odata.nextLink')) {
        $getMicrosoftTeamMembersResponse = Invoke-RestMethod -Uri $getMicrosoftTeamMembersResponse.'@odata.nextLink' -Method Get -Headers $authorization -Verbose:$false
        $microsoftTeamMembers += $getMicrosoftTeamMembersResponse.value
    }

    $microsoftTeamMembersGrouped = $microsoftTeamMembers | Group-Object -Property userId -AsHashTable -AsString

    Write-Information "Successfully queried Members of Microsoft Team $($team.displayName) ($($team.id)). Result count: $($microsoftTeamMembers.id.Count)"
}
catch{
   Write-Error "Could not query Members of Microsoft Team $($team.displayName) ($($team.id)). Error: $($_.Exception.Message)"
}

if($usersToAdd -ne $null){
    foreach($userToAdd in $usersToAdd){
        try{
            Write-Verbose "Adding user $($userToAdd.displayName) ($($userToAdd.userId)) to Microsoft Team $($team.displayName) ($($team.id))"

            # Somehow can't get this to work, but the example is provided
            # Use https://docs.microsoft.com/en-us/graph/api/team-post-members?view=graph-rest-1.0&tabs=http 
            # $baseAddUri = "https://graph.microsoft.com/"
            # $addUri = $baseAddUri + "v1.0/teams/$($team.id)/members"
            # $body = [PsCustomObject]@{
            #     '@odata.type'       = '#microsoft.graph.aadUserConversationMember'
            #     'roles'             = @("guest")
            #     'user@odata.bind'   = "https://graph.microsoft.com/v1.0/users('$($userToAdd.userId)')"                
            # } | ConvertTo-Json -Depth 10
            # $addMember = Invoke-RestMethod -Uri $addUri -Method Post -Headers $authorization -Verbose:$false

            # Since we can't add the users to the team using the team API, make use of the groups API
            # Use https://docs.microsoft.com/en-us/graph/api/team-post-members?view=graph-rest-1.0&tabs=http 
            $baseGraphUri = "https://graph.microsoft.com/"
            $addGroupMembershipUri = $baseGraphUri + "v1.0/groups/$($team.id)/members" + '/$ref'
            $body = @{ 
                "@odata.id"= "https://graph.microsoft.com/v1.0/users/$($userToAdd.userId)" 
            } | ConvertTo-Json -Depth 10
            $addMember = Invoke-RestMethod -Method POST -Uri $addGroupMembershipUri -Body $body -Headers $authorization -Verbose:$false

            Write-Information "Successfully added user $($userToAdd.displayName) ($($userToAdd.userId)) to Microsoft Team $($team.displayName) ($($team.id))"

            $Log = @{
                Action            = "GrantMembership" # optional. ENUM (undefined = default) 
                System            = "MicrosoftTeams" # optional (free format text) 
                Message           = "Successfully added user $($userToAdd.displayName) ($($userToAdd.userId)) to Microsoft Team $($team.displayName) ($($team.id))" # required (free format text) 
                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $($team.displayName) # optional (free format text)
                TargetIdentifier  = $($team.id) # optional (free format text)
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log
        }catch{
            $Log = @{
                Action            = "GrantMembership" # optional. ENUM (undefined = default) 
                System            = "MicrosoftTeams" # optional (free format text) 
                Message           = "Failed to add user $($userToAdd.displayName) ($($userToAdd.userId)) to Microsoft Team $($team.displayName) ($($team.id)). Error: $($_.Exception.Message)" # required (free format text) 
                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $($team.displayName) # optional (free format text)
                TargetIdentifier  = $($team.id) # optional (free format text)
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log

            Write-Error "Could not add user $($userToAdd.displayName) ($($userToAdd.userId)) to Microsoft Team $($team.displayName) ($($team.id)). Error: $($_.Exception.Message)"            
        }
    }
}

if($usersToRemove -ne $null){
    foreach($userToRemove in $usersToRemove){
         try{
            Write-Verbose "Removing user $($userToRemove.displayName) ($($userToRemove.userId)) from Microsoft Team $($team.displayName) ($($team.id))"

            $membershipId = $microsoftTeamMembersGrouped["$($userToRemove.userId)"].id

            if($null -eq $membershipId){
                throw "No membership found for user $($userToRemove.displayName) ($($userToRemove.userId)) to Microsoft Team $($team.displayName) ($($team.id))"
            }

            $baseRemoveUri = "https://graph.microsoft.com/"
            $removeUri = $baseRemoveUri + "v1.0/teams/$($team.id)/members/$membershipId"
            $removeMember = Invoke-RestMethod -Uri $removeUri -Method Delete -Headers $authorization -Verbose:$false

            Write-Information "Successfully removed user $($userToRemove.displayName) ($($userToRemove.userId)) from Microsoft Team $($team.displayName) ($($team.id))"

            $Log = @{
                Action            = "RevokeMembership" # optional. ENUM (undefined = default) 
                System            = "MicrosoftTeams" # optional (free format text) 
                Message           = "Successfully removed user $($userToRemove.displayName) ($($userToRemove.userId)) from Microsoft Team $($team.displayName) ($($team.id))" # required (free format text) 
                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $($team.displayName) # optional (free format text)
                TargetIdentifier  = $($team.id) # optional (free format text)
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log
        }catch{
            $Log = @{
                Action            = "RevokeMembership" # optional. ENUM (undefined = default) 
                System            = "MicrosoftTeams" # optional (free format text) 
                Message           = "Failed to remove user $($userToRemove.displayName) ($($userToRemove.userId)) from Microsoft Team $($team.displayName) ($($team.id)). Error: $($_.Exception.Message)" # required (free format text) 
                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $($team.displayName) # optional (free format text)
                TargetIdentifier  = $($team.id) # optional (free format text)
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log

            Write-Error "Could not remove user $($userToRemove.displayName) ($($userToRemove.userId)) from Microsoft Team $($team.displayName) ($($team.id)). Error: $($_.Exception.Message)"            
        }
    }
}
