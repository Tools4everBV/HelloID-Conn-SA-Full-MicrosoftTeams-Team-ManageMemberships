# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
#HelloID variables
#Note: when running this script inside HelloID; portalUrl and API credentials are provided automatically (generate and save API credentials first in your admin panel!)
$portalUrl = "https://CUSTOMER.helloid.com"
$apiKey = "API_KEY"
$apiSecret = "API_SECRET"
$delegatedFormAccessGroupNames = @("") #Only unique names are supported. Groups must exist!
$delegatedFormCategories = @("Microsoft Teams","Team Management") #Only unique names are supported. Categories will be created if not exists
$script:debugLogging = $false #Default value: $false. If $true, the HelloID resource GUIDs will be shown in the logging
$script:duplicateForm = $false #Default value: $false. If $true, the HelloID resource names will be changed to import a duplicate Form
$script:duplicateFormSuffix = "_tmp" #the suffix will be added to all HelloID resource names to generate a duplicate form with different resource names

#The following HelloID Global variables are used by this form. No existing HelloID global variables will be overriden only new ones are created.
#NOTE: You can also update the HelloID Global variable values afterwards in the HelloID Admin Portal: https://<CUSTOMER>.helloid.com/admin/variablelibrary
$globalHelloIDVariables = [System.Collections.Generic.List[object]]@();

#Global variable #1 >> AADtenantID
$tmpName = @'
AADtenantID
'@ 
$tmpValue = @'
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #2 >> AADAppId
$tmpName = @'
AADAppId
'@ 
$tmpValue = @'
'@ 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "False"});

#Global variable #3 >> AADAppSecret
$tmpName = @'
AADAppSecret
'@ 
$tmpValue = "" 
$globalHelloIDVariables.Add([PSCustomObject]@{name = $tmpName; value = $tmpValue; secret = "True"});


#make sure write-information logging is visual
$InformationPreference = "continue"
# Check for prefilled API Authorization header
if (-not [string]::IsNullOrEmpty($portalApiBasic)) {
    $script:headers = @{"authorization" = $portalApiBasic}
    Write-Information "Using prefilled API credentials"
} else {
    # Create authorization headers with HelloID API key
    $pair = "$apiKey" + ":" + "$apiSecret"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)
    $key = "Basic $base64"
    $script:headers = @{"authorization" = $Key}
    Write-Information "Using manual API credentials"
}
# Check for prefilled PortalBaseURL
if (-not [string]::IsNullOrEmpty($portalBaseUrl)) {
    $script:PortalBaseUrl = $portalBaseUrl
    Write-Information "Using prefilled PortalURL: $script:PortalBaseUrl"
} else {
    $script:PortalBaseUrl = $portalUrl
    Write-Information "Using manual PortalURL: $script:PortalBaseUrl"
}
# Define specific endpoint URI
$script:PortalBaseUrl = $script:PortalBaseUrl.trim("/") + "/"  
# Make sure to reveive an empty array using PowerShell Core
function ConvertFrom-Json-WithEmptyArray([string]$jsonString) {
    # Running in PowerShell Core?
    if($IsCoreCLR -eq $true){
        $r = [Object[]]($jsonString | ConvertFrom-Json -NoEnumerate)
        return ,$r  # Force return value to be an array using a comma
    } else {
        $r = [Object[]]($jsonString | ConvertFrom-Json)
        return ,$r  # Force return value to be an array using a comma
    }
}
function Invoke-HelloIDGlobalVariable {
    param(
        [parameter(Mandatory)][String]$Name,
        [parameter(Mandatory)][String][AllowEmptyString()]$Value,
        [parameter(Mandatory)][String]$Secret
    )
    $Name = $Name + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })
    try {
        $uri = ($script:PortalBaseUrl + "api/v1/automation/variables/named/$Name")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
    
        if ([string]::IsNullOrEmpty($response.automationVariableGuid)) {
            #Create Variable
            $body = @{
                name     = $Name;
                value    = $Value;
                secret   = $Secret;
                ItemType = 0;
            }    
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl + "api/v1/automation/variable")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $variableGuid = $response.automationVariableGuid
            Write-Information "Variable '$Name' created$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        } else {
            $variableGuid = $response.automationVariableGuid
            Write-Warning "Variable '$Name' already exists$(if ($script:debugLogging -eq $true) { ": " + $variableGuid })"
        }
    } catch {
        Write-Error "Variable '$Name', message: $_"
    }
}
function Invoke-HelloIDAutomationTask {
    param(
        [parameter(Mandatory)][String]$TaskName,
        [parameter(Mandatory)][String]$UseTemplate,
        [parameter(Mandatory)][String]$AutomationContainer,
        [parameter(Mandatory)][String][AllowEmptyString()]$Variables,
        [parameter(Mandatory)][String]$PowershellScript,
        [parameter()][String][AllowEmptyString()]$ObjectGuid,
        [parameter()][String][AllowEmptyString()]$ForceCreateTask,
        [parameter(Mandatory)][Ref]$returnObject
    )
    
    $TaskName = $TaskName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/automationtasks?search=$TaskName&container=$AutomationContainer")
        $responseRaw = (Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false) 
        $response = $responseRaw | Where-Object -filter {$_.name -eq $TaskName}
    
        if([string]::IsNullOrEmpty($response.automationTaskGuid) -or $ForceCreateTask -eq $true) {
            #Create Task
            $body = @{
                name                = $TaskName;
                useTemplate         = $UseTemplate;
                powerShellScript    = $PowershellScript;
                automationContainer = $AutomationContainer;
                objectGuid          = $ObjectGuid;
                variables           = (ConvertFrom-Json-WithEmptyArray($Variables));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/automationtasks/powershell")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
            $taskGuid = $response.automationTaskGuid
            Write-Information "Powershell task '$TaskName' created$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        } else {
            #Get TaskGUID
            $taskGuid = $response.automationTaskGuid
            Write-Warning "Powershell task '$TaskName' already exists$(if ($script:debugLogging -eq $true) { ": " + $taskGuid })"
        }
    } catch {
        Write-Error "Powershell task '$TaskName', message: $_"
    }
    $returnObject.Value = $taskGuid
}
function Invoke-HelloIDDatasource {
    param(
        [parameter(Mandatory)][String]$DatasourceName,
        [parameter(Mandatory)][String]$DatasourceType,
        [parameter(Mandatory)][String][AllowEmptyString()]$DatasourceModel,
        [parameter()][String][AllowEmptyString()]$DatasourceStaticValue,
        [parameter()][String][AllowEmptyString()]$DatasourcePsScript,        
        [parameter()][String][AllowEmptyString()]$DatasourceInput,
        [parameter()][String][AllowEmptyString()]$AutomationTaskGuid,
        [parameter(Mandatory)][Ref]$returnObject
    )
    $DatasourceName = $DatasourceName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })
    $datasourceTypeName = switch($DatasourceType) { 
        "1" { "Native data source"; break} 
        "2" { "Static data source"; break} 
        "3" { "Task data source"; break} 
        "4" { "Powershell data source"; break}
    }
    
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/datasource/named/$DatasourceName")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
      
        if([string]::IsNullOrEmpty($response.dataSourceGUID)) {
            #Create DataSource
            $body = @{
                name               = $DatasourceName;
                type               = $DatasourceType;
                model              = (ConvertFrom-Json-WithEmptyArray($DatasourceModel));
                automationTaskGUID = $AutomationTaskGuid;
                value              = (ConvertFrom-Json-WithEmptyArray($DatasourceStaticValue));
                script             = $DatasourcePsScript;
                input              = (ConvertFrom-Json-WithEmptyArray($DatasourceInput));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
      
            $uri = ($script:PortalBaseUrl +"api/v1/datasource")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
              
            $datasourceGuid = $response.dataSourceGUID
            Write-Information "$datasourceTypeName '$DatasourceName' created$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        } else {
            #Get DatasourceGUID
            $datasourceGuid = $response.dataSourceGUID
            Write-Warning "$datasourceTypeName '$DatasourceName' already exists$(if ($script:debugLogging -eq $true) { ": " + $datasourceGuid })"
        }
    } catch {
      Write-Error "$datasourceTypeName '$DatasourceName', message: $_"
    }
    $returnObject.Value = $datasourceGuid
}
function Invoke-HelloIDDynamicForm {
    param(
        [parameter(Mandatory)][String]$FormName,
        [parameter(Mandatory)][String]$FormSchema,
        [parameter(Mandatory)][Ref]$returnObject
    )
    
    $FormName = $FormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })
    try {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/forms/$FormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }
    
        if(([string]::IsNullOrEmpty($response.dynamicFormGUID)) -or ($response.isUpdated -eq $true)) {
            #Create Dynamic form
            $body = @{
                Name       = $FormName;
                FormSchema = (ConvertFrom-Json-WithEmptyArray($FormSchema));
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/forms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
    
            $formGuid = $response.dynamicFormGUID
            Write-Information "Dynamic form '$formName' created$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        } else {
            $formGuid = $response.dynamicFormGUID
            Write-Warning "Dynamic form '$FormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $formGuid })"
        }
    } catch {
        Write-Error "Dynamic form '$FormName', message: $_"
    }
    $returnObject.Value = $formGuid
}
function Invoke-HelloIDDelegatedForm {
    param(
        [parameter(Mandatory)][String]$DelegatedFormName,
        [parameter(Mandatory)][String]$DynamicFormGuid,
        [parameter()][Array][AllowEmptyString()]$AccessGroups,
        [parameter()][String][AllowEmptyString()]$Categories,
        [parameter(Mandatory)][String]$UseFaIcon,
        [parameter()][String][AllowEmptyString()]$FaIcon,
        [parameter()][String][AllowEmptyString()]$task,
        [parameter(Mandatory)][Ref]$returnObject
    )
    $delegatedFormCreated = $false
    $DelegatedFormName = $DelegatedFormName + $(if ($script:duplicateForm -eq $true) { $script:duplicateFormSuffix })
    try {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$DelegatedFormName")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        } catch {
            $response = $null
        }
    
        if([string]::IsNullOrEmpty($response.delegatedFormGUID)) {
            #Create DelegatedForm
            $body = @{
                name            = $DelegatedFormName;
                dynamicFormGUID = $DynamicFormGuid;
                isEnabled       = "True";
                useFaIcon       = $UseFaIcon;
                faIcon          = $FaIcon;
                task            = ConvertFrom-Json -inputObject $task;
            }
            if(-not[String]::IsNullOrEmpty($AccessGroups)) { 
                $body += @{
                    accessGroups    = (ConvertFrom-Json-WithEmptyArray($AccessGroups));
                }
            }
            $body = ConvertTo-Json -InputObject $body -Depth 100
    
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
    
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Information "Delegated form '$DelegatedFormName' created$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
            $delegatedFormCreated = $true
            $bodyCategories = $Categories
            $uri = ($script:PortalBaseUrl +"api/v1/delegatedforms/$delegatedFormGuid/categories")
            $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $bodyCategories
            Write-Information "Delegated form '$DelegatedFormName' updated with categories"
        } else {
            #Get delegatedFormGUID
            $delegatedFormGuid = $response.delegatedFormGUID
            Write-Warning "Delegated form '$DelegatedFormName' already exists$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormGuid })"
        }
    } catch {
        Write-Error "Delegated form '$DelegatedFormName', message: $_"
    }
    $returnObject.value.guid = $delegatedFormGuid
    $returnObject.value.created = $delegatedFormCreated
}

<# Begin: HelloID Global Variables #>
foreach ($item in $globalHelloIDVariables) {
	Invoke-HelloIDGlobalVariable -Name $item.name -Value $item.value -Secret $item.secret 
}
<# End: HelloID Global Variables #>


<# Begin: HelloID Data sources #>
<# Begin: DataSource "microsoft-teams-manage-team-members-Get-Users" #>
$tmpPsScript = @'
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
'@ 
$tmpModel = @'
[{"key":"userPrincipalName","type":0},{"key":"userId","type":0},{"key":"displayName","type":0},{"key":"surname","type":0},{"key":"givenName","type":0},{"key":"displayValue","type":0},{"key":"membershipId","type":0},{"key":"mail","type":0}]
'@ 
$tmpInput = @'
[]
'@ 
$dataSourceGuid_2 = [PSCustomObject]@{} 
$dataSourceGuid_2_Name = @'
microsoft-teams-manage-team-members-Get-Users
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_2_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_2) 
<# End: DataSource "microsoft-teams-manage-team-members-Get-Users" #>

<# Begin: DataSource "microsoft-teams-manage-team-members-Get-Team-Members" #>
$tmpPsScript = @'
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
'@ 
$tmpModel = @'
[{"key":"mail","type":0},{"key":"displayValue","type":0},{"key":"membershipId","type":0},{"key":"userId","type":0},{"key":"displayName","type":0}]
'@ 
$tmpInput = @'
[{"description":null,"translateDescription":false,"inputFieldType":1,"key":"team","type":0,"options":1}]
'@ 
$dataSourceGuid_1 = [PSCustomObject]@{} 
$dataSourceGuid_1_Name = @'
microsoft-teams-manage-team-members-Get-Team-Members
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_1_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_1) 
<# End: DataSource "microsoft-teams-manage-team-members-Get-Team-Members" #>

<# Begin: DataSource "microsoft-teams-manage-team-members-Get-Teams" #>
$tmpPsScript = @'
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
'@ 
$tmpModel = @'
[{"key":"securityIdentifier","type":0},{"key":"id","type":0},{"key":"mailEnabled","type":0},{"key":"displayName","type":0},{"key":"description","type":0},{"key":"mailNickName","type":0},{"key":"proxyAddresses","type":0},{"key":"displayValue","type":0},{"key":"mail","type":0}]
'@ 
$tmpInput = @'
[]
'@ 
$dataSourceGuid_0 = [PSCustomObject]@{} 
$dataSourceGuid_0_Name = @'
microsoft-teams-manage-team-members-Get-Teams
'@ 
Invoke-HelloIDDatasource -DatasourceName $dataSourceGuid_0_Name -DatasourceType "4" -DatasourceInput $tmpInput -DatasourcePsScript $tmpPsScript -DatasourceModel $tmpModel -returnObject ([Ref]$dataSourceGuid_0) 
<# End: DataSource "microsoft-teams-manage-team-members-Get-Teams" #>
<# End: HelloID Data sources #>

<# Begin: Dynamic Form "Microsoft Teams - Manage Team Members" #>
$tmpSchema = @"
[{"key":"dropDownTeam","templateOptions":{"label":"Team","required":true,"useObjects":false,"useDataSource":true,"useFilter":true,"options":["Option 1","Option 2","Option 3"],"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_0","input":{"propertyInputs":[]}},"valueField":"id","textField":"displayValue"},"type":"dropdown","summaryVisibility":"Show","textOrLabel":"text","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false},{"key":"dualListMembers","templateOptions":{"label":"Manage members","required":false,"filterable":true,"useDataSource":true,"dualList":{"options":[{"guid":"75ea2890-88f8-4851-b202-626123054e14","Name":"Apple"},{"guid":"0607270d-83e2-4574-9894-0b70011b663f","Name":"Pear"},{"guid":"1ef6fe01-3095-4614-a6db-7c8cd416ae3b","Name":"Orange"}],"optionKeyProperty":"userId","optionDisplayProperty":"displayValue"},"destinationDataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_1","input":{"propertyInputs":[{"propertyName":"team","otherFieldValue":{"otherFieldKey":"dropDownTeam"}}]}},"dataSourceConfig":{"dataSourceGuid":"$dataSourceGuid_2","input":{"propertyInputs":[]}}},"type":"duallist","summaryVisibility":"Show","sourceDataSourceIdentifierSuffix":"source-datasource","destinationDataSourceIdentifierSuffix":"destination-datasource","requiresTemplateOptions":true,"requiresKey":true,"requiresDataSource":false}]
"@ 

$dynamicFormGuid = [PSCustomObject]@{} 
$dynamicFormName = @'
Microsoft Teams - Manage Team Members
'@ 
Invoke-HelloIDDynamicForm -FormName $dynamicFormName -FormSchema $tmpSchema  -returnObject ([Ref]$dynamicFormGuid) 
<# END: Dynamic Form #>

<# Begin: Delegated Form Access Groups and Categories #>
$delegatedFormAccessGroupGuids = @()
if(-not[String]::IsNullOrEmpty($delegatedFormAccessGroupNames)){
    foreach($group in $delegatedFormAccessGroupNames) {
        try {
            $uri = ($script:PortalBaseUrl +"api/v1/groups/$group")
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
            $delegatedFormAccessGroupGuid = $response.groupGuid
            $delegatedFormAccessGroupGuids += $delegatedFormAccessGroupGuid
            
            Write-Information "HelloID (access)group '$group' successfully found$(if ($script:debugLogging -eq $true) { ": " + $delegatedFormAccessGroupGuid })"
        } catch {
            Write-Error "HelloID (access)group '$group', message: $_"
        }
    }
    if($null -ne $delegatedFormAccessGroupGuids){
        $delegatedFormAccessGroupGuids = ($delegatedFormAccessGroupGuids | Select-Object -Unique | ConvertTo-Json -Depth 100 -Compress)
    }
}
$delegatedFormCategoryGuids = @()
foreach($category in $delegatedFormCategories) {
    try {
        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories/$category")
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid
        
        Write-Information "HelloID Delegated Form category '$category' successfully found$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    } catch {
        Write-Warning "HelloID Delegated Form category '$category' not found"
        $body = @{
            name = @{"en" = $category};
        }
        $body = ConvertTo-Json -InputObject $body -Depth 100
        $uri = ($script:PortalBaseUrl +"api/v1/delegatedformcategories")
        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -ContentType "application/json" -Verbose:$false -Body $body
        $tmpGuid = $response.delegatedFormCategoryGuid
        $delegatedFormCategoryGuids += $tmpGuid
        Write-Information "HelloID Delegated Form category '$category' successfully created$(if ($script:debugLogging -eq $true) { ": " + $tmpGuid })"
    }
}
$delegatedFormCategoryGuids = (ConvertTo-Json -InputObject $delegatedFormCategoryGuids -Depth 100 -Compress)
<# End: Delegated Form Access Groups and Categories #>

<# Begin: Delegated Form #>
$delegatedFormRef = [PSCustomObject]@{guid = $null; created = $null} 
$delegatedFormName = @'
Microsoft Teams - Manage Team Members
'@
$tmpTask = @'
{"name":"Microsoft Teams - Manage Team Members","script":"# Set TLS to accept TLS, TLS 1.1 and TLS 1.2\r\n[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12\r\n\r\n$VerbosePreference = \"SilentlyContinue\"\r\n$InformationPreference = \"Continue\"\r\n$WarningPreference = \"Continue\"\r\n\r\n# Set from Global variables\r\n# $AADTenantID = \"\"\r\n# $AADAppId = \"\"\r\n# $AADAppSecret = \"\"\r\n\r\n# variables configured in form\r\n$team = $form.dropDownTeam\r\n$usersToAdd = $form.dualListMembers.leftToRight\r\n$usersToRemove = $form.dualListMembers.rightToLeft\r\n\r\n# Create authorization token and add to headers\r\ntry{\r\n    Write-Information \"Generating Microsoft Graph API Access Token\"\r\n\r\n    $baseUri = \"https://login.microsoftonline.com/\"\r\n    $authUri = $baseUri + \"$AADTenantID/oauth2/token\"\r\n\r\n    $body = @{\r\n        grant_type    = \"client_credentials\"\r\n        client_id     = \"$AADAppId\"\r\n        client_secret = \"$AADAppSecret\"\r\n        resource      = \"https://graph.microsoft.com\"\r\n    }\r\n\r\n    $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'\r\n    $accessToken = $Response.access_token;\r\n\r\n    #Add the authorization header to the request\r\n    $authorization = @{\r\n        Authorization  = \"Bearer $accesstoken\";\r\n        'Content-Type' = \"application/json\";\r\n        Accept         = \"application/json\";\r\n    }\r\n}\r\ncatch{\r\n    throw \"Could not generate Microsoft Graph API Access Token. Error: $($_.Exception.Message)\"\r\n}\r\n\r\n# Get current members because we need the membership id to remove it\r\ntry{\r\n    Write-Verbose \"Querying Members of Microsoft Team $($team.displayName) ($($team.id))\"\r\n\r\n    $baseSearchUri = \"https://graph.microsoft.com/\"\r\n    $searchUri = $baseSearchUri + \"v1.0/teams/$($team.id)/members\"\r\n\r\n    $getMicrosoftTeamMembersResponse = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false\r\n    $microsoftTeamMembers = $getMicrosoftTeamMembersResponse.value\r\n    while (![string]::IsNullOrEmpty($getMicrosoftTeamMembersResponse.'@odata.nextLink')) {\r\n        $getMicrosoftTeamMembersResponse = Invoke-RestMethod -Uri $getMicrosoftTeamMembersResponse.'@odata.nextLink' -Method Get -Headers $authorization -Verbose:$false\r\n        $microsoftTeamMembers += $getMicrosoftTeamMembersResponse.value\r\n    }\r\n\r\n    $microsoftTeamMembersGrouped = $microsoftTeamMembers | Group-Object -Property userId -AsHashTable -AsString\r\n\r\n    Write-Information \"Successfully queried Members of Microsoft Team $($team.displayName) ($($team.id)). Result count: $($microsoftTeamMembers.id.Count)\"\r\n}\r\ncatch{\r\n   Write-Error \"Could not query Members of Microsoft Team $($team.displayName) ($($team.id)). Error: $($_.Exception.Message)\"\r\n}\r\n\r\nif($usersToAdd -ne $null){\r\n    foreach($userToAdd in $usersToAdd){\r\n        try{\r\n            Write-Verbose \"Adding user $($userToAdd.displayName) ($($userToAdd.userId)) to Microsoft Team $($team.displayName) ($($team.id))\"\r\n\r\n            # Somehow can't get this to work, but the example is provided\r\n            # Use https://docs.microsoft.com/en-us/graph/api/team-post-members?view=graph-rest-1.0&tabs=http \r\n            # $baseAddUri = \"https://graph.microsoft.com/\"\r\n            # $addUri = $baseAddUri + \"v1.0/teams/$($team.id)/members\"\r\n            # $body = [PsCustomObject]@{\r\n            #     '@odata.type'       = '#microsoft.graph.aadUserConversationMember'\r\n            #     'roles'             = @(\"guest\")\r\n            #     'user@odata.bind'   = \"https://graph.microsoft.com/v1.0/users('$($userToAdd.userId)')\"                \r\n            # } | ConvertTo-Json -Depth 10\r\n            # $addMember = Invoke-RestMethod -Uri $addUri -Method Post -Headers $authorization -Verbose:$false\r\n\r\n            # Since we can't add the users to the team using the team API, make use of the groups API\r\n            # Use https://docs.microsoft.com/en-us/graph/api/team-post-members?view=graph-rest-1.0&tabs=http \r\n            $baseGraphUri = \"https://graph.microsoft.com/\"\r\n            $addGroupMembershipUri = $baseGraphUri + \"v1.0/groups/$($team.id)/members\" + '/$ref'\r\n            $body = @{ \r\n                \"@odata.id\"= \"https://graph.microsoft.com/v1.0/users/$($userToAdd.userId)\" \r\n            } | ConvertTo-Json -Depth 10\r\n            $addMember = Invoke-RestMethod -Method POST -Uri $addGroupMembershipUri -Body $body -Headers $authorization -Verbose:$false\r\n\r\n            Write-Information \"Successfully added user $($userToAdd.displayName) ($($userToAdd.userId)) to Microsoft Team $($team.displayName) ($($team.id))\"\r\n\r\n            $Log = @{\r\n                Action            = \"GrantMembership\" # optional. ENUM (undefined = default) \r\n                System            = \"MicrosoftTeams\" # optional (free format text) \r\n                Message           = \"Successfully added user $($userToAdd.displayName) ($($userToAdd.userId)) to Microsoft Team $($team.displayName) ($($team.id))\" # required (free format text) \r\n                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                TargetDisplayName = $($team.displayName) # optional (free format text)\r\n                TargetIdentifier  = $($team.id) # optional (free format text)\r\n            }\r\n            #send result back  \r\n            Write-Information -Tags \"Audit\" -MessageData $log\r\n        }catch{\r\n            $Log = @{\r\n                Action            = \"GrantMembership\" # optional. ENUM (undefined = default) \r\n                System            = \"MicrosoftTeams\" # optional (free format text) \r\n                Message           = \"Failed to add user $($userToAdd.displayName) ($($userToAdd.userId)) to Microsoft Team $($team.displayName) ($($team.id)). Error: $($_.Exception.Message)\" # required (free format text) \r\n                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                TargetDisplayName = $($team.displayName) # optional (free format text)\r\n                TargetIdentifier  = $($team.id) # optional (free format text)\r\n            }\r\n            #send result back  \r\n            Write-Information -Tags \"Audit\" -MessageData $log\r\n\r\n            Write-Error \"Could not add user $($userToAdd.displayName) ($($userToAdd.userId)) to Microsoft Team $($team.displayName) ($($team.id)). Error: $($_.Exception.Message)\"            \r\n        }\r\n    }\r\n}\r\n\r\nif($usersToRemove -ne $null){\r\n    foreach($userToRemove in $usersToRemove){\r\n         try{\r\n            Write-Verbose \"Removing user $($userToRemove.displayName) ($($userToRemove.userId)) from Microsoft Team $($team.displayName) ($($team.id))\"\r\n\r\n            $membershipId = $microsoftTeamMembersGrouped[\"$($userToRemove.userId)\"].id\r\n\r\n            if($null -eq $membershipId){\r\n                throw \"No membership found for user $($userToRemove.displayName) ($($userToRemove.userId)) to Microsoft Team $($team.displayName) ($($team.id))\"\r\n            }\r\n\r\n            $baseRemoveUri = \"https://graph.microsoft.com/\"\r\n            $removeUri = $baseRemoveUri + \"v1.0/teams/$($team.id)/members/$membershipId\"\r\n            $removeMember = Invoke-RestMethod -Uri $removeUri -Method Delete -Headers $authorization -Verbose:$false\r\n\r\n            Write-Information \"Successfully removed user $($userToRemove.displayName) ($($userToRemove.userId)) from Microsoft Team $($team.displayName) ($($team.id))\"\r\n\r\n            $Log = @{\r\n                Action            = \"RevokeMembership\" # optional. ENUM (undefined = default) \r\n                System            = \"MicrosoftTeams\" # optional (free format text) \r\n                Message           = \"Successfully removed user $($userToRemove.displayName) ($($userToRemove.userId)) from Microsoft Team $($team.displayName) ($($team.id))\" # required (free format text) \r\n                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                TargetDisplayName = $($team.displayName) # optional (free format text)\r\n                TargetIdentifier  = $($team.id) # optional (free format text)\r\n            }\r\n            #send result back  \r\n            Write-Information -Tags \"Audit\" -MessageData $log\r\n        }catch{\r\n            $Log = @{\r\n                Action            = \"RevokeMembership\" # optional. ENUM (undefined = default) \r\n                System            = \"MicrosoftTeams\" # optional (free format text) \r\n                Message           = \"Failed to remove user $($userToRemove.displayName) ($($userToRemove.userId)) from Microsoft Team $($team.displayName) ($($team.id)). Error: $($_.Exception.Message)\" # required (free format text) \r\n                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) \r\n                TargetDisplayName = $($team.displayName) # optional (free format text)\r\n                TargetIdentifier  = $($team.id) # optional (free format text)\r\n            }\r\n            #send result back  \r\n            Write-Information -Tags \"Audit\" -MessageData $log\r\n\r\n            Write-Error \"Could not remove user $($userToRemove.displayName) ($($userToRemove.userId)) from Microsoft Team $($team.displayName) ($($team.id)). Error: $($_.Exception.Message)\"            \r\n        }\r\n    }\r\n}","runInCloud":true}
'@ 

Invoke-HelloIDDelegatedForm -DelegatedFormName $delegatedFormName -DynamicFormGuid $dynamicFormGuid -AccessGroups $delegatedFormAccessGroupGuids -Categories $delegatedFormCategoryGuids -UseFaIcon "True" -FaIcon "fa fa-users" -task $tmpTask -returnObject ([Ref]$delegatedFormRef) 
<# End: Delegated Form #>

