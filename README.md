| :information_source: Information |
|:---------------------------|
| This repository contains the connector and configuration code only. The implementer is responsible to acquire the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements.       |
<br />

<p align="center">
  <img src="https://user-images.githubusercontent.com/69046642/175026144-bab6fb5f-f20a-4780-93ac-311ced56707d.png"
       width="400px", height="400px">
</p>


<!-- Description -->
## Description
This HelloID Service Automation Delegated Form can manage the groupmemberships for Azure users. The following options are available:
 1. Search and select the target team
 2. Choose the users to add or remove
 3. After confirmation the memberhips of the selected team are updated


## Versioning
| Version | Description | Date |
| - | - | - |
| 1.0.0   | Initial release | 2022/06/22  |


<!-- TABLE OF CONTENTS -->
## Table of Contents
- [Description](#description)
- [Versioning](#versioning)
- [Table of Contents](#table-of-contents)
- [Introduction](#introduction)
- [Requirements](#requirements)
- [Getting the Teams graph API access](#getting-the-teams-graph-api-access)
  - [Application Registration](#application-registration)
  - [Configuring App Permissions](#configuring-app-permissions)
  - [Authentication and Authorization](#authentication-and-authorization)
- [All-in-one PowerShell setup script](#all-in-one-powershell-setup-script)
  - [Getting started](#getting-started)
- [Post-setup configuration](#post-setup-configuration)
- [Manual resources](#manual-resources)
  - [Powershell data source 'microsoft-teams-manage-team-members-Get-Teams'](#powershell-data-source-microsoft-teams-manage-team-members-get-teams)
  - [Powershell data source 'microsoft-teams-manage-team-members-Get-Users'](#powershell-data-source-microsoft-teams-manage-team-members-get-users)
  - [Powershell data source 'microsoft-teams-manage-team-members-Get-Team-Members'](#powershell-data-source-microsoft-teams-manage-team-members-get-team-members)
  - [Delegated form task 'Microsoft Teams - Manage Team Members'](#delegated-form-task-microsoft-teams---manage-team-members)
- [Getting help](#getting-help)
- [HelloID Docs](#helloid-docs)


## Introduction
The interface to communicate with Microsoft Teams is through the Microsoft Graph API.

<!-- Requirements -->
## Requirements
This script uses the Microsoft Graph API and requires an App Registration with App permissions:
*	Read and write all users' full profiles by using <b><i>User.ReadWrite.All</i></b>
*	Read and write all groups by using <b><i>Group.ReadWrite.All</i></b>
*	Add and remove members from all teams by using <b><i>TeamMember.ReadWrite.All</i></b>

<!-- GETTING STARTED -->
## Getting the Teams graph API access

By using this connector you will have the ability to manage the memberships of a Microsoft Team.

### Application Registration
The first step to connect to Graph API and make requests, is to register a new <b>Azure Active Directory Application</b>. The application is used to connect to the API and to manage permissions.

* Navigate to <b>App Registrations</b> in Azure, and select “New Registration” (<b>Azure Portal > Azure Active Directory > App Registration > New Application Registration</b>).
* Next, give the application a name. In this example we are using “<b>HelloID PowerShell</b>” as application name.
* Specify who can use this application (<b>Accounts in this organizational directory only</b>).
* Specify the Redirect URI. You can enter any url as a redirect URI value. In this example we used http://localhost because it doesn't have to resolve.
* Click the “<b>Register</b>” button to finally create your new application.

Some key items regarding the application are the Application ID (which is the Client ID), the Directory ID (which is the Tenant ID) and Client Secret.

### Configuring App Permissions
The [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph) provides details on which permission are required for each permission type.

To assign your application the right permissions, navigate to <b>Azure Portal > Azure Active Directory >App Registrations</b>.
Select the application we created before, and select “<b>API Permissions</b>” or “<b>View API Permissions</b>”.
To assign a new permission to your application, click the “<b>Add a permission</b>” button.
From the “<b>Request API Permissions</b>” screen click “<b>Microsoft Graph</b>”.
For this connector the following permissions are used as <b>Application permissions</b>:
*	Read and Write all user’s full profiles by using <b><i>User.ReadWrite.All</i></b>
*	Read and Write all groups in an organization’s directory by using <b><i>Group.ReadWrite.All</i></b>
*	Add and remove members from all teams by using <b><i>TeamMember.ReadWrite.All</i></b>

Some high-privilege permissions can be set to admin-restricted and require an administrators consent to be granted.

To grant admin consent to our application press the “<b>Grant admin consent for TENANT</b>” button.

### Authentication and Authorization
There are multiple ways to authenticate to the Graph API with each has its own pros and cons, in this example we are using the Authorization Code grant type.

*	First we need to get the <b>Client ID</b>, go to the <b>Azure Portal > Azure Active Directory > App Registrations</b>.
*	Select your application and copy the Application (client) ID value.
*	After we have the Client ID we also have to create a <b>Client Secret</b>.
*	From the Azure Portal, go to <b>Azure Active Directory > App Registrations</b>.
*	Select the application we have created before, and select "<b>Certificates and Secrets</b>". 
*	Under “Client Secrets” click on the “<b>New Client Secret</b>” button to create a new secret.
*	Provide a logical name for your secret in the Description field, and select the expiration date for your secret.
*	It's IMPORTANT to copy the newly generated client secret, because you cannot see the value anymore after you close the page.
*	At least we need to get is the <b>Tenant ID</b>. This can be found in the Azure Portal by going to <b>Azure Active Directory > Custom Domain Names</b>, and then finding the .onmicrosoft.com domain.


## All-in-one PowerShell setup script
The PowerShell script "createform.ps1" contains a complete PowerShell script using the HelloID API to create the complete Form including user defined variables, tasks and data sources.

_Please note that this script asumes none of the required resources do exists within HelloID. The script does not contain versioning or source control_

### Getting started
Please follow the documentation steps on [HelloID Docs](https://docs.helloid.com/hc/en-us/articles/360017556559-Service-automation-GitHub-resources) in order to setup and run the All-in one Powershell Script in your own environment.


## Post-setup configuration
After the all-in-one PowerShell script has run and created all the required resources. The following items need to be configured according to your own environment
 1. Update the following [user defined variables](https://docs.helloid.com/hc/en-us/articles/360014169933-How-to-Create-and-Manage-User-Defined-Variables)

| Variable name                 | Description                                                   | Notes                        |
| ----------------------------- | ------------------------------------------------------------- | ---------------------------- |
| AADtenantID                   | Id of the Azure tenant                                        | **Define as Global Varaible**    |
| AADAppId                      | Id of the Azure app                                           | **Define as Global Varaible**    |
| AADAppSecret                  | Secret of the Azure app                                       | **Define as Global Varaible**    |

## Manual resources
This Delegated Form uses the following resources in order to run

### Powershell data source 'microsoft-teams-manage-team-members-Get-Teams'

### Powershell data source 'microsoft-teams-manage-team-members-Get-Users'

### Powershell data source 'microsoft-teams-manage-team-members-Get-Team-Members'

### Delegated form task 'Microsoft Teams - Manage Team Members'

## Getting help
_If you need help, feel free to ask questions on our [forum](https://forum.helloid.com/forum/helloid-connectors/service-automation/951-helloid-sa-microsoft-teams-manage-members)_

## HelloID Docs
The official HelloID documentation can be found at: https://docs.helloid.com/
