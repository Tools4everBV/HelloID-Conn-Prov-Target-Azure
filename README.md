
# HelloID-Conn-Prov-Target-Microsoft-Entra-ID

> [!IMPORTANT]
> This repository contains the connector and configuration code only. The implementer is responsible to acquire the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements.

<p align="center">
  <img src="https://github.com/Tools4everBV/HelloID-Conn-Prov-Target-EntraID/blob/main/Logo.png?raw=true">
</p>

## Table of contents

- [HelloID-Conn-Prov-Target-Microsoft-Entra-ID](#helloid-conn-prov-target-microsoft-entra-id)
  - [Table of contents](#table-of-contents)
  - [Requirements](#requirements)
  - [Remarks](#remarks)
    - [Account Creation Limitations](#account-creation-limitations)
    - [Supported Group Types](#supported-group-types)
    - [Managing Permissions in Teams](#managing-permissions-in-teams)
    - [Creating Guest Accounts](#creating-guest-accounts)
    - [Handling Null Values in Field Mapping](#handling-null-values-in-field-mapping)
      - [Example:](#example)
  - [Introduction](#introduction)
  - [Getting started](#getting-started)
    - [Provisioning PowerShell V2 connector](#provisioning-powershell-v2-connector)
      - [Correlation configuration](#correlation-configuration)
      - [Field mapping](#field-mapping)
    - [Connection settings](#connection-settings)
  - [Connector setup](#connector-setup)
    - [Application Registration](#application-registration)
    - [Configuring App Permissions](#configuring-app-permissions)
    - [Authentication and Authorization](#authentication-and-authorization)
  - [Getting help](#getting-help)
  - [HelloID docs](#helloid-docs)

## Requirements
1. **HelloID Environment**:
   - Set up your _HelloID_ environment.
   - Install the _HelloID_ Provisioning agent (cloud or on-prem).
1. **Graph API Credentials**:
   - Create an **App Registration** in Microsoft Entra ID.
   - Add API permissions for your app:
     - **Application permissions**:
       - `User.ReadWrite.All`: Read and write all user’s full profiles.
       - `Group.ReadWrite.All`: Read and write all groups in an organization’s directory.
       - `GroupMember.ReadWrite.All`: Read and write all group memberships.
       - `UserAuthenticationMethod.ReadWrite.All`: Read and write all users’ authentication methods.
   - Create access credentials for your app:
     - Create a **client secret** for your app.

## Remarks
### Account Creation Limitations
- The [Graph API](https://learn.microsoft.com/en-us/graph/api/user-post-users?view=graph-rest-1.0&tabs=http) has limitations when creating accounts. As a result, accounts may be created without all attributes. Since the correlation value is mandatory, HelloID can correlate the account when retrying the action.

### Supported Group Types
- The [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/api/resources/groups-overview?view=graph-rest-1.0) exclusively supports  Microsoft 365 and Security groups. Mail-enabled security groups and Distribution groups cannot be managed via this API. To manage these types of groups, use the [Exchange Online connector](https://github.com/Tools4everBV/HelloID-Conn-Prov-Target-ExchangeOnline).

### Managing Permissions in Teams
- The script for dynamically managing permissions in Teams is similar to that for Groups, with an added filter for Teams-enabled groups. This is because a Team is inherently an M365 group, allowing us to manage its members within the group context rather than within Teams itself.

### Creating Guest Accounts
- Direct creation of Microsoft Entra ID Guest accounts (with login names under the tenant domain) is only supported and preferred. "Invite as Guest" is not supported.
- By specifying the `userType` as 'Guest' in the mapping, Guest accounts with login names under the tenant domain can be created effortlessly.

### Authentication Methods 
- Granting and revoking `email` and `phone` authentication methods are supported.
- Chance mapping in the `grantPermissions.ps1` according to the HelloID person model.
- Configure `OnlySetWhenEmpty` and `RemoveWhenRevokingEntitlement` settings in `permissions.ps1` if needed.
  - Revoking authentication methods can give issues when the default method is revoking before others. This is the reason that our best practice is setting this value to `$false`.

### Handling Null Values in Field Mapping
- The script filters out all field mappings with the value `$null`. If the value in the HelloID person model is `$null`, it is also filtered out. If this behavior is not desired, change the mapping to complex and ensure you return a string with a `space` or `empty` when the value is `$null`. This way, the value is correctly handled by the script.

#### Example:
```javascript
function getCompanyName() {
  let companyName = Person.PrimaryContract.Employer.Name;
  if (companyName === null) {
    companyName = ' ';
  }
  return companyName;
}
getCompanyName();
```
**
### Limitations Without Exchange Online Connector
This connector is designed exclusively for Entra ID and does not integrate with Exchange Online. As a result, it has the following limitations compared to the built-in Azure AD connector:

#### ProxyAddress Expansion with Aliases
- It cannot expand ProxyAddress with additional aliases, which is crucial for managing multiple email addresses for a single user.

> [!NOTE]
> If the `mail` and `userPrincipalName` fields are different, the `mail` value will automatically become the primary SMTP address, and the `userPrincipalName` will be added as an alias.

#### Modifying showInAddressList
- It cannot modify the showInAddressList attribute, which determines whether a user appears in the global address list (GAL).

#### Mailbox Creation/Activation
- Mailboxes cannot be created or activated until a license is assigned, causing delays in email setup for new users.

## Introduction
_HelloID-Conn-Prov-Target-Microsoft-Entra-ID_ is a _target_ connector. _Microsoft_ provides a set of REST API's that allow you to programmatically interact with its data. The Microsoft Entra ID connector uses the API endpoints listed in the table below.

| Endpoint                                                                                                                                                        | Description                                |
| --------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------ |
| [/v1.0/users/{id}](https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http)                                                          | Get a user (GET)                           |
| [/v1.0/users](https://learn.microsoft.com/en-us/graph/api/user-post-users?view=graph-rest-1.0&tabs=http)                                                        | Create user (POST)                         |
| [/v1.0/users/{id}](https://learn.microsoft.com/en-us/graph/api/user-update?view=graph-rest-1.0&tabs=http)                                                       | Update user (PATCH)                        |
| [/v1.0/users/{id}](https://learn.microsoft.com/en-us/graph/api/user-delete?view=graph-rest-1.0&tabs=http)                                                       | Delete user (DELETE)                       |
| [/v1/groups](https://learn.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0&tabs=http)                                                              | List groups (GET)                          |
| [/v1/groups/{group-id}/members/$ref](https://learn.microsoft.com/en-us/graph/api/group-post-members?view=graph-rest-1.0&tabs=http)                              | Add member (POST)                          |
| [/v1/groups/{id}/members/{id}/$ref](https://learn.microsoft.com/en-us/graph/api/group-delete-members?view=graph-rest-1.0&tabs=http)                             | Remove member (DELETE)                     |
| [/v1/groups](https://learn.microsoft.com/en-us/graph/api/group-post-groups?view=graph-rest-1.0&tabs=http)                                                       | Create group (POST)                        |
| [/v1.0/users/{id}/authentication/emailMethods/{id}](https://learn.microsoft.com/nl-nl/graph/api/emailauthenticationmethod-get?view=graph-rest-1.0&tabs=http)    | Get emailAuthenticationMethod (GET)        |
| [/v1.0/users/{id}/authentication/emailMethods](https://learn.microsoft.com/nl-nl/graph/api/authentication-post-emailmethods?view=graph-rest-1.0&tabs=http)      | Create emailMethod (POST)                  |
| [/v1.0/users/{id}/authentication/emailMethods/{id}](https://learn.microsoft.com/nl-nl/graph/api/emailauthenticationmethod-update?view=graph-rest-1.0&tabs=http) | Update emailAuthenticationMethod (PATCH)   |
| [/v1.0/users/{id}/authentication/emailMethods/{id}](https://learn.microsoft.com/nl-nl/graph/api/emailauthenticationmethod-delete?view=graph-rest-1.0&tabs=http) | UDelete emailAuthenticationMethod (DELETE) |
| [/v1.0/users/{id}/authentication/phoneMethods/{id}](https://learn.microsoft.com/nl-nl/graph/api/phoneauthenticationmethod-get?view=graph-rest-1.0&tabs=http)    | Get phoneAuthenticationMethod (GET)        |
| [/v1.0/users/{id}/authentication/phoneMethods](https://learn.microsoft.com/nl-nl/graph/api/authentication-post-phonemethods?view=graph-rest-1.0&tabs=http)      | Create phoneMethod (POST)                  |
| [/v1.0/users/{id}/authentication/phoneMethods/{id}](https://learn.microsoft.com/nl-nl/graph/api/phoneauthenticationmethod-update?view=graph-rest-1.0&tabs=http) | Update phoneAuthenticationMethod (PATCH)   |
| [/v1.0/users/{id}/authentication/phoneMethods/{id}](https://learn.microsoft.com/nl-nl/graph/api/phoneauthenticationmethod-delete?view=graph-rest-1.0&tabs=http) | UDelete phoneAuthenticationMethod (DELETE) |


The following lifecycle actions are available:

| Action                                            | Description                                                                                |
| ------------------------------------------------- | ------------------------------------------------------------------------------------------ |
| create.ps1                                        | Create or correlate to an account                                                          |
| delete.ps1                                        | Delete an account                                                                          |
| disable.ps1                                       | Disable an account                                                                         |
| enable.ps1                                        | Enable an account                                                                          |
| update.ps1                                        | Update an account                                                                          |
| uniquenessCheck.json                              | Default _uniquenessCheck.json_                                                             |
| groups - permissions.ps1                          | List groups as permissions                                                                 |
| groups - grantPermission.ps1                      | Grant groupmembership to an account                                                        |
| groups - revokePermission.ps1                     | Revoke groupmembership from an account                                                     |
| groups - resources.ps1                            | Create groups from resources                                                               |
| groups - subPermissions.ps1                       | Grant/Revoke groupmembership from an account                                               |
| emailAuthenticationMethods - permissions.ps1      | List emailAuthenticationMethods as permissions                                             |
| emailAuthenticationMethods - grantPermission.ps1  | Grant emailAuthenticationMethod to an account                                              |
| emailAuthenticationMethods - revokePermission.ps1 | Revoke emailAuthenticationMethod from an account                                           |
| emailAuthenticationMethods - configuration.json   | Additional _configuration.json_ with settings specifically for emailAuthenticationMethods  |
| phoneAuthenticationMethods - permissions.ps1      | List phoneAuthenticationMethods as permissions                                             |
| phoneAuthenticationMethods - grantPermission.ps1  | Grant phoneAuthenticationMethod to an account                                              |
| phoneAuthenticationMethods - revokePermission.ps1 | Revoke phoneAuthenticationMethod from an account                                           |
| phoneAuthenticationMethods - configuration.json   | Additional _configuration.json_ with settings specifically for phoneAuthenticationMethods  |
| configuration.json                                | Default _configuration.json_                                                               |
| fieldMapping.json                                 | _fieldMapping.json_ for when using the the full account lifecycle                          |
| fieldMapping.correlateOnly.json                   | _fieldMapping.json_ for when only using the correlation and not the full account lifecycle |

## Getting started
By using this connector you will have the ability to seamlessly create and user accounts and groups in Microsoft Entra ID. Additionally, you can set the MFA phone or email settings.

Connecting to Microsoft the Microsoft Graph API is straightforward. Simply utilize the API Key and API Secret pair.
For further details, refer to the following pages in the Microsoft Docs:

- [Use the Microsoft Graph API](https://learn.microsoft.com/en-us/graph/use-the-api).
- [User Properties](https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0#properties).
- [Supported User Properties for Correlation](https://learn.microsoft.com/en-us/graph/aad-advanced-queries?tabs=http#user-properties).

### Provisioning PowerShell V2 connector

#### Correlation configuration
The correlation configuration is used to specify which properties will be used to match an existing account within _Microsoft Entra ID_ to a person in _HelloID_.

To properly setup the correlation:

1. Open the `Correlation` tab.

2. Specify the following configuration:

    | Setting                   | Value        |
    | ------------------------- | ------------ |
    | Enable correlation        | `True`       |
    | Person correlation field  | `ExternalId` |
    | Account correlation field | `employeeId` |

> [!IMPORTANT]
> The account correlation field is added to the create action. If you use a different value then `employeeId`, please make sure this is support by the [graph api](https://learn.microsoft.com/en-us/graph/api/user-post-users?view=graph-rest-1.0&tabs=http)


> [!TIP]
> _For more information on correlation, please refer to our correlation [documentation](https://docs.helloid.com/en/provisioning/target-systems/powershell-v2-target-systems/correlation.html) pages_.

#### Field mapping
The field mapping can be imported by using the _fieldMapping.json_ file.

### Connection settings
The following settings are required to connect to the API.

| Setting                                                        | Description                                                                                                                               | Mandatory |
| -------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- | --------- |
| App Registration Directory (tenant) ID                         | The ID to the Tenant in Microsoft Entra ID                                                                                                | Yes       |
| App Registration Application (client) ID                       | The ID to the App Registration in Microsoft Entra ID                                                                                      | Yes       |
| App Registration Client Secret                                 | The Client Secret to the App Registration in Microsoft Entra ID                                                                           | Yes       |
| Invite as Guest                                                | When toggled, this connector will create Guest accounts through invitation, allowing users to log in using their invited email addresses. | No        |
| Delete the account when revoking the entitlement               | When toggled, this delete accounts when revoking the account entitlement.                                                                 | No        |
| Set primary manager when an account is created                 | When toggled, this connector will calculate and set the manager upon creating an account.                                                 | No        |
| Update manager when the account updated operation is performed | When toggled, this connector will calculate and set the manager upon updating an account.                                                 | No        |
| IsDebug                                                        | When toggled, extra logging is shown. Note that this is only meant for debugging, please switch this off when in production.              | No        |


## Connector setup
### Application Registration
The first step to connect to the Graph API and make requests is to register a new **Microsoft Entra ID Application**. This application will be used to connect to the API and manage permissions.

Follow these steps:

1. **Navigate to App Registrations**:
   - Go to the Microsoft Entra ID Portal.
   - Navigate to **Microsoft Entra ID** > **App registrations**.
   - Click on **New registration**.

2. **Register the Application**:
   - **Name**: Enter a name for your application (e.g., "HelloID PowerShell").
   - **Supported Account Types**: Choose who can use this application (e.g., "Accounts in this organizational directory only").
   - **Redirect URI**: Choose the platform as `Web` and enter a redirect URI (e.g., `http://localhost`).

3. **Complete the Registration**:
   - Click the **Register** button to create your new application.

For more detailed instructions, please see the official Microsoft documentation: [Quickstart: Register an app in the Microsoft identity platform](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app?tabs=certificate).

### Configuring App Permissions
Next, configure the necessary API permissions for your Microsoft Entra ID application. For this connector, we use the **Microsoft Graph API**.

Follow these steps:

1. In your Microsoft Entra ID application, navigate to the **API Permissions** section.
2. Click on **Add a permission**.
3. Select **Microsoft Graph**.
4. Choose **Application permissions** and add the following:
   - `User.ReadWrite.All`: Read and write all user’s full profiles.
   - `Group.ReadWrite.All`: Read and write all groups in an organization’s directory.
   - `GroupMember.ReadWrite.All`: Read and write all group memberships.
   - `UserAuthenticationMethod.ReadWrite.All`: Read and write all users’ authentication methods.
5. Click **Add permissions**.
6. If required, click on **Grant admin consent for [Your Tenant]** to grant the necessary permissions.

For more detailed instructions, please see the official Microsoft documentation: [Quickstart: Configure a client application to access a web API](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-configure-app-access-web-apis).

### Authentication and Authorization
To authenticate to the Graph API using the Authorization Code grant type, you need to obtain the necessary credentials. We recommend using the Client secret.

Follow these steps:

1. **Get the Tenant ID**:
   - In the Microsoft Entra ID Portal, go to **Azure Active Directory** > **Overview**.
   - Copy the **Tenant ID** from the Overview page.

2. **Get the Client ID**:
   - Go to the Microsoft Entra ID Portal.
   - Navigate to **Azure Active Directory** > **App registrations**.
   - Select your application and copy the **Application (client) ID** value.

3. **Create a Client Secret**:
   - In the Microsoft Entra ID Portal, go to **Azure Active Directory** > **App registrations**.
   - Select the application you created earlier.
   - Navigate to **Certificates & secrets**.
   - Under **Client secrets**, click on **New client secret**.
   - Provide a description for your secret and select an expiration date.
   - Click **Add** and copy the newly generated client secret. **Important**: You cannot view the client secret value again after you close the page, so make sure to copy it immediately.

For more detailed instructions, please see the official Microsoft documentation: [Add credentials](https://learn.microsoft.com/en-us/graph/auth-register-app-v2#add-credentials).

## Getting help
> [!TIP]
> _For more information on how to configure a HelloID PowerShell connector, please refer to our [documentation](https://docs.helloid.com/en/provisioning/target-systems/powershell-v2-target-systems.html) pages_.

> [!TIP]
>  _If you need help, feel free to ask questions on our [forum](https://forum.helloid.com)_.

## HelloID docs
The official HelloID documentation can be found at: https://docs.helloid.com/
