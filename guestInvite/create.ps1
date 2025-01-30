#################################################
# HelloID-Conn-Prov-Target-Microsoft-Entra-GuestInvite-Create
# Create and update and optionally set manager or correlate to account
# PowerShell V2
#################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Set debug logging
switch ($actionContext.Configuration.isDebug) {
    $true { $VerbosePreference = "Continue" }
    $false { $VerbosePreference = "SilentlyContinue" }
}
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#region functions
function Resolve-MicrosoftGraphAPIError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object] $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }

        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception -is [System.Net.WebException] -and $ErrorObject.Exception.Response) {
            $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
            if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                $httpErrorObj.ErrorDetails = $streamReaderResponse
            }
        }

        try {
            $errorObjectConverted = $httpErrorObj.ErrorDetails | ConvertFrom-Json -ErrorAction Stop

            if ($errorObjectConverted.error_description) {
                $httpErrorObj.FriendlyMessage = $errorObjectConverted.error_description
            }
            elseif ($errorObjectConverted.error) {
                $httpErrorObj.FriendlyMessage = $errorObjectConverted.error.message
                if ($errorObjectConverted.error.code) {
                    $httpErrorObj.FriendlyMessage += " Error code: $($errorObjectConverted.error.code)."
                }
                if ($errorObjectConverted.error.details) {
                    if ($errorObjectConverted.error.details.message) {
                        $httpErrorObj.FriendlyMessage += " Details message: $($errorObjectConverted.error.details.message)"
                    }
                    if ($errorObjectConverted.error.details.code) {
                        $httpErrorObj.FriendlyMessage += " Details code: $($errorObjectConverted.error.details.code)."
                    }
                }
            }
            else {
                $httpErrorObj.FriendlyMessage = $ErrorObject
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }

        Write-Output $httpErrorObj
    }
}

function Convert-StringToBoolean($obj) {
    foreach ($property in $obj.PSObject.Properties) {
        $value = $property.Value
        if ($value -is [string]) {
            try {
                $obj.$($property.Name) = [System.Convert]::ToBoolean($value)
            }
            catch {
                # Handle cases where conversion fails
                $obj.$($property.Name) = $value
            }
        }
    }
    return $obj
}
#endregion functions

try {
    #region account
    # Define correlation
    $correlationField = $actionContext.CorrelationConfiguration.accountField
    $correlationValue = $actionContext.CorrelationConfiguration.accountFieldValue

    # Define account object
    $account = [PSCustomObject]$actionContext.Data.PsObject.Copy()

    # Define properties to query
    $accountPropertiesToQuery = @("id") + $account.PsObject.Properties.Name | Select-Object -Unique

    # Remove properties of account object with null-values
    $account.PsObject.Properties | ForEach-Object {
        # Remove properties with null-values
        if ($_.Value -eq $null) {
            $account.PsObject.Properties.Remove("$($_.Name)")
        }
    }
    # Convert the properties of account object containing "TRUE" or "FALSE" to boolean 
    $account = Convert-StringToBoolean $account
    #endRegion account

    #region Verify correlation configuration and properties
    $actionMessage = "verifying correlation configuration and properties"

    if ($actionContext.CorrelationConfiguration.Enabled -eq $true) {
        if ([string]::IsNullOrEmpty($correlationField)) {
            throw "Correlation is enabled but not configured correctly."
        }
        
        if ([string]::IsNullOrEmpty($correlationValue)) {
            throw "The correlation value for [$correlationField] is empty. This is likely a mapping issue."
        }
    }
    else {
        Write-Warning "Correlation is disabled."
    }
    #endregion Verify correlation configuration and properties
    
    #region Create access token
    $actionMessage = "creating access token"
    
    $createAccessTokenBody = @{
        grant_type    = "client_credentials"
        client_id     = $actionContext.Configuration.AppId
        client_secret = $actionContext.Configuration.AppSecret
        resource      = "https://graph.microsoft.com"
    }
    
    $createAccessTokenSplatParams = @{
        Uri         = "https://login.microsoftonline.com/$($actionContext.Configuration.TenantID)/oauth2/token"
        Headers     = $headers
        Body        = $createAccessTokenBody
        Method      = "POST"
        ContentType = "application/x-www-form-urlencoded"
        Verbose     = $false
        ErrorAction = "Stop"
    }
    
    $createAccessTokenResonse = Invoke-RestMethod @createAccessTokenSplatParams
    
    Write-Verbose "Created access token. Expires in: $($createAccessTokenResonse.expires_in | ConvertTo-Json)"
    #endregion Create access token
    
    #region Create headers
    $actionMessage = "creating headers"
    
    $headers = @{
        "Accept"          = "application/json"
        "Content-Type"    = "application/json;charset=utf-8"
        "Mwp-Api-Version" = "1.0"
    }
    
    Write-Verbose "Created headers. Result (without Authorization): $($headers | ConvertTo-Json)."

    # Add Authorization after printing splat
    $headers['Authorization'] = "Bearer $($createAccessTokenResonse.access_token)"
    #endregion Create headers
    
    if ($actionContext.CorrelationConfiguration.Enabled -eq $true) {
        #region Get account
        # API docs: https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
        $actionMessage = "querying account where [$($correlationField)] = [$($correlationValue)]"

        $getEntraIDAccountSplatParams = @{
            Uri         = "https://graph.microsoft.com/v1.0/users?`$filter=$correlationField eq '$correlationValue'&`$select=$($accountPropertiesToQuery -join ',')"
            Method      = "GET"
            Verbose     = $false
            ErrorAction = "Stop"
        }

        Write-Verbose "SplatParams: $($getEntraIDAccountSplatParams | ConvertTo-Json)"

        # Add Headers after printing splat
        $getEntraIDAccountSplatParams['Headers'] = $headers

        $getEntraIDAccountResponse = $null
        $getEntraIDAccountResponse = Invoke-RestMethod @getEntraIDAccountSplatParams
        $correlatedAccount = $getEntraIDAccountResponse.Value
            
        Write-Verbose "Queried account where [$($correlationField)] = [$($correlationValue)]. Result: $($correlatedAccount  | ConvertTo-Json)"
        #endregion Get account
    }
    else {
        $correlatedAccount = $null
    }

    #region Calulate action
    $actionMessage = "calculating action"
    if (($correlatedAccount | Measure-Object).count -eq 1) {
        $actionAccount = "Correlate"
    }
    elseif (($correlatedAccount | Measure-Object).count -eq 0) {
        $actionAccount = "GuestInvite"
    }
    elseif (($correlatedAccount | Measure-Object).count -gt 1) {
        $actionAccount = "MultipleFound"
    }
    #endregion Calulate action
    
    #region Process
    switch ($actionAccount) {
        "GuestInvite" {
            #region Create invitation
            # Microsoft docs: https://learn.microsoft.com/en-us/graph/api/invitation-post?view=graph-rest-1.0&tabs=http
            $actionMessage = "creating invitation"
            $baseUri = "https://graph.microsoft.com/"

            # GuestInvite account with only required fields
            $requiredFields = @("invitedUserEmailAddress", "invitedUserDisplayName", "invitedUserMessageInfo", "inviteRedirectUrl", "sendInvitationMessage")
            $createInvitationBody = @{}
            foreach ($accountProperty in $account.PsObject.Properties | Where-Object { $null -ne $_.Value -and $_.Name -in $requiredFields }) {
                [void]$createInvitationBody.Add($accountProperty.Name, $accountProperty.Value)
            }
            
            $createInvitationSplatParams = @{
                Uri         = "$($baseUri)/v1.0/invitations"
                Method      = "POST"
                Body        = ($createInvitationBody | ConvertTo-Json -Depth 10)
                ContentType = 'application/json; charset=utf-8'
                Verbose     = $false
                ErrorAction = "Stop"
            }

            Write-Verbose "SplatParams: $($createInvitationSplatParams | ConvertTo-Json)"

            if (-Not($actionContext.DryRun -eq $true)) {
                # Add header after printing splat
                $createInvitationSplatParams['Headers'] = $headers

                $createdInvitation = Invoke-RestMethod @createInvitationSplatParams

                #region Set AccountReference and add AccountReference to Data
                $outputContext.AccountReference = "$($createdInvitation.invitedUser.id)"
                $outputContext.Data.id = "$($createdInvitation.invitedUser.id)"
                #endregion Set AccountReference and add AccountReference to Data

                $outputContext.AuditLogs.Add([PSCustomObject]@{
                        # Action  = "" # Optional
                        Message = "Created invitiation for user with displayName [$($createInvitationBody.invitedUserDisplayName)] and emailAddress [$($createInvitationBody.invitedUserEmailAddress)] with AccountReference: $($outputContext.AccountReference | ConvertTo-Json)."
                        IsError = $false
                    })
            }
            else {
                Write-Warning "DryRun: Would create invitiation for user with displayName [$($createInvitationBody.invitedUserDisplayName)] and emailAddress [$($createInvitationBody.invitedUserEmailAddress)]."
            }            
            #endregion Create account

            #region Update account
            # API docs: https://learn.microsoft.com/en-us/graph/api/user-update?view=graph-rest-1.0&tabs=http
            $actionMessage = "updating created account with AccountReference: $($outputContext.AccountReference | ConvertTo-Json)"

            # Update account with all other fields than the required fields
            $updateAccountBody = [PSCustomObject]@{}
            foreach ($accountProperty in $account.PsObject.Properties | Where-Object { $null -ne $_.Value -and $_.Name -notin $requiredFields }) {
                $updateAccountBody | Add-Member -MemberType NoteProperty -Name $accountProperty.Name -Value $accountProperty.Value -Force
            }


            $updateAccountSplatParams = @{
                Uri         = "https://graph.microsoft.com/v1.0/users/$($outputContext.AccountReference)"
                Method      = "PATCH"
                Body        = ($updateAccountBody | ConvertTo-Json -Depth 10)
                ContentType = 'application/json; charset=utf-8'
                Verbose     = $false
                ErrorAction = "Stop"
            }

            Write-Verbose "SplatParams: $($updateAccountSplatParams | ConvertTo-Json)"

            if (-Not($actionContext.DryRun -eq $true)) {
                # Add Headers after printing splat
                $updateAccountSplatParams['Headers'] = $headers

                $updateAccountResponse = Invoke-RestMethod @updateAccountSplatParams

                $outputContext.AuditLogs.Add([PSCustomObject]@{
                        # Action  = "" # Optional
                        Message = "Updated created account with AccountReference: $($outputContext.AccountReference | ConvertTo-Json). Updated properties [$($updateAccountBody | ConvertTo-Json)."
                        IsError = $false
                    })
            }
            else {
                Write-Warning "DryRun: Would update created account. Updated properties [$($updateAccountBody | ConvertTo-Json)."
            }
            #endregion Update account

            #region Manager
            if ($actionContext.Configuration.setManagerOnCreate -eq $true) {
                #region Get manageraccount
                # API docs: https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
                # Managerreference of the Guest account is not present in targetconnector of this type and will be searched for in EntraID based on employeeId
                $actionMessage = "querying account where [employeeId] = [$($personContext.Manager.ExternalId)]"

                $getEntraIDManagerAccountSplatParams = @{
                    Uri         = "https://graph.microsoft.com/v1.0/users?`$filter=employeeId eq '$($personContext.Manager.ExternalId)'&`$select=$($accountPropertiesToQuery -join ',')"
                    Method      = "GET"
                    Verbose     = $false
                    ErrorAction = "Stop"
                }

                Write-Verbose "SplatParams: $($getEntraIDManagerAccountSplatParams | ConvertTo-Json)"

                # Add Headers after printing splat
                $getEntraIDManagerAccountSplatParams['Headers'] = $headers

                $getEntraIDManagerAccountResponse = $null
                $getEntraIDManagerAccountResponse = Invoke-RestMethod @getEntraIDManagerAccountSplatParams
                $correlatedManagerAccount = $getEntraIDManagerAccountResponse.Value
                    
                Write-Verbose "Queried account where [employeeId] = [$($personContext.Manager.ExternalId)]. Result: $($correlatedManagerAccount  | ConvertTo-Json)"
                
                if ($correlatedManagerAccount) {
                    $actionContext.References.ManagerAccount = $correlatedManagerAccount.id
                }
                #endregion Get account

                #region Calulate manager action
                if (-not[String]::IsNullOrEmpty(($actionContext.References.ManagerAccount))) {
                    $actionManager = "Set"
                }
                else {
                    $actionManager = "AccountReferenceEmpty"
                }

                #endregion Calulate manager action

                switch ($actionManager) {
                    "Set" {
                        #region Set Manager
                        # API docs: https://learn.microsoft.com/en-us/graph/api/user-post-manager?view=graph-rest-1.0&tabs=http
                        $actionMessage = "setting manager for created account with AccountReference: $($outputContext.AccountReference | ConvertTo-Json)"

                        $setManagerBody = @{
                            "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($actionContext.References.ManagerAccount)"
                        }

                        $setManagerSplatParams = @{
                            Uri         = "https://graph.microsoft.com/v1.0/users/$($outputContext.AccountReference)/manager/`$ref"
                            Method      = "PUT"
                            Body        = ($setManagerBody | ConvertTo-Json -Depth 10)
                            Verbose     = $false
                            ErrorAction = "Stop"
                        }

                        Write-Verbose "SplatParams: $($setManagerSplatParams | ConvertTo-Json)"

                        if (-Not($actionContext.DryRun -eq $true)) {
                            # Add Headers after printing splat
                            $setManagerSplatParams['Headers'] = $headers

                            $setManagerResponse = Invoke-RestMethod @setManagerSplatParams

                            #region Set ManagerId
                            $outputContext.Data | Add-Member -MemberType NoteProperty -Name "managerId" -Value $actionContext.References.ManagerAccount -Force
                            #endregion Set ManagerId

                            $outputContext.AuditLogs.Add([PSCustomObject]@{
                                    # Action  = "" # Optional
                                    Message = "Set manager for created account with AccountReference: $($outputContext.AccountReference | ConvertTo-Json). New value: $($actionContext.References.ManagerAccount | ConvertTo-Json)."
                                    IsError = $false
                                })
                        }
                        else {
                            Write-Warning "DryRun: Would set manager for created account. New value: $($actionContext.References.ManagerAccount | ConvertTo-Json)"
                        }
                        #endregion Set Manager

                        break
                    }

                    "AccountReferenceEmpty" {
                        #region No account found
                        $actionMessage = "setting manager for created account"

                        $outputContext.AuditLogs.Add([PSCustomObject]@{
                                # Action  = "" # Optional
                                Message = "Skipped setting manager for created account with AccountReference: $($outputContext.AccountReference | ConvertTo-Json). Reason: Manager AccountReference is empty."
                                IsError = $false
                            })
                        #endregion No account found
        
                        break
                    }
                }
            }
            #endregion Manager

            break
        }

        "Correlate" {
            #region Correlate account
            $actionMessage = "correlating to account"

            $outputContext.AccountReference = "$($correlatedAccount.id)"
            $outputContext.Data = $correlatedAccount.PsObject.Copy()

            $outputContext.AuditLogs.Add([PSCustomObject]@{
                    Action  = "CorrelateAccount" # Optionally specify a different action for this audit log
                    Message = "Correlated to account with AccountReference: $($outputContext.AccountReference | ConvertTo-Json) on [$($correlationField)] = [$($correlationValue)]."
                    IsError = $false
                })

            $outputContext.AccountCorrelated = $true
            #endregion Correlate account

            break
        }

        "MultipleFound" {
            #region Multiple accounts found
            $actionMessage = "correlating to account"

            # Throw terminal error
            throw "Multiple accounts found where [$($correlationField)] = [$($correlationValue)]. Please correct this so the persons are unique."
            #endregion Multiple accounts found

            break
        }
    }
    #endregion Process
}
catch {
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftGraphAPIError -ErrorObject $ex
        $auditMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        $warningMessage = "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
        $warningMessage = "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }

    Write-Warning $warningMessage

    $outputContext.AuditLogs.Add([PSCustomObject]@{
            # Action  = "" # Optional
            Message = $auditMessage
            IsError = $true
        })
}
finally {
    # Check if auditLogs contains errors, if no errors are found, set success to true
    if (-NOT($outputContext.AuditLogs.IsError -contains $true)) {
        $outputContext.Success = $true
    }

    # Check if accountreference is set, if not set, set this with default value as this must contain a value
    if ([String]::IsNullOrEmpty($outputContext.AccountReference) -and $actionContext.DryRun -eq $true) {
        $outputContext.AccountReference = "DryRun: Currently not available"
    }
}