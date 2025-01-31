#################################################
# HelloID-Conn-Prov-Target-Microsoft-Entra-ID-Permissions-perUserMfaState-Revoke
# disable per user mfa state (uses beta endpoint)
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

function New-AuthorizationHeaders {
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.Dictionary[[String], [String]]])]
    param(
        [parameter(Mandatory)]
        [string]
        $TenantId,

        [parameter(Mandatory)]
        [string]
        $ClientId,

        [parameter(Mandatory)]
        [string]
        $ClientSecret
    )
    try {
        Write-Verbose "Creating Access Token"
        $baseUri = "https://login.microsoftonline.com/"
        $authUri = $baseUri + "$TenantId/oauth2/token"
    
        $body = @{
            grant_type    = "client_credentials"
            client_id     = "$ClientId"
            client_secret = "$ClientSecret"
            resource      = "https://graph.microsoft.com"
        }
    
        $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
        $accessToken = $Response.access_token
    
        #Add the authorization header to the request
        Write-Verbose 'Adding Authorization headers'

        $headers = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
        $headers.Add('Authorization', "Bearer $accesstoken")
        $headers.Add('Accept', 'application/json')
        $headers.Add('Content-Type', 'application/json')
        # Needed to filter on specific attributes (https://docs.microsoft.com/en-us/graph/aad-advanced-queries)
        $headers.Add('ConsistencyLevel', 'eventual')

        Write-Output $headers  
    }
    catch {
        throw $_
    }
}

function Resolve-HTTPError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline
        )]
        [object]$ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            FullyQualifiedErrorId = $ErrorObject.FullyQualifiedErrorId
            MyCommand             = $ErrorObject.InvocationInfo.MyCommand
            RequestUri            = $ErrorObject.TargetObject.RequestUri
            ScriptStackTrace      = $ErrorObject.ScriptStackTrace
            ErrorMessage          = ''
        }
        if ($ErrorObject.Exception.GetType().FullName -eq 'Microsoft.Powershell.Commands.HttpResponseException') {
            $httpErrorObj.ErrorMessage = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            $httpErrorObj.ErrorMessage = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
        }
        Write-Output $httpErrorObj
    }
}
#endregion functions

try {
    $state = 'disabled'

    #region Verify account reference and required properties
    $actionMessage = "verifying account reference and required properties"
    if ([string]::IsNullOrEmpty($($actionContext.References.Account))) {
        throw "The account reference could not be found"
    }

    #endregion Verify account reference and required properties

    #region Create authorization headers
    $actionMessage = "creating authorization headers"

    $authorizationHeadersSplatParams = @{
        TenantId     = $actionContext.Configuration.TenantID
        ClientId     = $actionContext.Configuration.AppId
        ClientSecret = $actionContext.Configuration.AppSecret
    }

    $headers = New-AuthorizationHeaders @authorizationHeadersSplatParams

    Write-Verbose "Created authorization headers."
    #endregion Create authorization headers

    #region Get current PerUserMfaState
    # Microsoft docs: https://learn.microsoft.com/nl-nl/graph/api/PerUserMfaState-get?view=graph-rest-1.0&tabs=http
    $actionMessage = "querying perUserMfaState for account with AccountReference: $($actionContext.References.Account | ConvertTo-Json)"

    $baseUri = "https://graph.microsoft.com/"
    $getCurrentperUserMfaStateSplatParams = @{
        Uri         = "$($baseUri)/beta/users/$($actionContext.References.Account)/authentication/requirements"
        Headers     = $headers
        Method      = "GET"
        Verbose     = $false
        ErrorAction = "Stop"
    }

    $currentPerUserMfaState = $null
    $currentPerUserMfaState = (Invoke-RestMethod @getCurrentperUserMfaStateSplatParams)

    $currentPerUserMfaState = $currentPerUserMfaState.perUserMfaState

    Write-Verbose "Queried perUserMfaState for account with AccountReference: $($actionContext.References.Account | ConvertTo-Json). Result: $($currentPerUserMfaState | ConvertTo-Json)"
    #endregion Get current PerUserMfaState

    #region Calulate action
    $actionMessage = "calculating action"
    
    # Remove spaces
    
    if ($currentPerUserMfaState -ne $state) {
        $actionPerUserMfaState = "Update"
    }
    else {
        $actionPerUserMfaState = "NoChanges"
    }
    
    #endregion Calulate action

    #region Process
    switch ($actionPerUserMfaState) {
        "Update" {
            #region Update PerUserMfaState
            # Microsoft docs: https://learn.microsoft.com/nl-nl/graph/api/PerUserMfaState-update?view=graph-rest-1.0&tabs=http
            $actionMessage = "setting perUserMfaState to [$($state)] for account"
            $baseUri = "https://graph.microsoft.com/"

            $updatePerUserMfaStateBody = @{
                "perUserMfaState" = $state
            }

            $updatePerUserMfaStateSplatParams = @{
                Uri         = "$($baseUri)/beta/users/$($actionContext.References.Account)/authentication/requirements"
                Method      = "PATCH"
                Body        = ($updatePerUserMfaStateBody | ConvertTo-Json -Depth 10)
                Verbose     = $false
                ErrorAction = "Stop"
            }

            Write-Verbose "SplatParams: $($updatePerUserMfaStateSplatParams | ConvertTo-Json)"

            if (-Not($actionContext.DryRun -eq $true)) {
                # Add Headers after printing splat
                $updatePerUserMfaStateSplatParams['Headers'] = $headers

                $updatedPerUserMfaState = Invoke-RestMethod @updatePerUserMfaStateSplatParams

                $outputContext.AuditLogs.Add([PSCustomObject]@{
                        # Action  = "" # Optional
                        Message = "$state perUserMfaState [$($actionContext.References.Permission.Name)] for account with AccountReference: $($actionContext.References.Account | ConvertTo-Json). Old value: [$($currentPerUserMfaState)]."
                        IsError = $false
                    })
            }
            else {
                Write-Warning "DryRun: Would set perUserMfaState to [$($state)] for account with AccountReference: $($actionContext.References.Account | ConvertTo-Json). Old value: [$($currentPerUserMfaState)]."
            }
            #endregion Update PerUserMfaState

            break
        }

        "NoChanges" {
            #region No changes
            $actionMessage = "skipping setting perUserMfaState to [$($state)] for account"

            $outputContext.AuditLogs.Add([PSCustomObject]@{
                    # Action  = "" # Optional
                    Message = "Skipped setting perUserMfaState to [$($state)] as it is already set to the desired state."
                    IsError = $false
                })
            #endregion No changes

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
    if ($outputContext.AuditLogs.IsError -contains $true) {
        $outputContext.Success = $false
    }
    else {
        $outputContext.Success = $true
    }
}