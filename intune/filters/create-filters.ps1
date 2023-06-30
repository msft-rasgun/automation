# **************************************************************************************************
# This sample script is not supported under any Microsoft standard support program or service. 
# The sample script is provided AS IS without warranty of any kind. Microsoft further disclaims 
# all implied warranties including, without limitation, any implied warranties of merchantability 
# or of fitness for a particular purpose. The entire risk arising out of the use or performance 
# of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, 
# Or anyone else involved in the creation, production, or delivery of the scripts be liable for 
# any damages whatsoever (including, without limitation, damages for loss of business profits, 
# business interruption, loss of business information, or other pecuniary loss) arising out of 
# the use of or inability to use the sample scripts or documentation, even if Microsoft has been 
# advised of the possibility of such damages.
# **************************************************************************************************
#

<#
.SYNOPSIS
    Create assignment filters based on .json files
.DESCRIPTION
    Loop through all existing .json files and create/update assignment filters accordingly

    The script have two parameter sets: AppAuth and UserAuth. UserAuth is set by default, AppAuth is activated with the -AppAuth switch

    If UserAuth:
    User        [Required, UPN of user account used to connect to Graph]
    Tenant      [Optional, if not set, tenant will be derived from user UPN]
    
    If AppAuth:
    Tenant      [Required, specify the tenant name using either onmicrosoft.com or any verified domain name]
    AppSecret   [Required, specify the secret for authentication]
    
    In both scenarios:
    LogPath     [Optional, path where log file will be written. Default: The current directory]
    MaxLogSize  [Optional, specify the maximum log file size. Default: 2MB]
    HostOutput  [Optional, switch to enable output to host. Default: Off]
    VerboseLog  [Optional, switch to enable additional logging. Default: Off]
    AppClientID [Optional, specify the application/client ID for application registration used for authentication. Default: d1ddf0e4-d672-4dae-b554-9d5bdfd93547]

.EXAMPLE
    Create-Filters.ps1 -HostOutput -User admin@gundel.tech
    
    Connect using user authentication and write output to log file as well as host

.INPUTS
    None, you cannot pipe objects into the script

.OUTPUTS
    None

.NOTES
    Version 2023-06-30
#>

[CmdletBinding(DefaultParameterSetName = 'UserAuth')]
PARAM(
    [Parameter(ParameterSetName = 'AppAuth')]
    [Parameter(ParameterSetName = 'UserAuth')]
    [string]$LogPath = ".",

    [Parameter(ParameterSetName = 'AppAuth')]
    [Parameter(ParameterSetName = 'UserAuth')]
    [int]$MaxLogSize = 2MB,

    [Parameter(ParameterSetName = 'AppAuth')]
    [Parameter(ParameterSetName = 'UserAuth')]
    [switch]$HostOutput,

    [Parameter(ParameterSetName = 'AppAuth')]
    [Parameter(ParameterSetName = 'UserAuth')]
    [switch]$VerboseLog,

    [Parameter(Mandatory=$true, ParameterSetName = 'AppAuth')]
    [switch]$AppAuth,

    [Parameter(Mandatory=$true, ParameterSetName = 'UserAuth')]
    [string]$User,

    [Parameter(ParameterSetName = 'AppAuth')]
    [Parameter(ParameterSetName = 'UserAuth')]
    [string]$AppClientID = "d1ddf0e4-d672-4dae-b554-9d5bdfd93547",

    [Parameter(Mandatory=$true, ParameterSetName = 'AppAuth')]
    [Parameter(ParameterSetName = 'UserAuth')]
    [string]$Tenant,

    [Parameter(Mandatory=$true, ParameterSetName = 'AppAuth')]
    [string]$AppSecret
)

# **************************************************************************************************
#region functions

Function Write-Log{
    <#
    .SYNOPSIS
    This function is used to write a log file
    .DESCRIPTION
    This function write a log file in the format of Configuration Manager, which can be opened by CMTrace.exe
    .EXAMPLE
    Write-Log -Message "Message"
    Writes Message to the log file
    .EXAMPLE
    Write-Log -Message "Message" -severity 3
    Writes Message to the log file as an error
    .NOTES
    Version 2020-05-25
    #>
    PARAM(
        [Parameter(Mandatory=$true,ValueFromPipeline)][String]$Message,
        [String]$Path = (Join-Path -Path $LogPath -ChildPath "$LogName.log"),
        [int]$severity = 1,
        [string]$component = $LogName,
        [int]$thread = $PID,
        [string]$context,
        [string]$source
    )
    Process {
        if (Test-Path $Path){
            If ((Get-Item $Path).length -gt $MaxLogSize){
                $OldLogName = $Path.SubString(0,($Path.Length-1)) + "_"
                If (Test-Path $OldLogName){Remove-Item $OldLogName}
                Rename-Item $Path -NewName $OldLogName
            }
        }

        $TimeZoneBias = Get-WmiObject -Query "Select Bias from Win32_TimeZone"
        $Time= Get-Date -Format "HH:mm:ss.fff"
        $Date= Get-Date -Format "MM-dd-yyyy"

        "<![LOG[$Message]LOG]!><time=$([char]34)$Time$($TimeZoneBias.bias)$([char]34) date=$([char]34)$date$([char]34) component=$([char]34)$component$([char]34) context=$([char]34)$context$([char]34) type=$([char]34)$severity$([char]34) thread=$([char]34)$thread$([char]34) file=$([char]34)$source$([char]34)>"| Out-File -FilePath $Path -Append -NoClobber -Encoding default
        If ($HostOutput){
            switch ($severity){
                1 {$Color = "Green"}
                2 {$Color = "Yellow"}
                3 {$Color = "Red"}
            }
            Write-Host $Message -ForegroundColor $Color
        }
    }
}

function Get-AuthToken{
    <#
    .SYNOPSIS
    This function is getting and returning an authentication header
    .DESCRIPTION
    This function is getting and returning an authentication header
    .EXAMPLE
    Get-AuthToken -Tenant $Tenant -ClientID $IntuneClientId -User $User
    Get the auth token based on user authentication
    .EXAMPLE
    Get-AuthToken -Tenant $Tenant -Client $AppClientID -AppSecret $AppSecret
    Get the auth token based on application authentication
    .NOTES
    Version 2020-06-10
    #>
    PARAM(
        [Parameter(Mandatory=$true, ParameterSetName = 'AppAuth')]
        [Parameter(Mandatory=$true, ParameterSetName = 'UserAuth')]
        [string]$Tenant,
        [Parameter(Mandatory=$true, ParameterSetName = 'UserAuth')]
        [string]$User,
        [Parameter(Mandatory=$true, ParameterSetName = 'AppAuth')]
        [Parameter(Mandatory=$true, ParameterSetName = 'UserAuth')]
        [string]$ClientId,
        [Parameter(Mandatory=$true, ParameterSetName = 'AppAuth')]
        [string]$AppSecret
    )

    Process{
        Write-Log -Message "Checking for AzureAD module..."
        $AadModule = Get-Module -Name "AzureAD" -ListAvailable
        if ($null -eq $AadModule) {
            Write-Log -Message "AzureAD PowerShell module not found, looking for AzureADPreview" -Severity 2
            $AadModule = Get-Module -Name "AzureADPreview" -ListAvailable
        }

        if ($null -eq $AadModule) {
            Write-Log -Message "AzureAD Powershell module not installed..." -Severity 3
            Write-Log -Message "Install by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt" -Severity 3
            Write-Log -Message "Script can't continue..." -Severity 3
            exit
        }

        # Getting path to ActiveDirectory Assemblies
        # If the module count is greater than 1 find the latest version
        if ($AadModule.count -gt 1) {
            $Latest_Version = ($AadModule | Select-Object version | Sort-Object)[-1]
            $aadModule = $AadModule | Where-Object { $_.version -eq $Latest_Version.version }
        }

        Write-Log -Message "Found module: $($AadModule.Name)"

        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"

        [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
        [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null

        # Set variables to be used
        $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
        $resourceAppIdURI = "https://graph.microsoft.com"
        $authority = "https://login.microsoftonline.com/$Tenant"

        try {
            $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
            # https://msdn.microsoft.com/library/azure/microsoft.identitymodel.clients.activedirectory.promptbehavior.aspx
            # Change the prompt behaviour to force credentials each time: Auto, Always, Never, RefreshSession

            # Depending on authentication method, get the auth token
            If ($AppAuth){
                $ClientCredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential" -ArgumentList ($ClientId,$AppSecret) 
                $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI,$ClientCredential).Result
            }
            Else{
                $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Auto"
                $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")
                $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI,$ClientId,$redirectUri,$platformParameters,$userId).Result
            }
            
            # If the accesstoken is valid then create the authentication header
            if ($authResult.AccessToken) {
                # Creating header for Authorization token
                $authHeader = @{
                    'Content-Type' = 'application/json'
                    'Authorization' = "Bearer " + $authResult.AccessToken
                    'ExpiresOn' = $authResult.ExpiresOn
                }
                Write-Log -Message "Authentication header created"
                return $authHeader
            }
            else {
                Write-Log -Message "Authorization Access Token is null, please re-run authentication..." -Severity 3
                break
            }
        }
        catch {
            Write-Log -Message "Unable to get authentication header" -severity 3
            Write-Log -Message $_.Exception.Message -Severity 3
            Write-Log -Message $_.Exception.ItemName -Severity 3
            Write-Log -Message $_.Exception -Severity 3
            exit 1
        }   
    }
}

Function Get-MsGraphCollection{
    <#
    .SYNOPSIS
    This function is connecting to Microsoft Graph to get an collection
    .DESCRIPTION
    This function is connecting to Microsoft Graph to get an collection
    .EXAMPLE
    Get-MsGraphCollection -MSGraphPath "users/$Upn/memberOf/microsoft.graph.group"
    Returns the group membership of the user
    .NOTES
    Version 2020-06-10
    #>
    PARAM(
        [String]$MSGraphHost = "graph.microsoft.com",
        [String]$MSGraphVersion = "v1.0",
        [Parameter(Mandatory=$true,ValueFromPipeline)][String]$MSGraphPath
    )

    Process {
        $Collection = @()
        $FullUri = "https://$MSGraphHost/$MSGraphVersion/$MSGraphPath"
        $NextLink = $FullUri
        Do {
            try {
                If ($VerboseLog){Write-Log -Message "VERBOSE: GET $NextLink"}
                $Result = Invoke-RestMethod -Method Get -Uri $NextLink -Headers $AuthHeader
                $Collection += $Result.value
                $NextLink = $Result.'@odata.nextLink'
            } 
            catch {
                $ResponseStream = $_.Exception.Response.GetResponseStream()
                $ResponseReader = New-Object System.IO.StreamReader $ResponseStream
                $ResponseContent = $ResponseReader.ReadToEnd()
                Write-Log -Message "Request Failed: $($_.Exception.Message)`n$($_.ErrorDetails)" -severity 3
                Write-Log -Message "Request URL: $NextLink" -severity 3
                Write-Log -Message "Response Content:`n$ResponseContent" -severity 3
                break
            }
        } while ($null -ne $NextLink)
        Write-Log -Message "Got $($Collection.Count) object(s)"

        return $Collection
    }
}

function Get-MsGraphObject{
    <#
    .SYNOPSIS
    This function is connecting to Microsoft Graph to get an object
    .DESCRIPTION
    This function is connecting to Microsoft Graph to get an object
    .EXAMPLE
    Get-MsGraphObject -MSGraphPath "deviceManagement/deviceEnrollmentConfigurations" -MSGraphVersion "beta"
    Returns device enrollment configuration using the Beta version of the GraphAPI
    .NOTES
    Version 2020-06-10
    #>
    PARAM(
        [String]$MSGraphHost = "graph.microsoft.com",
        [String]$MSGraphVersion = "v1.0",
        [Parameter(Mandatory=$true,ValueFromPipeline)][String]$MSGraphPath,
        [switch]$IgnoreNotFound
    )
    Process{
        $FullUri = "https://$MSGraphHost/$MSGraphVersion/$MSGraphPath"
        If ($VerboseLog){Write-Log -Message "VERBOSE: GET $FullUri"}
        try {
            return Invoke-RestMethod -Method Get -Uri $FullUri -Headers $authToken
        } 
        catch {
            $Response = $_.Exception.Response
            if ($IgnoreNotFound -and $Response.StatusCode -eq "NotFound") {
                return $null
            }
            $ResponseStream = $Response.GetResponseStream()
            $ResponseReader = New-Object System.IO.StreamReader $ResponseStream
            $ResponseContent = $ResponseReader.ReadToEnd()
            Write-Log -Message "Request Failed: $($_.Exception.Message)`n$($_.ErrorDetails)" -severity 3
            Write-Log -Message "Request URL: $FullUri" -severity 3
            Write-Log -Message "Response Content:`n$ResponseContent" -severity 3
            break
        }
    }
}

function Invoke-MsGraphObject{
    <#
    .SYNOPSIS
    This function is connecting to Microsoft Graph to execute commands
    .DESCRIPTION
    This function is connecting to Microsoft Graph to execute commands
    .EXAMPLE
    Invoke-MsGraphObject -MSGraphPath "deviceManagement/deviceEnrollmentConfigurations" -MSGraphVersion "beta" -Method "GET"
    Returns device enrollment configuration using the Beta version of the GraphAPI
    .NOTES
    Version 2023-06-30
    #>
    PARAM(
        [String]$MSGraphHost = "graph.microsoft.com",
        [ValidateSet("v1.0","beta", IgnoreCase = $true)][String]$MSGraphVersion = "v1.0",
        [Parameter(Mandatory=$true,ValueFromPipeline)][String]$MSGraphPath,
        [ValidateSet("GET","POST","PATCH", IgnoreCase = $true)][Parameter(Mandatory=$true,ValueFromPipeline)][String]$Method,
        $Body,
        [String]$ContentType,
        [switch]$IgnoreNotFound
    )
    Process{
        $FullUri = "https://$MSGraphHost/$MSGraphVersion/$MSGraphPath"
        If ($VerboseLog){Write-Log -Message "VERBOSE: $Method - $FullUri"}
        $try = 1
        $maxTry = 5
        $sleep = 15
        Do {
            try {
                $success = $true
                If ($Method -eq "GET"){
                    return Invoke-RestMethod -Method $Method -Uri $FullUri -Headers $authToken
                }
                Else{
                    return Invoke-RestMethod -Method $Method -Uri $FullUri -Headers $authToken -Body $Body -ContentType $ContentType
                }
                
            } 
            catch {
                $success = $false
                $nonRetryStatusCode = @("BadRequest")
                $Response = $_.Exception.Response
                if ($IgnoreNotFound -and $Response.StatusCode -eq "NotFound") {
                    return $null
                }
                If ($try -gt $maxTry -or $Response.StatusCode -in $nonRetryStatusCode){
                    $ResponseStream = $Response.GetResponseStream()
                    $ResponseReader = New-Object System.IO.StreamReader $ResponseStream
                    $ResponseContent = $ResponseReader.ReadToEnd()
                    Write-Log -Message "Request Failed: $($_.Exception.Message)`n$($_.ErrorDetails)" -severity 3
                    Write-Log -Message "Request URL: $FullUri" -severity 3
                    Write-Log -Message "Request StatusCode: $($Response.StatusCode)" -severity 3
                    Write-Log -Message "Response Content:`n$ResponseContent" -severity 3
                    break
                }
                Write-Log -Message "$Method failed with error $($Response.StatusCode). Retry after sleeping for $sleep seconds, attempt $try of $maxTry" -severity 2
                Start-Sleep -Seconds $sleep
            }
            $try++
        } until ($success)
    }
}


#endregion

# **************************************************************************************************
#region initAuth

$Version = "2023-06-30"

# Set log file name to the script name
$LogName = $MyInvocation.MyCommand.Name.Substring(0,$MyInvocation.MyCommand.Name.LastIndexOf("."))

Write-Log -Message "***************************************************"
Write-Log -Message "Starting script version $Version"

If (-not $AppAuth -and -not $Tenant){
    try{
        $userUpn = New-Object "System.Net.Mail.MailAddress" -ArgumentList $User
        $Tenant = $userUpn.Host
    }
    catch{
        Write-Log -Message "Unable to verify user UPN" -Severity 3
        Write-Log -Message $_.Exception -Severity 3
        exit 1
    }
}

Write-Log -Message "Using tenant: $Tenant"

if($authToken){
    # Setting DateTime to Universal time to work in all timezones
    $DateTime = (Get-Date).ToUniversalTime()

    # If the authToken exists checking when it expires
    $TokenExpires = ($authToken.ExpiresOn.datetime - $DateTime).Minutes

    if($TokenExpires -le 0){
        Write-Log -Message "Authentication Token expired" $TokenExpires "minutes ago" -Severity 2
        If ($AppAuth){
            $authToken = Get-AuthToken -Tenant $Tenant -Client $AppClientID -AppSecret $AppSecret
        }
        else {
            $authToken = Get-AuthToken -Tenant $Tenant -ClientID $AppClientID -User $User
        }
    }
}
else {
    # Authentication doesn't exist, calling Get-AuthToken function
    If ($AppAuth){
        $authToken = Get-AuthToken -Tenant $Tenant -Client $AppClientID -AppSecret $AppSecret
    }
    else {
        $authToken = Get-AuthToken -Tenant $Tenant -ClientID $AppClientID -User $User
    }
}
#endregion


# Get current assignment filters
$ExistingFilters = (Invoke-MsGraphObject -MSGraphPath "deviceManagement/assignmentFilters" -MSGraphVersion "beta" -Method GET).value

# Get .json files
$Files = Get-ChildItem -Filter *.json

# Loop through files and POST/PATCH accordingly
foreach ($File in $Files){
    $json = $File | Get-Content | ConvertFrom-Json
    If ($json.displayName -in $ExistingFilters.displayName){
        #PATCH
        Write-Log -Message "Found existing filter: $($json.displayName). PATCH"
        $ExistingPolicy = $ExistingFilters | Where-Object {$_.displayName -eq $json.displayName}
        # Remove platform when patching since it cannot be updated
        $json = $json | Select-Object -Property * -ExcludeProperty platform
        $payload = $json | ConvertTo-JSON -Depth 20
        Invoke-MsGraphObject -MSGraphVersion beta -Method PATCH -Body $payload -ContentType "application/json" -MSGraphPath "deviceManagement/assignmentFilters/$($ExistingPolicy.id)"
    }
    Else{
        #POST
        Write-Log -Message "Filter not found: $($json.displayName). POST"
        $payload = $json | ConvertTo-JSON -Depth 20
        Invoke-MsGraphObject -MSGraphVersion beta -Method POST -Body $payload -ContentType "application/json" -MSGraphPath "deviceManagement/assignmentFilters"
    }
}

Write-Log -Message "Script end"