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
    Read configuration settings out of Intune and Azure AD and save them to .json files
.DESCRIPTION
    This script will authenticate to Azure AD and read the following configuration settings:
        Application Configuration policies
        Application Protection policies
        Compliance policies
        Conditional Access policies including Named Locations
        Device Configuration policies
        Enrollment Restriction policies

    The configuration will we saved as .json files in the folder specified

    In order to make the export more reader friendly, the Conditional Access policies have been changed to contain displayName rather than guid
        
.EXAMPLE
    Read-Config.ps1 -LogPath C:\Temp -ExportPath C:\ExportFolder
    Export configuration using the application registration service principal while writing the log file to C:\Temp

.EXAMPLE
    Read-Config.ps1 -HostOutput -LogPath C:\Temp -ExportPath C:\ExportFolder
    Export configuration using the application registration service principal while writing the log file to C:\Temp and to the console

.EXAMPLE
    Read-Config.ps1 -HostOutput -LogPath C:\Temp -ExportPath C:\ExportFolder -VerboseLog
    Export configuration using the application registration service principal while writing the log file to C:\Temp and to the console with additional logging enabled

.EXAMPLE
    Read-Config.ps1 -LogPath C:\Temp -ExportPath C:\ExportFolder -UserAuth -User user@domain.com
    Export configuration using user authentication while writing the log file to C:\Temp

.INPUTS
    None, you cannot pipe objects into the script

.OUTPUTS
    None

.NOTES
    Version 2023-06-01
#>

[CmdletBinding(DefaultParameterSetName = 'AppAuth')]
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
    [Parameter(Mandatory=$true, ParameterSetName = 'UserAuth')]
    [string]$ExportPath,

    [Parameter(Mandatory=$true, ParameterSetName = 'UserAuth')]
    [switch]$UserAuth,

    [Parameter(Mandatory=$true, ParameterSetName = 'UserAuth')]
    [string]$User,

    [Parameter(ParameterSetName = 'AppAuth')]
    [Parameter(ParameterSetName = 'UserAuth')]
    [string]$AppClientID = "d1ddf0e4-d672-4dae-b554-9d5bdfd93547",

    [Parameter(Mandatory=$true, ParameterSetName = 'AppAuth')]
    [Parameter(ParameterSetName = 'UserAuth')]
    [string]$Tenant,

    [Parameter(Mandatory=$true, ParameterSetName = 'AppAuth')]
    [string]$AppSecret,

    [Parameter(Mandatory=$true, ParameterSetName = 'AppAuth')]
    [Parameter(Mandatory=$true, ParameterSetName = 'UserAuth')]
    [ValidateSet('All','DeviceConfig','ConditionalAccess','Compliance','NamedLocations','EnrollmentRestrictions','AppConfig','AppProtection','Scripts','EndpointSecurity')]
    [String[]]$Export = 'All'
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
            If ($UserAuth){
                $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Auto"
                $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")
                $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI,$ClientId,$redirectUri,$platformParameters,$userId).Result
            }
            Else{
                $ClientCredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential" -ArgumentList ($ClientId,$AppSecret) 
                $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI,$ClientCredential).Result
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
    Version 2023-05-30
    #>
    PARAM(
        [String]$MSGraphHost = "graph.microsoft.com",
        [String]$MSGraphVersion = "v1.0",
        [Parameter(Mandatory=$true,ValueFromPipeline)][String]$MSGraphPath,
        [switch]$IgnoreNotFound,
        [switch]$StopOnError
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
            If ($StopOnError){
                break
            }
        }
    }
}

Function Get-ManagedAppProtection{
    <#
    .SYNOPSIS
    This function is getting managed application protection policies
    .DESCRIPTION
    This function is getting managed application protection policies
    .EXAMPLE
    Get-ManagedAppProtection -id $policy.id -OS "Android"
    Returns managed application policies for Android
    .NOTES
    Version 2023-06-01
    #>
    [cmdletbinding()]
    PARAM(
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$id,
        [Parameter(Mandatory=$true)][ValidateSet("Android","iOS","WIP_WE","WIP_MDM")][string]$OS
    )
    Process{
        $Resource = "groups/$($GroupID)?`$select=displayName"
        switch ($OS){
            "Android" {
                $Resource = "deviceAppManagement/androidManagedAppProtections('$id')/?`$expand=apps,assignments"
                return Get-MsGraphObject -MSGraphPath $Resource -MSGraphVersion "beta"
            }
            "iOS" {
                $Resource = "deviceAppManagement/iosManagedAppProtections('$id')/?`$expand=apps,assignments"
                return Get-MsGraphObject -MSGraphPath $Resource -MSGraphVersion "beta"
            }
            "WIP_WE" {
                $Resource = "deviceAppManagement/windowsInformationProtectionPolicies('$id')?`$expand=protectedAppLockerFiles,exemptAppLockerFiles,assignments"
                return Get-MsGraphObject -MSGraphPath $Resource -MSGraphVersion "beta"
            }
            "WIP_MDM" {
                $Resource = "deviceAppManagement/mdmWindowsInformationProtectionPolicies('$id')?`$expand=protectedAppLockerFiles,exemptAppLockerFiles,assignments"
                return Get-MsGraphObject -MSGraphPath $Resource -MSGraphVersion "beta"
            }
        }
    }
}


Function Get-ObjectDisplayName{
    <#
    .SYNOPSIS
    This function is getting the display name of an object, based on ID and object type
    .DESCRIPTION
    This function is getting the display name of an object, based on ID and object type
    .EXAMPLE
    Get-ObjectDisplayName -ID "12345678-1234-1234-1234-123456789012" -ObjectType "NamedLocation"
    Returns the displayname of the named location with ID "12345678-1234-1234-1234-123456789012"
    .NOTES
    Version 2022-10-05
    #>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$true)][ValidateSet("NamedLocation","User","Group","Role","Application","WindowsApp")][string]$ObjectType,
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$ID
    )
    Process{
        switch ($ObjectType){
            "NamedLocation" {
                $Resource = "identity/conditionalAccess/namedLocations/$($ID)?`$select=displayName"
                $Object = Get-MsGraphObject -MSGraphPath $Resource -IgnoreNotFound -MSGraphVersion "beta"
            }
            "User"{
                $Resource = "users/$($ID)?`$select=displayName"
                $Object = Get-MsGraphObject -MSGraphPath $Resource -IgnoreNotFound        
            }
            "Group"{
                $Resource = "groups/$($ID)?`$select=displayName"
                $Object = Get-MsGraphObject -MSGraphPath $Resource -IgnoreNotFound        
            }
            "Role"{
                $Resource = "roleManagement/directory/roleDefinitions/$($ID)?`$select=displayName"
                $Object = Get-MsGraphObject -MSGraphPath $Resource -IgnoreNotFound -MSGraphVersion "beta"        
            }
            "Application"{
                $Resource = "servicePrincipals?`$filter=appID eq '$ID'&`$select=displayName"
                $Object = (Get-MsGraphObject -MSGraphPath $Resource -IgnoreNotFound).Value
            }
            "WindowsApp"{
                $Resource = "deviceAppManagement/mobileApps?`$filter=id eq '$ID'&`$select=displayName"
                $Object = (Get-MsGraphObject -MSGraphPath $Resource -IgnoreNotFound -MSGraphVersion "beta").Value
            }

            Default {
                Write-Log -Message "Object type: $ObjectType unknown" -severity 3
            }
        }
        
        If ($null -eq $Object){
            return $ID
        }
        else{
            return $Object.displayName
        }
    }
}


Function IsGUID{
    PARAM(
         [Parameter(Mandatory=$true,ValueFromPipeline)][String]$InputStr
         )
    Process {         
        if ($InputStr -match '\w{8}-\w{4}-\w{4}-\w{4}-\w{12}$'){
            return $true
        }
        else{
            return $false
        }
    }
}

Function Add-ExportFolder{
    <#
    .SYNOPSIS
    This function creates the export folder is it doesn't exists
    .DESCRIPTION
    This function creates the export folder is it doesn't exists
    .EXAMPLE
    Add-ExportFolder -Folder "C:\Export"
    Creates C:\Export if it doesn't exists
    .NOTES
    Version 2020-06-10
    #>
    PARAM(
        [Parameter(Mandatory=$true,ValueFromPipeline)][String]$Folder
    )
    Process{
        if(Test-Path $Folder){
            If ($VerboseLog){Write-Log -Message "VERBOSE: Folder found: $Folder"}
        }
        else{
            Write-Log -Message "Path '$Folder' doesn't exist, creating" -Severity 2
            Try{
                New-Item -ItemType Directory -Path $Folder | Out-Null
            }
            Catch{
                Write-Log -Message "Failed to create directory $Folder" -Severity 3
                Write-Log -Message $_.Exception -Severity 3
                exit 1
            }
        }
    }
}

Function Export-JSONData(){
    <#
    .SYNOPSIS
    This function is used to export JSON data returned from Graph
    .DESCRIPTION
    This function is used to export JSON data returned from Graph
    .EXAMPLE
    Export-JSONData -JSON $JSON
    Export the JSON inputted on the function
    .NOTES
    NAME: Export-JSONData
    #>
    
    param (
        $JSON,
        $ExportPath,
        $bundleID
    )

    try {
        if($JSON -eq "" -or $null -eq $JSON){
            Write-Log -Message "No JSON specified, please specify valid JSON..." -Severity 3
        }
        elseif(!$ExportPath){
            Write-Log -Message "No export path parameter set, please provide a path to export the file" -Severity 3
        }
        elseif(!(Test-Path $ExportPath)){
            Write-Log -Message "$ExportPath doesn't exist, can't export JSON Data" -Severity 3
        }
        else {
            $JSON1 = ConvertTo-Json $JSON -Depth 10
            $JSON_Convert = $JSON1 | ConvertFrom-Json
            $displayName = $JSON_Convert.displayName
            # Updating display name to follow file naming conventions - https://msdn.microsoft.com/en-us/library/windows/desktop/aa365247%28v=vs.85%29.aspx
            $DisplayName = $DisplayName -replace '\<|\>|:|"|/|\\|\||\?|\*', "_"
            $DisplayName = $DisplayName -replace '–', "-"
            $Properties = ($JSON_Convert | Get-Member | Where-Object{ $_.MemberType -eq "NoteProperty" }).Name
            $FileName_JSON = "$DisplayName" + "_" + $(get-date -f yyyy-MM-dd_H-mm-ss) + ".json"

            $Object = New-Object System.Object
            foreach($Property in $Properties){
                $Object | Add-Member -MemberType NoteProperty -Name $Property -Value $JSON_Convert.$Property
            }
            If($bundleID){
                $Object | Add-Member -MemberType NoteProperty -name "bundleID" -Value $bundleID
            }

            $ExportFullFilename = Join-Path -Path $ExportPath -ChildPath $FileName_JSON
            $object | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ExportFullFilename
            Write-Log -Message "JSON created in $ExportFullFilename"
        }
    
    }
    
    catch {
        Write-Log -Message "Unable to export JSON" -Severity 3
        Write-Log -Message $_.Exception -Severity 3
    }
}

Function Add-AssignmentDisplayName{
    <#
    .SYNOPSIS
    This function adds a displayName to the assignment
    .DESCRIPTION
    To make the assignments easier to read, this function adds a displayName to the JSON. The input is the target object within the assignment
    .EXAMPLE
    Add-AssignmentDisplayName $Assignment.target
    .NOTES
    Version 2022-10-04
    #>
    PARAM(
        [Parameter(Mandatory=$true)][object]$target
    )
    Process{
        switch ($target.'@odata.type'){
            "#microsoft.graph.groupAssignmentTarget" {
                If (IsGUID -InputStr $target.groupId){
                    $displayName = Get-ObjectDisplayName -ID $assignment.target.groupId -ObjectType "Group"
                    Add-Member -InputObject $target -MemberType 'NoteProperty' -Name 'displayName' -Value $displayName
                }
            }
            "#microsoft.graph.exclusionGroupAssignmentTarget" {
                If (IsGUID -InputStr $target.groupId){
                    $displayName = Get-ObjectDisplayName -ID $assignment.target.groupId -ObjectType "Group"
                    Add-Member -InputObject $target -MemberType 'NoteProperty' -Name 'displayName' -Value $displayName
                }
            }
            "#microsoft.graph.allLicensedUsersAssignmentTarget" {
                Add-Member -InputObject $target -MemberType 'NoteProperty' -Name 'displayName' -Value 'All Users'
            }
            "#microsoft.graph.allDevicesAssignmentTarget" {
                Add-Member -InputObject $target -MemberType 'NoteProperty' -Name 'displayName' -Value 'All Devices'
            }
            default {
                Add-Member -InputObject $assignment.target -MemberType 'NoteProperty' -Name 'displayName' -Value 'Unknown assignment type'
            }
        }
        return $target
    }
}

#endregion

# **************************************************************************************************
#region initAuth

$Version = "2023-06-01"

# Set log file name to the script name
$LogName = $MyInvocation.MyCommand.Name.Substring(0,$MyInvocation.MyCommand.Name.LastIndexOf("."))

Write-Log -Message "***************************************************"
Write-Log -Message "Starting script version $Version"

If ($UserAuth -and -not $Tenant){
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
        If ($UserAuth){
            $authToken = Get-AuthToken -Tenant $Tenant -ClientID $AppClientID -User $User
        }
        else {
            $authToken = Get-AuthToken -Tenant $Tenant -Client $AppClientID -AppSecret $AppSecret
        }
    }
}
else {
    # Authentication doesn't exist, calling Get-AuthToken function
    If ($UserAuth){
        $authToken = Get-AuthToken -Tenant $Tenant -ClientID $AppClientID -User $User
    }
    else {
        $authToken = Get-AuthToken -Tenant $Tenant -Client $AppClientID -AppSecret $AppSecret
    }
}
#endregion

Add-ExportFolder -Folder $ExportPath

# **************************************************************************************************
#region Export configuration to .json files
If ($Export -contains "All" -or $Export -contains "DeviceConfig"){
    $DeviceConfigurationPolicies = Get-MsGraphObject -MSGraphPath "deviceManagement/deviceConfigurations?`$expand=assignments" -MSGraphVersion "beta"
    foreach($policy in $DeviceConfigurationPolicies.value){
        $ExportSubfolder = Join-Path -Path $ExportPath -ChildPath "DeviceConfigurationPolicy"
        switch ($policy.'@odata.type'){
            "#microsoft.graph.iosUpdateConfiguration" {$PolicyType = "iOSUpdatePolicy"}
            "#microsoft.graph.windowsUpdateForBusinessConfiguration" {$PolicyType = "WindowsUpdate"}
            "#microsoft.graph.windowsHealthMonitoringConfiguration" {$PolicyType = "DeviceMonitoring"}
            default {$PolicyType = ""}
        }
        $ExportSubfolder = Join-Path -Path $ExportSubfolder -ChildPath $PolicyType
        Add-ExportFolder -Folder $ExportSubfolder
        Write-Log -Message "Device Configuration Policy: $($policy.displayName)"
        If ($policy.assignments){
            foreach ($assignment in $policy.assignments){
                $assignment.target = Add-AssignmentDisplayName -target $assignment.target
            }
        }
        Export-JSONData -JSON $policy -ExportPath $ExportSubfolder
    }
}

If ($Export -contains "All" -or $Export -contains "DeviceConfig"){
    $Policies = Get-MsGraphObject -MSGraphPath "deviceManagement/configurationPolicies?`$filter=technologies has 'mdm'&`$expand=assignments" -MSGraphVersion "beta"
    if($Policies.value){
        foreach($policy in $Policies.value){
            $ExportSubfolder = Join-Path -Path $ExportPath -ChildPath "DeviceConfigurationPolicy"
            Add-ExportFolder -Folder $ExportSubfolder
            Write-Log -Message "Configuration Policy: $($policy.name)"
            If ($policy.assignments){
                foreach ($assignment in $policy.assignments){
                    $assignment.target = Add-AssignmentDisplayName -target $assignment.target
                }
            }
            $AllSettingsInstances = @()
            $PolicyBody = New-Object -TypeName PSObject

            Add-Member -InputObject $PolicyBody -MemberType 'NoteProperty' -Name 'displayName' -Value $Policy.name
            Add-Member -InputObject $PolicyBody -MemberType 'NoteProperty' -Name 'description' -Value $policy.description
            Add-Member -InputObject $PolicyBody -MemberType 'NoteProperty' -Name 'platforms' -Value $Policy.platforms
            Add-Member -InputObject $PolicyBody -MemberType 'NoteProperty' -Name 'technologies' -Value $policy.technologies
            Add-Member -InputObject $PolicyBody -MemberType 'NoteProperty' -Name 'assignments' -Value $policy.assignments
            
            # Checking if policy has a templateId associated
            if($policy.templateReference.templateId){
                Write-Log -Message "Found template reference"
                Add-Member -InputObject $PolicyBody -MemberType 'NoteProperty' -Name 'templateId' -Value $policy.templateReference.templateId
                Add-Member -InputObject $PolicyBody -MemberType 'NoteProperty' -Name 'templateDisplayVersion' -Value $policy.templateReference.templateDisplayVersion
                Add-Member -InputObject $PolicyBody -MemberType 'NoteProperty' -Name 'TemplateDisplayName' -Value $policy.templateReference.TemplateDisplayName
            }

            $SettingInstances = Get-MsGraphObject -MSGraphPath "deviceManagement/configurationPolicies('$($policy.id)')/settings?`$expand=settingDefinitions" -MSGraphVersion "beta"
            $Instances = $SettingInstances.value.settingInstance
            foreach($object in $Instances){
#                $Instance = New-Object -TypeName PSObject
#                Add-Member -InputObject $Instance -MemberType 'NoteProperty' -Name 'settingInstance' -Value $object
                $AllSettingsInstances += $object
            }
            Add-Member -InputObject $PolicyBody -MemberType 'NoteProperty' -Name 'settings' -Value @($AllSettingsInstances)
            Export-JSONData -JSON $PolicyBody -ExportPath $ExportSubfolder
        }
    }
}

If ($Export -contains "All" -or $Export -contains "Compliance"){
    $DeviceCompliancePolicies = Get-MsGraphObject -MSGraphPath "deviceManagement/deviceCompliancePolicies?`$expand=assignments" -MSGraphVersion "beta"
    foreach($policy in $DeviceCompliancePolicies.value){
        switch -Wildcard ($policy.'@odata.type'){
            "*windows*" {$PolicyType = "Compliance_Windows"}
            "*android*" {$PolicyType = "Compliance_Android"}
            "*iOS*" {$PolicyType = "Compliance_iOS"}
            "*MacOS*" {$PolicyType = "Compliance_MacOS"}
            default {$PolicyType = ""}
        }
#        $ExportSubfolder = Join-Path -Path $ExportPath -ChildPath "CompliancePolicy"
        $ExportSubfolder = Join-Path -Path $ExportPath -ChildPath $PolicyType
        Add-ExportFolder -Folder $ExportSubfolder
        Write-Log -Message "Device Compliance Policy: $($policy.displayName)"
        If ($policy.assignments){
            foreach ($assignment in $policy.assignments){
                $assignment.target = Add-AssignmentDisplayName -target $assignment.target
            }
        }
        Export-JSONData -JSON $policy -ExportPath $ExportSubfolder
    }
}

If ($Export -contains "All" -or $Export -contains "ConditionalAccess"){
    $conditionalAccessPolicies = Get-MsGraphObject -MSGraphPath "identity/conditionalAccess/policies" -MSGraphVersion "beta"
    foreach($policy in $conditionalAccessPolicies.value){
        $ExportSubfolder = Join-Path -Path $ExportPath -ChildPath "ConditionalAccessPolicy"
        Add-ExportFolder -Folder $ExportSubfolder
        Write-Log -Message "Conditional Access Policy: $($policy.displayName)"
        
        $Users = @()
        foreach ($User in $policy.conditions.users.includeUsers){
            If (IsGUID -InputStr $User){
                $Users += Get-ObjectDisplayName -ID $User -ObjectType "User"
            }
            Else{
                $Users += $User
            }
        }
        $policy.conditions.users.includeUsers = $Users

        $Users = @()
        foreach ($User in $policy.conditions.users.excludeUsers){
            If (IsGUID -InputStr $User){
                $Users += Get-ObjectDisplayName -ID $User -ObjectType "User"
            }
            Else{
                $Users += $User
            }
        }
        $policy.conditions.users.excludeUsers = $Users

        $Groups = @()
        foreach ($Group in $policy.conditions.users.includeGroups){
            If (IsGUID -InputStr $Group){
                $Groups += Get-ObjectDisplayName -ID $Group -ObjectType "Group"
            }
            Else{
                $Groups += $Group
            }
        }
        $policy.conditions.users.includeGroups = $Groups

        $Groups = @()
        foreach ($Group in $policy.conditions.users.excludeGroups){
            If (IsGUID -InputStr $Group){
                $Groups += Get-ObjectDisplayName -ID $Group -ObjectType "Group"
            }
            Else{
                $Groups += $Group
            }
        }
        $policy.conditions.users.excludeGroups = $Groups

        $Roles = @()
        foreach ($Role in $policy.conditions.users.includeRoles){
            If (IsGUID -InputStr $Role){
                $Roles += Get-ObjectDisplayName -ID $Role -ObjectType "Role"
            }
            Else{
                $Roles += $Role
            }
        }
        $policy.conditions.users.includeRoles = $Roles

        $Roles = @()
        foreach ($Role in $policy.conditions.users.excludeRoles){
            If (IsGUID -InputStr $Role){
                $Roles += Get-ObjectDisplayName -ID $Role -ObjectType "Role"
            }
            Else{
                $Roles += $Role
            }
        }
        $policy.conditions.users.excludeRoles = $Roles

        $Applications = @()
        foreach ($App in $policy.conditions.applications.includeApplications){
            If (IsGUID -InputStr $App){
                $Applications += Get-ObjectDisplayName -ID $App -ObjectType "Application"
            }
            Else{
                $Applications += $App
            }
        }
        $policy.conditions.applications.includeApplications = $Applications

        $Applications = @()
        foreach ($App in $policy.conditions.applications.excludeApplications){
            If (IsGUID -InputStr $App){
                $Applications += Get-ObjectDisplayName -ID $App -ObjectType "Application"
            }
            Else{
                $Applications += $App
            }
        }
        $policy.conditions.applications.excludeApplications = $Applications

        $Locations = @()
        foreach ($Location in $policy.conditions.locations.includeLocations){
            If (IsGUID -InputStr $Location){
                $Locations += Get-ObjectDisplayName -ID $Location -ObjectType "NamedLocation"
            }
            Else{
                $Locations += $Location
            }
        }
        If ($Locations.Count -ge 1){$policy.conditions.locations.includeLocations = $Locations}

        $Locations = @()
        foreach ($Location in $policy.conditions.locations.excludeLocations){
            If (IsGUID -InputStr $Location){
                $Locations += Get-ObjectDisplayName -ID $Location -ObjectType "NamedLocation"
            }
            Else{
                $Locations += $Location
            }
        }
        If ($Locations.Count -ge 1){$policy.conditions.locations.excludeLocations = $Locations}

        Export-JSONData -JSON $policy -ExportPath $ExportSubfolder
    }
}

If ($Export -contains "All" -or $Export -contains "NamedLocations"){
    $NamedLocations = Get-MsGraphObject -MSGraphPath "identity/conditionalAccess/namedLocations" -MSGraphVersion "beta"
    foreach($policy in $NamedLocations.value){
        $ExportSubfolder = Join-Path -Path $ExportPath -ChildPath "NamedLocations"
        Add-ExportFolder -Folder $ExportSubfolder
        Write-Log -Message "Named Location: $($policy.displayName)"
        Export-JSONData -JSON $policy -ExportPath $ExportSubfolder
    }
}

If ($Export -contains "All" -or $Export -contains "EnrollmentRestrictions"){
    $DeviceEnrollmentConfigurations = Get-MsGraphObject -MSGraphPath "deviceManagement/deviceEnrollmentConfigurations?`$expand=assignments" -MSGraphVersion "beta"
    foreach($policy in $DeviceEnrollmentConfigurations.value){
        $ExportSubfolder = Join-Path -Path $ExportPath -ChildPath "EnrollmentRestrictions"
        switch ($policy.'@odata.type'){
            "#microsoft.graph.deviceEnrollmentWindowsHelloForBusinessConfiguration" {$PolicyType = "WindowsHello"}
            "#microsoft.graph.windows10EnrollmentCompletionPageConfiguration" {$PolicyType = "EnrollmentStatusPage"}
            "#microsoft.graph.deviceEnrollmentPlatformRestrictionConfiguration" {$PolicyType = "PlatformRestrictions"}
            "#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration" {$PolicyType = "PlatformRestrictions"}
            "#microsoft.graph.deviceEnrollmentLimitConfiguration" {$PolicyType = "EnrollmentLimit"}
            default {$PolicyType = ""}
        }
        $ExportSubfolder = Join-Path -Path $ExportSubfolder -ChildPath $PolicyType
        Add-ExportFolder -Folder $ExportSubfolder
        Write-Log -Message "EnrollmentConfigType: $PolicyType, displayName: $($policy.displayName)"
        If ($policy.assignments){
            foreach ($assignment in $policy.assignments){
                $assignment.target = Add-AssignmentDisplayName -target $assignment.target
            }
        }
        If ($policy.selectedMobileAppIds){
            $ESPApps = @()
            foreach ($app in $policy.selectedMobileAppIds){
                If (IsGUID -InputStr $App){
                    $ESPApps += Get-ObjectDisplayName -ID $App -ObjectType "WindowsApp"
                }
                Else{
                    $ESPApps += $App
                }
            }
            $policy.selectedMobileAppIds = $ESPApps
        }
        Export-JSONData -JSON $policy -ExportPath $ExportSubfolder
    }
}

If ($Export -contains "All" -or $Export -contains "AppConfig"){
    $managedAppAppConfigPolicies = Get-MsGraphObject -MSGraphPath "deviceAppManagement/targetedManagedAppConfigurations?`$expand=apps,assignments" -MSGraphVersion "beta"
    foreach($policy in $managedAppAppConfigPolicies.value){
        $ExportSubfolder = Join-Path -Path $ExportPath -ChildPath "AppConfigurationPolicy"
        Add-ExportFolder -Folder $ExportSubfolder
        Write-Log -Message "(Managed App) App Configuration Policy: $($policy.displayName)"
        If ($policy.assignments){
            foreach ($assignment in $policy.assignments){
                $assignment.target = Add-AssignmentDisplayName -target $assignment.target
            }
        }
        Export-JSONData -JSON $policy -ExportPath $ExportSubfolder
    }
}

If ($Export -contains "All" -or $Export -contains "AppConfig"){
    $managedDeviceAppConfigPolicies = Get-MsGraphObject -MSGraphPath "deviceAppManagement/mobileAppConfigurations?`$expand=assignments" -MSGraphVersion "beta"
    foreach($policy in $managedDeviceAppConfigPolicies.value){
        $ExportSubfolder = Join-Path -Path $ExportPath -ChildPath "AppConfigurationPolicy"
        Add-ExportFolder -Folder $ExportSubfolder
        Write-Log -Message "(Managed Device) App Configuration Policy: $($policy.displayName)"
        If ($policy.assignments){
            foreach ($assignment in $policy.assignments){
                $assignment.target = Add-AssignmentDisplayName -target $assignment.target
            }
        }

        #If this is an Managed Device App Config for iOS, lookup the bundleID to support importing to a different tenant
        If($policy.'@odata.type' -eq "#microsoft.graph.iosMobileAppConfiguration"){
            $bundleID = (Get-MsGraphObject -MSGraphPath "deviceAppManagement/mobileApps?`$filter=id eq '$($policy.targetedMobileApps)'").value
            Export-JSONData -JSON $policy -ExportPath $ExportSubfolder -bundleID $bundleID.bundleID
        }
        Else{
            Export-JSONData -JSON $policy -ExportPath $ExportSubfolder
        }
    }
}

If ($Export -contains "All" -or $Export -contains "AppProtection"){
    $ManagedAppPolicies = (Get-MsGraphObject -MSGraphPath "deviceAppManagement/managedAppPolicies" -MSGraphVersion "beta").value | Where-Object { ($_.'@odata.type').contains("ManagedAppProtection") }
    foreach($policy in $ManagedAppPolicies){
        $ExportSubfolder = Join-Path -Path $ExportPath -ChildPath "AppProtectionPolicy"
        Add-ExportFolder -Folder $ExportSubfolder
        Write-Log -Message "Managed App Policy: $($policy.displayName)"
        if($policy.'@odata.type' -eq "#microsoft.graph.androidManagedAppProtection"){
            $AppProtectionPolicy = Get-ManagedAppProtection -id $policy.id -OS "Android"
            $AppProtectionPolicy | Add-Member -MemberType NoteProperty -Name '@odata.type' -Value "#microsoft.graph.androidManagedAppProtection"
            If ($AppProtectionPolicy.assignments){
                foreach ($assignment in $AppProtectionPolicy.assignments){
                    $assignment.target = Add-AssignmentDisplayName -target $assignment.target
                }
            }
            Export-JSONData -JSON $AppProtectionPolicy -ExportPath $ExportSubfolder
        }
        elseif($policy.'@odata.type' -eq "#microsoft.graph.iosManagedAppProtection"){
            $AppProtectionPolicy = Get-ManagedAppProtection -id $policy.id -OS "iOS"
            $AppProtectionPolicy | Add-Member -MemberType NoteProperty -Name '@odata.type' -Value "#microsoft.graph.iosManagedAppProtection"
            If ($AppProtectionPolicy.assignments){
                foreach ($assignment in $AppProtectionPolicy.assignments){
                    $assignment.target = Add-AssignmentDisplayName -target $assignment.target
                }
            }
            Export-JSONData -JSON $AppProtectionPolicy -ExportPath $ExportSubfolder
        }
    }
}

If ($Export -contains "All" -or $Export -contains "Scripts"){
    $Scripts = Get-MsGraphObject -MSGraphPath "deviceManagement/deviceShellScripts?`$expand=assignments" -MSGraphVersion "beta"
    foreach($Script in $Scripts.value){
        $ExportSubfolder = Join-Path -Path $ExportPath -ChildPath "Scripts"
        Add-ExportFolder -Folder $ExportSubfolder
        Write-Log -Message "ShellScript [MacOS]: $($Script.displayName)"
        If ($Script.assignments){
            foreach ($assignment in $Script.assignments){
                $assignment.target = Add-AssignmentDisplayName -target $assignment.target
            }
        }
        $ScriptContent = (Get-MsGraphObject -MSGraphPath "deviceManagement/deviceShellScripts/$($Script.id)" -MSGraphVersion "beta").scriptContent
        $DecodedScript = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($ScriptContent))
        $displayName = $Script.displayName
        # Updating display name to follow file naming conventions - https://msdn.microsoft.com/en-us/library/windows/desktop/aa365247%28v=vs.85%29.aspx
        $DisplayName = $DisplayName -replace '\<|\>|:|"|/|\\|\||\?|\*', "_"
        $FileName = "$DisplayName" + "_" + $(get-date -f yyyy-MM-dd_H-mm-ss) + ".ScriptContent"
        $ExportFullFilename = Join-Path -Path $ExportSubfolder -ChildPath $FileName
        Write-Log -Message "Export script content: $ExportFullFilename"
        $DecodedScript | Out-File -FilePath $ExportFullFilename
        Export-JSONData -JSON $Script -ExportPath $ExportSubfolder
    }
}

If ($Export -contains "All" -or $Export -contains "Scripts"){
    $Scripts = Get-MsGraphObject -MSGraphPath "deviceManagement/deviceManagementScripts?`$expand=assignments" -MSGraphVersion "beta"
    foreach($Script in $Scripts.value){
        $ExportSubfolder = Join-Path -Path $ExportPath -ChildPath "Scripts"
        Add-ExportFolder -Folder $ExportSubfolder
        Write-Log -Message "ShellScript [Win]: $($Script.displayName)"
        If ($Script.assignments){
            foreach ($assignment in $Script.assignments){
                $assignment.target = Add-AssignmentDisplayName -target $assignment.target
            }
        }
        $ScriptContent = (Get-MsGraphObject -MSGraphPath "deviceManagement/deviceManagementScripts/$($Script.id)" -MSGraphVersion "beta").scriptContent
        $DecodedScript = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($ScriptContent))
        $displayName = $Script.displayName
        # Updating display name to follow file naming conventions - https://msdn.microsoft.com/en-us/library/windows/desktop/aa365247%28v=vs.85%29.aspx
        $DisplayName = $DisplayName -replace '\<|\>|:|"|/|\\|\||\?|\*', "_"
        $FileName = "$DisplayName" + "_" + $(get-date -f yyyy-MM-dd_H-mm-ss) + ".ScriptContent"
        $ExportFullFilename = Join-Path -Path $ExportSubfolder -ChildPath $FileName
        Write-Log -Message "Export script content: $ExportFullFilename"
        $DecodedScript | Out-File -FilePath $ExportFullFilename
        Export-JSONData -JSON $Script -ExportPath $ExportSubfolder
    }
}


If ($Export -contains "All" -or $Export -contains "EndpointSecurity"){
    # Get all Endpoint Security Templates
    $Templates = Get-MsGraphObject -MSGraphPath "deviceManagement/templates?`$filter=(isof(%27microsoft.graph.securityBaselineTemplate%27))" -MSGraphVersion "beta"

    # Get all Endpoint Security Policies configured
    $ESPolicies = Get-MsGraphObject -MSGraphPath "deviceManagement/intents" -MSGraphVersion "beta"

    # Looping through all policies configured
    foreach($policy in $ESPolicies.value){
        $ExportSubfolder = Join-Path -Path $ExportPath -ChildPath "EndpointSecurity"
        Add-ExportFolder -Folder $ExportSubfolder

        Write-Log -Message "Endpoint Security Policy: $($policy.displayName)"
        $PolicyName = $policy.displayName
        $PolicyDescription = $policy.description
        $policyId = $policy.id
        $TemplateId = $policy.templateId
        $roleScopeTagIds = $policy.roleScopeTagIds

        $ES_Template = $Templates.value | ?  { $_.id -eq $policy.templateId }

        $TemplateDisplayName = $ES_Template.displayName
        $TemplateId = $ES_Template.id
        $versionInfo = $ES_Template.versionInfo

        if($TemplateDisplayName -eq "Endpoint detection and response"){
            Write-Log -Message "Export of 'Endpoint detection and response' policy not included in sample script..." -severity 3
        }
        else {
            # Creating object for JSON output
            $JSON = New-Object -TypeName PSObject

            Add-Member -InputObject $JSON -MemberType 'NoteProperty' -Name 'displayName' -Value "$PolicyName"
            Add-Member -InputObject $JSON -MemberType 'NoteProperty' -Name 'description' -Value "$PolicyDescription"
            Add-Member -InputObject $JSON -MemberType 'NoteProperty' -Name 'roleScopeTagIds' -Value $roleScopeTagIds
            Add-Member -InputObject $JSON -MemberType 'NoteProperty' -Name 'TemplateDisplayName' -Value "$TemplateDisplayName"
            Add-Member -InputObject $JSON -MemberType 'NoteProperty' -Name 'TemplateId' -Value "$TemplateId"
            Add-Member -InputObject $JSON -MemberType 'NoteProperty' -Name 'versionInfo' -Value "$versionInfo"

            # Getting all categories in specified Endpoint Security Template
            $Categories = Get-MsGraphObject -MSGraphPath "deviceManagement/templates/$TemplateId/categories" -MSGraphVersion "beta"

            # Looping through all categories within the Template
            $Settings = @()
            foreach($category in $Categories.value){
                $categoryId = $category.id
                $Settings += (Get-MsGraphObject -MSGraphPath "deviceManagement/intents/$policyId/categories/$categoryId/settings?`$expand=Microsoft.Graph.DeviceManagementComplexSettingInstance/Value" -MSGraphVersion "beta").value
            }

            # Adding All settings to settingsDelta ready for JSON export
            Add-Member -InputObject $JSON -MemberType 'NoteProperty' -Name 'settingsDelta' -Value @($Settings)
            Export-JSONData -JSON $JSON -ExportPath $ExportSubfolder
        }
    }
}

#endregion


Write-Log -Message "Script end"