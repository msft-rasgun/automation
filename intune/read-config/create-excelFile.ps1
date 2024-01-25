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
    Create Excel file with Intune and Entra ID configuration exporting into .json. Requires Excel to be installed on the machine where the script runs.
.DESCRIPTION
    This script will create an Excel file containing the configuration of Intune and Azure AD from Read-Config.ps1 in .json file format

    The script requires Excel to be installed and has been tested with Microsoft Apps 365 2210
        
.EXAMPLE
    Create-ExcelFile.ps1 -ExcelFilename C:\Temp\Export.xlsx -InputFolder C:\Export
    Create Excel file C:\Temp\Export.xlsx based on configuration export located in C:\Export

.EXAMPLE
    Create-ExcelFile.ps1 -NoHostOutput -ExcelFilename C:\Temp\Export.xlsx -InputFolder C:\Export
    Create Excel file C:\Temp\Export.xlsx based on configuration export located in C:\Export. Disable logging to the console

.EXAMPLE
    Create-ExcelFile.ps1 -ExcelFilename C:\Temp\Export.xlsx -InputFolder C:\Export -Force
    Create Excel file C:\Temp\Export.xlsx based on configuration export located in C:\Export. Replace C:\Temp\Export.xlsx if it already exists

.EXAMPLE
    Create-ExcelFile.ps1 -NoHostOutput -ExcelFilename C:\Temp\Export.xlsx -InputFolder C:\Export -VerboseLog
    Create Excel file C:\Temp\Export.xlsx based on configuration export located in C:\Export. Disable logging to the console and enabling verbose logging

.INPUTS
    None, you cannot pipe objects into the script

.OUTPUTS
    None

.NOTES
    Version 2024-01-23
#>

PARAM(
    [string]$LogPath = ".",
    [int]$MaxLogSize = 2MB,
    [switch]$NoHostOutput,
    [switch]$VerboseLog,
    [switch]$Force,
    [parameter(Mandatory=$true)][string]$ExcelFilename,
    [parameter(Mandatory=$true)][string]$InputFolder
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
    Version 2024-01-19
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
        If (-not $NoHostOutput){
            switch ($severity){
                1 {$Color = "Green"}
                2 {$Color = "Yellow"}
                3 {$Color = "Red"}
            }
            Write-Host $Message -ForegroundColor $Color
        }
    }
}


function ConvertTo-MultilineString{
    <#
    .SYNOPSIS
    This function is used to convert an array of strings into one string with the specificed separator
    .DESCRIPTION
    This function is used to convert an array of strings into one string with the specificed separator
    .EXAMPLE
    ConvertTo-MultilineString -Array @Array
    Returns a string with each item within the array separated with the NewLine character
    .EXAMPLE
    ConvertTo-MultilineString -Array @Array -Separator "|"
    Returns a string with each item within the array separated with the | character
    .NOTES
    Version 2020-06-10
    #>
    PARAM(
        [Array]$Array,
        [string]$separator = "`n"
    )
    Process {
        If ($array){
            $Str = ""
            foreach ($item in $array){
                If ($item.type -in ('image/jpeg','image/png')){
                    Write-Log -Message "Data type $($item.type) not supported" -severity 2
                    return "Data type $($item.type) not supported"
                }
                else{
                    $Str = $Str + $item + $separator
                }
            }
            #Remove trailing separator
            $Str = $str.Substring(0,$str.LastIndexOf($separator))
            return $str
        }
        Else{
            return $array
        }
    }
}


Function Add-PoliciesToExcel{
    <#
    .SYNOPSIS
    This function adds policies to the Excel file
    .DESCRIPTION
    Loop through policies .json export and write results to Excel
    .EXAMPLE
    Add-PoliciesToExcel -Subfolder C:\Export\CompliancePolicy
    Adds policy exports found in the subfolder to Excel
    .NOTES
    Version 2020-06-01
    #>
    PARAM(
        [Parameter(Mandatory=$true)][String]$Subfolder
    )
    # Add policies to Excel file
    If (!(Test-Path $Subfolder)){
        Write-Log -Message "Unable to find input folder: $Subfolder. Skip" -severity 3
        break
    }

    # Create new Worksheet
    $Worksheet = $Workbook.Worksheets.Add()
    $ParentFolderName = $Subfolder.Substring($Subfolder.LastIndexOf("\")+1)
    $Worksheet.Name = $ParentFolderName
    If ($VerboseLog){Write-Log -Message "VERBOSE: Created Worksheet $($Worksheet.Name)"}

    If ($VerboseLog){Write-Log -Message "VERBOSE: Loop through array and create headers"}
    $Headers = @('Policy Type','Policy Name','Setting','Value')
    $column = 1
    Foreach ($Header in $Headers){
        $Worksheet.Cells.Item(1,$column) = $Header
        $Worksheet.Cells.Item(1,$column).Font.Bold = $true
        $column++
    }

    # Loop through files and write result to Excel worksheet
    $row = 2
    $JSONs = Get-ChildItem $Subfolder -Filter "*.json"
    Foreach ($File in $JSONs){
        Write-Log -Message "Processing file: $($File.FullName)"
        $JSON = Get-Content ($File.FullName.Replace("[","``[")).Replace("]","``]") | ConvertFrom-Json
        $Configs = $JSON | Get-Member | Where-Object membertype -eq "NoteProperty" | Where-Object Name -NotIn $ExcludedDataTypes
        If ($ParentFolderName -eq "ConditionalAccessPolicy"){
            # Handle Conditional Access policies differently
            $Worksheet.Cells.Item($row,1) = ConvertTo-MultilineString -Array $JSON.state
            $Worksheet.Cells.Item($row,2) = $JSON.displayName
            $Worksheet.Cells.Item($row,3) = ConvertTo-MultilineString -Array $JSON.conditions.users.includeUsers
            $Worksheet.Cells.Item($row,4) = ConvertTo-MultilineString -Array $JSON.conditions.users.excludeUsers
            $Worksheet.Cells.Item($row,5) = ConvertTo-MultilineString -Array $JSON.conditions.users.includeGroups
            $Worksheet.Cells.Item($row,6) = ConvertTo-MultilineString -Array $JSON.conditions.users.excludeGroups
            $Worksheet.Cells.Item($row,7) = ConvertTo-MultilineString -Array $JSON.conditions.users.includeRoles
            $Worksheet.Cells.Item($row,8) = ConvertTo-MultilineString -Array $JSON.conditions.users.excludeRoles
            $Worksheet.Cells.Item($row,9) = ConvertTo-MultilineString -Array $JSON.conditions.applications.includeApplications
            $Worksheet.Cells.Item($row,10) = ConvertTo-MultilineString -Array $JSON.conditions.applications.excludeApplications
            $Worksheet.Cells.Item($row,11) = ConvertTo-MultilineString -Array $JSON.conditions.applications.includeUserActions
            $Worksheet.Cells.Item($row,12) = ConvertTo-MultilineString -Array $JSON.conditions.platforms.includePlatforms
            $Worksheet.Cells.Item($row,13) = ConvertTo-MultilineString -Array $JSON.conditions.platforms.excludePlatforms
            $Worksheet.Cells.Item($row,14) = ConvertTo-MultilineString -Array $JSON.conditions.locations.includeLocations
            $Worksheet.Cells.Item($row,15) = ConvertTo-MultilineString -Array $JSON.conditions.locations.excludeLocations
            $Worksheet.Cells.Item($row,16) = ConvertTo-MultilineString -Array $JSON.conditions.clientAppTypes
            $Worksheet.Cells.Item($row,17) = "Operator: $($JSON.grantControls.operator)`nbuiltInControls:$(ConvertTo-MultilineString -Array $JSON.grantControls.builtInControls -separator ",")`ncustomAuthenticationFactors:$(ConvertTo-MultilineString -Array $JSON.grantControls.customAuthenticationFactors -separator ",")`ntermsOfUse:$(ConvertTo-MultilineString -Array $JSON.grantControls.termsOfUse -separator ",")"
            $row++
        }
        Else{
            # If odata.type not set, use folder name. Otherwise, remove #microsoft.graph part of name
            If ($JSON.'@odata.type'){
                $PolicyType = $JSON.'@odata.type'.Replace("#microsoft.graph.","")
            }
            else{
                $PolicyType = $ParentFolderName
                If ($VerboseLog){Write-Log -Message "VERBOSE: Policy doesn't contain odata.type. Setting PolicyType to $PolicyType"}
            }
            
            # Loop through configuration within json file and write it to Excel
            Foreach ($Config in $Configs){
                If ($Config.Name -eq 'assignments'){
                    foreach ($target in $JSON.($Config.Name).target){
                        $Worksheet.Cells.Item($row,1) = $PolicyType
                        $Worksheet.Cells.Item($row,2) = $JSON.displayName
                        If ($target."@odata.type" -eq "#microsoft.graph.exclusionGroupAssignmentTarget"){
                            $Worksheet.Cells.Item($row,3) = "$($Config.Name) - exclude"
                        }
                        Else{
                            $Worksheet.Cells.Item($row,3) = "$($Config.Name) - include"
                        }
                        If ($target.displayName){
                            $Worksheet.Cells.Item($row,4) = $target.displayName
                        }
                        Else{
                            $Worksheet.Cells.Item($row,4) = $target.groupId
                        }
                        $row++
                    }
                }
                ElseIf ($Config.Name -eq 'omaSettings'){
                    foreach ($setting in $JSON.omaSettings){
                        $Worksheet.Cells.Item($row,1) = $PolicyType
                        $Worksheet.Cells.Item($row,2) = $JSON.displayName
                        $Worksheet.Cells.Item($row,3) = $Config.Name
                        $Worksheet.Cells.Item($row,4) = ConvertTo-MultilineString -array ($setting | Select-Object displayName, description, omaUri, value)
                        $row++
                    }
                }
                ElseIf ($Config.Name -eq 'settingsDelta'){
                    foreach ($setting in $JSON.settingsDelta){
                        If ($setting.'@odata.type' -eq '#microsoft.graph.deviceManagementComplexSettingInstance'){
                            $Worksheet.Cells.Item($row,1) = $PolicyType
                            $Worksheet.Cells.Item($row,2) = $JSON.displayName
                            $Worksheet.Cells.Item($row,3) = $Config.Name
                            If ($setting.value.valueJson.length -gt 4096){
                                $Worksheet.Cells.Item($row,4) = "Value too big for an Excel cell"
                            }
                            Else{
                                $Worksheet.Cells.Item($row,4) = ConvertTo-MultilineString -array ($setting.value | Select-Object definitionId, valueJson)
                            }
                            $row++
                        }
                        Else{
                            $Worksheet.Cells.Item($row,1) = $PolicyType
                            $Worksheet.Cells.Item($row,2) = $JSON.displayName
                            $Worksheet.Cells.Item($row,3) = $Config.Name
                            If ($setting.valueJson.length -gt 4096){
                                $Worksheet.Cells.Item($row,4) = "Value too big for an Excel cell"
                            }
                            Else{
                                $Worksheet.Cells.Item($row,4) = ConvertTo-MultilineString -array ($setting | Select-Object definitionId, valueJson)
                            }
                            $row++
                        }
                    }
                }
                ElseIf ($Config.Name -in ('encodedSettingXml','payloadJson','startMenuLayoutXml')){
                    $Worksheet.Cells.Item($row,1) = $PolicyType
                    $Worksheet.Cells.Item($row,2) = $JSON.displayName
                    $Worksheet.Cells.Item($row,3) = $Config.Name
                    $Worksheet.Cells.Item($row,4) = ConvertTo-MultilineString -array ([System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($JSON.($Config.Name))))
                    $row++
                }
                ElseIf ($JSON.($Config.Name) -notin $ExcludeValues){
                    $Worksheet.Cells.Item($row,1) = $PolicyType
                    $Worksheet.Cells.Item($row,2) = $JSON.displayName
                    $Worksheet.Cells.Item($row,3) = $Config.Name
                    $Worksheet.Cells.Item($row,4) = ConvertTo-MultilineString -array $JSON.($Config.Name)
                    $row++
                }
                Else{
                    If ($VerboseLog){Write-Log -Message "Configuration not written to Excel file: $($Config.Name)"}
                }
            }
        }
    }
}
#endregion


$Version = "2024-01-23"

# Set log file name to the script name
$LogName = $MyInvocation.MyCommand.Name.Substring(0,$MyInvocation.MyCommand.Name.LastIndexOf("."))

Write-Log -Message "***************************************************"
Write-Log -Message "Starting script version $Version"

If (!(Test-Path $InputFolder)){
    Write-Log -Message "Unable to find input folder: $InputFolder. Quitting" -severity 3
    exit 1
}

If ((Test-Path $ExcelFilename) -and -not $Force){
    Write-Log -Message "Excel file $ExcelFilename found. Specify another filename or use -Force to replace existing file" -severity 3
    exit 1
}

# Create Excel object
try{
    $Excel = New-Object -ComObject Excel.Application
    $Workbook = $Excel.Workbooks.Add()
    If ($VerboseLog){Write-Log -Message "VERBOSE: Created Excel object"}
}
catch{
    Write-Log -Message "Unable to create Excel object" -severity 3
    Write-Log -Message $_.Exception.Message -Severity 3
    Write-Log -Message $_.Exception.ItemName -Severity 3
    Write-Log -Message $_.Exception -Severity 3
    exit 1
}

# Exclude certain data types and values in Excel file
$ExcludedDataTypes = @('@odata.type','displayName','id','roleScopeTagIds','supportsScopeTags','version','createdDateTime','lastModifiedDateTime','assignments@odata.context')
$ExcludeValues = @($null,'notConfigured')

# Loop through folders
$Subfolders = Get-ChildItem -Path $InputFolder -Directory | Sort-Object -descending
Foreach ($Folder in $Subfolders){
    If ($VerboseLog){Write-Log -Message "VERBOSE: Process folder $($Folder.FullName)"}
    Add-PoliciesToExcel -Subfolder $Folder.FullName
    If ($Folder.Name -eq "EnrollmentRestrictions" -or $Folder.Name -eq "DeviceConfigurationPolicy"){
        # Enrollment restrictions are stored in a subfolder
        $EnrollmentRestrictions = Get-ChildItem -Path $Folder.FullName -Directory
        Foreach ($EnrollmentRestriction in $EnrollmentRestrictions){
            Add-PoliciesToExcel -Subfolder $EnrollmentRestriction.FullName
        }
    }
}

# Loop through Excel file and make formatting changes
Foreach ($sheet in $Workbook.Sheets){
    If ($sheet.name -eq "Sheet1"){
        If ($VerboseLog){Write-Log -Message "VERBOSE: Remove the default worksheet"}
        $sheet.delete()
    }
    Else{
        If ($VerboseLog){Write-Log -Message "VERBOSE: Formatting Excel worksheet: $($sheet.name)"}
        $usedRange = $sheet.UsedRange
        $usedRange.ColumnWidth = 100
        $usedRange.Autofilter() | Out-Null
        $usedRange.EntireColumn.AutoFit() | Out-Null
        $usedRange.EntireRow.AutoFit() | Out-Null
    }
}

# Remove existing file if exists
If (Test-Path $ExcelFilename){
    Write-Log -Message "Excel file $ExcelFilename already exists. Removing" -severity 2
    Remove-Item $ExcelFilename -Force
}

# Save file and exit Excel object
try{
    # Excel doesn't like relative paths, try to remediate
    If ($ExcelFilename.IndexOf(".") -eq 0){
        If ($VerboseLog){Write-Log -Message "VERBOSE: Found relative path $ExcelFilename"}
        $ExcelFilename = Join-Path -Path (Get-Location).Path -ChildPath $ExcelFilename
        If ($VerboseLog){Write-Log -Message "VERBOSE: New ExcelFilename $ExcelFilename"}
    }
    $Workbook.SaveAs($ExcelFilename)
    $Excel.Quit()
    Write-Log -Message "Save Excel file $ExcelFilename"
}
catch{
    Write-Log -Message "Unable to save Excel file $ExcelFilename" -severity 3
    Write-Log -Message $_.Exception.Message -Severity 3
    Write-Log -Message $_.Exception -Severity 3
}

Write-Log -Message "Script end"