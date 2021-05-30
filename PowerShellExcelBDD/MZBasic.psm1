$global:MZLogger = [System.Text.StringBuilder]::new()
$global:MZ_StringBuilder = [System.Text.StringBuilder]::new()

$global:SEP = [IO.Path]::DirectorySeparatorChar

function Add-Log {
    param (
        [String]$message
    )
    [void]$MZLogger.AppendLine($message)
}

function Set-MZLogPath {
    param (
        $LogPath,
        $Topic = 'dbcreate'
    )
    $ExecutionStartTime = Get-Date
    Write-MZLog "Execution Start Time: $($ExecutionStartTime.ToString())"
    if ($LogPath) {
        if (Test-Path  $LogPath) {
            $Script:LogPath = Join-Path $LogPath "${Topic}_$($ExecutionStartTime.ToString("dd-MM-yyyy_HH-mm")).log"
        }
    }
    else {
        $Script:LogPath = "${Topic}_$($ExecutionStartTime.ToString("dd-MM-yyyy_HH-mm")).log"
    }
    Write-MZDebug "LogPath $Script:LogPath"
}

function Set-MZLogFile {
    Set-Content -Path $LogPath -Value $(Get-MZLog)
}
<#
.Description
Show message and log with time stamp and color
.Example
Write-MZLog "message" -ForegroundColor Red -BackgroundColor Green
#>
function Write-MZLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true)]$message,
        $ForegroundColor = $null, #$Host.UI.RawUI.ForegroundColor,
        $BackgroundColor = $null #$Host.UI.RawUI.BackgroundColor,
    )
    $message = "$(Get-Date):$message"
    [void]$MZLogger.AppendLine($message)
    $msg = @{
        Object    = $message
        NoNewline = $false
    }
    if ($ForegroundColor) { $msg["ForegroundColor"] = $ForegroundColor }
    if ($BackgroundColor) { $msg["BackgroundColor"] = $BackgroundColor }

    Write-Host @msg
}

<#
.Description
Show message in Red and log
.Example
Write-MZRedLog "something wrong"
#>
function Write-MZRedLog {
    param (
        [String]$message
    )
    Write-MZLog $message -ForegroundColor Red
}

<#
.Description
Show Debug message and log debug message with time stamp
#>
function Write-MZDebug($message) {
    $message = "$(Get-Date): $message"
    if ($global:DebugMode) {
        Write-Host $message -ForegroundColor Gray
    }
    [void]$MZLogger.AppendLine("$message")
}

<#
.Description
Clear log 
#>
function Clear-MZLog {
    [void]$MZLogger.Clear()
}

<#
.Description
Add string with line end into global string builder
#>
function Write-MZStringBuilder($string) {
    [void]$MZ_StringBuilder.AppendLine($string) 
}

<#
.Description
Add string without line end into global string builder
#>
function Add-MZStringBuilder($string) {
    [void]$MZ_StringBuilder.Append($string) 
}

<#
.Description
Clear global string builder
#>
function Clear-MZStringBuilder {
    [void]$MZ_StringBuilder.Clear()
}


<#
.Description
Get string from global string builder
#>
function Get-MZStringFromBuilder {
    return  $MZ_StringBuilder.ToString()
}

function Get-MZLog {
    return $MZLogger.ToString()    
}



<#
.Description install module if it is not installed, then import
#>
function Install-AzPowershellModulesIfNotExists {
    param (
        [Parameter(Mandatory = $true)][String]$ModuleName,
        [String]$MinVersion
    )
    $TheModule = Get-InstalledModule -Name $ModuleName -ErrorAction SilentlyContinue
    try {
        if ($null -eq $TheModule) {
            Write-MZDebug "Module $ModuleName does not exist. Installing now..." -ForegroundColor Yellow
            Install-Module -Name $ModuleName  -Scope CurrentUser -Force
        }
        elseif ($MinVersion) {
            if ($TheModule.Version -lt $MinVersion) {
                Write-MZDebug "Module $ModuleName version $($TheModule.Version) exists, which need be updated to higher than $MinVersion."
                Update-Module -Name $ModuleName
            }
            else {
                Write-MZDebug "Module $ModuleName version $($TheModule.Version) exists, which meets min version $MinVersion."  
            }
        }
        else {
            Write-MZDebug "Module $ModuleName version $($TheModule.Version) exists."    
        }
    }
    catch {
        Write-MZRedLog $_.Exception.Message
    }
    Import-Module -Name $ModuleName
}
<#
.Description Show detail of an object
#>
function Show-Object($ob) {
    if ($null -eq $ob) {
        Write-Host "The Object is null."
        return
    }
    Write-Host "Object Type:$($ob.GetType())"
    switch ($ob.GetType().Name) {
        Hashtable { Write-Host ($ob | ConvertTo-Json) }
        OrderedDictionary { Write-Host ($ob | ConvertTo-Json) }
        Default {
            $ob.PSObject.Properties | ForEach-Object {
                Write-Host "$($_.Name):$($_.Value)"
            }
            if ($ob -is [pscustomobject]) {
                Write-Host ($ob | ConvertTo-Json)
            }
        }
    }
}

<#
.Description Return how many item in ob,
if no count attribute, return 1
#>
function Get-MZObjectCount ($ob) {
    if (-not $ob) {
        return 0
    }
    elseif ($ob -is [array]) {
        return $ob.count
    }
    elseif ($ob.GetType().Name -match "list") {
        return $ob.count
    }
    else {
        return 1
    }
}
<#
.Description
Check whether the param has any value
#>
function Test-MZHasValue($strValue) {
    if ($null -eq $strValue) {
        return $false
    }
    if ("" -eq $strValue.trim()) {
        return $false
    }      
    return $true
}

<#
.Description
Check the string is a num
#>
function Test-IsValidNum($strNum) {
    $Num = $strNum -as [int]
    if ($Num -le 0) {
        return $false
    }
    return ($Num -is [int])     
}

<#
.Description Check property according to the dynamic rule
#>
function Test-MZIsPropertyValid {
    param (
        [String]$PropertyName,
        [String]$PropertyValue,
        [Hashtable]$Rule
    )
    $IsValid = $true
    if ($null -eq $Rule ) {
        return $true
    }
    if ($Rule.NotNull) { 
        if (-not (Test-MZHasValue $PropertyValue)) {
            $IsValid = $false
            Write-MZRedLog "Value of $PropertyName is blank."
        }
    }
    if (Test-IsValidNum $Rule.MaxLength) { 
        if ($PropertyValue.length -gt $Rule.MaxLength) {
            $IsValid = $false
            Write-MZRedLog "Length of '$PropertyValue' exceeds max length $($Rule.MaxLength) of $PropertyName."
        }
    }

    if (Test-IsValidNum $Rule.MinLength) { 
        if ($PropertyValue.length -lt $Rule.MinLength) {
            $IsValid = $false
            Write-MZRedLog "Length of '$PropertyValue' doesn't meet min length $($Rule.MinLength) of $PropertyName."
        }
    }

    if ($Rule.FormatPattern) {
        if (($Rule.NotNull -eq $false) -and [String]::IsNullOrEmpty($PropertyValue)) {
            Write-MZLog "Value of $PropertyName is blank." 
        }
        elseif ($PropertyValue -notmatch $Rule.FormatPattern) {
            $IsValid = $false
            Write-MZRedLog "'$PropertyValue' doesn't match [$($Rule.FormatPattern)] Format of $PropertyName."
        }
    }

    if ($Rule.EmailAddress) {
        #check Email address, e.g. Service Owner
        $EmailArray = $PropertyValue -split ';'
        foreach ($Email in $EmailArray) {
            if (-not [bool]($Email -as [mailaddress])) {
                Write-MZRedLog "According to rule, $PropertyName-'$PropertyValue' should be email address, or email address list splitted by ';', but '$($Email)' is wrong email address format."
                $IsValid = $false
            }
        }
    }
    return $IsValid
}