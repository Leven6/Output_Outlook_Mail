[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Sender,

    [string]$OutputPath = (Join-Path $PWD 'outlook-export'),

    [switch]$Recurse,

    [string]$StartDate,

    [string]$EndDate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function ConvertTo-ExportDate {
    param(
        [Parameter(Mandatory = $true)][string]$Value,
        [Parameter(Mandatory = $true)][string]$ParameterName
    )

    $parsedDate = [datetime]::MinValue
    if (-not [datetime]::TryParseExact(
        $Value.Trim(),
        'yyyy-MM-dd',
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::None,
        [ref]$parsedDate
    )) {
        throw "$ParameterName must be a valid date in yyyy-MM-dd format (for example, 2026-10-01)."
    }

    return $parsedDate
}

function Get-SenderSmtpAddress {
    param([Parameter(Mandatory = $true)]$MailItem)

    try {
        if ($MailItem.SenderEmailType -eq 'EX') {
            $exchangeUser = $MailItem.Sender.GetExchangeUser()
            if ($null -ne $exchangeUser -and $exchangeUser.PrimarySmtpAddress) {
                return [string]$exchangeUser.PrimarySmtpAddress
            }

            $exchangeList = $MailItem.Sender.GetExchangeDistributionList()
            if ($null -ne $exchangeList -and $exchangeList.PrimarySmtpAddress) {
                return [string]$exchangeList.PrimarySmtpAddress
            }
        }

        return [string]$MailItem.SenderEmailAddress
    }
    catch {
        Write-Warning "Could not read the message sender: $($_.Exception.Message)"
        return $null
    }
}

function Get-SafeFileName {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Name
    )

    $invalidPattern = '[{0}]' -f [Regex]::Escape((-join [IO.Path]::GetInvalidFileNameChars()))
    $safeName = ($Name -replace $invalidPattern, '_').Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        return 'no-subject'
    }

    if ($safeName.Length -gt 100) {
        return $safeName.Substring(0, 100).TrimEnd()
    }

    return $safeName
}

function Get-EntryIdSuffix {
    param([Parameter(Mandatory = $true)][string]$EntryId)

    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($EntryId)
        $hash = $sha256.ComputeHash($bytes)
        return ([BitConverter]::ToString($hash) -replace '-', '').Substring(0, 10)
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-LatestExportedDate {
    param([Parameter(Mandatory = $true)][string]$Directory)

    $latestDate = $null
    foreach ($file in Get-ChildItem -LiteralPath $Directory -Filter '*.msg' -File) {
        if ($file.BaseName -notmatch '^(?<date>\d{8}-\d{6}) - ') {
            continue
        }

        $parsedDate = [datetime]::MinValue
        $parsed = [datetime]::TryParseExact(
            $Matches.date,
            'yyyyMMdd-HHmmss',
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None,
            [ref]$parsedDate
        )
        if ($parsed -and ($null -eq $latestDate -or $parsedDate -gt $latestDate)) {
            $latestDate = $parsedDate
        }
    }

    return $latestDate
}

function Export-OutlookFolder {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]$Folder,
        [Parameter(Mandatory = $true)][string]$TargetSender,
        [Parameter(Mandatory = $true)][string]$Destination,
        [Parameter(Mandatory = $true)][bool]$IncludeSubfolders,
        [AllowNull()][Nullable[datetime]]$Since,
        [Parameter(Mandatory = $true)][datetime]$Until
    )

    Write-Verbose "Scanning folder: $($Folder.FolderPath)"
    $items = $Folder.Items

    try {
        for ($index = 1; $index -le $items.Count; $index++) {
            $item = $null
            try {
                $item = $items.Item($index)
                if ($null -eq $item -or $item.Class -ne 43) {
                    continue
                }

                $sentOn = [datetime]$item.SentOn
                if (($null -ne $Since -and $sentOn -lt [datetime]$Since) -or $sentOn -gt $Until) {
                    continue
                }

                $senderAddress = Get-SenderSmtpAddress -MailItem $item
                if ($senderAddress -ine $TargetSender) {
                    continue
                }

                $sentDate = $sentOn.ToString('yyyyMMdd-HHmmss')
                $subject = Get-SafeFileName -Name ([string]$item.Subject)
                $suffix = Get-EntryIdSuffix -EntryId ([string]$item.EntryID)
                $filePath = Join-Path $Destination "$sentDate - $subject - $suffix.msg"

                if (Test-Path -LiteralPath $filePath) {
                    $script:AlreadyExportedCount++
                    continue
                }

                if ($PSCmdlet.ShouldProcess($filePath, 'Export Outlook message with attachments')) {
                    $item.SaveAs($filePath, 9)
                    $script:ExportedCount++
                    Write-Verbose "Exported: $filePath"
                }
            }
            catch {
                $script:FailedCount++
                Write-Warning "Skipped a message that could not be processed: $($_.Exception.Message)"
            }
            finally {
                if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item)
                }
            }
        }
    }
    finally {
        if ([Runtime.InteropServices.Marshal]::IsComObject($items)) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($items)
        }
    }

    if ($IncludeSubfolders) {
        foreach ($subfolder in $Folder.Folders) {
            try {
                Export-OutlookFolder -Folder $subfolder -TargetSender $TargetSender -Destination $Destination -IncludeSubfolders $true -Since $Since -Until $Until
            }
            finally {
                if ([Runtime.InteropServices.Marshal]::IsComObject($subfolder)) {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($subfolder)
                }
            }
        }
    }
}

$manualDateRange = -not [string]::IsNullOrWhiteSpace($StartDate) -or -not [string]::IsNullOrWhiteSpace($EndDate)
$exportSince = $null
$exportUntil = Get-Date
if (-not [string]::IsNullOrWhiteSpace($StartDate)) {
    $exportSince = ConvertTo-ExportDate -Value $StartDate -ParameterName 'StartDate'
}
if (-not [string]::IsNullOrWhiteSpace($EndDate)) {
    $exportUntil = (ConvertTo-ExportDate -Value $EndDate -ParameterName 'EndDate').AddTicks([TimeSpan]::TicksPerDay - 1)
}
if ($null -ne $exportSince -and $exportSince -gt $exportUntil) {
    throw 'StartDate must not be later than EndDate (or the current time when EndDate is omitted).'
}

$destinationDirectory = [IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path -LiteralPath $destinationDirectory)) {
    [void](New-Item -ItemType Directory -Path $destinationDirectory -Force)
}

$senderFolderName = Get-SafeFileName -Name $Sender.Trim()
$senderDirectory = Join-Path $destinationDirectory $senderFolderName
if (-not (Test-Path -LiteralPath $senderDirectory)) {
    [void](New-Item -ItemType Directory -Path $senderDirectory -Force)
}

if ($manualDateRange) {
    $startLabel = if ($null -eq $exportSince) { 'the earliest message' } else { $exportSince.ToString('yyyy-MM-dd HH:mm:ss') }
    Write-Host "Manual date range: $startLabel through $($exportUntil.ToString('yyyy-MM-dd HH:mm:ss'))."
    Write-Host 'Ignoring the latest exported date; existing files will still be skipped.'
}
else {
    $exportSince = Get-LatestExportedDate -Directory $senderDirectory
    if ($null -eq $exportSince) {
        Write-Host "No previous export found. Exporting messages through $($exportUntil.ToString('yyyy-MM-dd HH:mm:ss'))."
    }
    else {
        Write-Host "Latest existing message: $($exportSince.ToString('yyyy-MM-dd HH:mm:ss'))."
        Write-Host "Exporting from that time through $($exportUntil.ToString('yyyy-MM-dd HH:mm:ss'))."
    }
}

$outlook = $null
$namespace = $null
$inbox = $null
$script:ExportedCount = 0
$script:AlreadyExportedCount = 0
$script:FailedCount = 0

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace('MAPI')
    $inbox = $namespace.GetDefaultFolder(6)

    Export-OutlookFolder `
        -Folder $inbox `
        -TargetSender $Sender.Trim() `
        -Destination $senderDirectory `
        -IncludeSubfolders $Recurse.IsPresent `
        -Since $exportSince `
        -Until $exportUntil
}
finally {
    foreach ($comObject in @($inbox, $namespace, $outlook)) {
        if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
        }
    }
}

Write-Host "Done. Exported: $script:ExportedCount; already present: $script:AlreadyExportedCount; failed: $script:FailedCount."
Write-Host "Output: $senderDirectory"
