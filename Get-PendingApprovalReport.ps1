<#
.SYNOPSIS
    Scans SharePoint sites and libraries listed in a CSV file and generates an Excel report
    and an on-disk text log file for content items pending approval.

.DESCRIPTION
    Requires content approval (versioning/moderation) to be enabled on target lists.
    Groups CSV input by SiteUrl (case-insensitive) to minimize connection overhead.
    Logs execution progress, warnings, and errors to both console and an on-disk text log file.

.PARAMETER CsvPath
    Path to input CSV. Expected headers: SiteUrl, LibraryName
    Example CSV structure:
        SiteUrl,LibraryName
        https://contoso.sharepoint.com/sites/TeamA,Shared Documents
        https://contoso.sharepoint.com/sites/TeamA,Site Pages
        https://contoso.sharepoint.com/sites/TeamB,Contracts

.PARAMETER OutputPath
    Path for the output .xlsx report. Defaults to a timestamped file in the script directory.

.PARAMETER LogPath
    Path for the text log file. Defaults to a timestamped file matching OutputPath basename.

.PARAMETER ReadyForReviewFieldName
    Internal name of the "Ready for review" Yes/No column on document libraries. Default: "ReadyforReview".

.PARAMETER SitePagesLibraryName
    The LibraryName value in the CSV that should be treated as Site Pages. Default: "Site Pages".

.PARAMETER UseCertificateAuth
    Switch to use app-only certificate auth instead of interactive login. Default: $true.

.PARAMETER ClientId
    Entra ID App Registration Client ID.

.PARAMETER Tenant
    SharePoint / Entra Tenant Name (e.g., contoso.onmicrosoft.com).

.PARAMETER CertificatePath
    Path to .pfx certificate file.

.PARAMETER CertificatePassword
    Password for the .pfx certificate file.

.EXAMPLE
    .\Get-PendingApprovalReport.ps1 -CsvPath ".\sites.csv"

.EXAMPLE
    .\Get-PendingApprovalReport.ps1 -CsvPath ".\sites.csv" -UseCertificateAuth `
        -ClientId "00000000-0000-0000-0000-000000000000" -Tenant "contoso.onmicrosoft.com" `
        -CertificatePath ".\cert.pfx" -CertificatePassword "P@ssw0rd123"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CsvPath = ".\sites.csv",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$LogPath,

    [Parameter(Mandatory = $false)]
    [string]$ReadyForReviewFieldName = "ReadyforReview",

    [Parameter(Mandatory = $false)]
    [string]$SitePagesLibraryName = "Site Pages",

    [Parameter(Mandatory = $false)]
    [switch]$UseCertificateAuth = $true,

    [Parameter(Mandatory = $false)]
    [string]$ClientId = "YOUR_CLIENT_ID_HERE",

    [Parameter(Mandatory = $false)]
    [string]$Tenant = "YOUR_TENANT.onmicrosoft.com",

    [Parameter(Mandatory = $false)]
    [string]$CertificatePath = ".\cert.pfx",

    [Parameter(Mandatory = $false)]
    [string]$CertificatePassword = "YOUR_CERTIFICATE_PASSWORD"
)

# ---------------------------------------------------------------------------
# Defaults & Initialization
# ---------------------------------------------------------------------------
$runTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

if (-not $OutputPath) {
    $OutputPath = ".\PendingApprovalReport_$runTimestamp.xlsx"
}
if (-not $LogPath) {
    $LogPath = ".\PendingApprovalReport_$runTimestamp.log"
}

# Ensure parent directory for log file exists
$logDir = [System.IO.Path]::GetDirectoryName($LogPath)
if ($logDir -and -not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

# ---------------------------------------------------------------------------
# Logging Function
# ---------------------------------------------------------------------------
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO",

        [ConsoleColor]$ForegroundColor
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $formattedMsg = "[$timestamp] [$Level] $Message"

    # Write to text log file
    Add-Content -Path $LogPath -Value $formattedMsg -Encoding UTF8 -ErrorAction SilentlyContinue

    # Console color mapping
    if (-not $PSBoundParameters.ContainsKey('ForegroundColor')) {
        switch ($Level) {
            "INFO"    { $ForegroundColor = [ConsoleColor]::Gray }
            "WARNING" { $ForegroundColor = [ConsoleColor]::Yellow }
            "ERROR"   { $ForegroundColor = [ConsoleColor]::Red }
            "SUCCESS" { $ForegroundColor = [ConsoleColor]::Green }
        }
    }

    Write-Host $formattedMsg -ForegroundColor $ForegroundColor
}

# ---------------------------------------------------------------------------
# Module bootstrap
# ---------------------------------------------------------------------------
function Ensure-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Log "Installing module '$Name' (CurrentUser scope)..." -Level "WARNING"
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $Name -ErrorAction Stop
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Get-ModerationStatusText {
    param($StatusValue)
    switch ([int]$StatusValue) {
        0 { "Approved" }
        1 { "Rejected" }
        2 { "Pending" }
        3 { "Draft" }
        4 { "Scheduled" }
        default { "Unknown ($StatusValue)" }
    }
}

function Connect-TargetSite {
    param([string]$Url)
    if ($UseCertificateAuth) {
        if (-not ($ClientId -and $Tenant -and $CertificatePath -and $CertificatePassword)) {
            throw "UseCertificateAuth requires -ClientId, -Tenant, -CertificatePath and -CertificatePassword."
        }
        Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $Tenant `
            -CertificatePath $CertificatePath `
            -CertificatePassword (ConvertTo-SecureString $CertificatePassword -AsPlainText -Force) `
            -ErrorAction Stop
    }
    else {
        Connect-PnPOnline -Url $Url -Interactive -ErrorAction Stop
    }
}

# ---------------------------------------------------------------------------
# Execution Wrapper
# ---------------------------------------------------------------------------
try {
    Write-Log "==========================================================" -Level "INFO"
    Write-Log "Starting Pending Content Export Script" -Level "INFO"
    Write-Log "CSV Path: $CsvPath" -Level "INFO"
    Write-Log "Excel Output Path: $OutputPath" -Level "INFO"
    Write-Log "Text Log Path: $LogPath" -Level "INFO"
    Write-Log "Use Certificate Auth: $UseCertificateAuth" -Level "INFO"
    if ($UseCertificateAuth) {
        Write-Log "ClientId: $ClientId" -Level "INFO"
        Write-Log "Tenant: $Tenant" -Level "INFO"
        Write-Log "Certificate Path: $CertificatePath" -Level "INFO"
    }
    Write-Log "==========================================================" -Level "INFO"

    Ensure-Module -Name "PnP.PowerShell"
    Ensure-Module -Name "ImportExcel"

    if (-not (Test-Path $CsvPath)) {
        throw "CSV file not found at path: $CsvPath"
    }

    $rows = Import-Csv -Path $CsvPath
    if (-not $rows -or $rows.Count -eq 0) {
        throw "CSV file is empty or could not be parsed."
    }

    # Validate and normalise rows, skipping blanks early
    $validRows = [System.Collections.Generic.List[PSCustomObject]]::new()
    $rowNum = 0
    foreach ($row in $rows) {
        $rowNum++
        $siteUrl     = if ($row.SiteUrl) { ($row.SiteUrl).Trim() } else { "" }
        $libraryName = if ($row.LibraryName) { ($row.LibraryName).Trim() } else { "" }

        if (-not $siteUrl -or -not $libraryName) {
            Write-Log "Row [$rowNum]: Skipping - SiteUrl or LibraryName is blank." -Level "WARNING"
            continue
        }
        $validRows.Add([PSCustomObject]@{
            RowNum      = $rowNum
            SiteUrl     = $siteUrl
            LibraryName = $libraryName
        })
    }

    if ($validRows.Count -eq 0) {
        throw "No valid rows found in CSV."
    }

    # Group by SiteUrl (case-insensitive) so we connect once per unique site
    $siteGroups = $validRows | Group-Object -Property { $_.SiteUrl.ToLowerInvariant() }
    $totalSites  = $siteGroups.Count
    $siteIndex   = 0

    $results  = @()
    $errorLog = @()

    foreach ($group in $siteGroups) {
        $siteIndex++
        $siteUrl   = $group.Group[0].SiteUrl
        $libraries = $group.Group

        Write-Log "Site [$siteIndex/$totalSites]: $siteUrl ($($libraries.Count) library/libraries)" -Level "INFO" -ForegroundColor "Magenta"

        try {
            Connect-TargetSite -Url $siteUrl
            Write-Log "  Successfully connected to $siteUrl" -Level "INFO"

            $web     = Get-PnPWeb -ErrorAction Stop
            $baseUrl = ([System.Uri]$web.Url).GetLeftPart([System.UriPartial]::Authority)

            foreach ($entry in $libraries) {
                $libraryName = $entry.LibraryName
                Write-Log "  Processing library '$libraryName'..." -Level "INFO" -ForegroundColor "Cyan"

                try {
                    $list = Get-PnPList -Identity $libraryName -ErrorAction Stop

                    if (-not $list.EnableModeration) {
                        Write-Log "  Library '$libraryName' does not have Content Approval enabled. Skipping." -Level "WARNING"
                        continue
                    }

                    $isSitePages = $libraryName -ieq $SitePagesLibraryName

                    $fieldsToRetrieve = @("FileRef", "FileLeafRef", "_ModerationStatus")
                    $hasReadyField    = $false

                    if (-not $isSitePages) {
                        $readyField = Get-PnPField -List $list -Identity $ReadyForReviewFieldName -ErrorAction SilentlyContinue
                        if ($readyField) {
                            $fieldsToRetrieve += $ReadyForReviewFieldName
                            $hasReadyField     = $true
                        }
                        else {
                            Write-Log "  Field '$ReadyForReviewFieldName' not found on '$libraryName'. Filtering on Approval Status = Draft only." -Level "WARNING"
                        }
                    }

                    $items = Get-PnPListItem -List $list -PageSize 500 -Fields $fieldsToRetrieve

                    $foundInLibrary = 0
                    foreach ($item in $items) {
                        $moderationStatus = $item.FieldValues["_ModerationStatus"]
                        if ($null -eq $moderationStatus) { continue }

                        $statusText = Get-ModerationStatusText -StatusValue $moderationStatus
                        if ($statusText -ne "Draft") { continue }

                        if ($isSitePages) {
                            $readyValue = "N/A"
                        }
                        else {
                            if ($hasReadyField) {
                                $readyRaw   = $item.FieldValues[$ReadyForReviewFieldName]
                                $readyValue = if ($readyRaw -eq $true) { "Yes" } elseif ($readyRaw -eq $false) { "No" } else { "" }
                                if ($readyValue -ne "Yes") { continue }
                            }
                            else {
                                $readyValue = ""
                            }
                        }

                        $fileRef  = $item.FieldValues["FileRef"]
                        $fullLink = "$baseUrl$fileRef"

                        $results += [PSCustomObject]@{
                            "Site Url"           = $siteUrl
                            "Library"            = $libraryName
                            "Document/Page Link" = $fullLink
                            "Ready for Review"   = $readyValue
                            "Approval Status"    = $statusText
                        }
                        $foundInLibrary++
                    }

                    Write-Log "  Finished '$libraryName' ($foundInLibrary pending item(s) found)." -Level "SUCCESS"
                }
                catch {
                    $errMsg = "Error processing '$libraryName' on ${siteUrl}: $($_.Exception.Message)"
                    Write-Log "  $errMsg" -Level "ERROR"
                    $errorLog += [PSCustomObject]@{
                        "Site Url" = $siteUrl
                        "Library"  = $libraryName
                        "Error"    = $_.Exception.Message
                    }
                }
            }
        }
        catch {
            $connErrMsg = "Failed to connect to ${siteUrl}: $($_.Exception.Message)"
            Write-Log "  $connErrMsg" -Level "ERROR"
            foreach ($entry in $libraries) {
                $errorLog += [PSCustomObject]@{
                    "Site Url" = $siteUrl
                    "Library"  = $entry.LibraryName
                    "Error"    = "Connection failed: $($_.Exception.Message)"
                }
            }
        }
        finally {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
    }

    # ---------------------------------------------------------------------------
    # Report Output
    # ---------------------------------------------------------------------------
    Write-Log "==========================================================" -Level "INFO"
    Write-Log "Execution Summary:" -Level "INFO"
    Write-Log "  Total Unique Sites Processed : $totalSites" -Level "INFO"
    Write-Log "  Total Pending Items Found    : $($results.Count)" -Level "INFO"
    Write-Log "  Total Errors Encountered     : $($errorLog.Count)" -Level "INFO"

    if ($results.Count -gt 0) {
        $results | Export-Excel -Path $OutputPath -WorksheetName "Pending Approvals" `
            -AutoSize -TableStyle Medium2 -FreezeTopRow -BoldTopRow
        Write-Log "Excel Report saved successfully to: $OutputPath" -Level "SUCCESS"
    }
    else {
        Write-Log "No pending items found across all processed sites." -Level "WARNING"
    }

    if ($errorLog.Count -gt 0) {
        $errorLog | Export-Excel -Path $OutputPath -WorksheetName "Errors" `
            -AutoSize -TableStyle Medium6 -FreezeTopRow -BoldTopRow
        Write-Log "Errors log worksheet added to Excel Report ($($errorLog.Count) error(s))." -Level "WARNING"
    }

    Write-Log "Text log file written to: $LogPath" -Level "INFO"
    Write-Log "Script completed successfully." -Level "SUCCESS"
    Write-Log "==========================================================" -Level "INFO"
}
catch {
    Write-Log "CRITICAL SCRIPT FAILURE: $($_.Exception.Message)" -Level "ERROR"
    if ($_.ScriptStackTrace) {
        Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
    }
    throw $_
}
