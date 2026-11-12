#requires -Version 7.2
#requires -Modules PnP.PowerShell

<#
.SYNOPSIS
Analyzes the storage used by a SharePoint site's version history and, if requested,
removes previous file versions.

.DESCRIPTION
By default, the script only analyzes the site. Specify -DeleteAllVersions to
remove version history. The -WhatIf option analyzes the site and simulates
deletion without making changes to SharePoint.
Files with no previous versions are considered successfully analyzed.
By default, an error reading a file's version history stops the analysis.
When -DeleteAllVersions is specified, per-file history read errors are logged
and skipped automatically; versions are removed only from files analyzed
successfully. Library enumeration errors always prevent deletion.
The -ContinueOnVersionHistoryError parameter is retained for compatibility
and is no longer required.

The reported size is the sum of the sizes of previous versions returned by
SharePoint; it is not a guaranteed measure of actual or billable tenant storage.

The analysis includes non-hidden document libraries. Specify
-IncludeHiddenLibraries to include hidden libraries as well. The signed-in
account must be able to read all included files and, when deleting, remove
their versions. Reports are saved locally, including when -WhatIf is used.
Version histories are retrieved in batches; -BatchSize sets the maximum
number of files per request and defaults to 100.
Specify -FolderServerRelativeUrl to limit the analysis to an existing site
folder and its subfolders. For libraries with more than 5,000 items, the
script enumerates the library with Get-PnPFolderItem -List and filters for
the requested folder; for smaller libraries, it enumerates the folder
directly. The parameter accepts a server-relative path or an absolute HTTPS
folder URL. SharePoint navigation links are supported if they contain exactly
one valid id parameter. The URL must belong to the site specified by -SiteUrl
and must not include a fragment.

.EXAMPLE
.\Manage-SharePointVersionHistory.ps1 -SiteUrl https://contoso.sharepoint.com/sites/project -ClientId <AppId>

.EXAMPLE
.\Manage-SharePointVersionHistory.ps1 -SiteUrl https://contoso.sharepoint.com/sites/project -ClientId <AppId> -DeleteAllVersions -WhatIf

.EXAMPLE
.\Manage-SharePointVersionHistory.ps1 -SiteUrl https://contoso.sharepoint.com/sites/project -ClientId <AppId> -FolderServerRelativeUrl '/sites/project/Shared Documents/Archive'

.EXAMPLE
.\Manage-SharePointVersionHistory.ps1 -SiteUrl https://contoso.sharepoint.com/sites/project -ClientId <AppId> -FolderServerRelativeUrl 'https://contoso.sharepoint.com/sites/project/Shared%20Documents/Archive'

.EXAMPLE
.\Manage-SharePointVersionHistory.ps1 -SiteUrl https://contoso.sharepoint.com/sites/project -ClientId <AppId> -DeleteAllVersions -ContinueOnVersionHistoryError -WhatIf
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory)]
    [guid]$ClientId,

    [string]$OutputDirectory = (Join-Path $PSScriptRoot 'VersionHistoryReports'),

    [ValidateRange(1, 500)]
    [int]$BatchSize = 100,

    [ValidateNotNullOrEmpty()]
    [string]$FolderServerRelativeUrl,

    [switch]$DeleteAllVersions,

    [switch]$ContinueOnVersionHistoryError,

    [switch]$IncludeHiddenLibraries
)

$ErrorActionPreference = 'Stop'

function ConvertTo-SafeCsvRows {
    param (
        [object[]]$Rows
    )

    foreach ($Row in $Rows) {
        $SafeValues = [ordered]@{}
        foreach ($Property in $Row.PSObject.Properties) {
            $Value = $Property.Value
            if ($Value -is [string] -and $Value -match '^\s*[=+\-@]') {
                $Value = "'$Value"
            }
            $SafeValues[$Property.Name] = $Value
        }

        [pscustomobject]$SafeValues
    }
}

function Get-FileVersionBytes {
    param (
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Versions
    )

    [long]$TotalBytes = 0
    foreach ($Version in $Versions) {
        if ($Version.PSObject.Properties['Length']) {
            $TotalBytes += [long]$Version.Length
        }
        elseif ($Version.PSObject.Properties['Size']) {
            $TotalBytes += [long]$Version.Size
        }
        else {
            throw 'SharePoint did not return the version size.'
        }
    }

    return $TotalBytes
}

function Get-FileVersionBatchResults {
    param (
        [Parameter(Mandatory)]
        [object[]]$FileEntries,

        [Parameter(Mandatory)]
        $Connection
    )

    Write-Verbose "Retrieving version history for a batch of $($FileEntries.Count) files."
    if ($FileEntries.Count -eq 1) {
        $Entry = $FileEntries[0]
        try {
            $Versions = @(
                Get-PnPFileVersion `
                    -Url $Entry.FileUrl `
                    -Connection $Connection `
                    -ErrorAction Stop
            )
            $VersionBytes = Get-FileVersionBytes -Versions $Versions
            return [pscustomobject]@{
                Entry        = $Entry
                VersionCount = $Versions.Count
                VersionBytes = $VersionBytes
                Status       = 'Analyzed'
                Details      = ''
            }
        }
        catch {
            return [pscustomobject]@{
                Entry        = $Entry
                VersionCount = $null
                VersionBytes = $null
                Status       = 'Error'
                Details      = $_.Exception.Message
            }
        }
    }

    $Context = $Connection.Context
    $QueuedQueries = [System.Collections.Generic.List[object]]::new()
    try {
        foreach ($Entry in $FileEntries) {
            $ResourcePath = [Microsoft.SharePoint.Client.ResourcePath]::FromDecodedUrl($Entry.FileUrl)
            $File = $Context.Web.GetFileByServerRelativePath($ResourcePath)
            $VersionQuery = $Context.LoadQuery($File.Versions)
            $QueuedQueries.Add([pscustomobject]@{
                Entry = $Entry
                Query = $VersionQuery
            })
        }

        $Context.ExecuteQuery()

        foreach ($QueuedQuery in $QueuedQueries) {
            try {
                $Versions = @($QueuedQuery.Query)
                $VersionBytes = Get-FileVersionBytes -Versions $Versions
                [pscustomobject]@{
                    Entry        = $QueuedQuery.Entry
                    VersionCount = $Versions.Count
                    VersionBytes = $VersionBytes
                    Status       = 'Analyzed'
                    Details      = ''
                }
            }
            catch {
                Get-FileVersionBatchResults `
                    -FileEntries @($QueuedQuery.Entry) `
                    -Connection $Connection
            }
        }
    }
    catch {
        Write-Verbose "Batch of $($FileEntries.Count) files failed: $($_.Exception.Message). Splitting the batch or falling back to per-file requests."
        if ($FileEntries.Count -gt 1) {
            $SplitIndex = [Math]::Floor($FileEntries.Count / 2)
            Get-FileVersionBatchResults `
                -FileEntries @($FileEntries[0..($SplitIndex - 1)]) `
                -Connection $Connection
            Get-FileVersionBatchResults `
                -FileEntries @($FileEntries[$SplitIndex..($FileEntries.Count - 1)]) `
                -Connection $Connection
        }
        else {
            Get-FileVersionBatchResults `
                -FileEntries @($FileEntries[0]) `
                -Connection $Connection
        }
    }
}

$SiteUri = $null
if (-not [uri]::TryCreate($SiteUrl, [System.UriKind]::Absolute, [ref]$SiteUri) -or
    $SiteUri.Scheme -ne 'https' -or
    -not [string]::IsNullOrEmpty($SiteUri.UserInfo) -or
    -not [string]::IsNullOrEmpty($SiteUri.Query) -or
    -not [string]::IsNullOrEmpty($SiteUri.Fragment)) {
    throw 'SiteUrl must be an absolute HTTPS URL for a SharePoint site, without credentials, a query string, or a fragment.'
}
$SiteUrl = $SiteUri.AbsoluteUri.TrimEnd('/')

$NormalizedFolderUrl = $null
if (-not [string]::IsNullOrWhiteSpace($FolderServerRelativeUrl)) {
    $FolderInput = $FolderServerRelativeUrl.Trim()
    if ($FolderInput -match '[#\\]') {
        throw 'FolderServerRelativeUrl must not contain a fragment or backslashes.'
    }

    if ($FolderInput -match '^[A-Za-z][A-Za-z0-9+.-]*://') {
        $FolderUri = $null
        if (-not [uri]::TryCreate($FolderInput, [System.UriKind]::Absolute, [ref]$FolderUri) -or
            $FolderUri.Scheme -ne 'https' -or
            -not [string]::IsNullOrEmpty($FolderUri.UserInfo)) {
            throw 'FolderServerRelativeUrl must be a valid absolute HTTPS URL without credentials.'
        }

        if (-not [string]::Equals(
                $FolderUri.GetLeftPart([System.UriPartial]::Authority),
                $SiteUri.GetLeftPart([System.UriPartial]::Authority),
                [System.StringComparison]::OrdinalIgnoreCase
            )) {
            throw 'The FolderServerRelativeUrl URL must belong to the same site as SiteUrl.'
        }

        $FolderPath = $FolderUri.AbsolutePath
        if (-not [string]::IsNullOrEmpty($FolderUri.Query)) {
            $FolderIds = @(
                foreach ($QueryParameter in $FolderUri.Query.TrimStart('?').Split('&')) {
                    $QueryParts = $QueryParameter -split '=', 2
                    $QueryName = [uri]::UnescapeDataString(($QueryParts[0] -replace '\+', ' '))
                    if ($QueryName -eq 'id') {
                        if ($QueryParts.Count -eq 1) {
                            ''
                        }
                        else {
                            [uri]::UnescapeDataString(($QueryParts[1] -replace '\+', ' '))
                        }
                    }
                }
            )
            if ($FolderIds.Count -ne 1 -or [string]::IsNullOrWhiteSpace($FolderIds[0])) {
                throw 'The SharePoint navigation URL must contain exactly one valid id parameter.'
            }
            $FolderPath = $FolderIds[0]
        }

        $NormalizedFolderUrl = [uri]::UnescapeDataString($FolderPath).TrimEnd('/')
    }
    else {
        if ($FolderInput.Contains('?')) {
            throw 'A server-relative path must not contain a query string.'
        }
        $NormalizedFolderUrl = [uri]::UnescapeDataString($FolderInput).TrimEnd('/')
    }

    if (-not $NormalizedFolderUrl.StartsWith('/') -or
        $NormalizedFolderUrl.StartsWith('//') -or
        @($NormalizedFolderUrl.Split('/') | Where-Object { $_ -in @('.', '..') }).Count -gt 0) {
        throw 'FolderServerRelativeUrl must be a valid folder URL or server-relative path, without "." or ".." segments.'
    }

    $NormalizedSitePath = [uri]::UnescapeDataString($SiteUri.AbsolutePath).TrimEnd('/')
    if ($NormalizedSitePath -and
        -not $NormalizedFolderUrl.StartsWith(
            "$NormalizedSitePath/",
            [System.StringComparison]::OrdinalIgnoreCase
        )) {
        throw 'FolderServerRelativeUrl must belong to the site specified by SiteUrl.'
    }
}

$PnPModule = Get-Module -ListAvailable -Name PnP.PowerShell |
    Sort-Object Version -Descending |
    Select-Object -First 1
if (-not $PnPModule -or $PnPModule.Version -lt [version]'3.2.0') {
    throw 'PnP.PowerShell version 3.2.0 or later is required.'
}

New-Item -Path $OutputDirectory -ItemType Directory -Force -ErrorAction Stop -WhatIf:$false | Out-Null

$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$SiteName = @($SiteUri.Segments | Where-Object { $_ -ne '/' } | Select-Object -Last 1)
if ($SiteName.Count -eq 0) {
    $SiteName = @($SiteUri.Host)
}
$SafeSiteName = ([string]$SiteName[0]).TrimEnd('/') -replace '[\\/:*?"<>| ]', '_'
$FileReportPath = Join-Path $OutputDirectory "${SafeSiteName}_version_history_files_${Timestamp}.csv"
$SummaryReportPath = Join-Path $OutputDirectory "${SafeSiteName}_version_history_summary_${Timestamp}.csv"
$ActionReportPath = Join-Path $OutputDirectory "${SafeSiteName}_version_history_actions_${Timestamp}.csv"
$ErrorReportPath = Join-Path $OutputDirectory "${SafeSiteName}_version_history_errors_${Timestamp}.csv"

$FileRows = [System.Collections.Generic.List[object]]::new()
$LibraryRows = [System.Collections.Generic.List[object]]::new()
$AnalysisErrors = [System.Collections.Generic.List[object]]::new()
$ActionRows = [System.Collections.Generic.List[object]]::new()

Write-Host "Connecting interactively to site: $SiteUrl"
$Connection = Connect-PnPOnline `
    -Url $SiteUrl `
    -ClientId $ClientId `
    -Interactive `
    -ReturnConnection `
    -ErrorAction Stop

try {
    $Libraries = @(
        Get-PnPList `
            -Includes Id, Title, Hidden, BaseTemplate, ItemCount, RootFolder `
            -Connection $Connection `
            -ErrorAction Stop |
            Where-Object {
                $_.BaseTemplate -eq 101 -and
                ($IncludeHiddenLibraries -or -not $_.Hidden)
            }
    )

    if ($Libraries.Count -eq 0) {
        Write-Warning 'No document libraries were found for analysis.'
    }

    if ($NormalizedFolderUrl) {
        $MatchingLibraries = @(
            $Libraries |
                Where-Object {
                    $LibraryRoot = [uri]::UnescapeDataString(
                        [string]$_.RootFolder.ServerRelativeUrl
                    ).TrimEnd('/')
                    [string]::Equals(
                        $NormalizedFolderUrl,
                        $LibraryRoot,
                        [System.StringComparison]::OrdinalIgnoreCase
                    ) -or
                    $NormalizedFolderUrl.StartsWith(
                        "$LibraryRoot/",
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                } |
                Sort-Object {
                    [uri]::UnescapeDataString([string]$_.RootFolder.ServerRelativeUrl).Length
                } -Descending |
                Select-Object -First 1
        )
        if ($MatchingLibraries.Count -eq 0) {
            throw "The folder '$FolderServerRelativeUrl' does not belong to a document library included in the site."
        }
        $Libraries = $MatchingLibraries
        Get-PnPFolder `
            -Url $NormalizedFolderUrl `
            -Connection $Connection `
            -ErrorAction Stop | Out-Null
    }

    $LibraryIndex = 0
    foreach ($Library in $Libraries) {
        $LibraryIndex++
        [long]$LibraryFileCount = 0
        [long]$LibraryVersionCount = 0
        [long]$LibraryVersionBytes = 0
        [long]$ProcessedItems = 0
        $LibraryErrorCount = 0
        $FileBatch = [System.Collections.Generic.List[object]]::new()
        $ProcessFileBatch = {
            if ($FileBatch.Count -eq 0) {
                return
            }

            $BatchResults = @(
                Get-FileVersionBatchResults `
                    -FileEntries $FileBatch.ToArray() `
                    -Connection $Connection
            )
            foreach ($BatchResult in $BatchResults) {
                $Entry = $BatchResult.Entry
                $LibraryFileCount++
                if ($BatchResult.Status -eq 'Error') {
                    $LibraryErrorCount++
                    $AnalysisErrors.Add([pscustomobject]@{
                        SiteUrl = $SiteUrl
                        Library = $Library.Title
                        FileUrl = $Entry.FileUrl
                        Stage   = 'Version history'
                        Error   = $BatchResult.Details
                    })
                    $FileRows.Add([pscustomobject]@{
                        SiteUrl        = $SiteUrl
                        Library        = $Library.Title
                        FileUrl        = $Entry.FileUrl
                        VersionCount   = $null
                        VersionBytes   = $null
                        VersionMiB     = $null
                        AnalysisStatus = 'Error'
                        Details        = $BatchResult.Details
                    })
                    continue
                }

                $LibraryVersionCount += $BatchResult.VersionCount
                $LibraryVersionBytes += $BatchResult.VersionBytes
                $FileRows.Add([pscustomobject]@{
                    SiteUrl        = $SiteUrl
                    Library        = $Library.Title
                    FileUrl        = $Entry.FileUrl
                    VersionCount   = $BatchResult.VersionCount
                    VersionBytes   = $BatchResult.VersionBytes
                    VersionMiB     = [Math]::Round($BatchResult.VersionBytes / 1MB, 3)
                    AnalysisStatus = 'Analyzed'
                    Details        = ''
                })
            }
            $FileBatch.Clear()
        }

        Write-Host "Analyzing library $LibraryIndex/$($Libraries.Count): $($Library.Title)"
        try {
            if ($NormalizedFolderUrl) {
                $FolderSiteRelativeUrl = $NormalizedFolderUrl
                $NormalizedSitePath = [uri]::UnescapeDataString($SiteUri.AbsolutePath).TrimEnd('/')
                if ($NormalizedSitePath) {
                    $FolderSiteRelativeUrl = $FolderSiteRelativeUrl.Substring($NormalizedSitePath.Length).TrimStart('/')
                }

                if ($Library.ItemCount -gt 5000) {
                    Write-Host "The library has more than 5,000 items; enumerating its files and filtering for the requested folder."
                    $Items = @(
                        Get-PnPFolderItem `
                            -List $Library.Id `
                            -ItemType File `
                            -Connection $Connection `
                            -ErrorAction Stop |
                            Where-Object {
                                $ItemUrl = [uri]::UnescapeDataString([string]$_.ServerRelativeUrl)
                                $ItemUrl.StartsWith(
                                    "$NormalizedFolderUrl/",
                                    [System.StringComparison]::OrdinalIgnoreCase
                                )
                            }
                    )
                }
                else {
                    $Items = Get-PnPFolderItem `
                        -FolderSiteRelativeUrl $FolderSiteRelativeUrl `
                        -ItemType File `
                        -Recursive `
                        -Connection $Connection `
                        -ErrorAction Stop
                }
            }
            elseif ($Library.ItemCount -gt 5000) {
                $Items = Get-PnPFolderItem `
                    -List $Library.Id `
                    -ItemType File `
                    -Connection $Connection `
                    -ErrorAction Stop
            }
            else {
                $Items = Get-PnPListItem `
                    -List $Library.Id `
                    -PageSize 1000 `
                    -Fields @('FileRef', 'FSObjType') `
                    -Connection $Connection `
                    -ErrorAction Stop
            }

            foreach ($Item in $Items) {
                $ProcessedItems++
                if (-not $NormalizedFolderUrl -and $Library.ItemCount -gt 0 -and $Library.ItemCount -le 5000) {
                    $PercentComplete = [Math]::Min(
                        99,
                        [int](100 * $ProcessedItems / $Library.ItemCount)
                    )
                    Write-Progress `
                        -Activity 'Analyzing version history' `
                        -Status "$($Library.Title): $ProcessedItems/$($Library.ItemCount) items" `
                        -PercentComplete $PercentComplete
                }
                elseif ($NormalizedFolderUrl) {
                    Write-Progress `
                        -Activity 'Analyzing version history' `
                        -Status "$($Library.Title): $ProcessedItems items in folder"
                }

                if ($Item.PSObject.Properties['FileSystemObjectType'] -and $Item.FileSystemObjectType -ne 'File') {
                    continue
                }

                $FileUrl = if ($Item.PSObject.Properties['ServerRelativeUrl']) {
                    [string]$Item.ServerRelativeUrl
                }
                else {
                    [string]$Item['FileRef']
                }
                if ([string]::IsNullOrWhiteSpace($FileUrl)) {
                    $LibraryErrorCount++
                    $AnalysisErrors.Add([pscustomobject]@{
                        SiteUrl = $SiteUrl
                        Library = $Library.Title
                        FileUrl = ''
                        Stage   = 'Version history'
                        Error   = 'The item does not contain a file path.'
                    })
                    continue
                }

                $FileBatch.Add([pscustomobject]@{
                    Library = $Library.Title
                    FileUrl = $FileUrl
                })
                if ($FileBatch.Count -ge $BatchSize) {
                    . $ProcessFileBatch
                }
            }
            . $ProcessFileBatch
        }
        catch {
            $LibraryErrorCount++
            $ErrorDetail = $_.Exception.Message
            $AnalysisErrors.Add([pscustomobject]@{
                SiteUrl = $SiteUrl
                Library = $Library.Title
                FileUrl = ''
                Stage   = 'Library enumeration'
                Error   = $ErrorDetail
            })
            Write-Warning "Could not finish reading library '$($Library.Title)': $ErrorDetail"
        }
        finally {
            Write-Progress -Activity 'Analyzing version history' -Completed
        }

        $LibraryRows.Add([pscustomobject]@{
            SiteUrl           = $SiteUrl
            Library           = $Library.Title
            Files             = $LibraryFileCount
            PreviousVersions  = $LibraryVersionCount
            VersionHistoryBytes = $LibraryVersionBytes
            VersionHistoryMiB = [Math]::Round($LibraryVersionBytes / 1MB, 3)
            AnalysisErrors    = $LibraryErrorCount
        })
    }

    $TotalFiles = [long](($LibraryRows | Measure-Object -Property Files -Sum).Sum)
    $TotalVersions = [long](($LibraryRows | Measure-Object -Property PreviousVersions -Sum).Sum)
    $TotalBytes = [long](($LibraryRows | Measure-Object -Property VersionHistoryBytes -Sum).Sum)
    $AnalyzedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $SummaryRow = [pscustomobject]@{
        SiteUrl             = $SiteUrl
        AnalyzedAt          = $AnalyzedAt
        DocumentLibraries   = $Libraries.Count
        Files               = $TotalFiles
        PreviousVersions    = $TotalVersions
        VersionHistoryBytes = $TotalBytes
        VersionHistoryMiB   = [Math]::Round($TotalBytes / 1MB, 3)
        VersionHistoryGiB   = [Math]::Round($TotalBytes / 1GB, 6)
        AnalysisErrors      = $AnalysisErrors.Count
        IncludeHiddenLibraries = [bool]$IncludeHiddenLibraries
        Note                = 'Sum of the Length values for previous versions returned by SharePoint; this may not equal actual billable storage.'
    }

    ConvertTo-SafeCsvRows -Rows $FileRows |
        Export-Csv -LiteralPath $FileReportPath -NoTypeInformation -Encoding utf8BOM -Delimiter ';' -WhatIf:$false
    ConvertTo-SafeCsvRows -Rows $LibraryRows |
        Export-Csv -LiteralPath $SummaryReportPath -NoTypeInformation -Encoding utf8BOM -Delimiter ';' -WhatIf:$false
    ConvertTo-SafeCsvRows -Rows @($SummaryRow) |
        Export-Csv -LiteralPath ($SummaryReportPath -replace '_summary_', '_total_') -NoTypeInformation -Encoding utf8BOM -Delimiter ';' -WhatIf:$false

    $TotalReportPath = $SummaryReportPath -replace '_summary_', '_total_'
    Write-Host ''
    Write-Host "Files analyzed: $TotalFiles"
    Write-Host "Previous versions: $TotalVersions"
    Write-Host ("Total version history size: {0:F3} MiB ({1:F6} GiB)" -f $SummaryRow.VersionHistoryMiB, $SummaryRow.VersionHistoryGiB)
    Write-Host "Per-file report: $FileReportPath"
    Write-Host "Per-library report: $SummaryReportPath"
    Write-Host "Overall summary: $TotalReportPath"

    $CanContinueWithFileErrors = $DeleteAllVersions
    if ($AnalysisErrors.Count -gt 0) {
        ConvertTo-SafeCsvRows -Rows $AnalysisErrors |
            Export-Csv -LiteralPath $ErrorReportPath -NoTypeInformation -Encoding utf8BOM -Delimiter ';' -WhatIf:$false
        Write-Host "Error details: $ErrorReportPath"
        $LibraryEnumerationErrors = @($AnalysisErrors | Where-Object { $_.Stage -eq 'Library enumeration' })
        $VersionHistoryErrors = @($AnalysisErrors | Where-Object { $_.Stage -eq 'Version history' })
        $ErrorSummary = @(
            $AnalysisErrors |
                Select-Object -First 3 |
                ForEach-Object { "$($_.Library) [$($_.Stage)]: $($_.Error)" }
        ) -join ' | '

        if ($LibraryEnumerationErrors.Count -gt 0 -or
            -not $CanContinueWithFileErrors -or
            $VersionHistoryErrors.Count -ne $AnalysisErrors.Count) {
            throw "Analysis incomplete: $($AnalysisErrors.Count) errors. No versions were deleted. Details: $ErrorReportPath. $ErrorSummary"
        }

        Write-Warning "Continuing despite $($VersionHistoryErrors.Count) per-file errors; files with errors will be skipped. See report: $ErrorReportPath"
    }

    if ($DeleteAllVersions) {
        $FilesWithHistory = @($FileRows | Where-Object { $_.AnalysisStatus -eq 'Analyzed' -and $_.VersionCount -gt 0 })
        if ($FilesWithHistory.Count -eq 0) {
            Write-Host 'There are no previous versions to delete.'
        }
        else {
            $ActionScope = if ($NormalizedFolderUrl) {
                "in folder '$NormalizedFolderUrl' and its subfolders"
            }
            else {
                'in the site'
            }
            $Action = "Permanently delete all previous versions of $($FilesWithHistory.Count) files $ActionScope (current files will remain intact)"
            if ($PSCmdlet.ShouldProcess($SiteUrl, $Action)) {
                foreach ($File in $FilesWithHistory) {
                    try {
                        Remove-PnPFileVersion `
                            -Url $File.FileUrl `
                            -All `
                            -Force `
                            -Connection $Connection `
                            -ErrorAction Stop
                        $ActionRows.Add([pscustomobject]@{
                            SiteUrl      = $SiteUrl
                            Library      = $File.Library
                            FileUrl      = $File.FileUrl
                            VersionCount = $File.VersionCount
                            VersionBytes = $File.VersionBytes
                            Status       = 'Removed'
                            Details      = 'Removal command completed; current file was not deleted.'
                        })
                    }
                    catch {
                        $ActionRows.Add([pscustomobject]@{
                            SiteUrl      = $SiteUrl
                            Library      = $File.Library
                            FileUrl      = $File.FileUrl
                            VersionCount = $File.VersionCount
                            VersionBytes = $File.VersionBytes
                            Status       = 'Error'
                            Details      = $_.Exception.Message
                        })
                    }
                }
            }
            else {
                foreach ($File in $FilesWithHistory) {
                    $ActionRows.Add([pscustomobject]@{
                        SiteUrl      = $SiteUrl
                        Library      = $File.Library
                        FileUrl      = $File.FileUrl
                        VersionCount = $File.VersionCount
                        VersionBytes = $File.VersionBytes
                        Status       = 'WhatIf'
                        Details      = 'Removal simulated; no changes were made.'
                    })
                }
            }

            ConvertTo-SafeCsvRows -Rows $ActionRows |
                Export-Csv -LiteralPath $ActionReportPath -NoTypeInformation -Encoding utf8BOM -Delimiter ';' -WhatIf:$false
            Write-Host "Action report: $ActionReportPath"

            $ActionErrors = @($ActionRows | Where-Object { $_.Status -eq 'Error' }).Count
            if ($ActionErrors -gt 0) {
                throw "Removal completed with $ActionErrors errors. See the action report: $ActionReportPath"
            }
        }
    }
}
finally {
    Write-Progress -Activity 'Analyzing version history' -Completed
}
