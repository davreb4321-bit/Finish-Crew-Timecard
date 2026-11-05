#Requires -Version 7.2
<#
.SYNOPSIS
    RQI Courier: copies finished rough walks from SharePoint (WalkDrop) into the job folders on P:.

.DESCRIPTION
    Runs every 10 minutes on the 24-hour server (Task Scheduler). Each run:
      1. Gets RQI WALKS rows with Status "Ready to File" or "Folder Not Found".
      2. Finds the job folder from the job number:
           2601001 -> <JobRoot>\2026 Jobs\01 JANUARY\N2601001_*
      3. Copies N<job>-RWC.pdf into the job folder and the photos into "Rough Walk Photos".
         Existing files are never overwritten: a re-walk gets N<job>-RWC-2.pdf and the next
         free photo numbers.
      4. Sets Status to Filed (or Folder Not Found / Error) on the RQI WALKS row.
    Nothing is ever deleted, on P: or in SharePoint.

.PARAMETER ConfigPath
    Path to courier-config.json. Defaults to the file next to this script.

.PARAMETER WhatIf
    Shows what would be copied and updated, without copying or updating anything.

.EXAMPLE
    pwsh -File ".\Copy-RoughWalksToP.ps1" -WhatIf
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'courier-config.json'),
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'
Import-Module PnP.PowerShell

# ---------------------------------------------------------------- settings
$cfg = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json

$JobRoot         = $cfg.JobRoot                      # UNC path, e.g. \\fileserver\Public\Indianapolis\JOB FOLDERS
$PhotoFolderName = $cfg.PhotoFolderName              # e.g. Rough Walk Photos
$ListName        = $cfg.ListName                     # RQI WALKS
$LogFolder       = $cfg.LogFolder
$StagingRoot     = Join-Path $LogFolder 'staging'

$MonthNames = @('JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE',
                'JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER')

New-Item -ItemType Directory -Force -Path $LogFolder, $StagingRoot | Out-Null
$LogFile = Join-Path $LogFolder ("courier-{0:yyyy-MM-dd}.log" -f (Get-Date))

function Write-Log([string]$Message) {
    $line = "{0:yyyy-MM-dd HH:mm:ss}  {1}{2}" -f (Get-Date), $(if ($WhatIf) { '[WhatIf] ' } else { '' }), $Message
    Add-Content -LiteralPath $LogFile -Value $line
    Write-Host $line
}

# ---------------------------------------------------------------- helpers

# 2601001 -> @{ Path = <job folder> } or @{ Problem = <why not> }
function Find-JobFolder([string]$JobNumber) {
    if ($JobNumber -notmatch '^\d{7}$') {
        return @{ Problem = "Job number '$JobNumber' is not 7 digits." }
    }
    $month = [int]$JobNumber.Substring(2, 2)
    if ($month -lt 1 -or $month -gt 12) {
        return @{ Problem = "Job number '$JobNumber' has month $month." }
    }
    $yearFolder  = Join-Path $JobRoot ("20{0} Jobs" -f $JobNumber.Substring(0, 2))
    $monthFolder = Join-Path $yearFolder ("{0:D2} {1}" -f $month, $MonthNames[$month - 1])
    if (-not (Test-Path -LiteralPath $monthFolder -PathType Container)) {
        return @{ Problem = "Month folder not found: $monthFolder" }
    }

    $prefix = "N$JobNumber" + '_'
    $hits = @(Get-ChildItem -LiteralPath $monthFolder -Directory |
              Where-Object { $_.Name.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) })

    if ($hits.Count -eq 0) {
        return @{ Problem = "No folder starting with '$prefix' in $monthFolder" }
    }
    if ($hits.Count -gt 1) {
        return @{ Problem = "More than one folder starts with '$prefix' in ${monthFolder}: " + (($hits | ForEach-Object Name) -join ' | ') }
    }
    return @{ Path = $hits[0].FullName }
}

# N2601001-RWC.pdf -> N2601001-RWC.pdf, or N2601001-RWC-2.pdf, -3 ... if taken
function Get-FreePdfPath([string]$Folder, [string]$FileName) {
    $base = [IO.Path]::GetFileNameWithoutExtension($FileName)
    $ext  = [IO.Path]::GetExtension($FileName)
    $candidate = Join-Path $Folder $FileName
    $n = 2
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $Folder ("{0}-{1}{2}" -f $base, $n, $ext)
        $n++
    }
    return $candidate
}

# Highest NN already used by N<job>-RW-NN.* in the photo folder (0 if none)
function Get-LastPhotoNumber([string]$PhotoFolder, [string]$JobNumber) {
    if (-not (Test-Path -LiteralPath $PhotoFolder)) { return 0 }
    $max = 0
    $pattern = '^N' + $JobNumber + '-RW-(\d+)\.'
    foreach ($f in Get-ChildItem -LiteralPath $PhotoFolder -File) {
        if ($f.Name -match $pattern) {
            $num = [int]$Matches[1]
            if ($num -gt $max) { $max = $num }
        }
    }
    return $max
}

# Files directly in a WalkDrop folder (site-relative, e.g. WalkDrop/2601001-57). Empty if none.
function Get-DropFiles([string]$SiteRelativeFolder) {
    try {
        return @(Get-PnPFolderItem -FolderSiteRelativeUrl $SiteRelativeFolder -ItemType File |
                 Sort-Object Name)
    } catch {
        return @()
    }
}

# Download one SharePoint file to staging, copy it to $Target, check the size matches.
function Copy-DropFile($SpFile, [string]$StagingFolder, [string]$Target) {
    Get-PnPFile -Url $SpFile.ServerRelativeUrl -Path $StagingFolder -FileName $SpFile.Name -AsFile -Force | Out-Null
    $staged = Join-Path $StagingFolder $SpFile.Name
    $stagedSize = (Get-Item -LiteralPath $staged).Length

    $expected = $null
    try { $expected = [int64]$SpFile.Length } catch { }
    if ($expected -and $expected -ne $stagedSize) {
        throw "Download of $($SpFile.Name) is $stagedSize bytes, SharePoint says $expected."
    }

    if ($WhatIf) {
        Write-Log "  would copy $($SpFile.Name) -> $Target"
        return
    }
    Copy-Item -LiteralPath $staged -Destination $Target
    $copiedSize = (Get-Item -LiteralPath $Target).Length
    if ($copiedSize -ne $stagedSize) {
        throw "Copy of $($SpFile.Name) to $Target is $copiedSize bytes, expected $stagedSize."
    }
    Write-Log "  copied $($SpFile.Name) -> $Target"
}

function Set-WalkStatus($Item, [hashtable]$Values) {
    if ($WhatIf) {
        Write-Log ("  would set " + (($Values.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '))
        return
    }
    Set-PnPListItem -List $ListName -Identity $Item.Id -Values $Values | Out-Null
}

# ---------------------------------------------------------------- main
Connect-PnPOnline -Url $cfg.SiteUrl -ClientId $cfg.ClientId -Tenant $cfg.Tenant -Thumbprint $cfg.Thumbprint

$caml = @"
<View><Query><Where><Or>
  <Eq><FieldRef Name='WalkStatus'/><Value Type='Text'>Ready to File</Value></Eq>
  <Eq><FieldRef Name='WalkStatus'/><Value Type='Text'>Folder Not Found</Value></Eq>
</Or></Where><OrderBy><FieldRef Name='ID'/></OrderBy></Query></View>
"@

$items = @(Get-PnPListItem -List $ListName -Query $caml -PageSize 500)
if ($items.Count -gt 0) { Write-Log "$($items.Count) walk(s) to file." }

foreach ($item in $items) {
    $id         = $item.Id
    $jobNumber  = ([string]$item['Title']).Trim().TrimStart('N', 'n')
    $dropFolder = ([string]$item['DropFolder']).Trim('/')
    $oldStatus  = [string]$item['WalkStatus']
    $oldProblem = [string]$item['FilingError']
    Write-Log "Walk $id, job N$jobNumber, drop folder $dropFolder (status $oldStatus)"

    try {
        if ([string]::IsNullOrWhiteSpace($dropFolder)) { throw "Drop Folder is blank." }

        $found = Find-JobFolder $jobNumber
        if ($found.Problem) {
            Write-Log "  $($found.Problem)"
            # Only update (and so only alert) when something changed
            if ($oldStatus -ne 'Folder Not Found' -or $oldProblem -ne $found.Problem) {
                Set-WalkStatus $item @{ WalkStatus = 'Folder Not Found'; FilingError = $found.Problem }
            }
            continue
        }
        $jobFolder = $found.Path

        $pdfs   = @(Get-DropFiles $dropFolder | Where-Object { $_.Name -like '*.pdf' })
        $photos = @(Get-DropFiles "$dropFolder/$PhotoFolderName")
        if ($pdfs.Count -eq 0) { throw "No PDF in $dropFolder yet." }

        $staging = Join-Path $StagingRoot "$jobNumber-$id"
        New-Item -ItemType Directory -Force -Path $staging | Out-Null

        # PDF(s) into the job folder root, never overwriting
        $pdfNames = @()
        foreach ($pdf in $pdfs) {
            $target = Get-FreePdfPath $jobFolder $pdf.Name
            Copy-DropFile $pdf $staging $target
            $pdfNames += (Split-Path $target -Leaf)
        }

        # Photos into the photo subfolder, numbered after any that are already there
        if ($photos.Count -gt 0) {
            $photoFolder = Join-Path $jobFolder $PhotoFolderName
            if (-not $WhatIf) { New-Item -ItemType Directory -Force -Path $photoFolder | Out-Null }
            $next = (Get-LastPhotoNumber $photoFolder $jobNumber) + 1
            foreach ($photo in $photos) {
                $ext = [IO.Path]::GetExtension($photo.Name)
                if ([string]::IsNullOrEmpty($ext)) { $ext = '.jpg' }
                $target = Join-Path $photoFolder ("N{0}-RW-{1:D2}{2}" -f $jobNumber, $next, $ext)
                Copy-DropFile $photo $staging $target
                $next++
            }
        }

        Set-WalkStatus $item @{
            WalkStatus  = 'Filed'
            FiledOn     = (Get-Date)
            FiledPath   = $jobFolder
            PdfName     = ($pdfNames -join ', ')
            FilingError = ''
        }
        Write-Log "  filed: $($pdfNames -join ', ') and $($photos.Count) photo(s)."

        Remove-Item -LiteralPath $staging -Recurse -Force
    }
    catch {
        $msg = $_.Exception.Message
        Write-Log "  ERROR: $msg"
        try { Set-WalkStatus $item @{ WalkStatus = 'Error'; FilingError = "Courier: $msg" } } catch { Write-Log "  could not update the row: $($_.Exception.Message)" }
    }
}

# Keep 60 days of logs
Get-ChildItem -LiteralPath $LogFolder -Filter 'courier-*.log' -File |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-60) } |
    Remove-Item -Force

Disconnect-PnPOnline
