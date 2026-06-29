# ============================================================
#  School Photo Organiser - Windows app (GUI) version
#  Same logic as the original script, but in a friendly window
#  with buttons and pop-up messages instead of a console.
#  Can be run as-is, or compiled to a .exe with Build-PhotoOrganiser-EXE.ps1
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Give the app its own taskbar identity. Without this, when the script is run via
# powershell.exe the taskbar button groups under PowerShell and shows PowerShell's
# icon. Setting an explicit AppUserModelID makes the taskbar use OUR window icon.
try {
    Add-Type -Namespace Win32 -Name TaskbarId -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("shell32.dll", SetLastError=true)]
public static extern void SetCurrentProcessExplicitAppUserModelID(string AppID);
'@ -ErrorAction Stop
    [Win32.TaskbarId]::SetCurrentProcessExplicitAppUserModelID("School.PhotoOrganiser")
} catch { }

# ----- Settings -----
# If the CSV is this many days old (or older), warn before processing.
$MaxCsvAgeDays   = 3
$validExtensions = @(".png", ".jpg", ".jpeg")

# ----- Work out which folder the app is sitting in -----
# Works both as a .ps1 and as a compiled .exe.
function Get-AppFolder {
    if ($PSScriptRoot) { return $PSScriptRoot }
    try {
        $exe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if ($exe -and (Test-Path $exe)) { return (Split-Path -Parent $exe) }
    } catch { }
    return (Get-Location).Path
}

# ----- Saved settings (remembered across launches) -----
# Stored in the user's AppData so it works even if the app folder is read-only
# or on a network share.
$settingsDir  = Join-Path $env:APPDATA "SchoolPhotoOrganiser"
$settingsPath = Join-Path $settingsDir "settings.json"

function Get-AppSettings {
    $defaults = [PSCustomObject]@{ SkipColumnWarning = $false }
    try {
        if (Test-Path $settingsPath) {
            $loaded = Get-Content -Path $settingsPath -Raw | ConvertFrom-Json
            if ($null -ne $loaded -and ($loaded.PSObject.Properties.Name -contains 'SkipColumnWarning')) {
                $defaults.SkipColumnWarning = [bool]$loaded.SkipColumnWarning
            }
        }
    } catch { }
    return $defaults
}

function Save-AppSettings {
    param($Settings)
    try {
        if (-not (Test-Path $settingsDir)) {
            New-Item -ItemType Directory -Path $settingsDir -Force | Out-Null
        }
        $Settings | ConvertTo-Json | Set-Content -Path $settingsPath
    } catch { }
}

$appSettings = Get-AppSettings

# ============================================================
#  Build the window
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text          = "School Photo Organiser"
$form.Size          = New-Object System.Drawing.Size(760, 560)
$form.StartPosition = "CenterScreen"
$form.MinimumSize   = New-Object System.Drawing.Size(700, 500)
$form.Font          = New-Object System.Drawing.Font("Segoe UI", 9)

$lblHeader = New-Object System.Windows.Forms.Label
$lblHeader.Text     = "School Photo Organiser"
$lblHeader.Font     = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblHeader.AutoSize = $true
$lblHeader.Location = New-Object System.Drawing.Point(20, 15)
$form.Controls.Add($lblHeader)

$lblIntro = New-Object System.Windows.Forms.Label
$lblIntro.Text     = "Sorts the photos in the folder below into Staff / Left, and lists any missing photos."
$lblIntro.AutoSize = $true
$lblIntro.Location = New-Object System.Drawing.Point(22, 52)
$form.Controls.Add($lblIntro)

$lblFolderCap = New-Object System.Windows.Forms.Label
$lblFolderCap.Text     = "Photos folder:"
$lblFolderCap.AutoSize = $true
$lblFolderCap.Location = New-Object System.Drawing.Point(20, 84)
$form.Controls.Add($lblFolderCap)

$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Location = New-Object System.Drawing.Point(20, 106)
$txtFolder.Size     = New-Object System.Drawing.Size(600, 24)
$txtFolder.Text     = (Get-AppFolder)
$txtFolder.ReadOnly = $true
$txtFolder.Anchor   = "Top, Left, Right"
$form.Controls.Add($txtFolder)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text     = "Change..."
$btnBrowse.Location = New-Object System.Drawing.Point(630, 105)
$btnBrowse.Size     = New-Object System.Drawing.Size(95, 26)
$btnBrowse.Anchor   = "Top, Right"
$form.Controls.Add($btnBrowse)

$rtb = New-Object System.Windows.Forms.RichTextBox
$rtb.Location  = New-Object System.Drawing.Point(20, 146)
$rtb.Size      = New-Object System.Drawing.Size(705, 282)
$rtb.ReadOnly  = $true
$rtb.BackColor = [System.Drawing.Color]::White
$rtb.Font      = New-Object System.Drawing.Font("Consolas", 9)
$rtb.Anchor    = "Top, Bottom, Left, Right"
$form.Controls.Add($rtb)

# Live status line shown in the body of the window during processing.
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text      = ""
$lblStatus.Location  = New-Object System.Drawing.Point(20, 434)
$lblStatus.Size      = New-Object System.Drawing.Size(705, 20)
$lblStatus.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 95, 175)
$lblStatus.Anchor    = "Bottom, Left, Right"
$form.Controls.Add($lblStatus)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text     = "Organise Photos"
$btnRun.Location = New-Object System.Drawing.Point(20, 458)
$btnRun.Size     = New-Object System.Drawing.Size(160, 38)
$btnRun.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnRun.Anchor   = "Bottom, Left"
$form.Controls.Add($btnRun)

$btnOpenReport = New-Object System.Windows.Forms.Button
$btnOpenReport.Text     = "Open Missing Photos List"
$btnOpenReport.Location = New-Object System.Drawing.Point(190, 458)
$btnOpenReport.Size     = New-Object System.Drawing.Size(210, 38)
$btnOpenReport.Anchor   = "Bottom, Left"
$form.Controls.Add($btnOpenReport)

# Persistent checkbox: skip the "missing columns" warning. Remembered across launches.
$chkSkipWarning = New-Object System.Windows.Forms.CheckBox
$chkSkipWarning.Text     = "Don't warn me about optional CSV columns"
$chkSkipWarning.AutoSize = $true
$chkSkipWarning.Location = New-Object System.Drawing.Point(410, 465)
$chkSkipWarning.Anchor   = "Bottom, Left"
$chkSkipWarning.Checked  = [bool]$appSettings.SkipColumnWarning
$chkSkipWarning.Add_CheckedChanged({
    $appSettings.SkipColumnWarning = $chkSkipWarning.Checked
    Save-AppSettings $appSettings
})
$form.Controls.Add($chkSkipWarning)

# ----- Helper to write coloured lines into the results panel -----
function Write-Log {
    param([string]$Message, [System.Drawing.Color]$Color = [System.Drawing.Color]::Black)
    $rtb.SelectionStart = $rtb.TextLength
    $rtb.SelectionColor = $Color
    $rtb.AppendText($Message + "`r`n")
    $rtb.SelectionColor = $rtb.ForeColor
    $rtb.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# ----- Keeps the window responsive during long file loops and shows progress -----
$baseTitle = "School Photo Organiser"
function Update-Progress {
    param([int]$Done, [int]$Total, [string]$Action)
    if ($Total -gt 0) {
        $form.Text       = "$baseTitle  -  $Action $Done of $Total..."
        $lblStatus.Text  = "$Action $Done of $Total..."
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# ============================================================
#  The main job - same logic as the original script
# ============================================================
function Invoke-OrganisePhotos {
    param([string]$PhotosFolder, [bool]$SkipColumnWarning)

    $rtb.Clear()
    $lblStatus.Text = ""

    # --- Folder check ---
    if (-not (Test-Path $PhotosFolder -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not find the photos folder:`n`n$PhotosFolder`n`nPlease use 'Change...' to pick the folder that contains the photos.",
            "Photos folder not found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    # --- CSV exists? ---
    $csvPath = Join-Path $PhotosFolder "GridExport.CSV"
    if (-not (Test-Path $csvPath -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show(
            "GridExport.CSV was not found in:`n`n$PhotosFolder`n`nPlease export the student list from People Management in Compass and save it as 'GridExport.CSV' in this folder.",
            "GridExport.CSV not found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    # --- CSV age check (ignores bogus dates like 1/01/1970) ---
    $csvFile      = Get-Item -Path $csvPath
    $createdDate  = $csvFile.CreationTime
    $modifiedDate = $csvFile.LastWriteTime
    $minValidDate = Get-Date "2000-01-01"

    $validDates = @()
    if ($createdDate  -ge $minValidDate) { $validDates += $createdDate }
    if ($modifiedDate -ge $minValidDate) { $validDates += $modifiedDate }

    if ($validDates.Count -gt 0) {
        $oldestDate = ($validDates | Sort-Object)[0]
        $csvAgeDays = [math]::Floor(((Get-Date) - $oldestDate).TotalDays)

        if ($csvAgeDays -ge $MaxCsvAgeDays) {
            $createdStr  = if ($createdDate  -ge $minValidDate) { $createdDate.ToString('dd/MM/yyyy hh:mm tt') }  else { "(not available)" }
            $modifiedStr = if ($modifiedDate -ge $minValidDate) { $modifiedDate.ToString('dd/MM/yyyy hh:mm tt') } else { "(not available)" }

            $msg  = "This CSV appears to be about $csvAgeDays days old.`n`n"
            $msg += "Created:   $createdStr`n"
            $msg += "Modified:  $modifiedStr`n`n"

            $modifiedAgeDays = if ($modifiedDate -ge $minValidDate) { [math]::Floor(((Get-Date) - $modifiedDate).TotalDays) } else { $csvAgeDays }
            if ($modifiedAgeDays -lt $MaxCsvAgeDays) {
                $msg += "(If you've just replaced the file, sometimes Windows doesn't update the save date. That's okay.)`n`n"
            }

            $msg += "It looks like you need a new one. A fresh GridExport.CSV can be generated from People Management in Compass.`n`n"
            $msg += "Continue anyway using this old CSV?"

            $answer = [System.Windows.Forms.MessageBox]::Show(
                $msg, "This CSV looks out of date",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning,
                [System.Windows.Forms.MessageBoxDefaultButton]::Button2)

            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
                Write-Log "Cancelled - please export a new GridExport.CSV from People Management in Compass." ([System.Drawing.Color]::DarkOrange)
                return
            }
        }
    } else {
        Write-Log "Could not read a valid date from the CSV - skipping the age check." ([System.Drawing.Color]::Gray)
    }

    # --- Read CSV ---
    Write-Log "Reading GridExport.CSV..." ([System.Drawing.Color]::DarkCyan)
    $csvData = Import-Csv -Path $csvPath | Where-Object { $_."Import Identifier" -ne $null -and $_."Import Identifier" -ne "" }

    if (-not ($csvData | Get-Member -Name "Import Identifier" -ErrorAction SilentlyContinue)) {
        [System.Windows.Forms.MessageBox]::Show(
            "The CSV file does not contain a column called 'Import Identifier'.",
            "Missing required column",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    # --- Work out which columns this export actually contains ---
    $csvColumns = @()
    $firstRow = $csvData | Select-Object -First 1
    if ($firstRow) { $csvColumns = $firstRow.PSObject.Properties.Name }

    # Column name variations we know how to read (export headers can differ).
    $firstNameCols = @("FirstName", "First Name", "Preferred Name", "Legal First Name", "Given Name")
    $lastNameCols  = @("LastName", "Last Name", "Surname", "Family Name", "Legal Last Name")
    $fullNameCols  = @("Full Name", "FullName", "Name", "Student Name", "Legal Name")
    $yearCols      = @("Year Level (Current)", "Year Level", "Year")

    function Get-FirstValue {
        param($Row, [string[]]$Candidates, [string[]]$Columns)
        foreach ($c in $Candidates) {
            if ($Columns -contains $c) {
                $v = "$($Row.$c)".Trim()
                if ($v -ne "") { return $v }
            }
        }
        return ""
    }

    $hasFirst = @($firstNameCols | Where-Object { $csvColumns -contains $_ }).Count -gt 0
    $hasLast  = @($lastNameCols  | Where-Object { $csvColumns -contains $_ }).Count -gt 0
    $hasFull  = @($fullNameCols  | Where-Object { $csvColumns -contains $_ }).Count -gt 0
    $hasYear  = @($yearCols      | Where-Object { $csvColumns -contains $_ }).Count -gt 0

    # Warn (but let them continue) if the columns that make the report detailed are absent.
    $missingBits = @()
    if (-not $hasFirst) { $missingBits += "First name" }
    if (-not $hasLast)  { $missingBits += "Last name" }
    if (-not $hasYear)  { $missingBits += "Year Level (Current)" }

    if ($missingBits.Count -gt 0 -and -not $SkipColumnWarning) {
        $warn  = "The CSV is missing these columns:`n`n  - " + ($missingBits -join "`n  - ") + "`n`n"
        if ($hasFull) {
            $warn += "A 'Full Name' column was found, so student names will still be shown.`n`n"
        } else {
            $warn += "No name column was found, so students may be listed by ID only.`n`n"
        }
        $warn += "The Missing Photos report may be less detailed - for example, students may not be grouped by year level.`n`n"
        $warn += "You can still continue. For a complete report, export GridExport.CSV from People Management in Compass with First Name, Last Name and Year Level columns included.`n`n"
        $warn += "Continue anyway?"

        $ans = [System.Windows.Forms.MessageBox]::Show(
            $warn, "Some columns are missing",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button1)

        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log "Cancelled - export a CSV that includes name and year level columns for a full report." ([System.Drawing.Color]::DarkOrange)
            return
        }
    }

    $activeIds   = [System.Collections.Generic.HashSet[string]]::new()
    $hasGovtCode = ($csvData | Get-Member -Name "Govt Code 2" -ErrorAction SilentlyContinue) -ne $null

    foreach ($row in $csvData) {
        if ($row.Status -eq "Active") {
            $id = $row."Import Identifier".Trim()
            if ($id -ne "") { [void]$activeIds.Add($id) }
            if ($hasGovtCode) {
                $govtValue = $row."Govt Code 2".Trim()
                if ($govtValue -ne "") { [void]$activeIds.Add($govtValue) }
            }
        }
    }
    Write-Log "Loaded $($activeIds.Count) active student IDs from CSV." ([System.Drawing.Color]::Green)

    # --- Folder setup ---
    $leftFolder  = Join-Path $PhotosFolder "Left"
    $staffFolder = Join-Path $PhotosFolder "Staff"
    New-Item -ItemType Directory -Path $leftFolder  -Force | Out-Null
    New-Item -ItemType Directory -Path $staffFolder -Force | Out-Null

    # --- PHASE 1: Restore & rescue ---
    Write-Log "Checking 'Left' folder for returning students and misfiled staff..." ([System.Drawing.Color]::DarkCyan)
    $restoredCount = 0
    if (Test-Path $leftFolder) {
        $leftFiles = Get-ChildItem -Path $leftFolder -File | Where-Object { $validExtensions -contains $_.Extension.ToLower() }
        $leftTotal = @($leftFiles).Count
        $leftDone  = 0
        foreach ($file in $leftFiles) {
            $name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            $isStaff = ($name -match "^[Ee]\d{6,7}$") -or ($name -match "^\d{6,7}$") -or ($name -notmatch "^\d+$")
            if ($activeIds.Contains($name) -or $isStaff) {
                Move-Item -Path $file.FullName -Destination (Join-Path $PhotosFolder $file.Name) -Force
                $restoredCount++
            }
            $leftDone++
            if (($leftDone % 20) -eq 0) { Update-Progress $leftDone $leftTotal "Checking 'Left'" }
        }
    }

    # --- PHASE 2: Process photos ---
    Write-Log "Processing photos..." ([System.Drawing.Color]::DarkCyan)
    $kept = 0; $movedLeft = 0; $movedStaff = 0
    $foundPhotoIds = [System.Collections.Generic.HashSet[string]]::new()
    $files = Get-ChildItem -Path $PhotosFolder -File | Where-Object { $validExtensions -contains $_.Extension.ToLower() }
    $total = @($files).Count
    $done  = 0

    foreach ($file in $files) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

        if ($name -match "^[Ee]\d{6,7}$" -or ($name -match "^\d{6,7}$") -or ($name -notmatch "^\d+$")) {
            Move-Item -Path $file.FullName -Destination (Join-Path $staffFolder $file.Name) -Force
            $movedStaff++
            $done++
            if (($done % 20) -eq 0) { Update-Progress $done $total "Processing" }
            continue
        }

        if ($activeIds.Contains($name)) {
            $kept++
            [void]$foundPhotoIds.Add($name)
        } else {
            Move-Item -Path $file.FullName -Destination (Join-Path $leftFolder $file.Name) -Force
            $movedLeft++
        }
        $done++
        if (($done % 20) -eq 0) { Update-Progress $done $total "Processing" }
    }
    Update-Progress $total $total "Processing"

    # --- PHASE 3: Missing photo report ---
    Write-Log "Generating Missing Photos Report..." ([System.Drawing.Color]::DarkCyan)

    # (Column lists and Get-FirstValue were set up after the CSV was read.)
    $missingStudents = @()
    $rowTotal = @($csvData).Count
    $rowDone  = 0
    foreach ($row in $csvData) {
        if ($row.Status -eq "Active") {
            $id     = $row."Import Identifier".Trim()
            $govtId = if ($hasGovtCode) { $row."Govt Code 2".Trim() } else { "" }
            $hasPhoto = $foundPhotoIds.Contains($id) -or ($govtId -ne "" -and $foundPhotoIds.Contains($govtId))
            if (-not $hasPhoto) {
                # Prefer First + Last name; fall back to a Full Name column if present.
                $first = Get-FirstValue $row $firstNameCols $csvColumns
                $last  = Get-FirstValue $row $lastNameCols  $csvColumns
                $fullName = ("$first $last").Trim()
                if ($fullName -eq "") { $fullName = Get-FirstValue $row $fullNameCols $csvColumns }
                if ($fullName -eq "") { $fullName = "(name not found in CSV)" }

                $year = Get-FirstValue $row $yearCols $csvColumns
                if ($year -eq "") { $year = "Unknown Year" }

                $missingStudents += [PSCustomObject]@{
                    Name = $fullName
                    ID   = $id
                    Year = $year
                }
            }
        }
        $rowDone++
        if (($rowDone % 200) -eq 0) { Update-Progress $rowDone $rowTotal "Checking roster" }
    }
    $form.Text = $baseTitle
    $lblStatus.Text = ""

    if ($missingStudents.Count -gt 0) {
        $reportPath  = Join-Path $PhotosFolder "Missing-Photos.txt"
        $dateTimeStr = Get-Date -Format "dd/MM/yyyy hh:mm tt"
        $reportContent = [System.Collections.Generic.List[string]]::new()
        $reportContent.Add("Students Missing Photo - $dateTimeStr")
        $reportContent.Add("")

        # Sort year groups numerically; non-numeric groups (e.g. "Unknown Year") go last.
        $grouped = $missingStudents | Group-Object Year | Sort-Object {
            $digits = ($_.Name -replace '\D', '')
            if ($digits -eq "") { 9999 } else { [int]$digits }
        }
        foreach ($group in $grouped) {
            $reportContent.Add("$($group.Name)")
            foreach ($student in ($group.Group | Sort-Object Name)) {
                $reportContent.Add("$($student.Name) - $($student.ID)")
            }
            $reportContent.Add("")
        }
        Set-Content -Path $reportPath -Value $reportContent
        Write-Log "Report saved: Missing-Photos.txt" ([System.Drawing.Color]::Green)
    }

    # --- Results ---
    Write-Log ""
    Write-Log "--- Results ---" ([System.Drawing.Color]::DarkGoldenrod)
    if ($restoredCount -gt 0) { Write-Log "Photos moved from 'Left' for returning students: $restoredCount" }
    Write-Log "Student photos kept (matched):   $kept"
    Write-Log "Student photos moved to 'Left':  $movedLeft"
    Write-Log "Staff photos moved to 'Staff':   $movedStaff"
    if ($missingStudents.Count -gt 0) {
        Write-Log "Active students missing photos:  $($missingStudents.Count)" ([System.Drawing.Color]::Magenta)
    }

    $summary  = "All done!`n`n"
    $summary += "Student photos kept (matched): $kept`n"
    $summary += "Moved to 'Left': $movedLeft`n"
    $summary += "Moved to 'Staff': $movedStaff"
    if ($restoredCount -gt 0)         { $summary += "`nMoved from 'Left' for returning students: $restoredCount" }
    if ($missingStudents.Count -gt 0) { $summary += "`n`nMissing photos: $($missingStudents.Count) (see Missing-Photos.txt)" }

    [System.Windows.Forms.MessageBox]::Show(
        $summary, "Finished",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

# ============================================================
#  Button actions
# ============================================================
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Choose the folder that contains the photos"
    if (Test-Path $txtFolder.Text) { $dlg.SelectedPath = $txtFolder.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFolder.Text = $dlg.SelectedPath
    }
})

$btnRun.Add_Click({
    $btnRun.Enabled        = $false
    $btnBrowse.Enabled     = $false
    $btnOpenReport.Enabled = $false
    try {
        Invoke-OrganisePhotos -PhotosFolder $txtFolder.Text -SkipColumnWarning $chkSkipWarning.Checked
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Something went wrong:`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } finally {
        $btnRun.Enabled        = $true
        $btnBrowse.Enabled     = $true
        $btnOpenReport.Enabled = $true
        $form.Text             = $baseTitle
        $lblStatus.Text        = ""
    }
})

$btnOpenReport.Add_Click({
    $reportPath = Join-Path $txtFolder.Text "Missing-Photos.txt"
    if (Test-Path $reportPath) {
        try {
            Start-Process -FilePath $reportPath
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Could not open the file:`n`n$reportPath`n`n$($_.Exception.Message)",
                "Couldn't open report",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "There's no Missing-Photos.txt in this folder yet.`n`nRun the organiser first - the list is only created when there are active students without a photo.",
            "No report found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
})

# ============================================================
#  Window / taskbar icon
#  1. Prefer an icon.ico sitting next to the app (lets you swap the
#     logo later without editing this script).
#  2. Otherwise load the icon embedded in this script below. This always
#     works - including for the .exe and from network (UNC) paths, where
#     extracting the icon from the file can silently fail.
# ============================================================
$embeddedIconB64 = @'
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
'@

try {
    $icoPath = Join-Path (Get-AppFolder) "icon.ico"
    if (Test-Path $icoPath) {
        $form.Icon = New-Object System.Drawing.Icon($icoPath)
    } elseif ($embeddedIconB64.Trim() -ne "") {
        $iconBytes  = [Convert]::FromBase64String(($embeddedIconB64 -replace '\s', ''))
        $iconStream = New-Object System.IO.MemoryStream(,$iconBytes)
        $form.Icon  = New-Object System.Drawing.Icon($iconStream)
    }
} catch { }

# ============================================================
#  Show the window
# ============================================================
[void]$form.ShowDialog()
