<#
  .SYNOPSIS
  Utilizes GAM7 to automate Google Calendar Subscriptions for Google Group members.
  .DESCRIPTION
  Reads group/calendar pairs from a config.json file and uses GAM7 to subscribe all user members (including nested groups) to the calendars.
  .PARAMETER Config
  Launches an interactive mode to view/add Groups and Calendars to config.json
  .PARAMETER ConfigPath
  Path to the config.json file. Defaults to config.json in the script directory.
  .PARAMETER StateDir
  Directory for state_group_domain_tld.json files. Defaults to \state in the script directory.
  .PARAMETER MaxRetries
  Set an integer for GAM command retries.
  .PARAMETER AppTitle
  Used in app menu and as the Windows Event Log source name. Defaults to Deploy-CalendarSubscriptions

  .NOTES
  Requires a config.json file and requires GAM7.

#>
[CmdletBinding()]
param (
    [Parameter()]
    [switch]$Config,
    [string]$ConfigPath   = (Join-Path $PSScriptRoot "config.json"),
    [string]$StateDir     = (Join-Path $PSScriptRoot "state"),
    [string]$AppTitle     = "Deploy-CalendarSubscriptions",
    [int]$maxRetries      = 2 # Zero indexed, so total attempt is $MaxRetries + 1
)

# --- Helper function for Windows Event Logging ---
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Information", "Warning", "Error")]
        [string]$EntryType = "Information"
    )

    if (-not [System.Diagnostics.EventLog]::SourceExists($AppTitle)) {
        try {
            New-EventLog -LogName Application -Source $AppTitle -ErrorAction Stop
        } catch {
            Write-Output "[$EntryType] $Message (Source '$AppTitle' missing)"
            return
        }
    }
    Write-EventLog -LogName Application -Source $AppTitle -EntryType $EntryType -EventId 1001 -Message $Message
}

# --- INTERACTIVE CONFIG MANAGER ---
# --- Load Config File ---
function Read-Config {
  param([string]$ConfigPath)
  if (Test-Path $ConfigPath) {
    $data      = Get-Content $ConfigPath | ConvertFrom-Json
    $groups    = @($data.Groups)
    $calendars = @($data.Calendars)
    $DeployDays  = if ($null -ne $data.DeployDays) { $data.DeployDays } else { 7 }
  } else { # No config? No problem - we create an object
    $groups    = @()
    $calendars = @()
    $DeployDays  = 7
  }
  return @{ Groups = $groups; Calendars = $calendars; DeployDays = $DeployDays }
}

# --- Write the object to the config file ---
function Save-Config {
  param([string]$ConfigPath, [array]$Groups, [array]$Calendars, [int]$DeployDays = 7)
  $out = [PSCustomObject]@{ Groups = @($Groups); Calendars = @($Calendars); DeployDays = $DeployDays }
  $out | ConvertTo-Json -Depth 10 | Out-File $ConfigPath
  Write-Host "`nSaved: $ConfigPath" -ForegroundColor Green
}

# --- STATE MANAGEMENT ---
function Read-State {
  param([string]$StateDir, [string]$GroupEmail)
    $stateFile = Join-Path $StateDir "state-$($GroupEmail -replace '[^a-zA-Z0-9]', '-').json"
    if (Test-Path $stateFile) {
        return Get-Content $stateFile | ConvertFrom-Json
    }
    # No state file yet - return empty object
    return [PSCustomObject]@{}
}

function Save-State {
  param([string]$StateDir, [string]$GroupEmail, [PSCustomObject]$State)
    if (-not (Test-Path $StateDir)) {
        New-Item -ItemType Directory -Path $StateDir -Force | Out-Null
    }
    $stateFile = Join-Path $StateDir "state-$($GroupEmail -replace '[^a-zA-Z0-9]', '-').json"
    $State | ConvertTo-Json -Depth 10 | Out-File $stateFile
}

function Get-UsersNeedingSub {
    # Returns users from $Members who either have no state entry for this calendar,
    # or whose last deployment is older than $DeployDays.
    param(
        [array]$Members,
        [PSCustomObject]$State,
        [string]$CalendarId,
        [int]$DeployDays
    )
    $calState = $State.$CalendarId
    $now      = Get-Date

    return $Members | Where-Object {
        $email    = $_.email
        $lastSub = if ($calState -and $calState.$email) { [datetime]$calState.$email } else { $null }
        -not $lastSub -or ($now - $lastSub).TotalDays -gt $DeployDays
    }
}

function Update-State {
  # Stamps each subscribed user with the current datetime for this calendar.
  param(
    [PSCustomObject]$State,
    [string]$CalendarId,
    [array]$SubscribedUsers
  )
  if (-not $State.$CalendarId) {
    $State | Add-Member -NotePropertyName $CalendarId -NotePropertyValue ([PSCustomObject]@{}) -Force
  }
  $now = (Get-Date).ToString("o") # ISO 8601
  foreach ($user in $SubscribedUsers) {
    # Ensure we add the email as a property to the calendar sub-object
    if (-not $State.$CalendarId.$($user.email)) {
        $State.$CalendarId | Add-Member -NotePropertyName $user.email -NotePropertyValue $now -Force
    } else {
        $State.$CalendarId.$($user.email) = $now
    }
  }
  return $State
}

# --- Menu Helpers ---
function Write-Header {
  param([string]$Title)
  Clear-Host
  Write-Host ""
  Write-Host "=== $AppTitle - $Title ===" -ForegroundColor Cyan
  Write-Host ""
}

function Select-FromList {
  # Presents a numbered list and returns selected item
  param(
    [string]$Prompt,
    [array]$Items,
    [scriptBlock]$DisplayScript
  )
  if ($Items.Count -eq 0) { # Handling no items
    Write-Host "(none)" -ForegroundColor DarkGray
    return $null
  }
  for ($i = 0; $i -lt $Items.Count; $i++) {
    Write-Host "  [$($i + 1)] $(& $DisplayScript $Items[$i])"
  }
  Write-Host ""
  Write-Host "  [X] Cancel"
  $choice = Read-Host $Prompt
  if ($choice.ToUpper() -eq "X" -or $choice -eq "") { return $null }
  $index = [int]$choice - 1
  if ($index -ge 0 -and $index -lt $Items.Count) {
    return $Items[$index]
  }
  Write-Host "Invalid Selection" -ForegroundColor Red
  return $null
}

# --- Calendar mgmt ---
function Show-CalendarMenu {
  param([string]$ConfigPath)
  $cfg = Read-Config -ConfigPath $ConfigPath

  while ($true) {
    Write-Header "Manage Calendars"
    if ($cfg.Calendars.Count -eq 0) {
      Write-Host "  (no calendars defined)`n" -ForegroundColor DarkGray
    } else {
      $cfg.Calendars | ForEach-Object { Write-Host " - $($_.Label) ( $($_.Id) )" }
      Write-Host ""
    }

    Write-Host "[1] Add Calendar [2] Delete Calendar [X] Back/Cancel"
    $choice = Read-Host "`nSelection"

    switch ($choice.ToUpper()) {
      "1" {
        $id    = Read-Host "Enter Calendar ID (e.g. c_xxxx@group.calendar.google.com)"
        $label = Read-Host "Enter Label (e.g. Events)"
        if ($cfg.Calendars | Where-Object { $_.Id -eq $id }) { # Checks if duplicate calendars. Warns if found.
          Write-Host "A calendar with that ID already exists." -ForegroundColor Yellow
          Start-Sleep -Seconds 2
          continue
        }
        $cfg.Calendars += [PSCustomObject]@{ Id = $id; Label = $label }
        Save-Config -ConfigPath $ConfigPath -Groups $cfg.Groups -Calendars $cfg.Calendars -DeployDays $cfg.DeployDays
      }
      "2" {
        Write-Header "Delete Calendar"
        $cal = Select-FromList -Prompt "Select calendar to delete" -Items $cfg.Calendars -DisplayScript { param($c) "$($c.Label) ( $($c.Id) )" }
        if ($cal) {
          # Warn if any groups reference this calendar!
          $linked = @($cfg.Groups | Where-Object { $_.CalendarIds -contains $cal.Id })
          if ($linked.Count -gt 0) {
            Write-Host "WARNING: This calendar is linked to $($linked.Count) group(s):" -ForegroundColor Yellow
            $linked | ForEach-Object { Write-Host " - $($_.Label)" -ForegroundColor Yellow }
            $confirm = Read-Host "Delete and unlink from all groups? [Y/N]"
            if ($confirm.ToUpper() -ne "Y") { continue }
            # Unlink
            foreach ($group in $cfg.Groups) {
              $group.CalendarIds = @($group.CalendarIds | Where-Object { $_ -ne $cal.Id })
            }
          }
          $cfg.Calendars = @($cfg.Calendars | Where-Object { $_.Id -ne $cal.Id })
          Save-Config -ConfigPath $ConfigPath -Groups $cfg.Groups -Calendars $cfg.Calendars -DeployDays $cfg.DeployDays
          Write-Host "`nCalendar Deleted" -ForegroundColor Green
          Start-Sleep -Seconds 1
        }
      }
      "X" { return }
    }
  }
}

# --- Group mgmt ---
function Show-EditGroupMenu {
  param(
    [string]$ConfigPath,
    [PSCustomObject]$Group
    )
  $cfg = Read-Config -ConfigPath $ConfigPath

  while ($true) {
    Write-Header "Edit Group: $($Group.Label)"

    # Linked Calendars
    $linked = @($cfg.Calendars | Where-Object { $Group.CalendarIds -contains $_.Id })
    Write-Host "Linked Calendars:"
    if ($linked.Count -eq 0) { Write-Host " (none)" -ForegroundColor DarkGray }
    else { $linked | ForEach-Object { Write-Host "  - $($_.Label) ( $($_.Id) )"} }
    Write-Host ""

    Write-Host "[1] Link Calendar [2] Unlink Calendar [3] Delete Group [X] Back/Cancel"
    $choice = Read-Host "`nSelection"

    switch ($choice.ToUpper()) {
      "1" {
        Write-Header "Link Calendar to $($Group.Label)"
        $unlinked = @($cfg.Calendars | Where-Object { $Group.CalendarIds -notcontains $_.Id })
        if ($unlinked.Count -eq 0) {
          Write-Host "All available calendars are already linked." -ForegroundColor Yellow
          Start-Sleep -Seconds 2
          continue
        }
        $cal = Select-FromList -Prompt "`nSelect calendar to link" -Items $unlinked -DisplayScript { param($c) "$($c.Label) ( $($c.Id) )" }
        if ($cal) {
          $targetGroup = $cfg.Groups | Where-Object { $_.Email -eq $Group.Email }
          $targetGroup.CalendarIds = @($targetGroup.CalendarIds) + $cal.Id
          $Group.CalendarIds = $targetGroup.CalendarIds
          Save-Config -ConfigPath $ConfigPath -Groups $cfg.Groups -Calendars $cfg.Calendars -DeployDays $cfg.DeployDays
          Write-Host "`nLinked." -ForegroundColor Green
          Start-Sleep -Seconds 1
        }
      }
      "2" {
        Write-Header "Unlink Calendar from $($Group.Label)"
        if ($linked.Count -eq 0) {
          Write-Host "No calendars linked to this group." -ForegroundColor Yellow
          Start-Sleep -Seconds 2
          continue
        }
        $cal = Select-FromList -Prompt "`nSelect calendar to unlink" -Items $linked -DisplayScript { param($c) "$($c.Label) ( $($c.Id) )" }
        if ($cal) {
          $targetGroup = $cfg.Groups | Where-Object { $_.Email -eq $Group.Email }
          $targetGroup.CalendarIds = @($targetGroup.CalendarIds | Where-Object { $_ -ne $cal.Id })
          $Group.CalendarIds = $targetGroup.CalendarIds
          Save-Config -ConfigPath $ConfigPath -Groups $cfg.Groups -Calendars $cfg.Calendars -DeployDays $cfg.DeployDays
          Write-Host "`nUnlinked." -ForegroundColor Green
          Start-Sleep -Seconds 1
        }
      }
      "3" {
        $confirm = Read-Host "Delete group '$($Group.Label)'? [Y/N]"
        if ($confirm.ToUpper() -eq "Y") {
          $cfg.Groups = @($cfg.Groups | Where-Object { $_.Email -ne $Group.Email })
          Save-Config -ConfigPath $ConfigPath -Groups $cfg.Groups -Calendars $cfg.Calendars -DeployDays $cfg.DeployDays
          Write-Host "`nGroup deleted." -ForegroundColor Green
          Start-Sleep -Seconds 1
          return
        }
      }
      "X" { return }
    }
  }
}

function Show-GroupMenu {
  param([string]$ConfigPath)
  $cfg = Read-Config -ConfigPath $ConfigPath

  while($true) {
    Write-Header "Manage Groups"
    if ($cfg.Groups.Count -eq 0) {
      Write-Host "  (no groups defined)" -ForegroundColor DarkGray
    } else {
      $cfg.Groups | ForEach-Object {
        $calCount = @($_.CalendarIds).Count
        Write-Host "  - $($_.Label) ( $($_.Email) ) - $calCount calendar(s) linked"
      }
      Write-Host ""
    }

    Write-Host "[1] Add Group [2] Edit Group [X] Back/Cancel"
    $choice = Read-Host "`nSelection"

    switch ($choice.ToUpper()) {
      "1" {
        $email = Read-Host "Enter Group Email"
        $label = Read-Host "Enter Label (e.g. Marketing)"
        if ($cfg.Groups | Where-Object { $_.Email -eq $email }) { # Checks for duplicate groups, warns if found.
          Write-Host "A group with that email already exists." -ForegroundColor Yellow
          Start-Sleep -Seconds 2
          continue
        }
        $newGroup = [PSCustomObject]@{ Email = $email; Label = $label; CalendarIds = @() }
        $cfg.Groups += $newGroup
        Save-Config -ConfigPath $ConfigPath -Groups $cfg.Groups -Calendars $cfg.Calendars -DeployDays $cfg.DeployDays

        # Prompt to link calendars immediately
        if ($cfg.Calendars.Count -gt 0) {
          $link = Read-Host "Link calendars to this group? [Y/N]"
          if ($link.ToUpper() -eq "Y") {
            Show-EditGroupMenu -ConfigPath $ConfigPath -Group $newGroup
          }
        } else {
          Write-Host "No calendars defined yet. Add calendars first, then edit this group to link them."
          Start-Sleep -Seconds 3
        }
        $cfg = Read-Config -ConfigPath $ConfigPath
      }
      "2" {
        Write-Header "Edit Group"
        $cfg = Read-Config -ConfigPath $ConfigPath
        $group = Select-FromList -Prompt "`nSelect group to edit" -Items $cfg.Groups -DisplayScript { param($g) "$($g.Label) ( $($g.Email) )" }
        if ($group) {
          Show-EditGroupMenu -ConfigPath $ConfigPath -Group $group
          $cfg = Read-Config -ConfigPath $ConfigPath
        }
      }
      "X" { return }
    }
  }
}

# --- State Menus ---
function Show-StateMenu {
  param([string]$ConfigPath, [string]$StateDir)
  $cfg = Read-Config  -ConfigPath $ConfigPath

  while($true) {
    Write-Header "State Settings"
    Write-Host "  Current DeployDays: $($cfg.DeployDays)"
    Write-Host "  Users are re-subscribed if their last deployment is older than DeployDays." -ForegroundColor DarkGray
    Write-Host "  Set to 0 to always subscribe all users (disables state filtering)." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "[1] Update DeployDays  [2] Clear State Files  [X] Back/Cancel"
    $choice = Read-Host "`nSelection"

    switch ($choice.ToUpper()) {
      "1" {
        $days = Read-Host "Enter number of days between full re-subscribes (current: $($cfg.DeployDays))"
        if ($days -match '^\d+$') {
          $cfg.DeployDays = [int]$days
          Save-Config -ConfigPath $ConfigPath -Groups $cfg.Groups -Calendars $cfg.Calendars -DeployDays $cfg.DeployDays
          Write-Host "DeployDays updated to $($cfg.DeployDays)." -ForegroundColor Green
          Start-Sleep -Seconds 1
        } else {
          Write-Host "Invalid input. Please enter a whole number." -ForegroundColor Red
          Start-Sleep -Seconds 2
        }
      }
      "2" {
        # Wipe state files to force a full re-subscribe on next run
        $stateFiles = @(Get-ChildItem -Path $StateDir -Filter "state-*.json" -ErrorAction SilentlyContinue)
        if ($stateFiles.Count -eq 0) {
          Write-Host "No state files found." -ForegroundColor Yellow
          Start-Sleep -Seconds 2
          continue
        }
        Write-Host "This will delete $($stateFiles.Count) state file(s) and force a full re-subscribe on next run." -ForegroundColor Yellow
        $confirm = Read-Host "Continue? [Y/N]"
        if ($confirm.ToUpper() -eq "Y") {
          $stateFiles | Remove-Item -Force
          Write-Host "State files cleared." -ForegroundColor Green
          Start-Sleep -Seconds 1
        }
      }
      "X" { return }
    }
  }
}

# Main Menu
function Show-ConfigMenu {
  param([string]$ConfigPath)

  while ($true) {
    $cfg = Read-Config -ConfigPath $ConfigPath
    Write-Header "Main Menu"
    Write-Host "  Groups:     $(@($cfg.Groups).Count) defined"
    Write-Host "  Calendars:  $(@($cfg.Calendars).Count) defined"
    Write-Host "  Deploy Days:  $($cfg.DeployDays)"
    Write-Host "  Config:     $ConfigPath" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "[1] Manage Groups  [2] Manage Calendars  [3] State Settings  [Q] Quit"
    $choice = Read-Host "`nSelection"

    switch ($choice.ToUpper()) {
      "1" { Show-GroupMenu      -ConfigPath $ConfigPath }
      "2" { Show-CalendarMenu   -ConfigPath $ConfigPath }
      "3" { Show-StateMenu      -ConfigPath $ConfigPath -StateDir $StateDir }
      "Q" {
          Clear-Host
          return
      }
    }
  }
}

# --- Preflight Checks ---
function Start-Preflight {
  param([string]$ConfigPath, [string]$StateDir, [array]$Groups, [array]$Calendars)
  # Verify config exists...
  if (-not (Test-Path $ConfigPath)) {
    throw "No config found at '$ConfigPath'. Run with -Config to set up."
  }
  # Verify config isn't empty...
  if ((Get-Item $ConfigPath).Length -eq 0) {
    throw "Config file exists but is empty at '$ConfigPath'. Run with -Config to set up."
  }
  # Verify gam is in path...
  if (-not (Get-Command "gam" -ErrorAction SilentlyContinue)) {
    throw "GAM command not found in PATH. Ensure GAM is installed and accessible for the service account."
  }
  # Verify groups are defined (may not have any members!)
  if ($Groups.Count -eq 0) {
    throw "No groups defined in config. Run with -Config to set up."
  }
  # Verify calendars are defined (may not be associated with groups!)
  if ($Calendars.Count -eq 0) {
    throw "No calendars defined in config. Run with -Config to set up."
  }
  # Verify \state exists - creates it if not!
  if (-not (Test-Path $StateDir)) {
    New-Item -ItemType Directory -Path $StateDir -Force | Out-Null
  }
}

# --- DEPLOYMENT ENGINE ---
# Run w/ -Config param
if ($Config) {
  Show-ConfigMenu -ConfigPath $ConfigPath
  exit
}

# Temp data files
$TempCsv = [System.IO.Path]::GetTempFileName()
$DeltaCsv = [System.IO.Path]::GetTempFileName()

try {
  $cfg      = Read-Config -ConfigPath $ConfigPath
  $Groups   = @($cfg.Groups)
  $DeployDays = $cfg.DeployDays

  Start-Preflight -ConfigPath $ConfigPath -StateDir $StateDir -Groups $Groups -Calendars @($cfg.Calendars)

  foreach ($Group in $Groups) {
    $Calendars = @($cfg.Calendars | Where-Object { $Group.CalendarIds -contains $_.Id })

    if ($Calendars.Count -eq 0) {
      Write-Log "Skipping: '$($Group.Label)' - no calendars linked." -EntryType Warning
      continue
    }

    Write-Log "Starting GAM deployment: '$($Group.Label)' ($($Group.Email)) - $($calendars.Count) calendar(s)."
    <# ------
     GAM: 'redirect' writes directly to csv, ensuring there's no PS pipe formatting/artifacts to contend with
     'print group-members group ...' and 'recursive types user'
     ensures we grab all users that are members of this group and child groups.
    ------ #>
    $RetryCount = 0
    $CsvData    = @()

    while ($CsvData.Count -eq 0 -and $RetryCount -le $MaxRetries) {
      if ($RetryCount -gt 0) {
        Write-Log "Retry $RetryCount/$MaxRetries for '$($Group.Label)' - waiting 30s..." -EntryType Warning
        Start-Sleep -Seconds 30
      }
      gam redirect csv "$TempCsv" print group-members group "$($Group.Email)" recursive types user
      $GamExitCode = $LASTEXITCODE
      $CsvData = @(Import-Csv $TempCsv -ErrorAction SilentlyContinue)
      $RetryCount++

      if ($GamExitCode -ne 0) {
        Write-Log "GAM exited with code $GamExitCode for '$($Group.Label)'" -EntryType Warning
      }
    }

    # VALIDATE CSV CONTENT: Ensure we have more than just a header row
    if ($CsvData.Count -eq 0) {
      throw "No members found for '$($Group.Label)' after $($MaxRetries +1) attempts. Group may not exist or has no members."
    }

    Write-Log "Found $($CsvData.Count) members in '$($Group.Label)'."

    # Load state for this group
    $state = Read-State -StateDir $StateDir -GroupEmail $Group.Email

    foreach ($Calendar in $calendars) {
      if ($DeployDays -gt 0) {
        # Filter to only users who need subscribed
        $UsersToSub = @(Get-UsersNeedingSub -Members $CsvData -State $state -CalendarId $Calendar.Id -DeployDays $DeployDays)
      } else {
        # DeployDays = 0 means always deploy to everyone
        $UsersToSub = $CsvData
      }

      if ($usersToSub.Count -eq 0) {
        Write-Log "All members already subscribed to '$($Calendar.Label)' - skipping."
        continue
      }

      Write-Log "Adding '$($Calendar.Label)' to $($usersToSub.Count) member(s) of '$($Group.Label)'."

      # Write delta to temp CSV for GAM
      $UsersToSub | Export-Csv $deltaCsv -NoTypeInformation

      # Single-threaded call - use if experiencing API quota/throttling issues:
      # gam csv "$deltaCsv" gam user "~email" add calendar "$($Calendar.Id)" selected true

      # Multithreaded - adjust num_threads (current: 16) based on API quota and performance:
      gam config num_threads 16 csv "$deltaCsv" gam user "~email" add calendar "$($Calendar.Id)" selected true

      # Update state with newly subscribed users
      $State = Update-State -State $State -CalendarId $Calendar.Id -SubscribedUsers $UsersToSub
    }
    Save-State -StateDir $StateDir -GroupEmail $Group.Email -State $State
    Write-Log "Deployment complete for '$($Group.Label)'."
  }
} catch {
  Write-Log "CRITICAL ERROR: $($_.Exception.Message)" -EntryType Error
  throw $_
} finally {
  foreach ($f in @($TempCsv, $DeltaCsv)) {
    if ($f -and (Test-Path $f)) {
      Remove-Item $f -ErrorAction SilentlyContinue
    }
  }
}
