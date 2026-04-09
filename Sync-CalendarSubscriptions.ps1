<#
  .SYNOPSIS
  Utilizes GAM7 to automate Google Calendar Subscriptions for Google Group members.
  .DESCRIPTION
  Reads group/calendar pairs from a config.json file and uses GAM7 to subscribe all user members (including nested groups) to the calendars.
  .PARAMETER Config
  Launches an interactive mode to view/add Groups and Calendars to config.json
  .PARAMETER ConfigPath
  Path to the config.json file. Defaults to config.json in the script directory.
  .PARAMETER AppTitle
  Used in app menu and as the Windows Event Log source name. Defaults to Sync-CalendarSubscriptions

  .NOTES
  Requires a config.json file and requires GAM7.

#>
[CmdletBinding()]
param (
    [Parameter()]
    [switch]$Config,
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [string]$AppTitle = "Sync-CalendarSubscriptions"
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
  } else { # No config? No problem - we create an object
    $groups    = @()
    $calendars = @()
  }
  return @{ Groups = $groups; Calendars = $calendars }
}

# --- Write the object to the config file ---
function Save-Config {
  param([string]$ConfigPath, [array]$Groups, [array]$Calendars)
  $out = [PSCustomObject]@{ Groups = @($Groups); Calendars = @($Calendars) }
  $out | ConvertTo-Json -Depth 10 | Out-File $ConfigPath
  Write-Host "`nSaved: $ConfigPath" -ForegroundColor Green
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
        Save-Config -ConfigPath $ConfigPath -Groups $cfg.Groups -Calendars $cfg.Calendars
      }
      "2" {
        Write-Header "Delete Calendar"
        $cal = Select-FromList -Prompt "Select calender to delete" -Items $cfg.Calendars -DisplayScript { param($c) "$($c.Label) ( $($c.Id) )" }
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
          Save-Config -ConfigPath $ConfigPath -Groups $cfg.Groups -Calendars $cfg.Calendars
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
          Save-Config -ConfigPath $ConfigPath -Groups $cfg.Groups -Calendars $cfg.Calendars
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
          Save-Config -ConfigPath $ConfigPath -Groups $cfg.Groups -Calendars $cfg.Calendars
          Write-Host "`nUnlinked." -ForegroundColor Green
          Start-Sleep -Seconds 1
        }
      }
      "3" {
        $confirm = Read-Host "Delete group '$($Group.Label)'? [Y/N]"
        if ($confirm.ToUpper() -eq "Y") {
          $cfg.Groups = @($cfg.Groups | Where-Object { $_.Email -ne $Group.Email })
          Save-Config -ConfigPath $ConfigPath -Groups $cfg.Groups -Calendars $cfg.Calendars
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
        Save-Config -ConfigPath $ConfigPath -Groups $cfg.Groups -Calendars $cfg.Calendars

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

# Main Menu
function Show-ConfigMenu {
  param([string]$ConfigPath)

  while ($true) {
    $cfg = Read-Config -ConfigPath $ConfigPath
    Write-Header "Main Menu"

    Write-Host "  Groups:     $(@($cfg.Groups).Count) defined"
    Write-Host "  Calendars:  $(@($cfg.Calendars).Count) defined"
    Write-Host "  Config:     $ConfigPath" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "[1] Manage Groups [2] Manage Calendars [Q] Quit"
    $choice = Read-Host "`nSelection"

    switch ($choice.ToUpper()) {
      "1" { Show-GroupMenu    -ConfigPath $ConfigPath }
      "2" { Show-CalendarMenu -ConfigPath $ConfigPath }
      "Q" {
        Clear-Host
        return
      }
    }
  }
}

# --- NON-INTERACTIVE MODE ---
# --- Preflight Checks ---
function Start-Preflight {
  # Verify config is present
  param([string]$ConfigPath)
  if (-not (Test-Path $ConfigPath)) {
    throw "No config found at $ConfigPath. Run with -Config to set up."
  }
  # Verify config is not empty
  if ((Get-Item $ConfigPath).Length -eq 0) {
    throw "Config file exists but is empty at '$ConfigPath'. Run with -Config to set up."
  }
  # Verify GAM accessibility
  if (-not (Get-Command "gam" -ErrorAction SilentlyContinue)) {
    throw "GAM command not found in PATH. Ensure GAM is installed and accessible for the service account."
  }
}

# --- Entry Point ---
# Run w/ -Config param
if ($Config) {
  Show-ConfigMenu -ConfigPath $ConfigPath
  exit
}

$tempCsv = [System.IO.Path]::GetTempFileName()

try {
  Start-Preflight -ConfigPath $ConfigPath
  $cfg = Read-Config -ConfigPath $ConfigPath
  $groups = @($cfg.Groups)

  # Verifying config actually has groups
  if (@($groups).Count -eq 0) {
    throw "No groups defined in config. Run with -Config to set up."
  }
  # Verifying config actually has calendars
  if (@($cfg.Calendars).Count -eq 0) {
    throw "No calendars defined in config. Run with -Config to set up."
  }

  foreach ($Group in $groups) {
    $calendars = @($cfg.Calendars | Where-Object { $Group.CalendarIds -contains $_.Id })

    if ($calendars.Count -eq 0) {
      Write-Log "Skipping: '$($Group.Label)' - no calendars linked." -EntryType Warning
      continue
    }

    Write-Log "Starting GAM sync: '$($Group.Label)' ($($Group.Email)) - $($calendars.Count) calendar(s)."
    # GAM: 'redirect' writes directly to csv, ensuring there's no PS pipe formatting/artifacts to contend with
    # 'print group-members group ...' and 'recursive types user' ensures we grab all users that are members of this group and child groups.
    gam redirect csv "$tempCsv" print group-members group "$($Group.Email)" recursive types user

    # VALIDATE CSV CONTENT: Ensure we have more than just a header row
    $csvData = @(Import-Csv $tempCsv -ErrorAction SilentlyContinue)
    if (-not $csvData) {
        throw "No members found for group $($Group.Label) or group may not exist or has no members."
    }
    Write-Log "Found $($csvData.Count) users to process."

    foreach ($Calendar in $calendars) {
      Write-Log "Processing: Adding $($Calendar.Label) to members of $($Group.Label)."
      gam csv $tempCsv gam user "~email" add calendar "$($Calendar.Id)" selected true
    }
    Write-Log "Sync complete for '$($Group.Label)'."
  }
}
catch {
    Write-Log "CRITICAL ERROR: $($_.Exception.Message)" -EntryType Error
    throw $_
}
finally {
    if (Test-Path $tempCsv) {
        Remove-Item $tempCsv -ErrorAction SilentlyContinue
    }
}
