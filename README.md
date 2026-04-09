# Sync-CalendarSubscriptions

A PowerShell script that uses [GAM7](https://github.com/GAM-team/GAM) to automatically subscribe Google Group members to one or more Google Calendars. Group membership is resolved recursively, so users in nested child groups are included. Each group can be mapped to its own set of calendars.

---

## Requirements

- PowerShell 5.x
- [GAM7](https://github.com/GAM-team/GAM/wiki/How-to-Install-GAM7) installed and configured with appropriate Google Workspace admin credentials
- Windows (uses Windows Event Log for logging)

---

## Setup

### 1. Configure GAM7

Ensure GAM7 is installed and authorized for your Google Workspace domain before using this script. GAM7 must be accessible on the `PATH` of whichever account runs the script (including the service account used by Task Scheduler, if applicable).

See the [GAM7 installation guide](https://github.com/GAM-team/GAM/wiki/How-to-Install-GAM7) for details.

### 2. Create a config.json

Run the script with the `-Config` flag to launch the interactive configuration menu:

```powershell
.\Sync-GroupCalendars.ps1 -Config
```

Use the menu to add one or more Google Groups and Calendars separately. The config is saved as `config.json` in the same directory as the script.

To remove entries or make bulk edits, open `config.json` directly in any text editor.

```json
{
  "Groups": [
    {
      "Email": "my-group@domain.com",
      "Label": "My Group",
      "CalendarIds": [
        "c_abc123...@group.calendar.google.com"
      ]
    }
  ],
  "Calendars": [
    {
      "Id": "c_abc123...@group.calendar.google.com",
      "Label": "All Staff Events"
    }
  ]
}
```

Each group maps to its own list of calendars via `CalendarIds`. A calendar can be linked to multiple groups. The top-level `Calendars` array is the shared pool that groups reference from by ID.

---

## Usage

### Run manually

```powershell
.\Sync-GroupCalendars.ps1
```

### Open the config menu

```powershell
.\Sync-GroupCalendars.ps1 -Config
```

### Use a custom config path

```powershell
.\Sync-GroupCalendars.ps1 -ConfigPath "C:\Scripts\my-config.json"
```

### Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-Config` | Switch | — | Launches the interactive config menu |
| `-ConfigPath` | String | `.\config.json` | Path to the config file |
| `-AppTitle` | String | `Sync-CalendarSubscriptions` | Used as the Windows Event Log source name and in the Config Menu |

---

## How It Works

For each Group defined in `config.json`, the script:

1. Resolves which calendars are linked to that group via `CalendarIds`
2. Calls GAM7 to fetch all user members of the group, recursively resolving any nested child groups (`recursive types user`)
3. Validates that at least one user was returned
4. For each linked calendar, calls GAM7 to add the calendar to each user's account with `selected true` (visible by default)

Groups with no linked calendars are skipped, logging a warning.

If a user is already subscribed to a calendar, GAM7's `add calendar` is idempotent — it will not create duplicates or throw an error.

> **Note:** If a user is a member of multiple nested child groups within the same parent, they may appear more than once in the member list. This does not cause problems but will be reflected in the logged user count.

---

## Logging

All activity is written to the **Windows Event Log** under `Application` with the source `Sync-GroupCalendars` (or whatever `-EventLogSource` is set to).

| Event | Level |
|---|---|
| Sync started for a group | Information |
| User count found | Information |
| Calendar being processed | Information |
| Sync complete for a group | Information |
| No members found / group missing | Error |
| Any unhandled exception | Error |

To view logs:

```
Event Viewer → Windows Logs → Application → Source: Sync-GroupCalendars
```

> **First run:** Creating a new Event Log source requires administrator privileges. If the script is not run as an administrator on first use, it will fall back to `Write-Output` for that session. Run once as administrator (or pre-register the source) to initialize it permanently.

### Pre-register the Event Log source (run once as administrator)

```powershell
New-EventLog -LogName Application -Source "Sync-GroupCalendars"
```

---

## Scheduling

To run on a schedule, create a Task Scheduler job that calls:

```
Program: powershell.exe
Arguments: -NonInteractive -ExecutionPolicy Bypass -File "C:\Scripts\Sync-GroupCalendars\Sync-GroupCalendars.ps1"
```
_Make sure this points to the script on your own system!_

Ensure the task runs under an account that has:
- GAM7 on its `PATH`
- GAM7 credentials configured for that user profile
- Permission to write to the Windows Event Log (or the source pre-registered by an admin)

---

## File Structure

```
Sync-GroupCalendars/
├── Sync-GroupCalendars.ps1
└── config.json
```