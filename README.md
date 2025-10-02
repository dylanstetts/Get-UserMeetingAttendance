# Get-UserMeetingAttendance
 This PowerShell script uses the Microsoft Graph PowerShell SDK to extract detailed attendance reports from Microsoft Teams meetings for a specified user and date range.

## Features

- Connects to Microsoft Graph with required scopes
- Retrieves Teams meetings from a user's calendar
- Expands attendance reports to include join/leave times and durations
- Exports attendance data to CSV
- Logs failed meetings for review
- Optional debug logging of raw API responses

## Prerequisites

- PowerShell 7+
- Microsoft.Graph PowerShell SDK
- Admin consent for the following Graph scopes:
  - `OnlineMeetings.Read`
  - `Calendars.Read`
  - `OnlineMeetingArtifact.Read.All`
- Application Access Policy for Teams Calendar data is required
  - See more: https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy#configure-application-access-policy


Install the SDK if needed:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

## Usage

1. Clone or download this repository.
2. Open the script in PowerShell.
3. Modify the default parameters in the Main function or call it with custom arguments:

```powershell
Main -UserPrincipalName "user@domain.com" `
     -StartDate "2025-01-01T00:00:00Z" `
     -EndDate "2025-07-01T00:00:00Z" `
     -AttendanceOutput "Attendance.csv" `
     -FailedOutput "Failed.csv" `
     -EnableDebug
```

4. Run the script. It will:
 - Authenticate with Microsoft Graph
 - Retrieve Teams meetings for the user
 - Fetch and expand attendance reports
 - Export results to CSV files

## Output

- TeamsAttendanceReport.csv: Contains detailed attendance records
- FailedMeetings.csv: Meetings that failed to process
- debug_attendanceReport_<ID>.json (optional): Raw API responses for debugging (if switch $EnableDebugging is on)

## Script Structure

- Connect-ToGraph: Authenticates with Microsoft Graph
- Get-UserObjectId: Resolves user principal name to object ID
- Get-TeamsMeetings: Retrieves Teams meetings in a date range
- Get-OnlineMeetingByJoinUrl: Finds the online meeting object
- Get-ExpandedAttendanceRecords: Retrieves and expands attendance reports
- Export-AttendanceData: Outputs attendance to CSV
- Export-FailedMeetings: Outputs failed meetings to CSV
- Main: Orchestrates the full process
