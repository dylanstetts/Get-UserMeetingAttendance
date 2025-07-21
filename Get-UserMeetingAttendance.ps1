
<#
.SYNOPSIS
    Retrieves Microsoft Teams meeting attendance data using Microsoft Graph SDK.

.PARAMETER UserPrincipalName
    The UPN of the user whose meetings you want to analyze.

.PARAMETER StartDate
    The start date for the calendar view.

.PARAMETER EndDate
    The end date for the calendar view.

.PARAMETER AttendanceOutput
    Path to export attendance data CSV.

.PARAMETER FailedOutput
    Path to export failed meetings CSV.

.PARAMETER EnableDebug
    Optional switch to save raw JSON responses for debugging.
#>

param (
    [string]$UserPrincipalName = "testing@exchangelabs.online",
    [datetime]$StartDate = "2025-01-01T00:00:00Z",
    [datetime]$EndDate = "2025-07-01T00:00:00Z",
    [string]$AttendanceOutput = "TeamsAttendanceReport.csv",
    [string]$FailedOutput = "FailedMeetings.csv",
    [switch]$EnableDebug
)

# =========================
# Connect to Microsoft Graph
# =========================
function Connect-ToGraph {
    Connect-MgGraph -Scopes "OnlineMeetings.Read", "Calendars.Read", "OnlineMeetingArtifact.Read.All"
}

function Get-UserObjectId {
    param ([string]$UserPrincipalName)
    return (Get-MgUser -UserId $UserPrincipalName).Id
}

function Get-TeamsMeetings {
    param (
        [string]$UserId,
        [datetime]$StartDate,
        [datetime]$EndDate
    )
    $events = Get-MgUserCalendarView -UserId $UserId `
        -StartDateTime $StartDate `
        -EndDateTime $EndDate `
        -Top 1000

    return $events | Where-Object {
        $_.IsOnlineMeeting -eq $true -and $_.OnlineMeetingProvider -eq "teamsForBusiness"
    }
}

function Get-OnlineMeetingByJoinUrl {
    param (
        [string]$UserId,
        [string]$JoinUrl
    )
    $encodedUrl = [System.Web.HttpUtility]::UrlEncode($JoinUrl)
    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/onlineMeetings?`$filter=JoinWebUrl eq '$encodedUrl'"
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri
    return $response.value[0]
}

function Get-ExpandedAttendanceRecords {
    param (
        [string]$UserId,
        [string]$MeetingId,
        [switch]$EnableDebug
    )
    Write-Host "Fetching attendance reports for meeting ID: $MeetingId"
    $reportsUri = "https://graph.microsoft.com/v1.0/users/$UserId/onlineMeetings/$MeetingId/attendanceReports"

    try {
        $reportsResponse = Invoke-MgGraphRequest -Method GET -Uri $reportsUri
    } catch {
        Write-Warning "Failed to retrieve attendance reports for meeting ID: $MeetingId"
        Write-Host "Error: $($_.Exception.Message)"
        return @()
    }

    if (-not $reportsResponse.value -or $reportsResponse.value.Count -eq 0) {
        Write-Warning "No attendance reports found for meeting ID: $MeetingId"
        return @()
    }

    $records = @()

    foreach ($report in $reportsResponse.value) {
        $reportId = $report.Id
        Write-Host "Expanding attendance report ID: $reportId"

        $expandedUri = "https://graph.microsoft.com/v1.0/users/$UserId/onlineMeetings/$MeetingId/attendanceReports/$reportId" + "?" + "`$expand=attendanceRecords"

        try {
            $expandedReport = Invoke-MgGraphRequest -Method GET -Uri $expandedUri

            if ($EnableDebug) {
                $logPath = ".\debug_attendanceReport_$($reportId).json"
                $expandedReport | ConvertTo-Json -Depth 10 | Out-File -FilePath $logPath -Encoding utf8
                Write-Host "Saved raw expanded report to $logPath"
            }

            if ($expandedReport.attendanceRecords -and $expandedReport.attendanceRecords.Count -gt 0) {
                Write-Host "Found $($expandedReport.attendanceRecords.Count) attendance records in report ID: $reportId"
                $records += $expandedReport.attendanceRecords
            } else {
                Write-Warning "No attendance records in expanded report ID: $reportId"
            }
        } catch {
            Write-Warning "Failed to expand report ID: $reportId"
            Write-Host "Error: $($_.Exception.Message)"
            continue
        }
    }

    return $records
}

function Export-AttendanceData {
    param (
        [array]$AttendanceData,
        [string]$Path
    )
    $AttendanceData | Export-Csv -Path $Path -NoTypeInformation
}

function Export-FailedMeetings {
    param (
        [array]$FailedMeetings,
        [string]$Path
    )
    $FailedMeetings | ForEach-Object {
        [PSCustomObject]@{
            Subject       = $_.Subject
            Start         = $_.Start.DateTime
            Type          = $_.Type
            JoinUrl       = $_.OnlineMeeting?.JoinUrl
            OnlineMeeting = $_.OnlineMeeting?.ToString()
        }
    } | Export-Csv -Path $Path -NoTypeInformation
}

# =========================
# Main Execution Function
# =========================
function Main {

    Connect-ToGraph

    $userObjectId = Get-UserObjectId -UserPrincipalName $UserPrincipalName
    $meetings = Get-TeamsMeetings -UserId $UserPrincipalName -StartDate $StartDate -EndDate $EndDate

    $attendanceData = @()
    $failedMeetings = @()

    foreach ($meeting in $meetings) {
        Write-Host "`nProcessing: $($meeting.Subject)"

        if (-not $meeting.OnlineMeeting?.JoinUrl -or $meeting.Type -eq "seriesMaster") {
            Write-Warning "Skipping invalid or recurring meeting: $($meeting.Subject)"
            $failedMeetings += $meeting
            continue
        }

        try {
            $onlineMeeting = Get-OnlineMeetingByJoinUrl -UserId $userObjectId -JoinUrl $meeting.OnlineMeeting.JoinUrl

            if (-not $onlineMeeting) {
                Write-Warning "Online meeting not found: $($meeting.Subject)"
                $failedMeetings += $meeting
                continue
            }

            $records = Get-ExpandedAttendanceRecords -UserId $userObjectId -MeetingId $onlineMeeting.Id -EnableDebug:$EnableDebug

            if (-not $records -or $records.Count -eq 0) {
                Write-Warning "No attendance records found: $($meeting.Subject)"
                $failedMeetings += $meeting
                continue
            }

            foreach ($record in $records) {
                $attendanceData += [PSCustomObject]@{
                    MeetingID       = $onlineMeeting.Id
                    Subject         = $onlineMeeting.Subject
                    Organizer       = $onlineMeeting.participants.organizer.upn
                    AttendeeName    = $record.Identity.displayName
                    AttendeeUserId  = $record.Identity.id
                    Role            = $record.Role
                    JoinTime        = $record.attendanceIntervals.JoinDateTime
                    LeaveTime       = $record.attendanceIntervals.LeaveDateTime
                    DurationSeconds = $record.attendanceIntervals.DurationInSeconds
                }
            }

            Write-Host "Added $($records.Count) attendance records."
        } catch {
            Write-Warning "Error processing meeting: $($meeting.Subject)"
            Write-Host "Error: $($_.Exception.Message)"
            $failedMeetings += $meeting
        }
    }

    Export-AttendanceData -AttendanceData $attendanceData -Path $AttendanceOutput
    Export-FailedMeetings -FailedMeetings $failedMeetings -Path $FailedOutput

    Write-Host "`nScript completed. Attendance and error reports exported."
}

# =========================
# Call Main with Defaults
# =========================
Main -UserPrincipalName $UserPrincipalName `
     -StartDate $StartDate `
     -EndDate $EndDate `
     -AttendanceOutput $AttendanceOutput `
     -FailedOutput $FailedOutput `
     -EnableDebug:$EnableDebug
