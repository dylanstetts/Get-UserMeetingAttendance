# =========================
# Connect to Microsoft Graph
# =========================
function Connect-ToGraph {
    # Authenticate with Microsoft Graph using required scopes
    Connect-MgGraph -Scopes "OnlineMeetings.Read", "Calendars.Read", "OnlineMeetingArtifact.Read.All"
}


# =========================
# Get User Object ID
# =========================
function Get-UserObjectId {
    param ([string]$UserPrincipalName)
    # Retrieve the Azure AD object ID for the specified user
    return (Get-MgUser -UserId $UserPrincipalName).Id
}


# =========================
# Get Teams Meetings in Date Range
# =========================
function Get-TeamsMeetings {
    param (
        [string]$UserId,
        [datetime]$StartDate,
        [datetime]$EndDate
    )
    
    # Retrieve calendar events and filter for Teams meetings
    $events = Get-MgUserCalendarView -UserId $UserId `
        -StartDateTime $StartDate `
        -EndDateTime $EndDate `
        -Top 1000

    return $events | Where-Object {
        $_.IsOnlineMeeting -eq $true -and $_.OnlineMeetingProvider -eq "teamsForBusiness"
    }
}

# =========================
# Get Online Meeting by Join URL
# =========================
function Get-OnlineMeetingByJoinUrl {
    param (
        [string]$UserId,
        [string]$JoinUrl
    )
    # URL-encode the Join URL and query the online meeting
    $encodedUrl = [System.Web.HttpUtility]::UrlEncode($JoinUrl)
    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/onlineMeetings?`$filter=JoinWebUrl eq '$encodedUrl'"
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri
    return $response.value[0]
}


# =========================
# Get Expanded Attendance Records
# =========================
function Get-ExpandedAttendanceRecords {
    param (
        [string]$UserId,
        [string]$MeetingId,
        [switch]$EnableDebug
    )
    Write-Host "Fetching attendance reports for meeting ID: $MeetingId"
    $reportsUri = "https://graph.microsoft.com/v1.0/users/$UserId/onlineMeetings/$MeetingId/attendanceReports"

    try {
        # Retrieve all attendance reports for the meeting
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
        
        # Correctly construct the URI with $expand query parameter
        $expandedUri = "https://graph.microsoft.com/v1.0/users/$UserId/onlineMeetings/$MeetingId/attendanceReports/$reportId" + "?" + "`$expand=attendanceRecords"

        try {
            # Retrieve the expanded attendance report
            $expandedReport = Invoke-MgGraphRequest -Method GET -Uri $expandedUri
            
            # Optionally log the raw response for debugging
            if ($EnableDebug) {
                $logPath = ".\debug_attendanceReport_$($reportId).json"
                $expandedReport | ConvertTo-Json -Depth 10 | Out-File -FilePath $logPath -Encoding utf8
                Write-Host "Saved raw expanded report to $logPath"
            }
            
            # Append attendance records if present
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

# =========================
# Export Attendance Data
# =========================
function Export-AttendanceData {
    param (
        [array]$AttendanceData,
        [string]$Path
    )
    # Export attendance data to CSV
    $AttendanceData | Export-Csv -Path $Path -NoTypeInformation
}

# =========================
# Export Failed Meetings
# =========================
function Export-FailedMeetings {
    param (
        [array]$FailedMeetings,
        [string]$Path
    )
    
    # Export failed meetings with basic metadata
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
    param (
        [string]$UserPrincipalName = "dsadmin@exchangelabs.online",
        [datetime]$StartDate = "2025-01-01T00:00:00Z",
        [datetime]$EndDate = "2025-07-01T00:00:00Z",
        [string]$AttendanceOutput = "TeamsAttendanceReport.csv",
        [string]$FailedOutput = "FailedMeetings.csv",
        [switch]$EnableDebug
    )

    # Authenticate with Microsoft Graph
    Connect-ToGraph

    # Get user object ID
    $userObjectId = Get-UserObjectId -UserPrincipalName $UserPrincipalName

    # Retrieve Teams meetings in the specified date range
    $meetings = Get-TeamsMeetings -UserId $UserPrincipalName -StartDate $StartDate -EndDate $EndDate

    $attendanceData = @()
    $failedMeetings = @()

    foreach ($meeting in $meetings) {
        Write-Host "`nProcessing: $($meeting.Subject)"

        # Skip invalid or recurring series master meetings
        if (-not $meeting.OnlineMeeting?.JoinUrl -or $meeting.Type -eq "seriesMaster") {
            Write-Warning "Skipping invalid or recurring meeting: $($meeting.Subject)"
            $failedMeetings += $meeting
            continue
        }

        try {
            # Retrieve the online meeting object
            $onlineMeeting = Get-OnlineMeetingByJoinUrl -UserId $userObjectId -JoinUrl $meeting.OnlineMeeting.JoinUrl

            if (-not $onlineMeeting) {
                Write-Warning "Online meeting not found: $($meeting.Subject)"
                $failedMeetings += $meeting
                continue
            }

            # Retrieve expanded attendance records
            $records = Get-ExpandedAttendanceRecords -UserId $userObjectId -MeetingId $onlineMeeting.Id -EnableDebug:$EnableDebug

            if (-not $records -or $records.Count -eq 0) {
                Write-Warning "No attendance records found: $($meeting.Subject)"
                $failedMeetings += $meeting
                continue
            }

            # Append each attendance record to the output
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

    # Export results
    Export-AttendanceData -AttendanceData $attendanceData -Path $AttendanceOutput
    Export-FailedMeetings -FailedMeetings $failedMeetings -Path $FailedOutput

    Write-Host "`nScript completed. Attendance and error reports exported."
}

# =========================
# Call Main with Defaults
# =========================
Main
