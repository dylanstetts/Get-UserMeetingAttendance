# Teams Attendance Report - Usage Examples

This document provides comprehensive examples of how to use the Teams Attendance Report tool in various scenarios.

## Quick Start

### 1. First Time Setup

```powershell
# 1. Copy configuration template
Copy-Item config.template.json config.json

# 2. Edit config.json with your Azure app details
# Replace placeholders with actual values:
# - ApplicationId: Your Azure app ID
# - TenantId: Your Azure tenant ID
# - ClientSecret: Your Azure app secret
# - UserPrincipalName: Default user to analyze

# 3. Run with default settings
.\Get-TeamsAttendanceReport.ps1
```

### 2. Basic Usage Examples

```powershell
# Run for specific user and date range
.\Get-TeamsAttendanceReport.ps1 `
    -UserPrincipalName "john.doe@company.com" `
    -StartDate "2025-01-01" `
    -EndDate "2025-01-31"

# Run with custom output prefix
.\Get-TeamsAttendanceReport.ps1 `
    -OutputPrefix "JanuaryReport" `
    -UserPrincipalName "manager@company.com"

# Enable debug mode for troubleshooting
.\Get-TeamsAttendanceReport.ps1 `
    -EnableDebug `
    -LogLevel "Debug"
```

## Meeting Type Scenarios

### 1. All Meeting Types (Default)

```powershell
# Includes all available meeting types
.\Get-TeamsAttendanceReport.ps1 -UserPrincipalName "user@company.com"
```

### 2. Only Scheduled Meetings

```powershell
# Calendar-based meetings only
.\Get-TeamsAttendanceReport.ps1 `
    -UserPrincipalName "user@company.com" `
    -MeetingTypes @("Scheduled")
```

### 3. Calls and Instant Meetings

```powershell
# Focus on ad-hoc communications
.\Get-TeamsAttendanceReport.ps1 `
    -UserPrincipalName "user@company.com" `
    -MeetingTypes @("Instant", "OneOnOne")
```

### 4. Webinars and Large Events

```powershell
# Large-scale events only
.\Get-TeamsAttendanceReport.ps1 `
    -UserPrincipalName "events@company.com" `
    -MeetingTypes @("Webinar", "Townhall", "Broadcast")
```

## Configuration Examples

### 1. Custom Configuration File

Create a separate config for different environments:

```powershell
# Production environment
.\Get-TeamsAttendanceReport.ps1 -ConfigPath ".\config.production.json"

# Development environment  
.\Get-TeamsAttendanceReport.ps1 -ConfigPath ".\config.development.json"

# Specific project
.\Get-TeamsAttendanceReport.ps1 -ConfigPath ".\configs\project-alpha.json"
```

### 2. High-Volume Environment Configuration

For large tenants with many meetings:

```json
{
    "DefaultSettings": {
        "RequestDelayMs": 200,
        "MaxConcurrentRequests": 3,
        "EnableVerboseLogging": false
    }
}
```

```powershell
.\Get-TeamsAttendanceReport.ps1 -ConfigPath ".\config.high-volume.json"
```

### 3. Debug Configuration

For detailed troubleshooting:

```json
{
    "DefaultSettings": {
        "EnableDebug": true,
        "EnableVerboseLogging": true,
        "RequestDelayMs": 500
    }
}
```

## Date Range Scenarios

### 1. Last 30 Days (Default)

```powershell
# No date parameters - uses last 30 days
.\Get-TeamsAttendanceReport.ps1 -UserPrincipalName "user@company.com"
```

### 2. Specific Month

```powershell
# January 2025
.\Get-TeamsAttendanceReport.ps1 `
    -UserPrincipalName "user@company.com" `
    -StartDate "2025-01-01T00:00:00Z" `
    -EndDate "2025-01-31T23:59:59Z"
```

### 3. Quarter Report

```powershell
# Q1 2025
.\Get-TeamsAttendanceReport.ps1 `
    -UserPrincipalName "user@company.com" `
    -StartDate "2025-01-01T00:00:00Z" `
    -EndDate "2025-03-31T23:59:59Z" `
    -OutputPrefix "Q1_2025"
```

### 4. Weekly Report

```powershell
# Last week
$lastWeekStart = (Get-Date).AddDays(-7).Date
$lastWeekEnd = (Get-Date).Date

.\Get-TeamsAttendanceReport.ps1 `
    -UserPrincipalName "user@company.com" `
    -StartDate $lastWeekStart `
    -EndDate $lastWeekEnd `
    -OutputPrefix "Weekly"
```

## Troubleshooting Scenarios

### 1. Connection Issues

```powershell
# Test with minimal scope and debug enabled
.\Get-TeamsAttendanceReport.ps1 `
    -UserPrincipalName "test@company.com" `
    -StartDate (Get-Date).AddDays(-1) `
    -EndDate (Get-Date) `
    -EnableDebug `
    -LogLevel "Debug"
```

### 2. Permission Issues

```powershell
# Check what permissions are available
Connect-MgGraph -TenantId "your-tenant-id" -ClientId "your-app-id"
Get-MgContext | Select-Object Scopes
```

### 3. Rate Limiting Issues

```powershell
# Slow down requests for rate limiting
.\Get-TeamsAttendanceReport.ps1 `
    -UserPrincipalName "user@company.com" `
    -ConfigPath ".\config.slow.json"  # Config with RequestDelayMs: 1000
```

## Batch Processing Scenarios

### 1. Multiple Users

```powershell
# Process multiple users
$users = @(
    "user1@company.com",
    "user2@company.com", 
    "user3@company.com"
)

foreach ($user in $users) {
    Write-Host "Processing user: $user"
    .\Get-TeamsAttendanceReport.ps1 `
        -UserPrincipalName $user `
        -OutputPrefix "User_$($user.Split('@')[0])"
    
    # Wait between users to avoid rate limiting
    Start-Sleep -Seconds 30
}
```

### 2. Monthly Reports

```powershell
# Generate reports for each month in 2024
for ($month = 1; $month -le 12; $month++) {
    $startDate = Get-Date -Year 2024 -Month $month -Day 1
    $endDate = $startDate.AddMonths(1).AddDays(-1)
    
    .\Get-TeamsAttendanceReport.ps1 `
        -UserPrincipalName "manager@company.com" `
        -StartDate $startDate `
        -EndDate $endDate `
        -OutputPrefix "2024_Month_$($month.ToString('00'))"
}
```

### 3. Department Analysis

```powershell
# Analyze different departments
$departments = @{
    "Sales" = @("sales1@company.com", "sales2@company.com")
    "Marketing" = @("marketing1@company.com", "marketing2@company.com")
    "Engineering" = @("dev1@company.com", "dev2@company.com")
}

foreach ($dept in $departments.Keys) {
    foreach ($user in $departments[$dept]) {
        .\Get-TeamsAttendanceReport.ps1 `
            -UserPrincipalName $user `
            -OutputPrefix "$dept`_$($user.Split('@')[0])"
    }
}
```

## Advanced Scenarios

### 1. Scheduled Execution

Create a scheduled task to run reports automatically:

```powershell
# PowerShell script for scheduled execution
param(
    [string]$EmailRecipient = "reports@company.com"
)

try {
    # Run the report
    .\Get-TeamsAttendanceReport.ps1 `
        -UserPrincipalName "manager@company.com" `
        -OutputPrefix "Daily_$(Get-Date -Format 'yyyyMMdd')"
    
    # Email the results (requires email configuration)
    $attachments = Get-ChildItem -Filter "Daily_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 2
    
    Send-MailMessage `
        -To $EmailRecipient `
        -From "noreply@company.com" `
        -Subject "Daily Teams Attendance Report - $(Get-Date -Format 'yyyy-MM-dd')" `
        -Body "Please find attached the daily Teams attendance report." `
        -Attachments $attachments.FullName `
        -SmtpServer "smtp.company.com"
        
} catch {
    Write-Error "Failed to generate daily report: $($_.Exception.Message)"
    # Send error notification email
}
```

### 2. Custom Data Processing

Post-process the CSV data for specific insights:

```powershell
# Run the report first
.\Get-TeamsAttendanceReport.ps1 -UserPrincipalName "manager@company.com"

# Get the latest attendance file
$latestFile = Get-ChildItem -Filter "*Attendance*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

# Import and analyze the data
$data = Import-Csv $latestFile.FullName

# Calculate total meeting time per attendee
$attendeeSummary = $data | Group-Object AttendeeName | ForEach-Object {
    $totalMinutes = ($_.Group | Measure-Object DurationMinutes -Sum).Sum
    [PSCustomObject]@{
        AttendeeName = $_.Name
        TotalMeetings = $_.Count
        TotalMinutes = $totalMinutes
        TotalHours = [math]::Round($totalMinutes / 60, 2)
        AverageMinutesPerMeeting = [math]::Round($totalMinutes / $_.Count, 2)
    }
}

$attendeeSummary | Export-Csv "AttendanceSummary.csv" -NoTypeInformation
```

### 3. Integration with Power BI

Prepare data for Power BI analysis:

```powershell
# Run comprehensive report
.\Get-TeamsAttendanceReport.ps1 `
    -UserPrincipalName "analytics@company.com" `
    -StartDate "2025-01-01" `
    -EndDate "2025-12-31" `
    -OutputPrefix "PowerBI_Annual"

# Post-process for Power BI
$data = Import-Csv "PowerBI_Annual_Attendance_*.csv"

# Add calculated columns for Power BI
$enrichedData = $data | ForEach-Object {
    $_ | Add-Member -NotePropertyName "AttendanceDate" -NotePropertyValue (Get-Date $_.JoinTime -Format "yyyy-MM-dd") -PassThru |
         Add-Member -NotePropertyName "AttendanceHour" -NotePropertyValue (Get-Date $_.JoinTime -Format "HH") -PassThru |
         Add-Member -NotePropertyName "DayOfWeek" -NotePropertyValue (Get-Date $_.JoinTime).DayOfWeek -PassThru |
         Add-Member -NotePropertyName "Month" -NotePropertyValue (Get-Date $_.JoinTime -Format "yyyy-MM") -PassThru
}

$enrichedData | Export-Csv "PowerBI_Enhanced_Data.csv" -NoTypeInformation
```

## Environment-Specific Examples

### 1. Development Environment

```powershell
# Minimal data for testing
.\Get-TeamsAttendanceReport.ps1 `
    -UserPrincipalName "dev@company.com" `
    -StartDate (Get-Date).AddDays(-2) `
    -EndDate (Get-Date) `
    -MeetingTypes @("Scheduled") `
    -EnableDebug
```

### 2. Production Environment

```powershell
# Full production run with error handling
try {
    .\Get-TeamsAttendanceReport.ps1 `
        -ConfigPath ".\config.production.json" `
        -UserPrincipalName "production@company.com" `
        -LogLevel "Warning"
    
    Write-Host "Production report completed successfully"
} catch {
    Write-Error "Production report failed: $($_.Exception.Message)"
    # Log to monitoring system
}
```

### 3. Multi-Tenant Scenario

```powershell
# Different tenants with different configs
$tenants = @(
    @{Config = ".\config.tenant1.json"; User = "admin@tenant1.com"},
    @{Config = ".\config.tenant2.json"; User = "admin@tenant2.com"},
    @{Config = ".\config.tenant3.json"; User = "admin@tenant3.com"}
)

foreach ($tenant in $tenants) {
    $tenantName = Split-Path $tenant.Config -LeafBase
    .\Get-TeamsAttendanceReport.ps1 `
        -ConfigPath $tenant.Config `
        -UserPrincipalName $tenant.User `
        -OutputPrefix $tenantName
}
```

## Best Practices Examples

### 1. Error Handling Wrapper

```powershell
function Invoke-TeamsReportWithRetry {
    param(
        [string]$UserPrincipalName,
        [int]$MaxRetries = 3
    )
    
    for ($i = 1; $i -le $MaxRetries; $i++) {
        try {
            Write-Host "Attempt $i of $MaxRetries"
            .\Get-TeamsAttendanceReport.ps1 -UserPrincipalName $UserPrincipalName
            Write-Host "Report completed successfully"
            return
        } catch {
            Write-Warning "Attempt $i failed: $($_.Exception.Message)"
            if ($i -eq $MaxRetries) {
                throw "All retry attempts failed"
            }
            Start-Sleep -Seconds (30 * $i)  # Exponential backoff
        }
    }
}
```

### 2. Validation and Cleanup

```powershell
# Pre-execution validation
if (-not (Test-Path "config.json")) {
    throw "Configuration file not found. Please create config.json from template."
}

# Run the report
.\Get-TeamsAttendanceReport.ps1 -UserPrincipalName "user@company.com"

# Post-execution cleanup
Get-ChildItem -Path "logs" -Filter "*.log" | 
    Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-30)} |
    Remove-Item -Force

Get-ChildItem -Filter "debug_*.json" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -Skip 50 |
    Remove-Item -Force
```

### 3. Monitoring and Alerting

```powershell
# Report execution with monitoring
$startTime = Get-Date

try {
    .\Get-TeamsAttendanceReport.ps1 -UserPrincipalName "monitored@company.com"
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    # Log success metrics
    Write-Host "Report completed in $($duration.TotalMinutes) minutes"
    
    # Check output file size as quality metric
    $outputFile = Get-ChildItem -Filter "*Attendance*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($outputFile.Length -lt 1KB) {
        Write-Warning "Output file is suspiciously small - may indicate data issues"
    }
    
} catch {
    # Send alert for failures
    Write-Error "Report failed: $($_.Exception.Message)"
    # Integration with monitoring systems (Splunk, DataDog, etc.)
}
```

These examples demonstrate the flexibility and power of the Teams Attendance Report tool across various scenarios and environments. Adapt them to your specific needs and organizational requirements.