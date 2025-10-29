# Teams Attendance Report Tool v2.0

A comprehensive PowerShell tool for retrieving detailed attendance data from Microsoft Teams meetings, calls, webinars, and events using the Microsoft Graph API.

## Features

### Comprehensive Meeting Coverage
- **Scheduled Meetings**: Regular calendar-based Teams meetings
- **Instant Meetings**: Ad-hoc meetings started directly in Teams
- **1:1 Calls**: Direct calls between two users
- **Group Calls**: Multi-participant calls
- **Webinars**: Teams webinar events
- **Townhalls**: Large-scale broadcast events
- **Recurring Meetings**: Full support for meeting series

### Advanced Functionality
- **PowerShell Version Compatibility**: Works with both PowerShell 5.1 and 7+
- **Secure Configuration**: Sensitive credentials stored in separate config file
- **Comprehensive Logging**: Detailed logging of all API requests and responses
- **Duplicate Detection**: Automatic removal of duplicate attendance records
- **Proper DateTime Handling**: Fixes issues with System.Object time values
- **Rate Limiting**: Built-in throttling protection and retry logic
- **Debug Mode**: Optional raw API response saving for troubleshooting

### Data Quality Features
- **Deduplication**: Removes duplicate attendance entries automatically
- **DateTime Formatting**: Proper handling of various datetime formats
- **Error Handling**: Comprehensive error logging and failed meeting tracking
- **Data Validation**: Validates data integrity before export

## Prerequisites

### Required Software
- PowerShell 5.1 or higher (PowerShell 7+ recommended)
- Microsoft.Graph PowerShell SDK

### Required Azure App Registration
Create an Azure App Registration with the following **Application** permissions:
- `OnlineMeetings.Read.All`
- `OnlineMeetings.ReadWrite.All`
- `Calendars.Read.All`
- `CallRecords.Read.All`
- `User.Read.All`
- `Reports.Read.All`

### Additional Requirements
- **Application Access Policy** for Teams Calendar data
  - See: [Configure Application Access Policy](https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy#configure-application-access-policy)

## Installation

### 1. Install Microsoft Graph PowerShell SDK

```powershell
# Install for current user
Install-Module Microsoft.Graph -Scope CurrentUser

# Or install system-wide (requires admin)
Install-Module Microsoft.Graph -Scope AllUsers
```

### 2. Clone/Download Repository

```powershell
git clone <repository-url>
cd Get-MeetingAttendanceReport
```

### 3. Create Configuration File

Copy the template and add your credentials:

```powershell
Copy-Item config.template.json config.json
```

Edit `config.json` with your Azure app registration details:

```json
{
    "GraphConfiguration": {
        "ApplicationId": "your-app-id-here",
        "TenantId": "your-tenant-id-here",
        "ClientSecret": "your-client-secret-here"
    },
    "DefaultSettings": {
        "UserPrincipalName": "user@yourdomain.com",
        "AttendanceOutputPath": "TeamsAttendanceReport.csv",
        "FailedMeetingsOutputPath": "FailedMeetings.csv",
        "LogPath": "logs",
        "EnableDebug": false,
        "EnableVerboseLogging": true,
        "MaxConcurrentRequests": 5,
        "RequestDelayMs": 100
    }
}
```

## Usage

### Basic Usage

```powershell
# Run with default config settings
.\Get-TeamsAttendanceReport.ps1

# Specify user and date range
.\Get-TeamsAttendanceReport.ps1 -UserPrincipalName "user@domain.com" -StartDate "2025-01-01" -EndDate "2025-01-31"
```

### Advanced Usage

```powershell
# Enable debug mode with verbose logging
.\Get-TeamsAttendanceReport.ps1 -EnableDebug -LogLevel "Verbose"

# Specify meeting types to include
.\Get-TeamsAttendanceReport.ps1 -MeetingTypes @("Scheduled", "Instant", "OneOnOne")

# Use custom configuration file
.\Get-TeamsAttendanceReport.ps1 -ConfigPath ".\custom-config.json"

# Custom output file prefix
.\Get-TeamsAttendanceReport.ps1 -OutputPrefix "MonthlyReport"
```

### Parameter Reference

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| `ConfigPath` | string | Path to configuration file | `config.json` |
| `UserPrincipalName` | string | Target user's UPN | From config |
| `StartDate` | datetime | Start date for data retrieval | 30 days ago |
| `EndDate` | datetime | End date for data retrieval | Today |
| `OutputPrefix` | string | Prefix for output files | `TeamsReport` |
| `EnableDebug` | switch | Enable debug mode | `false` |
| `LogLevel` | string | Logging verbosity | `Information` |
| `MeetingTypes` | array | Meeting types to include | All types |

### Meeting Types

| Type | Description |
|------|-------------|
| `Scheduled` | Calendar-based meetings |
| `Instant` | Ad-hoc meetings |
| `OneOnOne` | Direct 1:1 calls |
| `Webinar` | Teams webinar events |
| `Townhall` | Large broadcast events |
| `Broadcast` | Live events |

## Output Files

### Attendance Report CSV
File: `{OutputPrefix}_Attendance_{timestamp}.csv`

Contains detailed attendance records with the following columns:
- `MeetingId`: Unique meeting identifier
- `Subject`: Meeting subject/title
- `MeetingStart`: Meeting start time
- `MeetingEnd`: Meeting end time
- `Organizer`: Meeting organizer email
- `AttendeeName`: Attendee display name
- `AttendeeEmail`: Attendee email address
- `AttendeeId`: Unique attendee identifier
- `Role`: Attendee role (Organizer, Presenter, Attendee)
- `JoinTime`: When attendee joined
- `LeaveTime`: When attendee left
- `DurationSeconds`: Attendance duration in seconds
- `DurationMinutes`: Attendance duration in minutes
- `MeetingType`: Type of meeting
- `Source`: Data source (Calendar, OnlineMeetings, etc.)
- `ProcessedAt`: When record was processed

### Failed Meetings Report CSV
File: `{OutputPrefix}_Failed_{timestamp}.csv`

Contains meetings that failed to process:
- `Subject`: Meeting subject
- `Start`: Meeting start time
- `Error`: Error description
- `Source`: Data source

### Log Files
Location: `logs/TeamsAttendanceReport_{timestamp}.log`

Contains detailed execution logs including:
- API request/response details
- Error messages and stack traces
- Performance metrics
- Processing statistics

## Configuration Options

### Graph Configuration
```json
"GraphConfiguration": {
    "ApplicationId": "your-app-id",
    "TenantId": "your-tenant-id", 
    "ClientSecret": "your-client-secret"
}
```

### Default Settings
```json
"DefaultSettings": {
    "UserPrincipalName": "default-user@domain.com",
    "AttendanceOutputPath": "TeamsAttendanceReport.csv",
    "FailedMeetingsOutputPath": "FailedMeetings.csv",
    "LogPath": "logs",
    "EnableDebug": false,
    "EnableVerboseLogging": true,
    "MaxConcurrentRequests": 5,
    "RequestDelayMs": 100
}
```

### Meeting Type Filters
```json
"MeetingTypes": {
    "IncludeScheduledMeetings": true,
    "IncludeInstantMeetings": true,
    "IncludeOneOnOneCalls": true,
    "IncludeWebinars": true,
    "IncludeTownhalls": true,
    "IncludeBroadcastEvents": true,
    "IncludeRecurringMeetings": true
}
```

### Data Retention
```json
"DataRetention": {
    "MaxLogAgeDays": 30,
    "MaxDebugFilesCount": 100,
    "CompressOldLogs": true
}
```

## Troubleshooting

### Common Issues

#### Authentication Errors
```
Error: Failed to connect to Microsoft Graph
```
**Solution**: Verify your app registration credentials in `config.json` and ensure proper permissions are granted.

#### Permission Errors
```
Error: Insufficient privileges to complete the operation
```
**Solution**: Ensure your Azure app has the required Application permissions and admin consent has been granted.

#### No Attendance Data
```
Warning: No attendance reports found for meeting
```
**Solution**: Attendance data is only available for meetings that have ended and had participants join.

#### Rate Limiting
```
Warning: Throttling detected
```
**Solution**: The script automatically handles throttling. Increase `RequestDelayMs` in config for proactive rate limiting.

### Debug Mode

Enable debug mode for detailed troubleshooting:

```powershell
.\Get-TeamsAttendanceReport.ps1 -EnableDebug -LogLevel "Debug"
```

This will:
- Save raw API responses to JSON files
- Provide detailed request/response logging
- Include performance metrics
- Show internal processing details

### Log Analysis

Check the log files in the `logs` directory for detailed execution information:

```powershell
# View latest log file
Get-ChildItem logs\*.log | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Get-Content
```

## Security Considerations

### Configuration Security
- Never commit `config.json` to version control
- Use `.gitignore` to exclude sensitive files
- Store client secrets securely
- Regularly rotate application secrets

### Permissions
- Use least-privilege principle
- Only grant necessary Graph API permissions
- Implement application access policies
- Monitor application usage

### Data Protection
- Log files may contain sensitive information
- Implement log retention policies
- Secure output files appropriately
- Consider data residency requirements

## Performance Optimization

### Rate Limiting
- Adjust `RequestDelayMs` based on tenant size
- Monitor for throttling responses
- Use `MaxConcurrentRequests` to control parallelism

### Large Datasets
- Process data in smaller date ranges
- Use pagination for large result sets
- Monitor memory usage for extensive reports

### Network Optimization
- Run from network location close to Microsoft 365
- Monitor API response times
- Implement connection pooling if available

## Version 2.0 Changes

### New Features
- **Multi-Version Support**: Compatible with PowerShell 5.1 and 7+
- **Enhanced Meeting Types**: Support for webinars, townhalls, and 1:1 calls
- **Comprehensive Logging**: Detailed request/response logging
- **Duplicate Removal**: Automatic deduplication of attendance records
- **DateTime Handling**: Proper formatting of time values
- **Configuration Management**: Secure credential storage
- **Rate Limiting**: Built-in throttling protection

### Breaking Changes
- Configuration moved to external JSON file
- Output file naming includes timestamps
- New required Graph API permissions
- Updated parameter structure

### Migration from v1.x
1. Create `config.json` from template
2. Update Azure app permissions
3. Update script invocation parameters
4. Test with small date range first

## Additional Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Teams Online Meetings API](https://docs.microsoft.com/en-us/graph/api/resources/onlinemeeting)
- [Application Access Policies](https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy)
- [PowerShell Module Documentation](https://docs.microsoft.com/en-us/powershell/microsoftgraph/)

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review log files for detailed error information
3. Search existing issues
4. Create a new issue with detailed information and log excerpts
