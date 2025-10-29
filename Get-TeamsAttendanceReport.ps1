<#
.SYNOPSIS
    Comprehensive Microsoft Teams meeting attendance data retrieval tool using Microsoft Graph SDK.
    
.DESCRIPTION
    This script retrieves detailed attendance data from various types of Microsoft Teams meetings,
    calls, webinars, and events. It supports both PowerShell 5.1 and 7+ with automatic version
    detection and handles different authentication methods securely.

.PARAMETER ConfigPath
    Path to the configuration JSON file. Defaults to 'config.json' in the script directory.

.PARAMETER UserPrincipalName
    The UPN of the user whose meetings you want to analyze. Overrides config file setting.

.PARAMETER StartDate
    The start date for the data retrieval period. Format: YYYY-MM-DDTHH:MM:SSZ

.PARAMETER EndDate
    The end date for the data retrieval period. Format: YYYY-MM-DDTHH:MM:SSZ

.PARAMETER OutputPrefix
    Prefix for output files. Timestamp will be appended automatically.

.PARAMETER EnableDebug
    Enable debug mode with detailed logging and raw API response saving.

.PARAMETER LogLevel
    Logging level: 'Error', 'Warning', 'Information', 'Verbose', 'Debug'

.PARAMETER MeetingTypes
    Array of meeting types to include: 'Scheduled', 'Instant', 'OneOnOne', 'Webinar', 'Townhall', 'Broadcast'

.EXAMPLE
    .\Get-TeamsAttendanceReport.ps1 -UserPrincipalName "user@domain.com" -StartDate "2025-01-01T00:00:00Z" -EndDate "2025-01-31T23:59:59Z"

.EXAMPLE
    .\Get-TeamsAttendanceReport.ps1 -ConfigPath ".\custom-config.json" -EnableDebug -LogLevel "Verbose"

.NOTES
    Author: Dylan Stetts
    Version: 2.0.0
    Requires: Microsoft.Graph PowerShell SDK
    
    Required Graph API Permissions:
    - OnlineMeetings.Read.All
    - OnlineMeetings.ReadWrite.All  
    - Calendars.Read.All
    - CallRecords.Read.All
    - User.Read.All
    - Reports.Read.All
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    
    [Parameter(Mandatory = $false)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory = $false)]
    [datetime]$StartDate,
    
    [Parameter(Mandatory = $false)]
    [datetime]$EndDate,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPrefix = "TeamsReport",
    
    [Parameter(Mandatory = $false)]
    [switch]$EnableDebug,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('Error', 'Warning', 'Information', 'Verbose', 'Debug')]
    [string]$LogLevel = 'Information',
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('Scheduled', 'Instant', 'OneOnOne', 'Webinar', 'Townhall', 'Broadcast')]
    [string[]]$MeetingTypes = @('Scheduled', 'Instant', 'OneOnOne', 'Webinar', 'Townhall')
)

# =========================
# Global Variables & Configuration
# =========================
$script:Config = $null
$script:LogPath = $null
$script:PowerShellVersion = $PSVersionTable.PSVersion.Major
$script:AttendanceData = [System.Collections.ArrayList]::new()
$script:FailedMeetings = [System.Collections.ArrayList]::new()
$script:ProcessedMeetingIds = [System.Collections.Generic.HashSet[string]]::new()
$script:RequestCount = 0
$script:StartTime = Get-Date

# =========================
# Configuration Management
# =========================
function Initialize-Configuration {
    [CmdletBinding()]
    param()
    
    try {
        if (-not (Test-Path $ConfigPath)) {
            throw "Configuration file not found at: $ConfigPath. Please create config.json from config.template.json"
        }
        
        $configContent = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        $script:Config = $configContent
        
        # Override config with script parameters if provided
        if ($UserPrincipalName) { $script:Config.DefaultSettings.UserPrincipalName = $UserPrincipalName }
        if ($EnableDebug) { $script:Config.DefaultSettings.EnableDebug = $true }
        
        # Set default dates if not provided
        if (-not $StartDate) { 
            $StartDate = (Get-Date).AddDays(-30)
        }
        if (-not $EndDate) { 
            $EndDate = Get-Date
        }
        
        # Add date properties to config (using Add-Member to create new properties)
        $script:Config.DefaultSettings | Add-Member -NotePropertyName "StartDate" -NotePropertyValue $StartDate -Force
        $script:Config.DefaultSettings | Add-Member -NotePropertyName "EndDate" -NotePropertyValue $EndDate -Force
        
        # Initialize logging
        Initialize-Logging
        
        Write-LogMessage "Configuration loaded successfully" "Information"
        Write-LogMessage "PowerShell Version: $script:PowerShellVersion" "Information"
        Write-LogMessage "Date Range: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))" "Information"
        
    } catch {
        Write-Error "Failed to initialize configuration: $($_.Exception.Message)"
        exit 1
    }
}

function Initialize-Logging {
    [CmdletBinding()]
    param()
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $script:LogPath = Join-Path $script:Config.DefaultSettings.LogPath "TeamsAttendanceReport_$timestamp.log"
    
    # Create log directory if it doesn't exist
    $logDir = Split-Path $script:LogPath -Parent
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    # Clean up old log files
    Invoke-LogCleanup
}

function Write-LogMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Error', 'Warning', 'Information', 'Verbose', 'Debug')]
        [string]$Level = 'Information',
        
        [Parameter(Mandatory = $false)]
        [hashtable]$AdditionalData = @{}
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Add additional data if provided
    if ($AdditionalData.Count -gt 0) {
        $additionalJson = $AdditionalData | ConvertTo-Json -Compress
        $logEntry += " | Data: $additionalJson"
    }
    
    # Write to console based on level
    switch ($Level) {
        'Error' { Write-Error $Message }
        'Warning' { Write-Warning $Message }
        'Information' { Write-Host $Message -ForegroundColor Green }
        'Verbose' { if ($VerbosePreference -ne 'SilentlyContinue') { Write-Host $Message -ForegroundColor Cyan } }
        'Debug' { if ($DebugPreference -ne 'SilentlyContinue') { Write-Host $Message -ForegroundColor Magenta } }
    }
    
    # Write to log file
    if ($script:LogPath) {
        Add-Content -Path $script:LogPath -Value $logEntry -Encoding UTF8
    }
}

function Invoke-LogCleanup {
    [CmdletBinding()]
    param()
    
    try {
        $logDir = Split-Path $script:LogPath -Parent
        $maxAge = $script:Config.DataRetention.MaxLogAgeDays
        $cutoffDate = (Get-Date).AddDays(-$maxAge)
        
        Get-ChildItem -Path $logDir -Filter "*.log" | 
            Where-Object { $_.LastWriteTime -lt $cutoffDate } |
            Remove-Item -Force
            
        Write-LogMessage "Log cleanup completed" "Verbose"
    } catch {
        Write-LogMessage "Log cleanup failed: $($_.Exception.Message)" "Warning"
    }
}

# =========================
# Graph Authentication
# =========================
function Connect-ToMicrosoftGraph {
    [CmdletBinding()]
    param()
    
    try {
        Write-LogMessage "Connecting to Microsoft Graph..." "Information"
        
        # Check if Microsoft.Graph module is available
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
            throw "Microsoft.Graph module not found. Please install with: Install-Module Microsoft.Graph -Scope CurrentUser"
        }
        
        # Import required modules based on PowerShell version
        if ($script:PowerShellVersion -ge 7) {
            Import-Module Microsoft.Graph.Authentication -Force
            Import-Module Microsoft.Graph.Users -Force
            Import-Module Microsoft.Graph.Calendar -Force
            Import-Module Microsoft.Graph.CloudCommunications -Force
        } else {
            # PowerShell 5.1 compatibility
            Import-Module Microsoft.Graph.Authentication -Force
            Import-Module Microsoft.Graph.Users -Force
            Import-Module Microsoft.Graph.Calendar -Force
            if (Get-Module -ListAvailable -Name Microsoft.Graph.CloudCommunications) {
                Import-Module Microsoft.Graph.CloudCommunications -Force
            }
        }
        
        # Create credential object
        $secureSecret = ConvertTo-SecureString -String $script:Config.GraphConfiguration.ClientSecret -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($script:Config.GraphConfiguration.ApplicationId, $secureSecret)
        
        # Connect to Graph
        $connectParams = @{
            TenantId = $script:Config.GraphConfiguration.TenantId
            ClientSecretCredential = $credential
            NoWelcome = $true
        }
        
        Connect-MgGraph @connectParams
        
        # Verify connection
        $context = Get-MgContext
        if (-not $context) {
            throw "Failed to establish Graph connection"
        }
        
        Write-LogMessage "Successfully connected to Microsoft Graph" "Information" @{
            TenantId = $context.TenantId
            AppId = $context.AppId
            Scopes = $context.Scopes -join ', '
        }
        
    } catch {
        Write-LogMessage "Failed to connect to Microsoft Graph: $($_.Exception.Message)" "Error"
        throw
    }
}

# =========================
# User and Meeting Discovery
# =========================
function Get-UserInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName
    )
    
    try {
        Write-LogMessage "Retrieving user information for: $UserPrincipalName" "Information"
        
        $user = Get-MgUser -UserId $UserPrincipalName -Property "Id,UserPrincipalName,DisplayName,Mail"
        
        if (-not $user) {
            throw "User not found: $UserPrincipalName"
        }
        
        Write-LogMessage "User found: $($user.DisplayName)" "Information" @{
            UserId = $user.Id
            Mail = $user.Mail
        }
        
        return $user
    } catch {
        Write-LogMessage "Failed to retrieve user information: $($_.Exception.Message)" "Error"
        throw
    }
}

function Get-ComprehensiveMeetingsData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [datetime]$StartDate,
        
        [Parameter(Mandatory = $true)]
        [datetime]$EndDate
    )
    
    try {
        Write-LogMessage "Retrieving comprehensive meetings data" "Information"
        
        $allMeetings = [System.Collections.ArrayList]::new()
        
        # Method 1: Calendar Events (Scheduled meetings)
        if ('Scheduled' -in $MeetingTypes) {
            $calendarMeetings = Get-CalendarMeetings -UserId $UserId -StartDate $StartDate -EndDate $EndDate
            if ($calendarMeetings) {
                # Ensure we have an array/collection
                if ($calendarMeetings -is [System.Collections.IEnumerable] -and $calendarMeetings -isnot [string]) {
                    foreach ($meeting in $calendarMeetings) {
                        $allMeetings.Add($meeting) | Out-Null
                    }
                } else {
                    $allMeetings.Add($calendarMeetings) | Out-Null
                }
            }
        }
        
        # Method 2: Online Meetings (All online meetings)
        if (('Instant' -in $MeetingTypes) -or ('OneOnOne' -in $MeetingTypes) -or ('Webinar' -in $MeetingTypes)) {
            $onlineMeetings = Get-OnlineMeetingsData -UserId $UserId -StartDate $StartDate -EndDate $EndDate
            if ($onlineMeetings) {
                # Ensure we have an array/collection
                if ($onlineMeetings -is [System.Collections.IEnumerable] -and $onlineMeetings -isnot [string]) {
                    foreach ($meeting in $onlineMeetings) {
                        $allMeetings.Add($meeting) | Out-Null
                    }
                } else {
                    $allMeetings.Add($onlineMeetings) | Out-Null
                }
            }
        }
        
        # Method 3: Direct calls and chat-based meetings
        if ('OneOnOne' -in $MeetingTypes) {
            $chatMeetings = Get-ChatBasedMeetings -UserId $UserId -StartDate $StartDate -EndDate $EndDate
            if ($chatMeetings) {
                # Ensure we have an array/collection
                if ($chatMeetings -is [System.Collections.IEnumerable] -and $chatMeetings -isnot [string]) {
                    foreach ($meeting in $chatMeetings) {
                        $allMeetings.Add($meeting) | Out-Null
                    }
                } else {
                    $allMeetings.Add($chatMeetings) | Out-Null
                }
            }
        }
        
        # Method 4: Call Records (1:1 calls and group calls)
        if ('OneOnOne' -in $MeetingTypes) {
            $callRecords = Get-CallRecordsData -UserId $UserId -StartDate $StartDate -EndDate $EndDate
            if ($callRecords) {
                # Ensure we have an array/collection
                if ($callRecords -is [System.Collections.IEnumerable] -and $callRecords -isnot [string]) {
                    foreach ($meeting in $callRecords) {
                        $allMeetings.Add($meeting) | Out-Null
                    }
                } else {
                    $allMeetings.Add($callRecords) | Out-Null
                }
            }
        }
        
        # Method 5: Broadcast Events (Townhalls, Live Events)
        if (('Townhall' -in $MeetingTypes) -or ('Broadcast' -in $MeetingTypes)) {
            $broadcastEvents = Get-BroadcastEventsData -UserId $UserId -StartDate $StartDate -EndDate $EndDate
            if ($broadcastEvents) {
                # Ensure we have an array/collection
                if ($broadcastEvents -is [System.Collections.IEnumerable] -and $broadcastEvents -isnot [string]) {
                    foreach ($meeting in $broadcastEvents) {
                        $allMeetings.Add($meeting) | Out-Null
                    }
                } else {
                    $allMeetings.Add($broadcastEvents) | Out-Null
                }
            }
        }
        
        Write-LogMessage "Retrieved $($allMeetings.Count) total meetings across all methods" "Information"
        
        return $allMeetings
    } catch {
        Write-LogMessage "Failed to retrieve comprehensive meetings data: $($_.Exception.Message)" "Error"
        throw
    }
}

function Get-MeetingTypeClassification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$CalendarEvent
    )
    
    $meetingType = "Scheduled"
    $classification = @{
        Type = "Scheduled"
        IsOneOnOne = $false
        IsInstant = $false
        IsRecurring = $false
        Category = "Regular Meeting"
    }
    
    try {
        # Check for 1:1 meetings
        if ($CalendarEvent.attendees -and $CalendarEvent.attendees.Count -le 2) {
            $classification.IsOneOnOne = $true
            $classification.Category = "One-on-One"
            $meetingType = "OneOnOne"
        }
        
        # Check for instant/ad-hoc meetings (short notice meetings)
        if ($CalendarEvent.start -and $CalendarEvent.organizer) {
            $meetingStart = [DateTime]::Parse($CalendarEvent.start.dateTime)
            $now = Get-Date
            
            # If meeting was created very close to start time, it might be instant
            # This is a heuristic since we don't have creation timestamp
            $daysDifference = ($meetingStart - $now).TotalDays
            if ([Math]::Abs($daysDifference) -lt 1) {
                $classification.IsInstant = $true
                $classification.Category = "Instant Meeting"
                $meetingType = "Instant"
            }
        }
        
        # Check for recurring meetings
        if ($CalendarEvent.type -eq "seriesMaster" -or $CalendarEvent.type -eq "occurrence") {
            $classification.IsRecurring = $true
            $classification.Category = "Recurring Meeting"
        }
        
        # Check for webinar-like meetings (large attendee count)
        if ($CalendarEvent.attendees -and $CalendarEvent.attendees.Count -gt 20) {
            $classification.Category = "Large Meeting/Webinar"
            $meetingType = "Webinar"
        }
        
        # Check for broadcast events (no attendees list, specific subject patterns)
        if ($CalendarEvent.subject) {
            $subject = $CalendarEvent.subject.ToLower()
            if ($subject -like "*townhall*" -or $subject -like "*all hands*" -or 
                $subject -like "*broadcast*" -or $subject -like "*live event*") {
                $classification.Category = "Broadcast Event"
                $meetingType = "Townhall"
            }
        }
        
        $classification.Type = $meetingType
        return $classification
        
    } catch {
        Write-LogMessage "Failed to classify meeting type: $($_.Exception.Message)" "Warning" @{
            Subject = $CalendarEvent.subject
        }
        return $classification
    }
}

function Get-CalendarMeetings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [datetime]$StartDate,
        
        [Parameter(Mandatory = $true)]
        [datetime]$EndDate
    )
    
    try {
        Write-LogMessage "Retrieving calendar meetings" "Verbose"
        
        $meetings = [System.Collections.ArrayList]::new()
        $pageSize = 100
        $skip = 0
        
        do {
            $requestUri = "https://graph.microsoft.com/v1.0/users/$UserId/calendarView" +
                          "?startDateTime=$($StartDate.ToString('yyyy-MM-ddTHH:mm:ss.fffZ'))" +
                          "&endDateTime=$($EndDate.ToString('yyyy-MM-ddTHH:mm:ss.fffZ'))" +
                          "&`$top=$pageSize&`$skip=$skip" +
                          "&`$select=id,subject,start,end,isOnlineMeeting,onlineMeetingProvider,onlineMeeting,type,organizer,attendees"
            
            $response = Invoke-GraphRequestWithRetry -Uri $requestUri -Method "GET"
            
            if ($response.value) {
                $teamsEvents = $response.value | Where-Object {
                    $_.isOnlineMeeting -eq $true -and 
                    ($_.onlineMeetingProvider -eq "teamsForBusiness" -or $_.onlineMeetingProvider -eq "teams")
                }
                
                foreach ($event in $teamsEvents) {
                    # Classify the meeting type
                    $classification = Get-MeetingTypeClassification -CalendarEvent $event
                    
                    $meetings.Add([PSCustomObject]@{
                        Id = $event.id
                        Subject = $event.subject
                        Start = $event.start.dateTime
                        End = $event.end.dateTime
                        Type = "Calendar"
                        MeetingType = $classification.Type
                        Category = $classification.Category
                        IsOneOnOne = $classification.IsOneOnOne
                        IsInstant = $classification.IsInstant
                        IsRecurring = $classification.IsRecurring
                        JoinUrl = $event.onlineMeeting.joinUrl
                        Organizer = $event.organizer.emailAddress.address
                        OnlineMeetingId = $null
                        Source = "Calendar"
                        AttendeeCount = if ($event.attendees) { $event.attendees.Count } else { 0 }
                    }) | Out-Null
                }
            }
            
            $skip += $pageSize
        } while ($response.value -and $response.value.Count -eq $pageSize)
        
        Write-LogMessage "Found $($meetings.Count) calendar meetings" "Verbose"
        return $meetings
    } catch {
        Write-LogMessage "Failed to retrieve calendar meetings: $($_.Exception.Message)" "Warning"
        return @()
    }
}

function Get-OnlineMeetingsData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [datetime]$StartDate,
        
        [Parameter(Mandatory = $true)]
        [datetime]$EndDate
    )
    
    try {
        Write-LogMessage "Retrieving online meetings data (including ad-hoc meetings)" "Verbose"
        
        $meetings = [System.Collections.ArrayList]::new()
        
        # Method 1: Note that listing all online meetings for a user is not supported
        # The online meetings API requires specific meeting IDs or join URLs
        try {
            Write-LogMessage "Direct online meetings API not accessible: API requires specific meeting IDs or join URLs" "Verbose"
            # Commenting out the direct API call since it's not supported
            # $requestUri = "https://graph.microsoft.com/v1.0/users/$UserId/onlineMeetings"
            # $response = Invoke-GraphRequestWithRetry -Uri $requestUri -Method "GET"
        } catch {
            Write-LogMessage "Direct online meetings API not accessible: $($_.Exception.Message)" "Verbose"
        }
        
        # Method 2: Try to get meetings from Teams activity (chat-based calls)
        try {
            Write-LogMessage "Attempting to retrieve Teams chat-based meetings" "Verbose"
            
            # Get chats where the user participated
            $chatsUri = "https://graph.microsoft.com/v1.0/users/$UserId/chats" +
                       "?`$expand=members" +
                       "&`$select=id,topic,chatType,createdDateTime,lastUpdatedDateTime"
            
            $chatsResponse = Invoke-GraphRequestWithRetry -Uri $chatsUri -Method "GET"
            
            if ($chatsResponse.value) {
                foreach ($chat in $chatsResponse.value) {
                    if ($chat.lastUpdatedDateTime) {
                        $chatDate = [DateTime]::Parse($chat.lastUpdatedDateTime)
                        if ($chatDate -ge $StartDate -and $chatDate -le $EndDate) {
                            # Check if this chat had any calls/meetings
                            try {
                                # Use simple request without filter and select to avoid BadRequest errors, but limit with $top
                                $messagesUri = "https://graph.microsoft.com/v1.0/users/$UserId/chats/$($chat.id)/messages?`$top=50"
                                
                                $messagesResponse = Invoke-GraphRequestWithRetry -Uri $messagesUri -Method "GET"
                                
                                if ($messagesResponse.value) {
                                    # Client-side filtering for system event messages
                                    $systemEventMessages = $messagesResponse.value | Where-Object { 
                                        $_.messageType -eq 'systemEventMessage' 
                                    }
                                    
                                    foreach ($message in $systemEventMessages) {
                                        if ($message.eventDetail -and 
                                            ($message.eventDetail.'@odata.type' -like "*callStarted*" -or 
                                             $message.eventDetail.'@odata.type' -like "*callEnded*")) {
                                            
                                            $meetings.Add([PSCustomObject]@{
                                                Id = "chat-call-$($chat.id)-$($message.id)"
                                                Subject = if ($chat.topic) { "Chat: $($chat.topic)" } else { "Chat Call" }
                                                Start = $message.createdDateTime
                                                End = $null
                                                Type = "ChatCall"
                                                MeetingType = "call"
                                                JoinUrl = $null
                                                Organizer = $null
                                                OnlineMeetingId = $null
                                                Source = "TeamsChats"
                                            }) | Out-Null
                                        }
                                    }
                                }
                            } catch {
                                Write-LogMessage "Could not retrieve messages for chat $($chat.id): $($_.Exception.Message)" "Debug"
                            }
                        }
                    }
                }
                Write-LogMessage "Processed $($chatsResponse.value.Count) chats for call activities" "Verbose"
            }
        } catch {
            Write-LogMessage "Teams chats API not accessible: $($_.Exception.Message)" "Verbose"
        }
        
        # Method 3: Try to get meetings from user's presence/activity logs
        try {
            Write-LogMessage "Attempting to retrieve user presence-based meetings" "Verbose"
            
            # Get user's Teams activity/presence changes that might indicate calls
            $activityUri = "https://graph.microsoft.com/v1.0/users/$UserId/presence"
            $presenceResponse = Invoke-GraphRequestWithRetry -Uri $activityUri -Method "GET"
            
            # Note: This gives current presence, not historical data
            # For historical presence data, we'd need different approaches
            Write-LogMessage "Current presence retrieved (historical presence data requires different APIs)" "Debug"
            
        } catch {
            Write-LogMessage "Presence API not accessible: $($_.Exception.Message)" "Verbose"
        }
        
        Write-LogMessage "Found $($meetings.Count) total online/chat meetings" "Verbose"
        return $meetings
    } catch {
        Write-LogMessage "Failed to retrieve online meetings: $($_.Exception.Message)" "Warning"
        return @()
    }
}

function Get-ChatBasedMeetings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [datetime]$StartDate,
        
        [Parameter(Mandatory = $true)]
        [datetime]$EndDate
    )
    
    try {
        Write-LogMessage "Retrieving chat-based meetings and direct calls" "Verbose"
        
        $chatMeetings = [System.Collections.ArrayList]::new()
        
        # Method 1: Get user's chats and look for call activities
        try {
            Write-LogMessage "Analyzing Teams chats for call activities" "Verbose"
            
            # Get user's chats
            $chatsUri = "https://graph.microsoft.com/v1.0/users/$UserId/chats" +
                       "?`$select=id,topic,chatType,createdDateTime,lastUpdatedDateTime" +
                       "&`$top=50"
            
            $chatsResponse = Invoke-GraphRequestWithRetry -Uri $chatsUri -Method "GET"
            
            if ($chatsResponse.value) {
                foreach ($chat in $chatsResponse.value) {
                    # Check if chat was active in our date range
                    $chatUpdated = [DateTime]::Parse($chat.lastUpdatedDateTime)
                    if ($chatUpdated -ge $StartDate -and $chatUpdated -le $EndDate) {
                        
                        # Get chat members to determine if it's a direct call
                        $membersUri = "https://graph.microsoft.com/v1.0/users/$UserId/chats/$($chat.id)/members"
                        try {
                            $membersResponse = Invoke-GraphRequestWithRetry -Uri $membersUri -Method "GET"
                            $memberCount = if ($membersResponse.value) { $membersResponse.value.Count } else { 0 }
                            
                            # Look for call-related messages in this chat
                            # Use simple request without select to avoid BadRequest errors, but limit with $top
                            $messagesUri = "https://graph.microsoft.com/v1.0/users/$UserId/chats/$($chat.id)/messages?`$top=50"
                            
                            $messagesResponse = Invoke-GraphRequestWithRetry -Uri $messagesUri -Method "GET"
                            
                            if ($messagesResponse.value) {
                                # Client-side filtering for messages in our date range
                                $relevantMessages = $messagesResponse.value | Where-Object {
                                    $messageDate = [DateTime]::Parse($_.createdDateTime)
                                    $messageDate -ge $StartDate -and $messageDate -le $EndDate
                                }
                                
                                foreach ($message in $relevantMessages) {
                                    
                                    # Check for call-related system messages
                                    if ($message.messageType -eq "systemEventMessage" -and $message.eventDetail) {
                                        $eventType = $message.eventDetail.'@odata.type'
                                        if ($eventType -like "*callStarted*" -or 
                                            $eventType -like "*callEnded*" -or
                                            $eventType -like "*call*") {
                                            
                                            $callType = if ($memberCount -eq 2) { "Direct Call" } else { "Group Call" }
                                            $subject = if ($chat.topic) { "Chat Call: $($chat.topic)" } else { $callType }
                                            
                                            $chatMeetings.Add([PSCustomObject]@{
                                                Id = "chat-call-$($chat.id)-$($message.id)"
                                                Subject = $subject
                                                Start = $message.createdDateTime
                                                End = $null
                                                Type = "ChatCall"
                                                MeetingType = $callType.ToLower().Replace(" ", "")
                                                JoinUrl = $null
                                                Organizer = $null
                                                OnlineMeetingId = $null
                                                Source = "TeamsChatCalls"
                                            }) | Out-Null
                                        }
                                    }
                                    
                                    # Also check for call-related content in regular messages
                                    if ($message.body -and $message.body.content) {
                                        $content = $message.body.content.ToLower()
                                        if ($content -like "*started a call*" -or 
                                            $content -like "*ended a call*" -or
                                            $content -like "*joined the call*" -or
                                            $content -like "*left the call*") {
                                            
                                            $callType = if ($memberCount -eq 2) { "Direct Call" } else { "Group Call" }
                                            $subject = if ($chat.topic) { "Chat Activity: $($chat.topic)" } else { "$callType Activity" }
                                            
                                            $chatMeetings.Add([PSCustomObject]@{
                                                Id = "chat-activity-$($chat.id)-$($message.id)"
                                                Subject = $subject
                                                Start = $message.createdDateTime
                                                End = $null
                                                Type = "ChatActivity"
                                                MeetingType = "call"
                                                JoinUrl = $null
                                                Organizer = $null
                                                OnlineMeetingId = $null
                                                Source = "TeamsChatActivity"
                                            }) | Out-Null
                                        }
                                    }
                                }
                            }
                        } catch {
                            Write-LogMessage "Could not analyze chat $($chat.id): $($_.Exception.Message)" "Debug"
                        }
                    }
                }
                
                Write-LogMessage "Analyzed $($chatsResponse.value.Count) chats, found $($chatMeetings.Count) chat-based meetings" "Verbose"
            }
        } catch {
            Write-LogMessage "Could not retrieve Teams chats: $($_.Exception.Message)" "Warning"
        }
        
        # Method 2: Look for Teams notifications/activities (alternative approach)
        try {
            Write-LogMessage "Attempting to retrieve Teams activity notifications" "Verbose"
            
            # Get user notifications which might include call activities
            $notificationsUri = "https://graph.microsoft.com/v1.0/users/$UserId/teamwork/installedApps" +
                               "?`$expand=teamsAppDefinition"
            
            # This is more of an experimental approach
            # $notificationsResponse = Invoke-GraphRequestWithRetry -Uri $notificationsUri -Method "GET"
            
            Write-LogMessage "Teams notifications approach not fully implemented" "Debug"
            
        } catch {
            Write-LogMessage "Teams notifications API not accessible: $($_.Exception.Message)" "Debug"
        }
        
        Write-LogMessage "Found $($chatMeetings.Count) chat-based meetings/calls" "Verbose"
        return $chatMeetings
        
    } catch {
        Write-LogMessage "Failed to retrieve chat-based meetings: $($_.Exception.Message)" "Warning"
        return @()
    }
}

function Get-CallRecordsData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [datetime]$StartDate,
        
        [Parameter(Mandatory = $true)]
        [datetime]$EndDate
    )
    
    try {
        Write-LogMessage "Retrieving call records data (including direct calls and chat calls)" "Verbose"
        
        $calls = [System.Collections.ArrayList]::new()
        
        # Method 1: Try Call Records API (requires CallRecords.Read.All permission)
        try {
            Write-LogMessage "Attempting to access Call Records API" "Verbose"
            
            # Get call records without complex filters to avoid BadRequest errors
            # Note: This API may still require special permissions
            $requestUri = "https://graph.microsoft.com/v1.0/communications/callRecords"
            
            $response = Invoke-GraphRequestWithRetry -Uri $requestUri -Method "GET"
            
            if ($response.value) {
                # Client-side filtering for date range
                $relevantCalls = $response.value | Where-Object {
                    $callStart = [DateTime]::Parse($_.startDateTime)
                    $callStart -ge $StartDate -and $callStart -le $EndDate
                }
                
                foreach ($call in $relevantCalls) {
                    # Process call records that fall within our date range
                    $userInvolved = $false
                    $isDirectCall = $false
                    $callType = "unknown"
                    $participantCount = 0
                    
                    # Try to get sessions for this call (separate API call)
                    try {
                        $sessionsUri = "https://graph.microsoft.com/v1.0/communications/callRecords/$($call.id)/sessions"
                        $sessionsResponse = Invoke-GraphRequestWithRetry -Uri $sessionsUri -Method "GET"
                        
                        if ($sessionsResponse.value) {
                            # Check sessions to see if user was involved
                            foreach ($session in $sessionsResponse.value) {
                                # Check if user was involved in this session
                                if ($session.caller -and $session.caller.user -and $session.caller.user.id -eq $UserId) {
                                    $userInvolved = $true
                                }
                                if ($session.callee -and $session.callee.user -and $session.callee.user.id -eq $UserId) {
                                    $userInvolved = $true
                                }
                            }
                            
                            # Estimate participant count from sessions
                            $participantCount = $sessionsResponse.value.Count
                        }
                    } catch {
                        # If sessions API fails, assume user was involved since the call record exists in their tenant
                        $userInvolved = $true
                        $participantCount = 1
                        Write-LogMessage "Could not retrieve sessions for call $($call.id): $($_.Exception.Message)" "Debug"
                    }
                    
                    # Determine call type based on session count
                    if ($participantCount -eq 1) {
                        $callType = "Direct Call"
                        $isDirectCall = $true
                    } elseif ($participantCount -gt 1) {
                        $callType = "Group Call"
                    }
                    
                    # Only include calls where the user was involved
                    if ($userInvolved) {
                        $calls.Add([PSCustomObject]@{
                            Id = $call.id
                            Subject = if ($isDirectCall) { "Direct Call" } else { "Teams Call ($participantCount sessions)" }
                            Start = $call.startDateTime
                            End = $call.endDateTime
                            Type = "Call"
                            MeetingType = $callType
                            JoinUrl = $null
                            Organizer = $null
                            OnlineMeetingId = $null
                            Source = "CallRecords"
                        }) | Out-Null
                    }
                }
                Write-LogMessage "Found $($calls.Count) call records from Call Records API" "Verbose"
            }
        } catch {
            Write-LogMessage "Call Records API not accessible (may require additional permissions): $($_.Exception.Message)" "Warning"
        }
        
        # Method 2: Alternative approach - Get user's Teams activities from audit logs (if available)
        try {
            Write-LogMessage "Attempting to retrieve Teams activities from audit logs" "Verbose"
            
            # This would require Security.Read.All or AuditLog.Read.All permissions
            # Note: This is an alternative approach that might capture more call data
            $auditUri = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits" +
                       "?`$filter=activityDisplayName eq 'Started call' or activityDisplayName eq 'Ended call'" +
                       "&`$select=id,activityDateTime,activityDisplayName,initiatedBy,targetResources"
            
            # This API might not be available or might require special permissions
            # Commenting out for now as it's experimental
            # $auditResponse = Invoke-GraphRequestWithRetry -Uri $auditUri -Method "GET"
            
            Write-LogMessage "Audit logs approach not implemented (requires additional permissions)" "Debug"
            
        } catch {
            Write-LogMessage "Audit logs API not accessible: $($_.Exception.Message)" "Verbose"
        }
        
        # Method 3: Try to get user's Teams device usage (might indicate call activity)
        try {
            Write-LogMessage "Attempting to retrieve Teams device usage reports" "Verbose"
            
            # Get Teams device usage reports which might indicate call activity
            $reportsUri = "https://graph.microsoft.com/v1.0/reports/getTeamsDeviceUsageUserDetail(period='D30')"
            
            # This might give us insights into user activity but not specific call records
            # $reportsResponse = Invoke-GraphRequestWithRetry -Uri $reportsUri -Method "GET"
            
            Write-LogMessage "Device usage reports approach not fully implemented" "Debug"
            
        } catch {
            Write-LogMessage "Teams reports API not accessible: $($_.Exception.Message)" "Verbose"
        }
        
        Write-LogMessage "Found $($calls.Count) call records" "Verbose"
        return $calls
    } catch {
        Write-LogMessage "Failed to retrieve call records: $($_.Exception.Message)" "Warning"
        return @()
    }
}

function Get-BroadcastEventsData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [datetime]$StartDate,
        
        [Parameter(Mandatory = $true)]
        [datetime]$EndDate
    )
    
    try {
        Write-LogMessage "Retrieving broadcast events data" "Verbose"
        
        $events = [System.Collections.ArrayList]::new()
        
        # Note: This is a placeholder for broadcast events API
        # The actual implementation would depend on available Graph API endpoints
        Write-LogMessage "Broadcast events retrieval not yet implemented in Graph API" "Warning"
        
        return $events
    } catch {
        Write-LogMessage "Failed to retrieve broadcast events: $($_.Exception.Message)" "Warning"
        return @()
    }
}

# =========================
# Attendance Data Processing
# =========================
function Get-MeetingAttendanceData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Meeting
    )
    
    try {
        Write-LogMessage "Processing attendance for: $($Meeting.Subject)" "Verbose" @{
            MeetingId = $Meeting.Id
            Source = $Meeting.Source
        }
        
        # Skip if already processed
        if ($script:ProcessedMeetingIds.Contains($Meeting.Id)) {
            Write-LogMessage "Meeting already processed, skipping: $($Meeting.Id)" "Verbose"
            return @()
        }
        
        $attendanceRecords = @()
        
        # Determine the online meeting ID
        $onlineMeetingId = $null
        if ($Meeting.OnlineMeetingId) {
            $onlineMeetingId = $Meeting.OnlineMeetingId
        } elseif ($Meeting.JoinUrl) {
            $onlineMeeting = Get-OnlineMeetingByJoinUrl -UserId $UserId -JoinUrl $Meeting.JoinUrl
            if ($onlineMeeting) {
                $onlineMeetingId = $onlineMeeting.id
            }
        }
        
        if (-not $onlineMeetingId) {
            Write-LogMessage "No online meeting ID found for: $($Meeting.Subject)" "Warning"
            $script:FailedMeetings.Add([PSCustomObject]@{
                Subject = $Meeting.Subject
                Start = $Meeting.Start
                Error = "No online meeting ID found"
                Source = $Meeting.Source
            }) | Out-Null
            return @()
        }
        
        # Get attendance reports
        $attendanceReports = Get-AttendanceReports -UserId $UserId -OnlineMeetingId $onlineMeetingId
        
        foreach ($report in $attendanceReports) {
            $expandedRecords = Get-ExpandedAttendanceRecords -UserId $UserId -OnlineMeetingId $onlineMeetingId -ReportId $report.id
            
            foreach ($record in $expandedRecords) {
                foreach ($interval in $record.attendanceIntervals) {
                    $attendanceRecords += [PSCustomObject]@{
                        MeetingId = $onlineMeetingId
                        Subject = $Meeting.Subject
                        MeetingStart = Format-DateTime $Meeting.Start
                        MeetingEnd = Format-DateTime $Meeting.End
                        Organizer = $Meeting.Organizer
                        AttendeeName = $record.identity.displayName
                        AttendeeEmail = $record.identity.email
                        AttendeeId = $record.identity.id
                        Role = $record.role
                        JoinTime = Format-DateTime $interval.joinDateTime
                        LeaveTime = Format-DateTime $interval.leaveDateTime
                        DurationSeconds = $interval.durationInSeconds
                        DurationMinutes = [math]::Round($interval.durationInSeconds / 60, 2)
                        MeetingType = $Meeting.MeetingType
                        Source = $Meeting.Source
                        ReportId = $report.id
                        ProcessedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
            }
        }
        
        $script:ProcessedMeetingIds.Add($Meeting.Id) | Out-Null
        
        Write-LogMessage "Processed $($attendanceRecords.Count) attendance records for: $($Meeting.Subject)" "Verbose"
        
        return $attendanceRecords
    } catch {
        Write-LogMessage "Failed to process meeting attendance: $($_.Exception.Message)" "Error" @{
            MeetingId = $Meeting.Id
            Subject = $Meeting.Subject
        }
        
        $script:FailedMeetings.Add([PSCustomObject]@{
            Subject = $Meeting.Subject
            Start = $Meeting.Start
            Error = $_.Exception.Message
            Source = $Meeting.Source
        }) | Out-Null
        
        return @()
    }
}

function Get-OnlineMeetingByJoinUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [string]$JoinUrl
    )
    
    try {
        $encodedUrl = [System.Web.HttpUtility]::UrlEncode($JoinUrl)
        $requestUri = "https://graph.microsoft.com/v1.0/users/$UserId/onlineMeetings?`$filter=joinWebUrl eq '$encodedUrl'"
        
        $response = Invoke-GraphRequestWithRetry -Uri $requestUri -Method "GET"
        
        if ($response.value -and $response.value.Count -gt 0) {
            return $response.value[0]
        }
        
        return $null
    } catch {
        Write-LogMessage "Failed to find online meeting by join URL: $($_.Exception.Message)" "Warning"
        return $null
    }
}

function Get-AttendanceReports {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [string]$OnlineMeetingId
    )
    
    try {
        $requestUri = "https://graph.microsoft.com/v1.0/users/$UserId/onlineMeetings/$OnlineMeetingId/attendanceReports"
        $response = Invoke-GraphRequestWithRetry -Uri $requestUri -Method "GET"
        
        if ($response.value) {
            return $response.value
        }
        
        return @()
    } catch {
        Write-LogMessage "Failed to retrieve attendance reports: $($_.Exception.Message)" "Warning"
        return @()
    }
}

function Get-ExpandedAttendanceRecords {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [string]$OnlineMeetingId,
        
        [Parameter(Mandatory = $true)]
        [string]$ReportId
    )
    
    try {
        $requestUri = "https://graph.microsoft.com/v1.0/users/$UserId/onlineMeetings/$OnlineMeetingId/attendanceReports/$ReportId" +
                      "?`$expand=attendanceRecords"
        
        $response = Invoke-GraphRequestWithRetry -Uri $requestUri -Method "GET"
        
        # Save debug information if enabled
        if ($script:Config.DefaultSettings.EnableDebug) {
            $debugPath = Join-Path $script:Config.DefaultSettings.LogPath "debug_attendanceReport_$ReportId.json"
            $response | ConvertTo-Json -Depth 10 | Out-File -FilePath $debugPath -Encoding UTF8
        }
        
        if ($response.attendanceRecords) {
            return $response.attendanceRecords
        }
        
        return @()
    } catch {
        Write-LogMessage "Failed to expand attendance records: $($_.Exception.Message)" "Warning"
        return @()
    }
}

# =========================
# Utility Functions
# =========================
function Invoke-GraphRequestWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        
        [Parameter(Mandatory = $true)]
        [string]$Method,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Headers = @{},
        
        [Parameter(Mandatory = $false)]
        [object]$Body = $null,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 3,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryDelaySeconds = 2
    )
    
    $script:RequestCount++
    
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            Write-LogMessage "Graph API Request [$attempt/$MaxRetries]: $Method $Uri" "Debug" @{
                RequestId = $script:RequestCount
                Attempt = $attempt
            }
            
            $params = @{
                Uri = $Uri
                Method = $Method
                Headers = $Headers
            }
            
            if ($Body) {
                $params.Body = $Body
            }
            
            $response = Invoke-MgGraphRequest @params
            
            Write-LogMessage "Graph API Request successful" "Debug" @{
                RequestId = $script:RequestCount
                StatusCode = "Success"
            }
            
            # Add delay between requests to avoid throttling
            if ($script:Config.DefaultSettings.RequestDelayMs -gt 0) {
                Start-Sleep -Milliseconds $script:Config.DefaultSettings.RequestDelayMs
            }
            
            return $response
        } catch {
            $errorMessage = $_.Exception.Message
            Write-LogMessage "Graph API Request failed (attempt $attempt): $errorMessage" "Warning" @{
                RequestId = $script:RequestCount
                Attempt = $attempt
                Error = $errorMessage
            }
            
            # Check if it's a throttling error
            if ($_.Exception.Response.StatusCode -eq 429 -or $errorMessage -like "*throttle*") {
                $retryAfter = 60 # Default retry after 60 seconds for throttling
                if ($_.Exception.Response.Headers["Retry-After"]) {
                    $retryAfter = [int]$_.Exception.Response.Headers["Retry-After"]
                }
                Write-LogMessage "Throttling detected, waiting $retryAfter seconds before retry" "Warning"
                Start-Sleep -Seconds $retryAfter
            } else {
                Start-Sleep -Seconds ($RetryDelaySeconds * $attempt)
            }
            
            if ($attempt -eq $MaxRetries) {
                throw
            }
        }
    }
}

function Format-DateTime {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$DateTime
    )
    
    if (-not $DateTime) {
        return $null
    }
    
    try {
        # Handle different datetime formats
        if ($DateTime -is [string]) {
            $parsedDate = [DateTime]::Parse($DateTime)
            return $parsedDate.ToString("yyyy-MM-dd HH:mm:ss")
        } elseif ($DateTime -is [DateTime]) {
            return $DateTime.ToString("yyyy-MM-dd HH:mm:ss")
        } elseif ($DateTime.GetType().Name -eq "PSCustomObject" -and $DateTime.dateTime) {
            $parsedDate = [DateTime]::Parse($DateTime.dateTime)
            return $parsedDate.ToString("yyyy-MM-dd HH:mm:ss")
        } else {
            return $DateTime.ToString()
        }
    } catch {
        Write-LogMessage "Failed to format datetime: $($_.Exception.Message)" "Warning" @{
            Input = $DateTime.ToString()
            Type = $DateTime.GetType().Name
        }
        return $DateTime.ToString()
    }
}

function Remove-DuplicateRecords {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.ArrayList]$Records
    )
    
    try {
        Write-LogMessage "Removing duplicate records from $($Records.Count) total records" "Information"
        
        $uniqueRecords = [System.Collections.ArrayList]::new()
        $seenKeys = [System.Collections.Generic.HashSet[string]]::new()
        
        foreach ($record in $Records) {
            # Create a unique key based on meeting ID, attendee ID, and join time
            $key = "$($record.MeetingId)|$($record.AttendeeId)|$($record.JoinTime)"
            
            if (-not $seenKeys.Contains($key)) {
                $uniqueRecords.Add($record) | Out-Null
                $seenKeys.Add($key) | Out-Null
            }
        }
        
        $duplicatesRemoved = $Records.Count - $uniqueRecords.Count
        Write-LogMessage "Removed $duplicatesRemoved duplicate records, $($uniqueRecords.Count) unique records remain" "Information"
        
        return $uniqueRecords
    } catch {
        Write-LogMessage "Failed to remove duplicates: $($_.Exception.Message)" "Error"
        return $Records
    }
}

function Export-AttendanceReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [System.Collections.ArrayList]$AttendanceData = @(),
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )
    
    try {
        if ($AttendanceData.Count -eq 0) {
            Write-LogMessage "No attendance data to export" "Warning"
            return
        }
        
        # Remove duplicates before export
        $uniqueData = Remove-DuplicateRecords -Records $AttendanceData
        
        # Export to CSV
        $uniqueData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        
        Write-LogMessage "Exported $($uniqueData.Count) attendance records to: $OutputPath" "Information"
    } catch {
        Write-LogMessage "Failed to export attendance report: $($_.Exception.Message)" "Error"
        throw
    }
}

function Export-FailedMeetingsReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [System.Collections.ArrayList]$FailedMeetings = @(),
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )
    
    try {
        if ($FailedMeetings.Count -eq 0) {
            Write-LogMessage "No failed meetings to export" "Information"
            return
        }
        
        $FailedMeetings | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        
        Write-LogMessage "Exported $($FailedMeetings.Count) failed meetings to: $OutputPath" "Information"
    } catch {
        Write-LogMessage "Failed to export failed meetings report: $($_.Exception.Message)" "Error"
        throw
    }
}

function Show-ExecutionSummary {
    [CmdletBinding()]
    param()
    
    $endTime = Get-Date
    $duration = $endTime - $script:StartTime
    
    $summary = @{
        StartTime = $script:StartTime.ToString("yyyy-MM-dd HH:mm:ss")
        EndTime = $endTime.ToString("yyyy-MM-dd HH:mm:ss")
        Duration = $duration.ToString("hh\:mm\:ss")
        TotalRequests = $script:RequestCount
        AttendanceRecords = $script:AttendanceData.Count
        FailedMeetings = $script:FailedMeetings.Count
        ProcessedMeetings = $script:ProcessedMeetingIds.Count
        PowerShellVersion = $script:PowerShellVersion
        LogFile = $script:LogPath
    }
    
    Write-Host "`n=== Execution Summary ===" -ForegroundColor Yellow
    foreach ($key in $summary.Keys) {
        Write-Host "$key`: $($summary[$key])" -ForegroundColor Cyan
    }
    Write-Host "=========================" -ForegroundColor Yellow
    
    Write-LogMessage "Execution completed" "Information" $summary
}

# =========================
# Main Execution Function
# =========================
function Start-TeamsAttendanceReport {
    [CmdletBinding()]
    param()
    
    try {
        # Initialize configuration and logging
        Initialize-Configuration
        
        # Connect to Microsoft Graph
        Connect-ToMicrosoftGraph
        
        # Get user information
        $user = Get-UserInformation -UserPrincipalName $script:Config.DefaultSettings.UserPrincipalName
        
        # Get comprehensive meetings data
        $meetings = Get-ComprehensiveMeetingsData -UserId $user.Id -StartDate $script:Config.DefaultSettings.StartDate -EndDate $script:Config.DefaultSettings.EndDate
        
        Write-LogMessage "Processing $($meetings.Count) meetings for attendance data" "Information"
        
        # Process each meeting for attendance data
        foreach ($meeting in $meetings) {
            $attendanceRecords = Get-MeetingAttendanceData -UserId $user.Id -Meeting $meeting
            
            if ($attendanceRecords) {
                # Safely add attendance records
                if ($attendanceRecords -is [System.Collections.IEnumerable] -and $attendanceRecords -isnot [string]) {
                    foreach ($record in $attendanceRecords) {
                        $script:AttendanceData.Add($record) | Out-Null
                    }
                } else {
                    $script:AttendanceData.Add($attendanceRecords) | Out-Null
                }
            }
            
            # Rate limiting
            if ($script:Config.DefaultSettings.RequestDelayMs -gt 0) {
                Start-Sleep -Milliseconds $script:Config.DefaultSettings.RequestDelayMs
            }
        }
        
        # Generate output file names with timestamp
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $attendanceOutputPath = "$OutputPrefix`_Attendance_$timestamp.csv"
        $failedOutputPath = "$OutputPrefix`_Failed_$timestamp.csv"
        
        # Export reports
        Export-AttendanceReport -AttendanceData $script:AttendanceData -OutputPath $attendanceOutputPath
        Export-FailedMeetingsReport -FailedMeetings $script:FailedMeetings -OutputPath $failedOutputPath
        
        # Show execution summary
        Show-ExecutionSummary
        
    } catch {
        Write-LogMessage "Critical error in main execution: $($_.Exception.Message)" "Error"
        Write-Error "Script execution failed: $($_.Exception.Message)"
        exit 1
    } finally {
        # Disconnect from Graph
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Write-LogMessage "Disconnected from Microsoft Graph" "Information"
        } catch {
            Write-LogMessage "Error disconnecting from Graph: $($_.Exception.Message)" "Warning"
        }
    }
}

# =========================
# Script Entry Point
# =========================
if ($MyInvocation.InvocationName -ne '.') {
    Start-TeamsAttendanceReport
}