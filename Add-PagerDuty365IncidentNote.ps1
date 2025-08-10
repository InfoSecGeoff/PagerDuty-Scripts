<#
.SYNOPSIS
Adds a note to an existing PagerDuty incident

.DESCRIPTION
This script adds a note to a PagerDuty incident using the REST API.
Can find incidents by ID, title, or dedup key.

IMPORTANT: 
- The UserEmail parameter must be a valid PagerDuty user's email address.
- Search by title only works on incident titles, NOT incident body/description content.

.PARAMETER ApiToken
PagerDuty REST API token (required)

.PARAMETER IncidentId
The ID of the incident to add a note to (required unless using other search methods)

.PARAMETER Note
The content of the note to add (required)

.PARAMETER UserEmail
Email address of the user adding the note (required)
MUST be a valid PagerDuty user email address with appropriate permissions

.PARAMETER FindIncidentByTitle
Switch to search for incident by title instead of using ID

.PARAMETER IncidentTitle
Title or partial title to search for (used with -FindIncidentByTitle)
Note: This searches the incident TITLE only, not the body/description

.PARAMETER FindIncidentByDedupKey
Switch to search for incident by dedup key instead of using ID

.PARAMETER DedupKey
Dedup key from incident creation to search for (used with -FindIncidentByDedupKey)

.PARAMETER IncludeResolved
Include resolved incidents in the search (by default only shows triggered/acknowledged)

.PARAMETER ListIncidents
Just list all available incidents without adding a note

.EXAMPLE
# List all active incidents to find the right ID
.\Add-PagerDutyIncidentNote.ps1 -ApiToken "your_api_token" -ListIncidents -UserEmail "analyst@company.com" -Note "dummy"

.EXAMPLE
# List ALL incidents including resolved
.\Add-PagerDutyIncidentNote.ps1 -ApiToken "your_api_token" -ListIncidents -IncludeResolved -UserEmail "analyst@company.com" -Note "dummy"

.EXAMPLE
# Add note using incident ID directly
.\Add-PagerDutyIncidentNote.ps1 -ApiToken "your_api_token" -IncidentId "P1234567" -Note "SOC investigating" -UserEmail "analyst@company.com"

.EXAMPLE
# Search by title (searches incident titles only, not body content)
.\Add-PagerDutyIncidentNote.ps1 -ApiToken "your_api_token" -FindIncidentByTitle -IncidentTitle "disk space" -Note "Investigation update" -UserEmail "soc@company.com"

.EXAMPLE
# Search including resolved incidents
.\Add-PagerDutyIncidentNote.ps1 -ApiToken "your_api_token" -FindIncidentByTitle -IncidentTitle "alert" -IncludeResolved -Note "Post-mortem note" -UserEmail "soc@company.com"

.NOTES
- API token: Found in PagerDuty under User Settings > API Access Keys
- UserEmail MUST be a valid PagerDuty user email (not just any email address)
- Common error 1008 means "Requester User Not Found" - check the email address
- The script searches incident TITLES only, not body/description content
- Use -ListIncidents to see all available incidents and their IDs
- Use -Verbose for additional debug information

Author: Geoff Tankersley
Version: 2.1
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$ApiToken,
    
    [Parameter(Mandatory=$false)]
    [string]$IncidentId,
    
    [Parameter(Mandatory=$true)]
    [string]$Note,
    
    [Parameter(Mandatory=$true)]
    [string]$UserEmail,
    
    [Parameter(Mandatory=$false)]
    [switch]$FindIncidentByTitle,
    
    [Parameter(Mandatory=$false)]
    [string]$IncidentTitle = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$FindIncidentByDedupKey,
    
    [Parameter(Mandatory=$false)]
    [string]$DedupKey = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeResolved,
    
    [Parameter(Mandatory=$false)]
    [switch]$ListIncidents
)

function Add-PagerDutyNote {
    param(
        [string]$ApiToken,
        [string]$IncidentId,
        [string]$Note,
        [string]$UserEmail
    )

    $Uri = "https://api.pagerduty.com/incidents/$IncidentId/notes"
    
    # From header is REQUIRED and must be a valid PagerDuty user email
    $Headers = @{
        "Authorization" = "Token token=$ApiToken"
        "Content-Type" = "application/json"
        "Accept" = "application/vnd.pagerduty+json;version=2"
        "From" = $UserEmail  # This is MANDATORY and must be a valid PagerDuty user email
    }
    
    $Payload = @{
        note = @{
            content = $Note
        }
    }
    

    $JsonPayload = $Payload | ConvertTo-Json -Depth 3
    
    try {
        Write-Host "Adding note to incident $IncidentId..." -ForegroundColor Yellow
        Write-Host "Note preview: $($Note.Substring(0, [Math]::Min(100, $Note.Length)))$(if($Note.Length -gt 100){'...'})" -ForegroundColor Cyan
        Write-Host "From user: $UserEmail" -ForegroundColor Gray
        
        Write-Verbose "Headers: $($Headers | ConvertTo-Json)"
        Write-Verbose "JSON Payload: $JsonPayload"
        
        # Make the API call
        $Response = Invoke-RestMethod -Uri $Uri -Method Post -Body $JsonPayload -Headers $Headers
        
        Write-Host "✓ Note added successfully!" -ForegroundColor Green
        Write-Host "Note ID: $($Response.note.id)" -ForegroundColor Green
        Write-Host "Created at: $($Response.note.created_at)" -ForegroundColor Green
        Write-Host "Created by: $($Response.note.user.summary)" -ForegroundColor Green
        
        return $Response.note
    }
    catch {
        $StatusCode = $_.Exception.Response.StatusCode
        Write-Error "Failed to add note to PagerDuty incident: $StatusCode"
        
        # Error handling
        if ($_.ErrorDetails.Message) {
            $errorJson = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
            if ($errorJson) {
                Write-Host "API Error:" -ForegroundColor Red
                Write-Host "  Code: $($errorJson.error.code)" -ForegroundColor Red
                Write-Host "  Message: $($errorJson.error.message)" -ForegroundColor Red
                if ($errorJson.error.errors) {
                    Write-Host "  Details: $($errorJson.error.errors -join ', ')" -ForegroundColor Red
                }
                
                # Common error codes
                if ($errorJson.error.code -eq 1008) {
                    Write-Host "`nNote: Error 1008 means 'Requester User Not Found'. Ensure that:" -ForegroundColor Yellow
                    Write-Host "  - The email address '$UserEmail' belongs to a valid PagerDuty user" -ForegroundColor Yellow
                    Write-Host "  - The user has appropriate permissions to add notes to incidents" -ForegroundColor Yellow
                }
            }
            else {
                Write-Host "Raw error: $($_.ErrorDetails.Message)" -ForegroundColor Red
            }
        }
        throw
    }
}

function Test-PagerDutyApiToken {
    param([string]$ApiToken)
    
    $Headers = @{
        "Authorization" = "Token token=$ApiToken"
        "Accept" = "application/vnd.pagerduty+json;version=2"
    }
    
    try {
        Write-Host "Testing API token..." -ForegroundColor Yellow
        Write-Verbose "Token: $($ApiToken.Substring(0, [Math]::Min(10, $ApiToken.Length)))..."
        
        $Response = Invoke-RestMethod -Uri "https://api.pagerduty.com/abilities" -Method Get -Headers $Headers
        Write-Host "✓ API token is valid" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "API token validation failed: $($_.Exception.Message)"
        if ($_.ErrorDetails.Message) {
            Write-Host "Error details: $($_.ErrorDetails.Message)" -ForegroundColor Red
        }
        return $false
    }
}

function Find-PagerDutyIncidentByDedupKey {
    param(
        [string]$ApiToken,
        [string]$DedupKey
    )
    
    $Headers = @{
        "Authorization" = "Token token=$ApiToken"
        "Accept" = "application/vnd.pagerduty+json;version=2"
    }
    
    Write-Host "Searching for incident with dedup key: '$DedupKey'..." -ForegroundColor Yellow
    
    # Search recent incidents
    $Since = (Get-Date).AddDays(-7).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $Uri = "https://api.pagerduty.com/incidents?since=$Since&limit=100&include[]=first_trigger_log_entries"
    
    try {
        $Response = Invoke-RestMethod -Uri $Uri -Method Get -Headers $Headers
        
        Write-Host "Checking $($Response.incidents.Count) recent incidents..." -ForegroundColor Gray
        
        foreach ($incident in $Response.incidents) {
            # Check first trigger log entry
            if ($incident.first_trigger_log_entry -and 
                $incident.first_trigger_log_entry.channel -and 
                $incident.first_trigger_log_entry.channel.details -and
                $incident.first_trigger_log_entry.channel.details.dedup_key -eq $DedupKey) {
                Write-Host "Found incident: $($incident.id) - $($incident.title)" -ForegroundColor Green
                return $incident.id
            }
            
            # Incident_key fallback
            if ($incident.incident_key -eq $DedupKey) {
                Write-Host "Found incident via incident_key: $($incident.id) - $($incident.title)" -ForegroundColor Green
                return $incident.id
            }
        }
        
        # Pagination
        $offset = 100
        while ($Response.more -and $offset -lt 500) {
            $PageUri = "https://api.pagerduty.com/incidents?since=$Since&limit=100&offset=$offset&include[]=first_trigger_log_entries"
            $Response = Invoke-RestMethod -Uri $PageUri -Method Get -Headers $Headers
            
            foreach ($incident in $Response.incidents) {
                if ($incident.first_trigger_log_entry?.channel?.details?.dedup_key -eq $DedupKey -or
                    $incident.incident_key -eq $DedupKey) {
                    Write-Host "Found incident: $($incident.id) - $($incident.title)" -ForegroundColor Green
                    return $incident.id
                }
            }
            $offset += 100
        }
        
        Write-Warning "No incident found with dedup key: '$DedupKey'"
        return $null
    }
    catch {
        Write-Error "Failed to search by dedup key: $($_.Exception.Message)"
        throw
    }
}

function Find-PagerDutyIncident {
    param(
        [string]$ApiToken,
        [string]$IncidentTitle,
        [switch]$IncludeResolved
    )
    
    $Headers = @{
        "Authorization" = "Token token=$ApiToken"
        "Accept" = "application/vnd.pagerduty+json;version=2"
    }
    
    try {
        Write-Host "Searching for incidents..." -ForegroundColor Yellow
        Write-Host "Search term: '$IncidentTitle'" -ForegroundColor Gray
        Write-Host "Note: Search only works on incident titles, not body content" -ForegroundColor Gray
        
        # Get incidents - try with statuses parameter first
        $BaseUri = "https://api.pagerduty.com/incidents"
        
        # Try to get incidents with specific statuses
        if (-not $IncludeResolved) {
            # Try the proper way with array parameters
            $Uri = "${BaseUri}?statuses[]=triggered&statuses[]=acknowledged&limit=100"
            Write-Verbose "Trying with status parameters: $Uri"
        }
        else {
            # Get all incidents including resolved
            $Uri = "${BaseUri}?limit=100"
        }
        
        Write-Verbose "Fetching incidents from: $Uri"
        
        $FilteredIncidents = $null
        try {
            $Response = Invoke-RestMethod -Uri $Uri -Method Get -Headers $Headers
            Write-Host "Retrieved $($Response.incidents.Count) incidents from API" -ForegroundColor Gray
            
            # If we used status filters in the URL, use the results directly
            if (-not $IncludeResolved) {
                $FilteredIncidents = $Response.incidents
                Write-Host "Active incidents from API: $($FilteredIncidents.Count)" -ForegroundColor Gray
            }
        }
        catch {
            # If status parameter fails, fall back to getting all and filtering
            Write-Verbose "Status parameter failed, falling back to fetch all and filter locally"
            $Uri = "${BaseUri}?limit=100"
            $Response = Invoke-RestMethod -Uri $Uri -Method Get -Headers $Headers
            Write-Host "Retrieved $($Response.incidents.Count) incidents (fallback method)" -ForegroundColor Gray
        }
        
        # Apply status filter locally if we haven't already
        if ($null -eq $FilteredIncidents) {
            if (-not $IncludeResolved) {
                $FilteredIncidents = $Response.incidents | Where-Object { 
                    $_.status -eq 'triggered' -or $_.status -eq 'acknowledged' 
                }
                Write-Host "Active incidents after local filtering: $($FilteredIncidents.Count)" -ForegroundColor Gray
                
                # Debug: Show status values we're seeing
                $uniqueStatuses = $Response.incidents | Select-Object -ExpandProperty status -Unique
                Write-Verbose "Unique status values found: $($uniqueStatuses -join ', ')"
            }
            else {
                $FilteredIncidents = $Response.incidents
            }
        }
        
        # Show all incidents with their actual titles
        if ($FilteredIncidents.Count -gt 0) {
            Write-Host "`nShowing all available incidents:" -ForegroundColor Cyan
            Write-Host "ID          | Status       | Title" -ForegroundColor Gray
            Write-Host "----------- | ------------ | -----" -ForegroundColor Gray
            
            $FilteredIncidents | ForEach-Object {
                $statusColor = switch ($_.status) {
                    'triggered' { 'Red' }
                    'acknowledged' { 'Yellow' }
                    'resolved' { 'DarkGray' }
                    default { 'Gray' }
                }
                
                $titleDisplay = if ($_.title.Length -gt 60) { 
                    $_.title.Substring(0, 57) + "..." 
                } else { 
                    $_.title 
                }
                
                Write-Host "$($_.incident_number.ToString().PadRight(11)) | " -NoNewline
                Write-Host "$($_.status.PadRight(12))" -ForegroundColor $statusColor -NoNewline
                Write-Host " | $titleDisplay"
            }
        }
        
        # Filter by title if provided
        if ($IncidentTitle) {
            $MatchingIncidents = $FilteredIncidents | Where-Object { 
                $_.title -like "*$IncidentTitle*" 
            }
            
            if ($MatchingIncidents.Count -eq 0) {
                Write-Warning "`nNo incidents found with '$IncidentTitle' in the title"
                Write-Host "Remember: Search only works on incident titles, not body/description content" -ForegroundColor Yellow
                
                # Suggest using incident number instead
                Write-Host "`nTip: You can use the incident ID directly. For example:" -ForegroundColor Cyan
                Write-Host "  -IncidentId 'P1234567'" -ForegroundColor Gray
                
                return $null
            }
            elseif ($MatchingIncidents.Count -eq 1) {
                $Incident = $MatchingIncidents[0]
                Write-Host "`n✓ Found matching incident:" -ForegroundColor Green
                Write-Host "  ID: $($Incident.id)" -ForegroundColor Green
                Write-Host "  Title: $($Incident.title)" -ForegroundColor Green
                Write-Host "  Status: $($Incident.status)" -ForegroundColor Green
                return $Incident.id
            }
            else {
                Write-Host "`nMultiple incidents found matching '$IncidentTitle':" -ForegroundColor Yellow
                $MatchingIncidents | ForEach-Object { 
                    Write-Host "  $($_.id) - $($_.title) [$($_.status)]" -ForegroundColor Cyan
                }
                
                # Return the most recent one
                $MostRecent = $MatchingIncidents | Sort-Object created_at -Descending | Select-Object -First 1
                Write-Host "`nUsing most recent: $($MostRecent.id)" -ForegroundColor Green
                return $MostRecent.id
            }
        }
        else {
            Write-Host "`nNo search term provided. Please specify an incident ID from the list above." -ForegroundColor Yellow
            return $null
        }
    }
    catch {
        Write-Error "Failed to search for incidents: $($_.Exception.Message)"
        if ($_.ErrorDetails.Message) {
            Write-Host "Error details:" -ForegroundColor Red
            Write-Host $_.ErrorDetails.Message -ForegroundColor Red
        }
        throw
    }
}

function Get-PagerDutyIncidentInfo {
    param(
        [string]$ApiToken,
        [string]$IncidentId
    )
    
    $Uri = "https://api.pagerduty.com/incidents/$IncidentId"
    
    $Headers = @{
        "Authorization" = "Token token=$ApiToken"
        "Accept" = "application/vnd.pagerduty+json;version=2"
    }
    
    try {
        $Response = Invoke-RestMethod -Uri $Uri -Method Get -Headers $Headers
        return $Response.incident
    }
    catch {
        Write-Error "Failed to get incident info: $($_.Exception.Message)"
        return $null
    }
}

# Main 
Write-Host "`n=== PagerDuty Note Addition Script ===" -ForegroundColor Cyan
Write-Host "User Email: $UserEmail" -ForegroundColor Gray

# Test API token 
if (-not (Test-PagerDutyApiToken -ApiToken $ApiToken)) {
    Write-Error "API token validation failed. Please check your token."
    exit 1
}

try {
    # Determine the incident ID using various methods
    $ActualIncidentId = $IncidentId
    
    if ($ListIncidents) {
        Write-Host "`nListing incidents..." -ForegroundColor Cyan
        $null = Find-PagerDutyIncident -ApiToken $ApiToken -IncludeResolved:$IncludeResolved
        Write-Host "`nUse one of the incident IDs above with -IncidentId parameter" -ForegroundColor Yellow
        exit 0
    }
    
    if ($FindIncidentByDedupKey -and $DedupKey) {
        Write-Host "`nFinding incident by dedup key..." -ForegroundColor Cyan
        $ActualIncidentId = Find-PagerDutyIncidentByDedupKey -ApiToken $ApiToken -DedupKey $DedupKey
        if (-not $ActualIncidentId) {
            Write-Error "Could not find incident with dedup key: '$DedupKey'"
            exit 1
        }
    }
    elseif ($FindIncidentByTitle -and $IncidentTitle) {
        Write-Host "`nFinding incident by title..." -ForegroundColor Cyan
        $ActualIncidentId = Find-PagerDutyIncident -ApiToken $ApiToken -IncidentTitle $IncidentTitle -IncludeResolved:$IncludeResolved
        if (-not $ActualIncidentId) {
            Write-Error "Could not find incident matching title: '$IncidentTitle'"
            exit 1
        }
    }
    elseif (-not $ActualIncidentId) {
        Write-Error "Must provide either -IncidentId, -FindIncidentByTitle with -IncidentTitle, -FindIncidentByDedupKey with -DedupKey, or -ListIncidents"
        exit 1
    }
    
    # Verify incident exists
    Write-Host "`nVerifying incident..." -ForegroundColor Cyan
    $IncidentInfo = Get-PagerDutyIncidentInfo -ApiToken $ApiToken -IncidentId $ActualIncidentId
    
    if (-not $IncidentInfo) {
        Write-Error "Incident $ActualIncidentId not found"
        exit 1
    }
    
    Write-Host "✓ Incident verified: $($IncidentInfo.title)" -ForegroundColor Green
    Write-Host "  Status: $($IncidentInfo.status)" -ForegroundColor Gray
    Write-Host "  Service: $($IncidentInfo.service.summary)" -ForegroundColor Gray
    Write-Host "  Created: $($IncidentInfo.created_at)" -ForegroundColor Gray

    Write-Host "`nAdding note..." -ForegroundColor Cyan
    $Result = Add-PagerDutyNote -ApiToken $ApiToken -IncidentId $ActualIncidentId -Note $Note -UserEmail $UserEmail
    
    Write-Host "`n✓ Script completed successfully!" -ForegroundColor Green
    
    return $Result
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    exit 1
}
