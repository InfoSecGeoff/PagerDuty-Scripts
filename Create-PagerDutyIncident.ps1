<#
.SYNOPSIS
Creates a new incident in PagerDuty using the Events API v2

.DESCRIPTION
This script creates a new incident in PagerDuty with customizable parameters suitable for SOC operations.
Requires an Integration Key from your PagerDuty service.

.PARAMETER IntegrationKey
The Integration Key from your PagerDuty service (required)

.PARAMETER Summary
Brief description of the incident (required)

.PARAMETER Severity
Incident severity level: critical, error, warning, or info (default: error)

.PARAMETER Source
Source of the incident (default: computer name)

.PARAMETER CustomDetails
Hashtable of additional custom details

.PARAMETER DedupKey
Deduplication key for grouping related incidents (auto-generated if not provided)

.PARAMETER Component
Component affected (default: Security Operations)

.PARAMETER Group
Group responsible (default: SOC)

.EXAMPLE
.\Create-PagerDutyIncident.ps1 -IntegrationKey "YOUR_INTEGRATION_KEY" -Summary "Suspicious login detected" -Severity "critical"

.EXAMPLE
$CustomData = @{"ip_address" = "192.168.1.100"; "user" = "jdoe"}
.\Create-PagerDutyIncident.ps1 -IntegrationKey "YOUR_KEY" -Summary "Failed login attempts" -CustomDetails $CustomData

.NOTES 

- API token: Found in PagerDuty under Services > Your Service > Integrations > Events API V2

Author: Geoff Tankersley
Version: 2.1
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$IntegrationKey,
    
    [Parameter(Mandatory=$true)]
    [string]$Summary,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("critical", "error", "warning", "info")]
    [string]$Severity = "error",
    
    [Parameter(Mandatory=$false)]
    [string]$Source = $env:COMPUTERNAME,
    
    [Parameter(Mandatory=$false)]
    [hashtable]$CustomDetails = @{},
    
    [Parameter(Mandatory=$false)]
    [string]$DedupKey = $null,
    
    [Parameter(Mandatory=$false)]
    [string]$Component = "Security Operations",
    
    [Parameter(Mandatory=$false)]
    [string]$Group = "SOC"
)

function New-PagerDutyIncident {
    param(
        [string]$IntegrationKey,
        [string]$Summary,
        [string]$Severity,
        [string]$Source,
        [hashtable]$CustomDetails,
        [string]$DedupKey,
        [string]$Component,
        [string]$Group
    )
    
    # PagerDuty Events API v2 endpoint
    $Uri = "https://events.pagerduty.com/v2/enqueue"
    
    # Generate dedup_key if necessary
    if ([string]::IsNullOrEmpty($DedupKey)) {
        $DedupKey = [System.Guid]::NewGuid().ToString()
    }
    
    $Payload = @{
        routing_key = $IntegrationKey
        event_action = "trigger"
        dedup_key = $DedupKey
        payload = @{
            summary = $Summary
            source = $Source
            severity = $Severity
            component = $Component
            group = $Group
            custom_details = $CustomDetails
        }
    }
    
    $JsonPayload = $Payload | ConvertTo-Json -Depth 3
    
    # Headers
    $Headers = @{
        "Content-Type" = "application/json"
        "User-Agent" = "PowerShell-SOC-Script/1.0"
    }
    
    try {
        Write-Host "Creating PagerDuty incident..." -ForegroundColor Yellow
        Write-Host "Summary: $Summary" -ForegroundColor Cyan
        Write-Host "Severity: $Severity" -ForegroundColor Cyan
        Write-Host "Dedup Key: $DedupKey" -ForegroundColor Cyan
        
        # Debug: Show the JSON payload
        Write-Host "JSON Payload:" -ForegroundColor Gray
        Write-Host $JsonPayload -ForegroundColor Gray
        
        # Make the API call
        $Response = Invoke-RestMethod -Uri $Uri -Method Post -Body $JsonPayload -Headers $Headers
        
        Write-Host "✓ Incident created successfully!" -ForegroundColor Green
        Write-Host "Dedup Key: $($Response.dedup_key)" -ForegroundColor Green
        Write-Host "Status: $($Response.status)" -ForegroundColor Green
        Write-Host "Message: $($Response.message)" -ForegroundColor Green
        
        return $Response
    }
    catch [Microsoft.PowerShell.Commands.HttpResponseException] {
        $StatusCode = $_.Exception.Response.StatusCode
        $ReasonPhrase = $_.Exception.Response.ReasonPhrase
        
        Write-Error "Failed to create PagerDuty incident: $StatusCode $ReasonPhrase"
        
        try {
            $ErrorContent = $_.ErrorDetails.Message
            if ($ErrorContent) {
                Write-Host "API Error Details:" -ForegroundColor Red
                Write-Host $ErrorContent -ForegroundColor Red
            }
        }
        catch {
            Write-Host "Could not retrieve error details" -ForegroundColor Red
        }
        throw
    }
    catch {
        Write-Error "Failed to create PagerDuty incident: $($_.Exception.Message)"
        throw
    }
}

$SecurityDetails = @{
    "alert_time" = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss UTC")
    "analyst" = $env:USERNAME
    "detection_method" = "SOC Monitoring"
}

$AllCustomDetails = $SecurityDetails + $CustomDetails

# Create incident
try {
    $Result = New-PagerDutyIncident -IntegrationKey $IntegrationKey -Summary $Summary -Severity $Severity -Source $Source -CustomDetails $AllCustomDetails -DedupKey $DedupKey -Component $Component -Group $Group
    
    return $Result
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    exit 1
}
