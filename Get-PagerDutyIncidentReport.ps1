<#
.SYNOPSIS
    Generates a comprehensive HTML report of PagerDuty incidents over the past 30 days.

.DESCRIPTION
    Get-PagerDutyIncidentReport collects incident data from PagerDuty API and generates a detailed
    HTML analytics report including:
    - Incident status breakdown (triggered, acknowledged, resolved)
    - Response and resolution time metrics
    - Service and client distribution analysis
    - Assignee workload distribution
    - Temporal patterns (daily trends, hourly distribution, day of week)
    - Most common alert types
    - Urgency distribution
    
    The report includes interactive charts and tables for comprehensive SOC performance analysis.

.PARAMETER ApiKey
    Required. PagerDuty API token with read access to incidents.
    Generate at: https://[your-subdomain].pagerduty.com/api_keys
    
.PARAMETER OutputPath
    Optional. Directory where the HTML report will be saved.
    Default: Current directory (.\)

.EXAMPLE
    .\Get-PagerDutyIncidentReport.ps1 -ApiKey "u+ABC123xyz"
    
    Generates report using the provided API key and saves to current directory.

.EXAMPLE
    .\Get-PagerDutyIncidentReport.ps1 -ApiKey "u+ABC123xyz" -OutputPath "C:\Reports\"
    
    Generates report and saves to specified directory.

.OUTPUTS
    Generates an HTML file named "PagerDuty_SOC-Report.html" in the specified output path.

.NOTES
    File Name      : Get-PagerDutyIncidentReport.ps1
    Author         : Geoff Tankersley
    Prerequisite   : PowerShell 5.1 or higher
    Requirements   : 
    - Internet connectivity to PagerDuty API
    - Valid PagerDuty API token with incident read permissions
    
    API Rate Limits:
    - The script implements pagination and includes a 50ms delay between log entry requests
    - Typical execution time: 2-5 minutes depending on incident volume
    
    Report Features:
    - Auto-opens in default browser upon completion
    - Responsive design for mobile/desktop viewing
    - Interactive charts using Chart.js
    - Client extraction from incident titles and service names

.LINK
    https://developer.pagerduty.com/api-reference/
    
.LINK
    https://developer.pagerduty.com/docs/rest-api-v2/authentication/
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$ApiKey,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\"
)

# Global variables
$script:PagerDutyApiToken = $ApiKey
$script:PagerDutyApiUrl = "https://api.pagerduty.com"

function Get-PagerDutyIncidentData {
    param(
        [DateTime]$StartDate,
        [DateTime]$EndDate
    )
    
    $headers = @{
        "Authorization" = "Token token=$script:PagerDutyApiToken"
        "Accept" = "application/vnd.pagerduty+json;version=2"
    }
    
    $allIncidents = @()
    $offset = 0
    $limit = 100
    $hasMore = $true
    
    while ($hasMore) {
        $queryParams = @()
        $queryParams += "limit=$limit"
        $queryParams += "offset=$offset"
        $queryParams += "since=$($StartDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
        $queryParams += "until=$($EndDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
        $queryParams += "include[]=services"
        $queryParams += "include[]=assignments"
        $queryParams += "include[]=acknowledgers"
        $queryParams += "include[]=assignees"
        $queryParams += "include[]=first_trigger_log_entries"
        $queryParams += "include[]=acknowledgements"
        $queryParams += "statuses[]=triggered"
        $queryParams += "statuses[]=acknowledged" 
        $queryParams += "statuses[]=resolved"
        
        $queryString = $queryParams -join "&"
        $url = "$script:PagerDutyApiUrl/incidents?$queryString"
        
        try {
            $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
            $allIncidents += $response.incidents
            
            $hasMore = $response.more
            $offset += $limit
            
            Write-Progress -Activity "Collecting incidents" -Status "Retrieved $($allIncidents.Count) incidents" -PercentComplete (($offset / ($offset + 100)) * 100)
        }
        catch {
            Write-Error "Failed to retrieve incidents: $($_.Exception.Message)"
            break
        }
    }
    
    Write-Progress -Activity "Collecting incidents" -Completed
    
    # Get log entries for acknowledgment times
    Write-Host "Collecting acknowledgment data from log entries..." -ForegroundColor Cyan
    $logEntryCount = 0
    
    for ($i = 0; $i -lt $allIncidents.Count; $i++) {
        $incident = $allIncidents[$i]
        
        try {
            $logUrl = "$script:PagerDutyApiUrl/incidents/$($incident.id)/log_entries"
            $logResponse = Invoke-RestMethod -Uri $logUrl -Method Get -Headers $headers
            
            # Find acknowledge log entries
            $ackLogEntries = $logResponse.log_entries | Where-Object { $_.type -eq "acknowledge_log_entry" }
            
            if ($ackLogEntries.Count -gt 0) {
                $firstAck = $ackLogEntries | Sort-Object created_at | Select-Object -First 1
                $incident | Add-Member -NotePropertyName "first_acknowledgment" -NotePropertyValue @{
                    at = $firstAck.created_at
                    acknowledger = $firstAck.agent
                } -Force
                $logEntryCount++
            }
        }
        catch {
            continue
        }
        
        Start-Sleep -Milliseconds 50
        
        if (($i + 1) % 10 -eq 0) {
            Write-Progress -Activity "Collecting log entries" -Status "Processed $($i + 1) of $($allIncidents.Count) incidents" -PercentComplete (($i + 1) / $allIncidents.Count * 100)
        }
    }
    
    Write-Progress -Activity "Collecting log entries" -Completed
    Write-Host "Found acknowledgment data in $logEntryCount incidents" -ForegroundColor Green
    
    return $allIncidents
}

function Get-IncidentAnalytics {
    param([array]$Incidents)
    
    $analytics = @{}
    
    # Basic counts
    $analytics.TotalIncidents = $Incidents.Count
    $analytics.TriggeredIncidents = ($Incidents | Where-Object { $_.status -eq "triggered" }).Count
    $analytics.AcknowledgedIncidents = ($Incidents | Where-Object { $_.status -eq "acknowledged" }).Count
    $analytics.ResolvedIncidents = ($Incidents | Where-Object { $_.status -eq "resolved" }).Count
    
    # Service analysis
    $serviceStats = $Incidents | Group-Object { $_.service.summary } | Sort-Object Count -Descending
    $analytics.TopServices = $serviceStats | Select-Object -First 10 | ForEach-Object {
        @{ Name = $_.Name; Count = $_.Count; Percentage = [math]::Round(($_.Count / $Incidents.Count) * 100, 2) }
    }
    
    # Client analysis
    $clientStats = @{}
    foreach ($incident in $Incidents) {
        $clientName = "Unknown Client"
        
        if ($incident.title -match "(?i)Client Name:\s*([^,\n\r\|]+)") {
            $clientName = $matches[1].Trim()
        }
        elseif ($incident.title -match "(?i)ame\.client[:\s]*([^,\n\r\s\|]+)") {
            $clientName = $matches[1].Trim()
        }
        elseif ($incident.title -match "(?i)AME-([^-\s,\n\r\|]+)") {
            $clientName = $matches[1].Trim()
        }
        elseif ($incident.service.summary -match "^([^-\s]+)") {
            $clientName = $matches[1].Trim()
        }
        
        if ($clientStats.ContainsKey($clientName)) {
            $clientStats[$clientName]++
        } else {
            $clientStats[$clientName] = 1
        }
    }
    
    $analytics.ClientDistribution = $clientStats.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
        @{ Client = $_.Key; Count = $_.Value; Percentage = [math]::Round(($_.Value / $Incidents.Count) * 100, 2) }
    }
    
    # Alert title analysis
    $titleStats = $Incidents | Group-Object { $_.title } | Sort-Object Count -Descending
    $analytics.TopAlertTitles = $titleStats | Select-Object -First 15 | ForEach-Object {
        @{ Title = $_.Name; Count = $_.Count; Percentage = [math]::Round(($_.Count / $Incidents.Count) * 100, 2) }
    }
    
    # Urgency distribution
    $urgencyStats = $Incidents | Group-Object urgency
    $analytics.UrgencyDistribution = $urgencyStats | ForEach-Object {
        @{ Urgency = $_.Name; Count = $_.Count; Percentage = [math]::Round(($_.Count / $Incidents.Count) * 100, 2) }
    }
    
    # Assignee analysis
    $assigneeStats = @{}
    foreach ($incident in $Incidents) {
        $hasAssignee = $false
        
        if ($incident.assignments -and $incident.assignments.Count -gt 0) {
            foreach ($assignment in $incident.assignments) {
                if ($assignment.assignee -and $assignment.assignee.summary) {
                    $assigneeName = $assignment.assignee.summary
                    if ($assigneeStats.ContainsKey($assigneeName)) {
                        $assigneeStats[$assigneeName]++
                    } else {
                        $assigneeStats[$assigneeName] = 1
                    }
                    $hasAssignee = $true
                }
            }
        }
        
        if (-not $hasAssignee -and $incident.acknowledgers -and $incident.acknowledgers.Count -gt 0) {
            foreach ($acknowledger in $incident.acknowledgers) {
                if ($acknowledger.summary) {
                    $assigneeName = $acknowledger.summary
                    if ($assigneeStats.ContainsKey($assigneeName)) {
                        $assigneeStats[$assigneeName]++
                    } else {
                        $assigneeStats[$assigneeName] = 1
                    }
                    $hasAssignee = $true
                    break
                }
            }
        }
        
        if (-not $hasAssignee -and $incident.assignees -and $incident.assignees.Count -gt 0) {
            foreach ($assignee in $incident.assignees) {
                if ($assignee.summary) {
                    $assigneeName = $assignee.summary
                    if ($assigneeStats.ContainsKey($assigneeName)) {
                        $assigneeStats[$assigneeName]++
                    } else {
                        $assigneeStats[$assigneeName] = 1
                    }
                    $hasAssignee = $true
                    break
                }
            }
        }
        
        if (-not $hasAssignee) {
            if ($assigneeStats.ContainsKey("Unassigned")) {
                $assigneeStats["Unassigned"]++
            } else {
                $assigneeStats["Unassigned"] = 1
            }
        }
    }
    
    $analytics.AssigneeDistribution = $assigneeStats.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
        @{ Assignee = $_.Key; Count = $_.Value; Percentage = [math]::Round(($_.Value / $Incidents.Count) * 100, 2) }
    }
    
    # Resolution time analysis
    $resolvedIncidents = $Incidents | Where-Object { $_.status -eq "resolved" -and $_.resolved_at -and $_.created_at }
    
    if ($resolvedIncidents.Count -gt 0) {
        $resolutionTimes = $resolvedIncidents | ForEach-Object {
            ([DateTime]$_.resolved_at - [DateTime]$_.created_at).TotalMinutes
        }
        $analytics.AvgResolutionTimeMinutes = [math]::Round(($resolutionTimes | Measure-Object -Average).Average, 2)
        $analytics.MedianResolutionTimeMinutes = [math]::Round(($resolutionTimes | Sort-Object)[[math]::Floor($resolutionTimes.Count / 2)], 2)
    } else {
        $analytics.AvgResolutionTimeMinutes = 0
        $analytics.MedianResolutionTimeMinutes = 0
    }
    
    # Acknowledgment time analysis
    $acknowledgedIncidents = $Incidents | Where-Object { 
        ($_.acknowledgements -and $_.acknowledgements.Count -gt 0) -or $_.first_acknowledgment
    }
    
    if ($acknowledgedIncidents.Count -gt 0) {
        $ackTimes = @()
        
        foreach ($incident in $acknowledgedIncidents) {
            try {
                $createdTime = [DateTime]$incident.created_at
                $ackTime = $null
                
                if ($incident.first_acknowledgment -and $incident.first_acknowledgment.at) {
                    $ackTime = [DateTime]$incident.first_acknowledgment.at
                }
                elseif ($incident.acknowledgements -and $incident.acknowledgements.Count -gt 0 -and $incident.acknowledgements[0].at) {
                    $ackTime = [DateTime]$incident.acknowledgements[0].at
                }
                
                if ($ackTime) {
                    $timeDiff = ($ackTime - $createdTime).TotalMinutes
                    if ($timeDiff -gt 0 -and $timeDiff -lt 10080) {
                        $ackTimes += $timeDiff
                    }
                }
            }
            catch {
                continue
            }
        }
        
        if ($ackTimes.Count -gt 0) {
            $analytics.AvgAckTimeMinutes = [math]::Round(($ackTimes | Measure-Object -Average).Average, 2)
            $analytics.MedianAckTimeMinutes = [math]::Round(($ackTimes | Sort-Object)[[math]::Floor($ackTimes.Count / 2)], 2)
        } else {
            $analytics.AvgAckTimeMinutes = 0
            $analytics.MedianAckTimeMinutes = 0
        }
    } else {
        $analytics.AvgAckTimeMinutes = 0
        $analytics.MedianAckTimeMinutes = 0
    }
    
    # Daily incident trend
    $dailyStats = $Incidents | Group-Object { ([DateTime]$_.created_at).Date.ToString("yyyy-MM-dd") } | Sort-Object Name
    $analytics.DailyTrend = $dailyStats | ForEach-Object {
        @{ Date = $_.Name; Count = $_.Count }
    }
    
    # Hourly distribution
    $hourlyStats = $Incidents | Group-Object { ([DateTime]$_.created_at).Hour } | Sort-Object { [int]$_.Name }
    $analytics.HourlyDistribution = $hourlyStats | ForEach-Object {
        @{ Hour = [int]$_.Name; Count = $_.Count }
    }
    
    # Day of week distribution
    $dowStats = $Incidents | Group-Object { ([DateTime]$_.created_at).DayOfWeek } | Sort-Object { 
        switch ($_.Name) {
            "Monday" { 1 }; "Tuesday" { 2 }; "Wednesday" { 3 }; "Thursday" { 4 }
            "Friday" { 5 }; "Saturday" { 6 }; "Sunday" { 7 }
        }
    }
    $analytics.DayOfWeekDistribution = $dowStats | ForEach-Object {
        @{ DayOfWeek = $_.Name; Count = $_.Count }
    }
    
    return $analytics
}

function Generate-SOCReportHTML {
    param(
        [hashtable]$Analytics,
        [string]$ReportTitle,
        [string]$DateRange
    )
    
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$ReportTitle</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js"></script>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f7fa; color: #333; line-height: 1.6; }
        .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px; margin-bottom: 30px; text-align: center; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; font-weight: 300; }
        .header p { font-size: 1.2em; opacity: 0.9; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .stat-card { background: white; padding: 25px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); text-align: center; border-left: 4px solid #667eea; }
        .stat-number { font-size: 2.5em; font-weight: bold; color: #667eea; margin-bottom: 5px; }
        .stat-label { color: #666; font-size: 1.1em; }
        .section { background: white; margin-bottom: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); overflow: hidden; }
        .section-header { background: #667eea; color: white; padding: 20px; font-size: 1.3em; font-weight: 500; }
        .section-content { padding: 25px; }
        .chart-container { position: relative; height: 400px; margin-bottom: 20px; }
        .small-chart { height: 300px; }
        .table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        .table th, .table td { padding: 12px; text-align: left; border-bottom: 1px solid #eee; }
        .table th { background: #f8f9fa; font-weight: 600; color: #555; }
        .table tr:hover { background: #f8f9fa; }
        .percentage { font-weight: bold; color: #667eea; }
        .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 30px; }
        .grid-3 { display: grid; grid-template-columns: repeat(auto-fit, minmax(350px, 1fr)); gap: 30px; }
        .status-triggered { color: #dc3545; font-weight: bold; }
        .status-ack { color: #ffc107; font-weight: bold; }
        .status-resolved { color: #28a745; font-weight: bold; }
        .urgency-high { color: #dc3545; font-weight: bold; }
        .urgency-low { color: #6c757d; }
        .footer { text-align: center; padding: 20px; color: #666; border-top: 1px solid #eee; margin-top: 30px; }
        .time-metric { display: inline-block; margin: 10px 15px; padding: 15px; background: #f8f9fa; border-radius: 8px; border-left: 3px solid #667eea; }
        .time-value { font-size: 1.8em; font-weight: bold; color: #667eea; display: block; }
        .time-label { color: #666; font-size: 0.9em; }
        @media (max-width: 768px) { .grid-2, .grid-3 { grid-template-columns: 1fr; } }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🛡️ $ReportTitle</h1>
            <p>Analysis Period: $DateRange | Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        </div>

        <!-- Key Metrics -->
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-number">$($Analytics.TotalIncidents)</div>
                <div class="stat-label">Total Incidents</div>
            </div>
            <div class="stat-card">
                <div class="stat-number status-triggered">$($Analytics.TriggeredIncidents)</div>
                <div class="stat-label">Active (Triggered)</div>
            </div>
            <div class="stat-card">
                <div class="stat-number status-ack">$($Analytics.AcknowledgedIncidents)</div>
                <div class="stat-label">Acknowledged</div>
            </div>
            <div class="stat-card">
                <div class="stat-number status-resolved">$($Analytics.ResolvedIncidents)</div>
                <div class="stat-label">Resolved</div>
            </div>
        </div>

        <!-- Response Time Metrics -->
        <div class="section">
            <div class="section-header">⏱️ Response & Resolution Metrics</div>
            <div class="section-content">
                <div style="text-align: center;">
                    <div class="time-metric">
                        <span class="time-value">$($Analytics.AvgAckTimeMinutes)</span>
                        <span class="time-label">Avg Response Time (min)</span>
                    </div>
                    <div class="time-metric">
                        <span class="time-value">$($Analytics.MedianAckTimeMinutes)</span>
                        <span class="time-label">Median Response Time (min)</span>
                    </div>
                    <div class="time-metric">
                        <span class="time-value">$($Analytics.AvgResolutionTimeMinutes)</span>
                        <span class="time-label">Avg Resolution Time (min)</span>
                    </div>
                    <div class="time-metric">
                        <span class="time-value">$($Analytics.MedianResolutionTimeMinutes)</span>
                        <span class="time-label">Median Resolution Time (min)</span>
                    </div>
                </div>
            </div>
        </div>

        <!-- Incident Trend Chart -->
        <div class="section">
            <div class="section-header">📈 Incident Trend Over Time</div>
            <div class="section-content">
                <div class="chart-container">
                    <canvas id="trendChart"></canvas>
                </div>
            </div>
        </div>

        <div class="grid-2">
            <!-- Top Services -->
            <div class="section">
                <div class="section-header">🏢 Top Services</div>
                <div class="section-content">
                    <table class="table">
                        <thead>
                            <tr><th>Service</th><th>Incidents</th><th>%</th></tr>
                        </thead>
                        <tbody>
"@

    foreach ($service in $Analytics.TopServices) {
        $html += "<tr><td>$($service.Name)</td><td>$($service.Count)</td><td class='percentage'>$($service.Percentage)%</td></tr>"
    }

    $html += @"
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- Client Distribution -->
            <div class="section">
                <div class="section-header">🏢 Clients Affected</div>
                <div class="section-content">
                    <table class="table">
                        <thead>
                            <tr><th>Client</th><th>Incidents</th><th>%</th></tr>
                        </thead>
                        <tbody>
"@

    if ($Analytics.ClientDistribution -and $Analytics.ClientDistribution.Count -gt 0) {
        foreach ($client in ($Analytics.ClientDistribution | Select-Object -First 10)) {
            $escapedClient = $client.Client -replace '"', '&quot;' -replace '<', '&lt;' -replace '>', '&gt;'
            $html += "<tr><td>$escapedClient</td><td>$($client.Count)</td><td class='percentage'>$($client.Percentage)%</td></tr>"
        }
    } else {
        $html += "<tr><td colspan='3' style='text-align: center; color: #666;'>No client data available</td></tr>"
    }

    $html += @"
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- Assignee Distribution -->
        <div class="section">
            <div class="section-header">👥 Incident Distribution by Assignee</div>
            <div class="section-content">
                <div class="chart-container">
                    <canvas id="assigneeChart"></canvas>
                </div>
            </div>
        </div>

        <!-- Most Common Alerts -->
        <div class="section">
            <div class="section-header">🚨 Most Common Alert Types</div>
            <div class="section-content">
                <table class="table">
                    <thead>
                        <tr><th>Alert Title</th><th>Count</th><th>%</th></tr>
                    </thead>
                    <tbody>
"@

    foreach ($alert in $Analytics.TopAlertTitles) {
        $escapedTitle = $alert.Title -replace '"', '&quot;' -replace '<', '&lt;' -replace '>', '&gt;'
        $html += "<tr><td>$escapedTitle</td><td>$($alert.Count)</td><td class='percentage'>$($alert.Percentage)%</td></tr>"
    }

    $html += @"
                    </tbody>
                </table>
            </div>
        </div>

        <div class="grid-3">
            <!-- Urgency Distribution -->
            <div class="section">
                <div class="section-header">⚡ Urgency Distribution</div>
                <div class="section-content">
                    <div class="chart-container small-chart">
                        <canvas id="urgencyChart"></canvas>
                    </div>
                </div>
            </div>

            <!-- Hourly Distribution -->
            <div class="section">
                <div class="section-header">🕐 Incidents by Hour</div>
                <div class="section-content">
                    <div class="chart-container small-chart">
                        <canvas id="hourlyChart"></canvas>
                    </div>
                </div>
            </div>

            <!-- Day of Week Distribution -->
            <div class="section">
                <div class="section-header">📅 Incidents by Day of Week</div>
                <div class="section-content">
                    <div class="chart-container small-chart">
                        <canvas id="dowChart"></canvas>
                    </div>
                </div>
            </div>
        </div>

        <div class="footer">
            <p>Report generated by PagerDuty SOC Analytics Tool | Data as of $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        </div>
    </div>

    <script>
        // Trend Chart
        const trendCtx = document.getElementById('trendChart').getContext('2d');
        new Chart(trendCtx, {
            type: 'line',
            data: {
                labels: [$(($Analytics.DailyTrend | ForEach-Object { "'$($_.Date)'" }) -join ', ')],
                datasets: [{
                    label: 'Incidents per Day',
                    data: [$(($Analytics.DailyTrend | ForEach-Object { $_.Count }) -join ', ')],
                    borderColor: '#667eea',
                    backgroundColor: 'rgba(102, 126, 234, 0.1)',
                    borderWidth: 3,
                    fill: true,
                    tension: 0.4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { beginAtZero: true, grid: { color: '#f0f0f0' } },
                    x: { grid: { color: '#f0f0f0' } }
                }
            }
        });

        // Assignee Chart
        const assigneeCtx = document.getElementById('assigneeChart').getContext('2d');
        new Chart(assigneeCtx, {
            type: 'doughnut',
            data: {
                labels: [$(($Analytics.AssigneeDistribution | Select-Object -First 8 | ForEach-Object { "'$($_.Assignee)'" }) -join ', ')],
                datasets: [{
                    data: [$(($Analytics.AssigneeDistribution | Select-Object -First 8 | ForEach-Object { $_.Count }) -join ', ')],
                    backgroundColor: ['#667eea', '#764ba2', '#f093fb', '#f5576c', '#4facfe', '#00f2fe', '#43e97b', '#38f9d7']
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'bottom' } }
            }
        });

        // Urgency Chart
        const urgencyCtx = document.getElementById('urgencyChart').getContext('2d');
        new Chart(urgencyCtx, {
            type: 'pie',
            data: {
                labels: [$(($Analytics.UrgencyDistribution | ForEach-Object { "'$($_.Urgency)'" }) -join ', ')],
                datasets: [{
                    data: [$(($Analytics.UrgencyDistribution | ForEach-Object { $_.Count }) -join ', ')],
                    backgroundColor: ['#dc3545', '#ffc107']
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'bottom' } }
            }
        });

        // Hourly Distribution Chart
        const hourlyCtx = document.getElementById('hourlyChart').getContext('2d');
        new Chart(hourlyCtx, {
            type: 'bar',
            data: {
                labels: [$(($Analytics.HourlyDistribution | ForEach-Object { $_.Hour }) -join ', ')],
                datasets: [{
                    label: 'Incidents',
                    data: [$(($Analytics.HourlyDistribution | ForEach-Object { $_.Count }) -join ', ')],
                    backgroundColor: 'rgba(102, 126, 234, 0.6)',
                    borderColor: '#667eea',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { beginAtZero: true },
                    x: { title: { display: true, text: 'Hour of Day' } }
                }
            }
        });

        // Day of Week Chart
        const dowCtx = document.getElementById('dowChart').getContext('2d');
        new Chart(dowCtx, {
            type: 'bar',
            data: {
                labels: [$(($Analytics.DayOfWeekDistribution | ForEach-Object { "'$($_.DayOfWeek)'" }) -join ', ')],
                datasets: [{
                    label: 'Incidents',
                    data: [$(($Analytics.DayOfWeekDistribution | ForEach-Object { $_.Count }) -join ', ')],
                    backgroundColor: 'rgba(118, 75, 162, 0.6)',
                    borderColor: '#764ba2',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: { y: { beginAtZero: true } }
            }
        });
    </script>
</body>
</html>
"@

    return $html
}

# Main execution
Write-Host "Starting PagerDuty 30-Day SOC Report Generation..." -ForegroundColor Yellow

$endDate = Get-Date
$startDate = $endDate.AddDays(-30)
$dateRange = "$($startDate.ToString('yyyy-MM-dd')) to $($endDate.ToString('yyyy-MM-dd'))"
$reportTitle = "PagerDuty SOC Activity Report - Last 30 Days"

Write-Host "Collecting incident data for the past 30 days..." -ForegroundColor Cyan
$allIncidents = Get-PagerDutyIncidentData -StartDate $startDate -EndDate $endDate

if (-not $allIncidents -or $allIncidents.Count -eq 0) {
    Write-Warning "No incidents found for the specified time period."
    exit 1
}

Write-Host "Processing $($allIncidents.Count) incidents..." -ForegroundColor Cyan
$analytics = Get-IncidentAnalytics -Incidents $allIncidents

Write-Host "Generating HTML report..." -ForegroundColor Cyan
$htmlContent = Generate-SOCReportHTML -Analytics $analytics -ReportTitle $reportTitle -DateRange $dateRange

$reportPath = Join-Path $OutputPath "PagerDuty_SOC-Report.html"
$htmlContent | Out-File -FilePath $reportPath -Encoding UTF8

Write-Host "Report generated successfully!" -ForegroundColor Green
Write-Host "Report saved to: $reportPath" -ForegroundColor Cyan

try {
    Start-Process $reportPath
    Write-Host "Report opened in default browser." -ForegroundColor Green
}
catch {
    Write-Host "Report saved. Open manually: $reportPath" -ForegroundColor Yellow
}
