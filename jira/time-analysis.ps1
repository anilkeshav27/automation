# =============================================================================
# Jira Time-in-Status per Developer
# =============================================================================
# Fetches all issues matching a JQL query, reads each issue's changelog,
# and computes how long each developer spent in each status transition.
#
# Output: CSV file with columns:
#   Developer, IssueKey, FromStatus, ToStatus, TimeSpentHours, TransitionDate
#
# Setup: set these environment variables before running
#
#   $env:JIRA_BASE_URL   = "https://your-jira.example.com"
#   $env:JIRA_USERNAME   = "your-username"
#   $env:JIRA_PASSWORD   = "your-password-or-token"
#   $env:JIRA_PROJECTS   = "PROJECT1,PROJECT2,PROJECT3"
#   $env:JIRA_COMPONENTS = "component-a,component-b"    # optional, leave empty to skip
#   $env:JIRA_STATUSES   = "In Progress,In Review,Done" # optional, leave empty to skip
#   $env:JIRA_DEVELOPERS = "dev1@example.com,dev2@example.com"  # optional, leave empty to skip
#   $env:OUTPUT_FILE     = "time_in_status.csv"         # optional, default: time_in_status.csv
#   $env:MAX_RESULTS     = "500"                        # optional, default: 500
#
# Then run:
#   .\jira_time_in_status.ps1
# =============================================================================

# =============================================================================
# Read config from environment variables
# =============================================================================
$JiraBaseUrl = $env:JIRA_BASE_URL
$Username    = $env:JIRA_USERNAME
$Password    = $env:JIRA_PASSWORD
$OutputFile  = if ($env:OUTPUT_FILE)  { $env:OUTPUT_FILE }      else { "time_in_status.csv" }
$MaxResults  = if ($env:MAX_RESULTS)  { [int]$env:MAX_RESULTS } else { 500 }

# Parse comma-separated lists into arrays
function Parse-EnvList {
    param([string]$Value)
    if (-not $Value -or $Value -eq "") { return @() }
    return $Value -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
}

$Projects   = Parse-EnvList $env:JIRA_PROJECTS
$Components = Parse-EnvList $env:JIRA_COMPONENTS
$Statuses   = Parse-EnvList $env:JIRA_STATUSES
$Developers = Parse-EnvList $env:JIRA_DEVELOPERS

# =============================================================================
# Validate required variables
# =============================================================================
$missing = @()
if (-not $JiraBaseUrl)      { $missing += "JIRA_BASE_URL" }
if (-not $Username)         { $missing += "JIRA_USERNAME" }
if (-not $Password)         { $missing += "JIRA_PASSWORD" }
if ($Projects.Count -eq 0)  { $missing += "JIRA_PROJECTS" }

if ($missing.Count -gt 0) {
    Write-Error "Missing required environment variables: $($missing -join ', ')"
    exit 1
}

# =============================================================================
# Build JQL
# =============================================================================
$projectList = ($Projects | ForEach-Object { "`"$_`"" }) -join ", "
$jqlParts    = @("project in ($projectList)")

if ($Components.Count -gt 0) {
    $componentList = ($Components | ForEach-Object { "`"$_`"" }) -join ", "
    $jqlParts += "component in ($componentList)"
}

if ($Statuses.Count -gt 0) {
    $statusList = ($Statuses | ForEach-Object { "`"$_`"" }) -join ", "
    $jqlParts += "status in ($statusList)"
}

if ($Developers.Count -gt 0) {
    $devList = ($Developers | ForEach-Object { "`"$_`"" }) -join ", "
    $jqlParts += "assignee in ($devList)"
}

$jqlParts += "updated >= startOfWeek()"

$JQL = $jqlParts -join " AND "
Write-Host "JQL: $JQL" -ForegroundColor Cyan

# =============================================================================
# Auth header (Basic Auth)
# =============================================================================
$pair         = "${Username}:${Password}"
$encodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))
$headers      = @{
    "Authorization" = "Basic $encodedCreds"
    "Accept"        = "application/json"
    "Content-Type"  = "application/json"
}

# =============================================================================
# Fetch issues with changelog (paginated)
# =============================================================================
function Fetch-Issues {
    param([string]$Jql, [int]$MaxResults)

    $allIssues = @()
    $startAt   = 0
    $pageSize  = 100

    do {
        $url = "$JiraBaseUrl/rest/api/2/search" +
               "?jql=$([Uri]::EscapeDataString($Jql))" +
               "&startAt=$startAt" +
               "&maxResults=$pageSize" +
               "&fields=summary,assignee,status" +
               "&expand=changelog"

        try {
            $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        } catch {
            Write-Error "Failed to fetch issues: $_"
            exit 1
        }

        $batch    = $response.issues
        $allIssues += $batch
        $total    = $response.total
        $startAt  += $batch.Count

        Write-Host "  Fetched $startAt / $total issues..."

    } while ($startAt -lt $total -and $startAt -lt $MaxResults -and $batch.Count -gt 0)

    Write-Host "Total issues fetched: $($allIssues.Count)" -ForegroundColor Green
    return $allIssues
}

# =============================================================================
# Parse status transitions from changelog
# =============================================================================
function Parse-Transitions {
    param($Issue)

    $transitions = @()
    $histories   = $Issue.changelog.histories

    foreach ($history in $histories) {
        $author    = $history.author.displayName
        $createdAt = $history.created

        foreach ($item in $history.items) {
            if ($item.field -ne "status") { continue }

            $transitions += [PSCustomObject]@{
                Developer    = $author
                IssueKey     = $Issue.key
                FromStatus   = $item.fromString
                ToStatus     = $item.toString
                TransitionAt = $createdAt
            }
        }
    }

    $transitions = $transitions | Sort-Object TransitionAt
    return $transitions
}

# =============================================================================
# Compute time spent in each status (gap between consecutive transitions)
# =============================================================================
function Compute-TimeSpent {
    param($Transitions)

    $rows = @()

    for ($i = 0; $i -lt $Transitions.Count; $i++) {
        $t = $Transitions[$i]

        if ($i -eq 0) {
            $timeSpentHours = "N/A"
        } else {
            $prevTime       = [datetime]$Transitions[$i - 1].TransitionAt
            $currentTime    = [datetime]$t.TransitionAt
            $delta          = $currentTime - $prevTime
            $timeSpentHours = [math]::Round($delta.TotalHours, 2)
        }

        $rows += [PSCustomObject]@{
            Developer      = $t.Developer
            IssueKey       = $t.IssueKey
            FromStatus     = $t.FromStatus
            ToStatus       = $t.ToStatus
            TimeSpentHours = $timeSpentHours
            TransitionDate = $t.TransitionAt
        }
    }

    return $rows
}

# =============================================================================
# Print summary table in terminal
# =============================================================================
function Print-Summary {
    param($Rows)

    $summary = @{}

    foreach ($row in $Rows) {
        if ($row.TimeSpentHours -eq "N/A") { continue }
        $key = "$($row.Developer)|$($row.FromStatus)"
        if (-not $summary.ContainsKey($key)) { $summary[$key] = 0.0 }
        $summary[$key] += [double]$row.TimeSpentHours
    }

    Write-Host "`n--- Summary: Total hours per developer per status ---" -ForegroundColor Yellow
    Write-Host ("{0,-35} {1,-25} {2,12}" -f "Developer", "Status", "Total Hours")
    Write-Host ("-" * 75)

    foreach ($key in ($summary.Keys | Sort-Object)) {
        $parts  = $key -split "\|"
        $dev    = $parts[0]
        $status = $parts[1]
        $hours  = [math]::Round($summary[$key], 2)
        Write-Host ("{0,-35} {1,-25} {2,12}" -f $dev, $status, $hours)
    }
}

# =============================================================================
# Main
# =============================================================================
$issues  = Fetch-Issues -Jql $JQL -MaxResults $MaxResults
$allRows = @()

foreach ($issue in $issues) {
    $transitions = Parse-Transitions -Issue $issue
    $rows        = Compute-TimeSpent -Transitions $transitions
    $allRows    += $rows
}

if ($allRows.Count -eq 0) {
    Write-Host "No transitions found." -ForegroundColor Red
    exit 0
}

$allRows | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
Write-Host "`nCSV written to: $OutputFile  ($($allRows.Count) rows)" -ForegroundColor Green

Print-Summary -Rows $allRows