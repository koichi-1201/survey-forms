<#
.SYNOPSIS
    HR管理ツール 動作監視ヘルスチェック v1.0

.DESCRIPTION
    GitHub Actions上で実行される定期動作監視スクリプト。
    以下の項目をチェック：
    1. HR管理ツールHTMLの到達性
    2. Azure AD トークン取得
    3. SharePoint Site への接続
    4. 主要 SharePoint Lists へのアクセス
    5. リスト件数の異常変動検知

.NOTES
    実行環境: GitHub Actions (Ubuntu + PowerShell Core)
    Version: 1.0
    Updated: 2026-06-26
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'
$script:StartTime = Get-Date
$script:Result = @{
    StartedAt      = $script:StartTime.ToString("o")
    Version        = '1.0'
    OverallStatus  = 'unknown'
    Checks         = @()
    FailedChecks   = @()
    RunUrl         = $env:GH_RUN_URL
}

# ============================================================
# ユーティリティ
# ============================================================
function Add-CheckResult {
    param(
        [string]$Name,
        [string]$Status,
        [string]$Message,
        [object]$Detail = $null,
        [double]$DurationMs = 0
    )
    $entry = @{
        Name = $Name
        Status = $Status
        Message = $Message
        Detail = $Detail
        DurationMs = [Math]::Round($DurationMs, 1)
        CheckedAt = (Get-Date).ToString("o")
    }
    $script:Result.Checks += $entry
    if ($Status -ne 'success') {
        $script:Result.FailedChecks += $entry
    }
    $icon = switch ($Status) {
        'success' { '[OK]' }
        'warning' { '[WARN]' }
        default   { '[FAIL]' }
    }
    Write-Host "  $icon [$Name] $Message"
}

function Get-AccessToken {
    $body = @{
        client_id     = $env:AZURE_CLIENT_ID
        client_secret = $env:AZURE_CLIENT_SECRET
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    }
    $url = "https://login.microsoftonline.com/$env:AZURE_TENANT_ID/oauth2/v2.0/token"
    $res = Invoke-RestMethod -Uri $url -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded'
    return $res.access_token
}

# ============================================================
# チェック1: HR管理ツールHTML到達性
# ============================================================
function Test-HtmlReachability {
    Write-Host "[1] HR管理ツールHTML到達性チェック..."
    if (-not $env:HR_TOOL_URL) {
        Add-CheckResult -Name 'HTML到達性' -Status 'success' -Message 'HR_TOOL_URL 未設定（スキップ）'
        return
    }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $res = Invoke-WebRequest -Uri $env:HR_TOOL_URL -Method HEAD -TimeoutSec 30 -ErrorAction Stop
        $sw.Stop()
        if ($res.StatusCode -eq 200) {
            Add-CheckResult -Name 'HTML到達性' -Status 'success' -Message "200 OK (応答 $($sw.ElapsedMilliseconds)ms)" -DurationMs $sw.ElapsedMilliseconds
        } else {
            Add-CheckResult -Name 'HTML到達性' -Status 'warning' -Message "ステータス $($res.StatusCode)" -DurationMs $sw.ElapsedMilliseconds
        }
    } catch {
        $sw.Stop()
        Add-CheckResult -Name 'HTML到達性' -Status 'failed' -Message $_.Exception.Message -DurationMs $sw.ElapsedMilliseconds
    }
}

# ============================================================
# チェック2: Azure AD トークン取得
# ============================================================
function Test-AzureAdAuth {
    Write-Host "[2] Azure AD認証チェック..."
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $token = Get-AccessToken
        $sw.Stop()
        if ($token) {
            Add-CheckResult -Name 'Azure AD認証' -Status 'success' -Message "トークン取得成功 ($($sw.ElapsedMilliseconds)ms)" -DurationMs $sw.ElapsedMilliseconds
            return $token
        }
    } catch {
        $sw.Stop()
        Add-CheckResult -Name 'Azure AD認証' -Status 'failed' -Message $_.Exception.Message -DurationMs $sw.ElapsedMilliseconds
    }
    return $null
}

# ============================================================
# チェック3: SharePoint Site 接続
# ============================================================
function Test-SharePointSite {
    param($Token)
    Write-Host "[3] SharePoint Site接続チェック..."
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $uri = [System.Uri]$env:SP_SITE_URL
        $apiUrl = "https://graph.microsoft.com/v1.0/sites/$($uri.Host):$($uri.AbsolutePath)"
        $headers = @{ Authorization = "Bearer $Token" }
        $res = Invoke-RestMethod -Uri $apiUrl -Method GET -Headers $headers -TimeoutSec 30
        $sw.Stop()
        Add-CheckResult -Name 'SharePoint Site' -Status 'success' -Message "Site ID 取得成功 ($($sw.ElapsedMilliseconds)ms)" -Detail @{ SiteId = $res.id } -DurationMs $sw.ElapsedMilliseconds
        return $res.id
    } catch {
        $sw.Stop()
        Add-CheckResult -Name 'SharePoint Site' -Status 'failed' -Message $_.Exception.Message -DurationMs $sw.ElapsedMilliseconds
    }
    return $null
}

# ============================================================
# チェック4: 主要リストアクセス・件数取得
# ============================================================
function Test-SharePointLists {
    param($Token, $SiteId)
    Write-Host "[4] SharePoint Lists 件数チェック..."
    $headers = @{ Authorization = "Bearer $Token" }
    $listName = $env:SP_LIST_NAME
    $targetLists = @(
        $listName,
        "${listName}_マスター",
        "${listName}_職歴",
        "${listName}_異動履歴",
        "${listName}_アンケート",
        "${listName}_アンケート回答",
        "${listName}_申請",
        "${listName}_マイナンバー閲覧ログ"
    )
    $counts = @{}
    foreach ($name in $targetLists) {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $listsUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists?`$filter=displayName eq '$name'"
            $listsRes = Invoke-RestMethod -Uri $listsUrl -Method GET -Headers $headers -TimeoutSec 30
            if (-not $listsRes.value -or $listsRes.value.Count -eq 0) {
                Add-CheckResult -Name "リスト[$name]" -Status 'warning' -Message 'リストが存在しない（未作成）'
                continue
            }
            $listId = $listsRes.value[0].id
            $itemsUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$listId/items?`$top=999"
            $itemsRes = Invoke-RestMethod -Uri $itemsUrl -Method GET -Headers $headers -TimeoutSec 30
            $count = $itemsRes.value.Count
            $counts[$name] = $count
            $sw.Stop()
            Add-CheckResult -Name "リスト[$name]" -Status 'success' -Message "$count 件 ($($sw.ElapsedMilliseconds)ms)" -Detail @{ Count = $count } -DurationMs $sw.ElapsedMilliseconds
        } catch {
            $sw.Stop()
            Add-CheckResult -Name "リスト[$name]" -Status 'failed' -Message $_.Exception.Message -DurationMs $sw.ElapsedMilliseconds
        }
    }
    return $counts
}

# ============================================================
# チェック5: 件数の異常変動検知
# ============================================================
function Test-CountAnomaly {
    param($Token, $SiteId, $CurrentCounts)
    Write-Host "[5] 件数変動チェック..."
    try {
        $headers = @{ Authorization = "Bearer $Token" }
        $masterListName = "$env:SP_LIST_NAME`_マスター"
        $listsUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists?`$filter=displayName eq '$masterListName'"
        $listsRes = Invoke-RestMethod -Uri $listsUrl -Method GET -Headers $headers
        if (-not $listsRes.value) { Add-CheckResult -Name '件数変動' -Status 'warning' -Message 'マスターリストなし'; return }
        $masterListId = $listsRes.value[0].id
        $itemsUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$masterListId/items?expand=fields&`$top=999"
        $itemsRes = Invoke-RestMethod -Uri $itemsUrl -Method GET -Headers $headers
        $prev = $itemsRes.value | Where-Object { $_.fields.MasterKey -eq 'toolHealth' } | Select-Object -First 1
        if (-not $prev) {
            Add-CheckResult -Name '件数変動' -Status 'success' -Message '初回実行（変動なし）'
            return
        }
        $prevData = $prev.fields.MasterValue | ConvertFrom-Json
        $prevCounts = $prevData.LastCounts
        if (-not $prevCounts) {
            Add-CheckResult -Name '件数変動' -Status 'success' -Message '前回件数情報なし'
            return
        }
        $anomalies = @()
        foreach ($key in $CurrentCounts.Keys) {
            $prev = $prevCounts.$key
            $cur = $CurrentCounts[$key]
            if ($prev -gt 10 -and $cur -lt ($prev * 0.7)) {
                $anomalies += "[$key] $prev -> $cur (-$(([math]::Round((1-$cur/$prev)*100,0)))%)"
            }
        }
        if ($anomalies.Count -gt 0) {
            Add-CheckResult -Name '件数変動' -Status 'failed' -Message "急減検知: $($anomalies -join ', ')" -Detail @{ Anomalies = $anomalies }
        } else {
            Add-CheckResult -Name '件数変動' -Status 'success' -Message '異常変動なし'
        }
    } catch {
        Add-CheckResult -Name '件数変動' -Status 'warning' -Message "チェック失敗: $($_.Exception.Message)"
    }
}

# ============================================================
# SharePoint への記録
# ============================================================
function Save-HealthStatus {
    param($Token, $SiteId, $Counts)
    Write-Host "[6] SharePoint へステータス記録..."
    try {
        $headers = @{ Authorization = "Bearer $Token"; 'Content-Type' = 'application/json' }
        $masterListName = "$env:SP_LIST_NAME`_マスター"
        $listsUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists?`$filter=displayName eq '$masterListName'"
        $listsRes = Invoke-RestMethod -Uri $listsUrl -Method GET -Headers $headers
        if (-not $listsRes.value) { Write-Warning 'マスターリストなし'; return }
        $masterListId = $listsRes.value[0].id

        $script:Result.CompletedAt = (Get-Date).ToString("o")
        $script:Result.DurationSec = [Math]::Round(((Get-Date) - $script:StartTime).TotalSeconds, 1)
        $script:Result.LastCounts = $Counts

        $failed = $script:Result.FailedChecks.Count
        if ($failed -eq 0) {
            $script:Result.OverallStatus = 'healthy'
        } elseif ($failed -le 2) {
            $script:Result.OverallStatus = 'degraded'
        } else {
            $script:Result.OverallStatus = 'unhealthy'
        }

        $valueJson = ($script:Result | ConvertTo-Json -Depth 20 -Compress)

        $itemsUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$masterListId/items?expand=fields&`$top=999"
        $itemsRes = Invoke-RestMethod -Uri $itemsUrl -Method GET -Headers $headers
        $existing = $itemsRes.value | Where-Object { $_.fields.MasterKey -eq 'toolHealth' } | Select-Object -First 1

        $fieldsObj = @{
            MasterKey = 'toolHealth'
            MasterValue = $valueJson
            UpdatedAt = (Get-Date).ToString("o")
            Title = 'toolHealth'
        }
        $fieldsJson = $fieldsObj | ConvertTo-Json -Depth 5

        if ($existing) {
            $url = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$masterListId/items/$($existing.id)/fields"
            Invoke-RestMethod -Uri $url -Method PATCH -Headers $headers -Body $fieldsJson | Out-Null
            Write-Host "  toolHealth ステータス更新"
        } else {
            $url = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$masterListId/items"
            $body = @{ fields = $fieldsObj } | ConvertTo-Json -Depth 5
            Invoke-RestMethod -Uri $url -Method POST -Headers $headers -Body $body | Out-Null
            Write-Host "  toolHealth ステータス新規作成"
        }

        $historyExisting = $itemsRes.value | Where-Object { $_.fields.MasterKey -eq 'toolHealthHistory' } | Select-Object -First 1
        $history = @()
        if ($historyExisting -and $historyExisting.fields.MasterValue) {
            try { $history = $historyExisting.fields.MasterValue | ConvertFrom-Json } catch { $history = @() }
            if ($history -isnot [array]) { $history = @($history) }
        }
        $entry = @{
            CompletedAt = $script:Result.CompletedAt
            OverallStatus = $script:Result.OverallStatus
            FailedChecks = $script:Result.FailedChecks.Count
            TotalChecks = $script:Result.Checks.Count
            DurationSec = $script:Result.DurationSec
        }
        $history = @($entry) + $history | Select-Object -First 100

        $historyFields = @{
            MasterKey = 'toolHealthHistory'
            MasterValue = ($history | ConvertTo-Json -Depth 10 -Compress)
            UpdatedAt = (Get-Date).ToString("o")
            Title = 'toolHealthHistory'
        }
        $historyJson = $historyFields | ConvertTo-Json -Depth 5

        if ($historyExisting) {
            $url = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$masterListId/items/$($historyExisting.id)/fields"
            Invoke-RestMethod -Uri $url -Method PATCH -Headers $headers -Body $historyJson | Out-Null
        } else {
            $url = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$masterListId/items"
            $body = @{ fields = $historyFields } | ConvertTo-Json -Depth 5
            Invoke-RestMethod -Uri $url -Method POST -Headers $headers -Body $body | Out-Null
        }
        Write-Host "  履歴を記録（最新100件保持）"
    } catch {
        Write-Warning "ステータス記録エラー: $($_.Exception.Message)"
    }
}

# ============================================================
# Teams通知
# ============================================================
function Send-HealthNotification {
    if (-not $env:TEAMS_WEBHOOK_URL) { return }
    if ($script:Result.OverallStatus -eq 'healthy') { return }

    $themeColor = if ($script:Result.OverallStatus -eq 'degraded') { 'FFA500' } else { 'DC3545' }
    $label = if ($script:Result.OverallStatus -eq 'degraded') { '一部障害' } else { '重大障害' }

    $failedList = ($script:Result.FailedChecks | ForEach-Object { "- [$($_.Name)] $($_.Message)" }) -join "`n"

    $payload = @{
        '@type' = 'MessageCard'
        '@context' = 'https://schema.org/extensions'
        themeColor = $themeColor
        summary = "HR ヘルスチェック $label"
        sections = @(@{
            activityTitle = "HR管理ツール ヘルスチェック異常"
            activitySubtitle = "GitHub Actions / $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')"
            text = "総合判定: **$($script:Result.OverallStatus.ToUpper())**`n`n**失敗項目:**`n$failedList"
            facts = @(
                @{ name = '実行環境'; value = 'GitHub Actions (Cloud)' }
                @{ name = '総チェック数'; value = "$($script:Result.Checks.Count) 件" }
                @{ name = '失敗数'; value = "$($script:Result.FailedChecks.Count) 件" }
                @{ name = '所要時間'; value = "$($script:Result.DurationSec) 秒" }
                @{ name = 'ログURL'; value = $env:GH_RUN_URL }
            )
            markdown = $true
        })
    } | ConvertTo-Json -Depth 10
    try {
        Invoke-RestMethod -Uri $env:TEAMS_WEBHOOK_URL -Method POST -Body $payload -ContentType 'application/json' | Out-Null
        Write-Host "  Teams通知送信成功"
    } catch {
        Write-Warning "Teams通知失敗: $($_.Exception.Message)"
    }
}

# ============================================================
# メイン
# ============================================================
Write-Host "================================================"
Write-Host "HR管理ツール ヘルスチェック 開始"
Write-Host "================================================"

$outputDir = Join-Path $PWD 'health-output'
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }

$required = @('AZURE_TENANT_ID','AZURE_CLIENT_ID','AZURE_CLIENT_SECRET','SP_SITE_URL','SP_LIST_NAME')
foreach ($var in $required) {
    if (-not (Get-Item env:$var -ErrorAction SilentlyContinue)) {
        Write-Error "環境変数が未設定: $var"
        $script:Result.OverallStatus = 'failed'
        $script:Result | ConvertTo-Json -Depth 20 | Out-File (Join-Path $outputDir 'health-status.json') -Encoding UTF8
        exit 2
    }
}

Test-HtmlReachability
$token = Test-AzureAdAuth
if ($token) {
    $siteId = Test-SharePointSite -Token $token
    if ($siteId) {
        $counts = Test-SharePointLists -Token $token -SiteId $siteId
        Test-CountAnomaly -Token $token -SiteId $siteId -CurrentCounts $counts
        Save-HealthStatus -Token $token -SiteId $siteId -Counts $counts
    }
}

$failed = $script:Result.FailedChecks.Count
if ($failed -eq 0) {
    $script:Result.OverallStatus = 'healthy'
} elseif ($failed -le 2) {
    $script:Result.OverallStatus = 'degraded'
} else {
    $script:Result.OverallStatus = 'unhealthy'
}
$script:Result.CompletedAt = (Get-Date).ToString("o")
$script:Result.DurationSec = [Math]::Round(((Get-Date) - $script:StartTime).TotalSeconds, 1)

$script:Result | ConvertTo-Json -Depth 20 | Out-File (Join-Path $outputDir 'health-status.json') -Encoding UTF8

Send-HealthNotification

Write-Host "================================================"
Write-Host "ヘルスチェック完了"
Write-Host "総合判定: $($script:Result.OverallStatus.ToUpper())"
Write-Host "失敗項目: $failed / $($script:Result.Checks.Count)"
Write-Host "所要時間: $($script:Result.DurationSec) 秒"
Write-Host "================================================"

if ($script:Result.OverallStatus -eq 'healthy') { exit 0 }
elseif ($script:Result.OverallStatus -eq 'degraded') { exit 1 }
else { exit 2 }
