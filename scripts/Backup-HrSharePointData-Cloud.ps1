<#
.SYNOPSIS
    HR管理ツール SharePoint 自動バックアップ（GitHub Actions版）v2.0 Cloud

.DESCRIPTION
    GitHub Actions 上で動作するクラウドネイティブ版バックアップスクリプト。
    環境変数から設定を取得し、Azure AD アプリ認証で SharePoint Lists をバックアップします。

.NOTES
    実行環境: GitHub Actions (Ubuntu latest + PowerShell Core)
    認証: Azure AD アプリ認証
    保存先: GitHub Actions Artifacts + (任意で SharePoint Document Library)
    Version: 2.0 Cloud
    Updated: 2026-06-26
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$script:StartTime = Get-Date
$script:LogEntries = @()
$script:BackupResult = @{
    StartedAt    = $script:StartTime.ToString("o")
    Version      = '2.0 Cloud'
    Status       = 'running'
    Lists        = @()
    Errors       = @()
    TotalItems   = 0
    DurationSec  = 0
    ServerName   = 'GitHub Actions'
}

# ============================================================
# ユーティリティ関数
# ============================================================
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "[$ts] [$Level] $Message"
    Write-Host $entry
    $script:LogEntries += $entry
}

function Get-AccessToken {
    Write-Log "Azure AD トークンを取得中..."
    $body = @{
        client_id     = $env:AZURE_CLIENT_ID
        client_secret = $env:AZURE_CLIENT_SECRET
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    }
    $url = "https://login.microsoftonline.com/$env:AZURE_TENANT_ID/oauth2/v2.0/token"
    $response = Invoke-RestMethod -Uri $url -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded'
    Write-Log "トークン取得成功"
    return $response.access_token
}

function Get-SiteId {
    param($Token, $SiteUrl)
    $uri = [System.Uri]$SiteUrl
    $apiUrl = "https://graph.microsoft.com/v1.0/sites/$($uri.Host):$($uri.AbsolutePath)"
    $headers = @{ Authorization = "Bearer $Token" }
    $response = Invoke-RestMethod -Uri $apiUrl -Method GET -Headers $headers
    Write-Log "Site ID 取得: $($response.id)"
    return $response.id
}

function Get-ListItems {
    param($Token, $SiteId, $ListName)
    Write-Log "リスト [$ListName] のバックアップ開始..."
    $headers = @{ Authorization = "Bearer $Token" }

    # リスト ID 取得
    $listsUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists?`$filter=displayName eq '$ListName'"
    $listsRes = Invoke-RestMethod -Uri $listsUrl -Method GET -Headers $headers
    if (-not $listsRes.value -or $listsRes.value.Count -eq 0) {
        Write-Log "リスト [$ListName] が見つかりません" 'WARN'
        return $null
    }
    $listId = $listsRes.value[0].id

    # 全アイテム取得（ページング対応）
    $allItems = @()
    $nextUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$listId/items?expand=fields&`$top=999"
    while ($nextUrl) {
        $itemsRes = Invoke-RestMethod -Uri $nextUrl -Method GET -Headers $headers
        $allItems += $itemsRes.value
        $nextUrl = $itemsRes.'@odata.nextLink'
    }
    Write-Log "リスト [$ListName]: $($allItems.Count) 件取得完了"
    return @{
        ListId   = $listId
        ListName = $ListName
        Items    = $allItems
    }
}

function Save-BackupFile {
    param($Data, $BackupRoot, $ListName, $Timestamp)
    $safeFileName = $ListName -replace '[\\/:*?"<>|]', '_'
    $fileName = "${safeFileName}_${Timestamp}.json"
    $filePath = Join-Path $BackupRoot $fileName
    $Data | ConvertTo-Json -Depth 20 -Compress | Out-File -FilePath $filePath -Encoding UTF8
    $fileInfo = Get-Item $filePath
    return @{
        Path     = $filePath
        FileName = $fileName
        SizeKB   = [Math]::Round($fileInfo.Length / 1KB, 2)
    }
}

function Update-StatusInSharePoint {
    param($Token, $SiteId, $ListName, $StatusObj)
    Write-Log "SharePoint Listへのステータス記録..."
    $headers = @{
        Authorization = "Bearer $Token"
        'Content-Type' = 'application/json'
    }
    # マスターリストの ID を取得
    $masterListName = "${ListName}_マスター"
    $listsUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists?`$filter=displayName eq '$masterListName'"
    $listsRes = Invoke-RestMethod -Uri $listsUrl -Method GET -Headers $headers
    if (-not $listsRes.value) {
        Write-Log "マスターリストが見つかりません" 'WARN'
        return
    }
    $masterListId = $listsRes.value[0].id

    # backupStatus キーの既存アイテムを検索
    $itemsUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$masterListId/items?expand=fields&`$top=999"
    $itemsRes = Invoke-RestMethod -Uri $itemsUrl -Method GET -Headers $headers
    $existing = $itemsRes.value | Where-Object { $_.fields.MasterKey -eq 'backupStatus' } | Select-Object -First 1

    $valueJson = ($StatusObj | ConvertTo-Json -Depth 20 -Compress)
    $fields = @{
        MasterKey   = 'backupStatus'
        MasterValue = $valueJson
        UpdatedAt   = (Get-Date).ToString("o")
        Title       = 'backupStatus'
    } | ConvertTo-Json -Depth 5

    if ($existing) {
        # 更新
        $url = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$masterListId/items/$($existing.id)/fields"
        Invoke-RestMethod -Uri $url -Method PATCH -Headers $headers -Body $fields | Out-Null
        Write-Log "ステータスを更新しました"
    } else {
        # 新規作成
        $url = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$masterListId/items"
        $body = @{ fields = ($fields | ConvertFrom-Json) } | ConvertTo-Json -Depth 5
        Invoke-RestMethod -Uri $url -Method POST -Headers $headers -Body $body | Out-Null
        Write-Log "ステータスを新規作成しました"
    }

    # 履歴に追記
    $historyExisting = $itemsRes.value | Where-Object { $_.fields.MasterKey -eq 'backupHistory' } | Select-Object -First 1
    $history = @()
    if ($historyExisting -and $historyExisting.fields.MasterValue) {
        try { $history = $historyExisting.fields.MasterValue | ConvertFrom-Json } catch { $history = @() }
        if ($history -isnot [array]) { $history = @($history) }
    }
    $history = @($StatusObj) + $history | Select-Object -First 100
    $historyFields = @{
        MasterKey   = 'backupHistory'
        MasterValue = ($history | ConvertTo-Json -Depth 20 -Compress)
        UpdatedAt   = (Get-Date).ToString("o")
        Title       = 'backupHistory'
    } | ConvertTo-Json -Depth 5
    if ($historyExisting) {
        $url = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$masterListId/items/$($historyExisting.id)/fields"
        Invoke-RestMethod -Uri $url -Method PATCH -Headers $headers -Body $historyFields | Out-Null
    } else {
        $url = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$masterListId/items"
        $body = @{ fields = ($historyFields | ConvertFrom-Json) } | ConvertTo-Json -Depth 5
        Invoke-RestMethod -Uri $url -Method POST -Headers $headers -Body $body | Out-Null
    }
    Write-Log "履歴を保存しました（最新100件保持）"
}

function Send-TeamsNotification {
    param($WebhookUrl, $Result)
    if (-not $WebhookUrl) {
        Write-Log "Teams Webhook URL が設定されていないため通知をスキップ" 'INFO'
        return
    }
    $themeColor = switch ($Result.Status) {
        'success' { '28A745' }
        'partial' { 'FFA500' }
        default   { 'DC3545' }
    }
    $icon = switch ($Result.Status) {
        'success' { '✅' }
        'partial' { '⚠️' }
        default   { '🚨' }
    }
    $payload = @{
        '@type'    = 'MessageCard'
        '@context' = 'https://schema.org/extensions'
        themeColor = $themeColor
        summary    = "HRデータバックアップ $($Result.Status)"
        sections   = @(@{
            activityTitle    = "$icon SharePoint データバックアップ"
            activitySubtitle = "GitHub Actions / $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')"
            facts            = @(
                @{ name = '実行環境'; value = 'GitHub Actions (Cloud)' }
                @{ name = 'バックアップ対象'; value = "$($Result.Lists.Count) リスト" }
                @{ name = '総アイテム数'; value = "$($Result.TotalItems) 件" }
                @{ name = '所要時間'; value = "$($Result.DurationSec) 秒" }
                @{ name = 'エラー'; value = "$($Result.Errors.Count) 件" }
            )
            markdown         = $true
        })
    } | ConvertTo-Json -Depth 10
    try {
        Invoke-RestMethod -Uri $WebhookUrl -Method POST -Body $payload -ContentType 'application/json' | Out-Null
        Write-Log "Teams通知送信成功"
    } catch {
        Write-Log "Teams通知失敗: $($_.Exception.Message)" 'ERROR'
    }
}

# ============================================================
# メイン処理
# ============================================================
try {
    Write-Log "================================================"
    Write-Log "HR SharePoint バックアップ開始 (Cloud版)"
    Write-Log "================================================"

    # 環境変数チェック
    foreach ($var in @('AZURE_TENANT_ID','AZURE_CLIENT_ID','AZURE_CLIENT_SECRET','SP_SITE_URL','SP_LIST_NAME')) {
        if (-not (Get-Item env:$var -ErrorAction SilentlyContinue)) {
            throw "環境変数が未設定: $var"
        }
    }

    # バックアップ出力ディレクトリ
    $backupRoot = if ($env:BACKUP_FOLDER) { $env:BACKUP_FOLDER } else { Join-Path $PWD 'backup-output' }
    if (-not (Test-Path $backupRoot)) { New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null }

    $timestamp = $script:StartTime.ToString("yyyyMMdd_HHmmss")
    $sessionDir = Join-Path $backupRoot $timestamp
    New-Item -ItemType Directory -Path $sessionDir -Force | Out-Null

    # トークン取得
    $token = Get-AccessToken
    $siteId = Get-SiteId -Token $token -SiteUrl $env:SP_SITE_URL

    # バックアップ対象リスト
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

    # 各リストをバックアップ
    foreach ($name in $targetLists) {
        try {
            $listData = Get-ListItems -Token $token -SiteId $siteId -ListName $name
            if ($listData) {
                $backupFile = Save-BackupFile -Data $listData -BackupRoot $sessionDir -ListName $name -Timestamp $timestamp
                $script:BackupResult.Lists += @{
                    ListName = $name
                    Count    = $listData.Items.Count
                    File     = $backupFile.FileName
                    SizeKB   = $backupFile.SizeKB
                    Status   = 'success'
                }
                $script:BackupResult.TotalItems += $listData.Items.Count
            }
        } catch {
            $errMsg = "リスト [$name] バックアップ失敗: $($_.Exception.Message)"
            Write-Log $errMsg 'ERROR'
            $script:BackupResult.Errors += $errMsg
            $script:BackupResult.Lists += @{
                ListName = $name
                Count    = 0
                Status   = 'error'
                Error    = $_.Exception.Message
            }
        }
    }

    # 結果判定
    if ($script:BackupResult.Errors.Count -eq 0) {
        $script:BackupResult.Status = 'success'
    } elseif ($script:BackupResult.Lists | Where-Object { $_.Status -eq 'success' }) {
        $script:BackupResult.Status = 'partial'
    } else {
        $script:BackupResult.Status = 'failed'
    }
    $script:BackupResult.DurationSec = [Math]::Round(((Get-Date) - $script:StartTime).TotalSeconds, 1)
    $script:BackupResult.CompletedAt = (Get-Date).ToString("o")

    Write-Log "================================================"
    Write-Log "バックアップ完了: $($script:BackupResult.Status)"
    Write-Log "総アイテム数: $($script:BackupResult.TotalItems)"
    Write-Log "所要時間: $($script:BackupResult.DurationSec) 秒"
    Write-Log "================================================"

    # SharePoint への状態記録
    try {
        Update-StatusInSharePoint -Token $token -SiteId $siteId -ListName $listName -StatusObj $script:BackupResult
    } catch {
        Write-Log "SharePoint状態記録に失敗（バックアップ自体は成功）: $($_.Exception.Message)" 'WARN'
    }

} catch {
    $script:BackupResult.Status = 'failed'
    $script:BackupResult.Errors += $_.Exception.Message
    Write-Log "致命的エラー: $($_.Exception.Message)" 'ERROR'
    Write-Log $_.ScriptStackTrace 'ERROR'
} finally {
    # ステータスJSON保存
    $script:BackupResult.LogTail = $script:LogEntries | Select-Object -Last 50
    $statusPath = Join-Path $backupRoot "backup-status.json"
    $script:BackupResult | ConvertTo-Json -Depth 20 | Out-File -FilePath $statusPath -Encoding UTF8
    Write-Log "ステータスファイル: $statusPath"

    # ログ保存
    $logPath = Join-Path $backupRoot "backup-log_$timestamp.log"
    $script:LogEntries | Out-File -FilePath $logPath -Encoding UTF8

    # Teams通知
    if ($env:TEAMS_WEBHOOK_URL) {
        Send-TeamsNotification -WebhookUrl $env:TEAMS_WEBHOOK_URL -Result $script:BackupResult
    }

    # 終了コード
    if ($script:BackupResult.Status -eq 'success') { exit 0 }
    elseif ($script:BackupResult.Status -eq 'partial') { exit 1 }
    else { exit 2 }
}
