<#
.SYNOPSIS
    人事情報管理ツール SharePoint データ自動バックアップスクリプト v2.0

.DESCRIPTION
    Azure AD アプリ認証を使用して Microsoft Graph API 経由で
    SharePoint List 全件をバックアップします。
    Windows タスクスケジューラから定期実行することを想定。

.PARAMETER ConfigPath
    設定ファイル(JSON)のパス（既定: 同フォルダの backup-config.json）

.EXAMPLE
    .\Backup-HrSharePointData.ps1
    .\Backup-HrSharePointData.ps1 -ConfigPath "C:\HRBackup\config.json"

.NOTES
    Version : 2.0
    Author  : HR管理ツール
    Updated : 2026-06-26
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "backup-config.json")
)

# ============================================================
# 初期設定
# ============================================================
$ErrorActionPreference = 'Stop'
$script:StartTime = Get-Date
$script:ScriptVersion = "2.0"
$script:LogEntries = @()
$script:BackupResult = @{
    StartedAt    = $script:StartTime.ToString("o")
    Version      = $script:ScriptVersion
    Status       = 'running'
    Lists        = @()
    Errors       = @()
    TotalItems   = 0
    DurationSec  = 0
}

# ============================================================
# 関数定義
# ============================================================
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "[$timestamp] [$Level] $Message"
    Write-Host $entry
    $script:LogEntries += $entry
}

function Get-AccessToken {
    param($Config)
    Write-Log "Azure AD トークンを取得中..."
    $tokenBody = @{
        client_id     = $Config.ClientId
        client_secret = $Config.ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }
    $tokenUrl = "https://login.microsoftonline.com/$($Config.TenantId)/oauth2/v2.0/token"
    $response = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $tokenBody -ContentType 'application/x-www-form-urlencoded'
    Write-Log "トークン取得成功"
    return $response.access_token
}

function Get-SiteId {
    param($Token, $SiteUrl)
    $uri = [System.Uri]$SiteUrl
    $hostname = $uri.Host
    $path = $uri.AbsolutePath
    $apiUrl = "https://graph.microsoft.com/v1.0/sites/${hostname}:${path}"
    $headers = @{ Authorization = "Bearer $Token" }
    $response = Invoke-RestMethod -Uri $apiUrl -Method GET -Headers $headers
    Write-Log "SharePoint Site ID 取得: $($response.id)"
    return $response.id
}

function Get-ListItems {
    param($Token, $SiteId, $ListName)
    Write-Log "リスト「$ListName」のアイテム取得を開始..."
    $headers = @{ Authorization = "Bearer $Token" }
    # リスト ID 取得
    $encodedName = [System.Web.HttpUtility]::UrlEncode($ListName)
    $listsUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/lists?`$filter=displayName eq '$ListName'"
    $listsRes = Invoke-RestMethod -Uri $listsUrl -Method GET -Headers $headers
    if (-not $listsRes.value -or $listsRes.value.Count -eq 0) {
        Write-Log "リスト「$ListName」が見つかりません" 'WARN'
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
    Write-Log "リスト「$ListName」: $($allItems.Count) 件取得"
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
    Write-Log "保存: $fileName ($([Math]::Round($fileInfo.Length / 1KB, 2)) KB)"
    return @{
        Path     = $filePath
        FileName = $fileName
        SizeKB   = [Math]::Round($fileInfo.Length / 1KB, 2)
    }
}

function Compress-BackupFiles {
    param($SourcePath, $ZipPath)
    Write-Log "ZIP圧縮中: $ZipPath"
    if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }
    Compress-Archive -Path "$SourcePath\*" -DestinationPath $ZipPath -CompressionLevel Optimal
    $zipInfo = Get-Item $ZipPath
    Write-Log "ZIP作成完了: $([Math]::Round($zipInfo.Length / 1MB, 2)) MB"
    return $zipInfo.Length
}

function Remove-OldBackups {
    param($BackupRoot, $RetentionDays)
    Write-Log "古いバックアップを削除中（保持期間: $RetentionDays 日）..."
    $cutoff = (Get-Date).AddDays(-$RetentionDays)
    $removed = 0
    Get-ChildItem -Path $BackupRoot -Filter "*.zip" | Where-Object {
        $_.LastWriteTime -lt $cutoff
    } | ForEach-Object {
        Write-Log "  削除: $($_.Name)"
        Remove-Item $_.FullName -Force
        $removed++
    }
    # JSONフォルダも古いものを削除
    Get-ChildItem -Path $BackupRoot -Directory | Where-Object {
        $_.LastWriteTime -lt $cutoff -and $_.Name -match '^\d{8}_\d{6}$'
    } | ForEach-Object {
        Remove-Item $_.FullName -Recurse -Force
        $removed++
    }
    Write-Log "古いバックアップ削除: $removed 件"
}

function Send-TeamsNotification {
    param($WebhookUrl, $Result, $Config)
    if (-not $WebhookUrl) { return }
    Write-Log "Teams通知を送信中..."

    $severity = if ($Result.Status -eq 'success') { 'good' }
                elseif ($Result.Status -eq 'partial') { 'warning' }
                else { 'attention' }
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

    $facts = @(
        @{ name = '実行サーバー'; value = $env:COMPUTERNAME }
        @{ name = 'バックアップ対象'; value = "$($Result.Lists.Count) リスト" }
        @{ name = '総アイテム数'; value = "$($Result.TotalItems) 件" }
        @{ name = '所要時間'; value = "$($Result.DurationSec) 秒" }
        @{ name = '保存先'; value = $Config.BackupRoot }
    )
    if ($Result.Errors.Count -gt 0) {
        $facts += @{ name = 'エラー'; value = "$($Result.Errors.Count) 件発生" }
    }

    $payload = @{
        '@type' = 'MessageCard'
        '@context' = 'https://schema.org/extensions'
        themeColor = $themeColor
        summary = "HRデータバックアップ $($Result.Status)"
        sections = @(@{
            activityTitle = "$icon SharePoint データバックアップ"
            activitySubtitle = "人事情報管理システム / $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')"
            text = if ($Result.Status -eq 'success') { 'バックアップが正常に完了しました。' }
                   elseif ($Result.Status -eq 'partial') { '一部のリストでエラーが発生しました。' }
                   else { 'バックアップに失敗しました。ログを確認してください。' }
            facts = $facts
            markdown = $true
        })
    } | ConvertTo-Json -Depth 10

    try {
        Invoke-RestMethod -Uri $WebhookUrl -Method POST -Body $payload -ContentType 'application/json' | Out-Null
        Write-Log "Teams通知送信成功"
    } catch {
        Write-Log "Teams通知送信失敗: $($_.Exception.Message)" 'ERROR'
    }
}

function Save-StatusJson {
    param($Result, $StatusPath, $LogEntries)
    $Result.LogTail = $LogEntries | Select-Object -Last 50
    $Result.CompletedAt = (Get-Date).ToString("o")
    $Result | ConvertTo-Json -Depth 20 | Out-File -FilePath $StatusPath -Encoding UTF8
    Write-Log "ステータスファイル保存: $StatusPath"
}

# ============================================================
# メイン処理
# ============================================================
try {
    Write-Log "================================================"
    Write-Log "HR SharePoint バックアップ開始 v$script:ScriptVersion"
    Write-Log "================================================"

    # 設定読み込み
    if (-not (Test-Path $ConfigPath)) {
        throw "設定ファイルが見つかりません: $ConfigPath"
    }
    $config = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Write-Log "設定読み込み: $ConfigPath"

    # 必須パラメータチェック
    foreach ($required in @('TenantId','ClientId','ClientSecret','SiteUrl','ListName','BackupRoot')) {
        if (-not $config.$required) { throw "設定が不足: $required" }
    }

    # バックアップフォルダ準備
    if (-not (Test-Path $config.BackupRoot)) {
        New-Item -ItemType Directory -Path $config.BackupRoot -Force | Out-Null
    }
    $timestamp = $script:StartTime.ToString("yyyyMMdd_HHmmss")
    $sessionDir = Join-Path $config.BackupRoot $timestamp
    New-Item -ItemType Directory -Path $sessionDir -Force | Out-Null

    # トークン取得
    $token = Get-AccessToken -Config $config
    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
    # サイトID
    $siteId = Get-SiteId -Token $token -SiteUrl $config.SiteUrl

    # バックアップ対象リスト構築
    $targetLists = @(
        $config.ListName,
        "$($config.ListName)_マスター",
        "$($config.ListName)_職歴",
        "$($config.ListName)_異動履歴",
        "$($config.ListName)_アンケート",
        "$($config.ListName)_アンケート回答",
        "$($config.ListName)_申請",
        "$($config.ListName)_マイナンバー閲覧ログ"
    )
    if ($config.AdditionalLists) {
        $targetLists += $config.AdditionalLists
    }

    # 各リストをバックアップ
    foreach ($listName in $targetLists) {
        try {
            $listData = Get-ListItems -Token $token -SiteId $siteId -ListName $listName
            if ($listData) {
                $backupFile = Save-BackupFile -Data $listData -BackupRoot $sessionDir -ListName $listName -Timestamp $timestamp
                $script:BackupResult.Lists += @{
                    ListName = $listName
                    Count = $listData.Items.Count
                    File = $backupFile.FileName
                    SizeKB = $backupFile.SizeKB
                    Status = 'success'
                }
                $script:BackupResult.TotalItems += $listData.Items.Count
            }
        } catch {
            $errorMsg = "リスト「$listName」バックアップ失敗: $($_.Exception.Message)"
            Write-Log $errorMsg 'ERROR'
            $script:BackupResult.Errors += $errorMsg
            $script:BackupResult.Lists += @{
                ListName = $listName
                Count = 0
                Status = 'error'
                Error = $_.Exception.Message
            }
        }
    }

    # ZIP圧縮
    if ($config.CreateZip -eq $true) {
        $zipPath = Join-Path $config.BackupRoot "HR_Backup_$timestamp.zip"
        $zipSize = Compress-BackupFiles -SourcePath $sessionDir -ZipPath $zipPath
        $script:BackupResult.ZipFile = (Split-Path $zipPath -Leaf)
        $script:BackupResult.ZipSizeMB = [Math]::Round($zipSize / 1MB, 2)
        if ($config.DeleteJsonAfterZip -eq $true) {
            Remove-Item $sessionDir -Recurse -Force
            Write-Log "JSON個別ファイルを削除（ZIP保管のみ）"
        }
    }

    # 古いバックアップ削除
    if ($config.RetentionDays -and $config.RetentionDays -gt 0) {
        Remove-OldBackups -BackupRoot $config.BackupRoot -RetentionDays $config.RetentionDays
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
    Write-Log "================================================"
    Write-Log "バックアップ完了: $($script:BackupResult.Status)"
    Write-Log "総アイテム数: $($script:BackupResult.TotalItems)"
    Write-Log "所要時間: $($script:BackupResult.DurationSec) 秒"
    Write-Log "================================================"

} catch {
    $script:BackupResult.Status = 'failed'
    $script:BackupResult.Errors += $_.Exception.Message
    Write-Log "致命的エラー: $($_.Exception.Message)" 'ERROR'
    Write-Log $_.ScriptStackTrace 'ERROR'
} finally {
    # ステータスファイル保存
    $statusPath = Join-Path $config.BackupRoot "backup-status.json"
    Save-StatusJson -Result $script:BackupResult -StatusPath $statusPath -LogEntries $script:LogEntries

    # ログファイル保存
    $logPath = Join-Path $config.BackupRoot "backup-log_$($script:StartTime.ToString('yyyyMMdd_HHmmss')).log"
    $script:LogEntries | Out-File -FilePath $logPath -Encoding UTF8

    # Teams通知
    if ($config.TeamsWebhookUrl) {
        Send-TeamsNotification -WebhookUrl $config.TeamsWebhookUrl -Result $script:BackupResult -Config $config
    }

    # 終了コード
    if ($script:BackupResult.Status -eq 'success') { exit 0 }
    elseif ($script:BackupResult.Status -eq 'partial') { exit 1 }
    else { exit 2 }
}
