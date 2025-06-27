# 必須以系統管理員身份執行
# 瑞思派送後可驗證此檔案是否存在來判斷腳本有無異常  C:\WINDOWS\Temp\upgrade7z_ok
# 執行log在 C:\Windows\Temp\upgrade7z.log


$log = "$env:TEMP\upgrade7z.log"
Get-Date -Format "yyyy-MM-dd HH:mm:ss" | Out-File $log

$account = whoami
"Account: $account" | Out-File $log -Append

$runAsAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$status = if ($runAsAdmin) { "Administrator (UAC elevated)" } else { "Not Administrator" }
"User Status: $status" | Out-File $log -Append



# 清掉之前驗證檔案
Add-Content -Path $log -Value "清掉之前驗證檔案"

$flagOkPath = "$($env:SystemRoot)\Temp\upgrade7z_ok"
if (Test-Path $flagOkPath) {
    try {
        Remove-Item -Path "$($env:SystemRoot)\Temp\upgrade7z_ok" -Force
    } catch {
        Write-Host "⚠️ 無法刪除舊驗證檔案，請用系統管理員身份執行。" -ForegroundColor Red
        Add-Content -Path $log -Value "⚠️ 無法刪除舊驗證檔案，請用系統管理員身份執行。"
        exit
    }
}


# 檢查註冊表找 7-Zip 安裝資訊
Write-Host "`n🔍 檢查是否安裝 7-Zip..." -ForegroundColor Cyan
Add-Content -Path $log -Value "`n🔍 檢查是否安裝 7-Zip..."
$registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
)

$sevenZip = $null
foreach ($path in $registryPaths) {
    Get-ChildItem $path -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $key = Get-ItemProperty $_.PSPath
            if ($key.DisplayName -like "7-Zip*") {
                $sevenZip = $key
            }
        } catch {}
    }
}

# 如果登錄找不到就用常見路徑的實體檔檢測
if (-not $sevenZip) {
    Write-Host "`n🔍❌ 登錄檔找不到 7-Zip，改用實體檔案偵測..."
    Add-Content $log -Value "\n`n🔍❌ 登錄檔找不到 7-Zip，改用實體檔案偵測..."
    $possiblePaths = @(
        "C:\\Program Files\\7-Zip\\7zFM.exe",
        "C:\\Program Files (x86)\\7-Zip\\7zFM.exe"
    )

    foreach ($exe in $possiblePaths) {
        if (Test-Path $exe) {
            $installSource = if ($exe -like "*x86*") { "exe (x86)" } else { "exe (x64)" }
            $installLocation = Split-Path $exe
            $installedVersion = (Get-Item $exe).VersionInfo.ProductVersion
            $sevenZip = @{ DisplayVersion = $installedVersion; InstallLocation = $installLocation; UninstallString = "$installLocation\\Uninstall.exe" }
            Write-Host "`n🔍📂 從實體路徑判斷已安裝 7-Zip，類型: $installSource, 路徑: $installLocation"
            Add-Content $log -Value "`n🔍📂 從實體路徑判斷已安裝 7-Zip，類型: $installSource, 路徑: $installLocation"
            break
        }
    }
}

if (-not $sevenZip) {
    Write-Host "❌ 未偵測到 7-Zip，結束。" -ForegroundColor Yellow
    Add-Content -Path $log -Value "❌ 未偵測到 7-Zip，結束。"
    # 新增驗證檔案
    New-Item -Path "$flagOkPath" -ItemType File
    exit
} 

$installedVersion = $sevenZip.DisplayVersion
$installSource = if ($sevenZip.UninstallString -match "msiexec") { "msi" } else { "exe" }

# 解析原安裝路徑
$installLocation = $sevenZip.InstallLocation
if (-not $installLocation) {
    # 若無 InstallLocation，從 UninstallString 判斷
    if ($installSource -eq "exe" -and $sevenZip.UninstallString -match '"(.+?)\\Uninstall\.exe') {
        $installLocation = $matches[1]
    } elseif ($installSource -eq "msi") {
        # MSI 通常無明確 InstallLocation，可用預設值推測，MSI用不到此值
        $installLocation = if ([Environment]::Is64BitOperatingSystem) { "C:\Program Files\7-Zip" } else { "C:\Program Files (x86)\7-Zip" }
    }
}


Write-Host "✅ 已安裝 7-Zip $installedVersion ($installSource)"
Write-Host "📂 安裝路徑: $installLocation"
Write-Host "🌐 嘗試連線 7-Zip 官方網站..."

Add-Content -Path $log -Value "✅ 已安裝 7-Zip $installedVersion ($installSource)"
Add-Content -Path $log -Value "📂 安裝路徑: $installLocation"
Add-Content -Path $log -Value "🌐 嘗試連線 7-Zip 官方網站..."


try {
    $response = Invoke-WebRequest "https://www.7-zip.org/" -UseBasicParsing
} catch {
    Write-Host "❌ 無法連線 7-Zip 官網，結束。" -ForegroundColor Red
    Add-Content -Path $log -Value "❌ 無法連線 7-Zip 官網，結束。"
    exit
}

# 抓取最新版版本號
if ($response.Content -match "Download 7-Zip ([\d\.]+)") {
    $latestVersion = $Matches[1]
    Write-Host "🌟 最新版本：$latestVersion"
    Add-Content -Path $log -Value "🌟 最新版本：$latestVersion"
} else {
    Write-Host "❌ 無法取得最新版資訊。" -ForegroundColor Red
    Add-Content -Path $log -Value "❌ 無法取得最新版資訊。"
    exit
}

if ($installedVersion -eq $latestVersion) {
    Write-Host "🆗 已是最新版，無需更新。" -ForegroundColor Green
    Add-Content -Path $log -Value "🆗 已是最新版，無需更新。"
    # 新增驗證檔案
    New-Item -Path "$flagOkPath" -ItemType File
    exit
}

# 判斷架構
$arch = if ([Environment]::Is64BitOperatingSystem) { "x64" } else { "x86" }

# 組合下載網址與檔名
$versionNoDot = $latestVersion.Replace(".", "")
$extension = if ($installSource -eq "msi") { "msi" } else { "exe" }
$filename = "7z$versionNoDot-$arch.$extension"
$url = "https://www.7-zip.org/a/$filename"
$tempPath = "$env:TEMP\$filename"

# 先刪除舊檔案（若存在）
if (Test-Path $tempPath) {
    Write-Host "🗑 刪除舊檔案 $tempPath..."
    Add-Content -Path $log -Value "🗑 刪除舊檔案 $tempPath..."
    try {
        Remove-Item -Path $tempPath -Force
    } catch {
        Write-Host "⚠️ 無法刪除檔案，可能被鎖定，請稍後重試。" -ForegroundColor Red
        Add-Content -Path $log -Value "⚠️ 無法刪除檔案，可能被鎖定，請稍後重試。"
        exit
    }
}

Write-Host "⬇️ 下載 $url ..."
Add-Content -Path $log -Value "⬇️ 下載 $url ..."
Invoke-WebRequest -Uri $url -OutFile $tempPath

if (-not (Test-Path $tempPath)) {
    Write-Host "❌ 下載失敗。" -ForegroundColor Red
    Add-Content -Path $log -Value "❌ 下載失敗。"
    exit
}

Write-Host "🚀 執行靜默安裝..."
Add-Content -Path $log -Value "🚀 執行靜默安裝..."
if ($extension -eq "msi") {
    Start-Process msiexec.exe -ArgumentList "/i `"$tempPath`" /quiet /norestart" -Wait
} else {
    Start-Process -FilePath $tempPath -ArgumentList "/S /D=""$installLocation""" -Wait
}

# 安裝完後 → 動態等待 Uninstall 註冊表 DisplayVersion 更新成功

$maxWaitSeconds = 120   # 最長等待 120 秒
$waitInterval = 5      # 每次檢查間隔 5 秒
$elapsed = 0

Write-Host "`n⏳ 等待 7-Zip 完成安裝..."
Add-Content -Path $log -Value "`n⏳ 等待 7-Zip 完成安裝..."

do {
    Start-Sleep -Seconds $waitInterval
    $elapsed += $waitInterval

    # 重新讀取註冊表版本
    $newSevenZip = $null
    foreach ($path in $registryPaths) {
		Get-ChildItem $path -ErrorAction SilentlyContinue | ForEach-Object {
			try {
				$key = Get-ItemProperty $_.PSPath
				if ($key.DisplayName -like "7-Zip*") {
					$newSevenZip = $key
				}
			} catch {}
		}
    }
	if ($newSevenZip) {
        $newVersion = $newSevenZip.DisplayVersion
    } else {
		$newVersion = $null
		$newVersion = (Get-Item $exe).VersionInfo.ProductVersion
	} 

	
    Write-Host "   -> 目前版本: $newVersion (等待中...)"
    Add-Content -Path $log -Value "   -> 目前版本: $newVersion (等待中...)"

    if ($newVersion -eq $latestVersion) {
        break
    }

} while ($elapsed -lt $maxWaitSeconds)

# 最終結果判斷
if ($newVersion -eq $latestVersion) {
    Write-Host "🎉 更新成功！目前版本: $newVersion" -ForegroundColor Green
    Add-Content -Path $log -Value "🎉 更新成功！目前版本: $newVersion"
    # 新增驗證檔案
    New-Item -Path "$flagOkPath" -ItemType File
} else {
    Write-Host "⚠️ 等待超時 ($maxWaitSeconds 秒)，更新版本尚未確認成功，請手動檢查。" -ForegroundColor Red
    Add-Content -Path $log -Value "⚠️ 等待超時 ($maxWaitSeconds 秒)，更新版本尚未確認成功，請手動檢查。"
}

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"


Write-Host "✅ 更新完成！" -ForegroundColor Green
Add-Content -Path $log -Value "✅ 更新完成！"
Add-Content -Path $log -Value $timestamp
