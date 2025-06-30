# 必須以系統管理員身份執行
# 瑞思派送後可驗證此檔案是否存在來判斷腳本有無異常  C:\WINDOWS\Temp\upgrade7z_ok
# 執行log在 C:\Windows\Temp\upgrade7z.log
# 2025/06/16 : 改版，加上驗證檔案
# 2025/06/27 : 登錄檔找不到就直接找常見路徑，避免system身分看不到Local User的軟體清單
# 2025/06/27 : 調整log寫入方式，使用Out-File替換掉原本的Add-Content，避免寫入失敗
# 2025/06/27 : 重複片段重構成func


# ===[ 宣告變數 ]===
$log = "$env:TEMP\upgrade7z.log"


# ===[ 宣告函數 ]===
# 輸出/寫入log
function Write-Log {
    param (
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
    Out-File -FilePath "$env:TEMP\upgrade7z.log" -Append -InputObject $Message
}

# 取得最新版本
function Get-7ZipLatestVersion {
	try {
		$response = Invoke-WebRequest "https://www.7-zip.org/" -UseBasicParsing
	} catch {
		Write-Log "❌ 無法連線 7-Zip 官網，結束。" "Red"
		return $false
	}

	# 抓取最新版版本號
	if ($response.Content -match "Download 7-Zip ([\d\.]+)") {
		$latestVersion = $Matches[1]
		return $latestVersion
	} else {
		Write-Log "❌ 無法取得最新版資訊。" "Red"
		return $false
	}
}

# 取得本地7Zip資訊
function Get-7ZipInstalledInfo {
	# 註冊表位置
	$registryPaths = @(
		"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\",
		"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
	)
	# 實體常見路徑
	$possiblePaths = @(
		"C:\\Program Files\\7-Zip\\7zFM.exe",
		"C:\\Program Files (x86)\\7-Zip\\7zFM.exe"
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
		# Write-Log "\n`n🔍❌ 登錄檔找不到 7-Zip，改用實體檔案偵測..."

		foreach ($exe in $possiblePaths) {
			if (Test-Path $exe) {
				
				$installType = if (Test-Path "$installLocation\\Uninstall.exe") { "exe" } else { "msi" }
				$installLocation = Split-Path $exe
				$installedVersion = (Get-Item $exe).VersionInfo.ProductVersion
				$uninstallString = if (Test-Path "$installLocation\\Uninstall.exe") { "$installLocation\\Uninstall.exe" } else { "msiexec  /I{XXXX}" }
				
				$sevenZip = @{ DisplayVersion = $installedVersion; InstallLocation = $installLocation; UninstallString = $uninstallString }
				# Write-Log "`n🔍📂 從實體路徑判斷已安裝 7-Zip，類型: $installType, 路徑: $installLocation"
				break
			}
		}
	}

	return $sevenZip
}


# ===[預處理]===

# 重設log
Out-File -FilePath "$env:TEMP\upgrade7z.log" -InputObject ("開始執行於 " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
Write-Host "開始執行於 " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")

$account = whoami
Write-Log "Account: $account" 

$runAsAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$status = if ($runAsAdmin) { "Administrator (UAC elevated)" } else { "Not Administrator" }
Write-Log "User Status: $status"


# 清掉之前驗證檔案
Write-Log "清掉之前驗證檔案"

$flagOkPath = "$($env:SystemRoot)\Temp\upgrade7z_ok"
if (Test-Path $flagOkPath) {
    try {
        Remove-Item -Path "$($env:SystemRoot)\Temp\upgrade7z_ok" -Force
    } catch {
        Write-Log "⚠️ 無法刪除舊驗證檔案，請用系統管理員身份執行。" "Red"
        exit
    }
}

# ===[主要邏輯]===


Write-Log "`n🔍 檢查是否安裝 7-Zip..." "Cyan"

$sevenZip = Get-7ZipInstalledInfo

if (-not $sevenZip) {
    Write-Log "❌ 未偵測到 7-Zip，結束。" "Yellow"
    # 新增驗證檔案
    New-Item -Path "$flagOkPath" -ItemType File | Out-Null
    exit
} 

# 解析原安裝路徑
$installLocation = $sevenZip.InstallLocation
if (-not $installLocation) {
    # 若無 InstallLocation，從 UninstallString 判斷
    if ($installType -eq "exe" -and $sevenZip.UninstallString -match '"(.+?)\\Uninstall\.exe') {
        $installLocation = $matches[1]
    } elseif ($installType -eq "msi") {
        # MSI 通常無明確 InstallLocation，可用預設值推測，MSI用不到此值
        $installLocation = if ([Environment]::Is64BitOperatingSystem) { "C:\Program Files\7-Zip" } else { "C:\Program Files (x86)\7-Zip" }
    }
}

$installedVersion = $sevenZip.DisplayVersion
$installType = if ($sevenZip.UninstallString -match "msiexec") { "msi" } else { "exe" }

Write-Log "✅ 已安裝 7-Zip $installedVersion ($installType)"
Write-Log "📂 安裝路徑: $installLocation"
Write-Log "🌐 嘗試連線 7-Zip 官方網站..."


# 抓取最新版版本號
$latestVersion = Get-7ZipLatestVersion
if ($latestVersion) {
    Write-Host  "🌟 最新版本：$latestVersion"
}
else {
    Write-Log ("❌ 取得版本失敗，結束於"+(Get-Date -Format "yyyy-MM-dd HH:mm:ss")) "Red"
}


if ($installedVersion -eq $latestVersion) {
    Write-Log "🆗 已是最新版，無需更新。" "Green"
    # 新增驗證檔案
    New-Item -Path "$flagOkPath" -ItemType File | Out-Null
    exit
}


# 組合下載網址與檔名
$versionNoDot = $latestVersion.Replace(".", "")
$extension = if ($installType -eq "msi") { "msi" } else { "exe" }
# 判斷架構
$arch = if ([Environment]::Is64BitOperatingSystem) { "x64" } else { "x86" }
$filename = "7z$versionNoDot-$arch.$extension"
$url = "https://www.7-zip.org/a/$filename"
$tempPath = "$env:TEMP\$filename"

# 先刪除舊檔案（若存在）
if (Test-Path $tempPath) {
    Write-Log  "🗑 刪除舊檔案 $tempPath..."
    try {
        Remove-Item -Path $tempPath -Force
    } catch {
        Write-Log "⚠️ 無法刪除檔案，可能被鎖定，請稍後重試。" "Red"
        exit
    }
}

Write-Log "⬇️ 下載 $url ..."
Invoke-WebRequest -Uri $url -OutFile $tempPath

if (-not (Test-Path $tempPath)) {
    Write-Log "❌ 下載失敗。" "Red"
    exit
}

Write-Log "🚀 執行靜默安裝..."
if ($extension -eq "msi") {
    $proc = Start-Process msiexec.exe -ArgumentList "/i `"$tempPath`" /quiet /norestart " -Wait
	if ($proc.ExitCode -eq 0) {
		Write-Log "`n⏳ MSI安裝成功"
	} else {
		if ($proc.ExitCod -eq 1618) {
			Write-Log "⚠️ 正在進行其他安裝，請稍後重試，錯誤碼：$($proc.ExitCode)" "Red"
			exit
		} 
		if ($proc.ExitCode -and $proc.ExitCode -gt 0) {
			Write-Log "`n⏳ 安裝異常，錯誤碼：$($proc.ExitCode)" "Red"
			exit
		}
	}
} else {
    Start-Process -FilePath $tempPath -ArgumentList "/S /D=""$installLocation""" -Wait
}

# 安裝完後 → 動態等待 Uninstall 註冊表 DisplayVersion 更新成功

$maxWaitSeconds = 120   # 最長等待 120 秒
$waitInterval = 5      # 每次檢查間隔 5 秒
$elapsed = 0

Write-Log "`n⏳ 等待 7-Zip 完成安裝..."

do {
    Start-Sleep -Seconds $waitInterval
    $elapsed += $waitInterval

	$newSevenZip = Get-7ZipInstalledInfo
	$newVersion = "等待中..."
	if ($newSevenZip) {
        $newVersion = ($newSevenZip.DisplayVersion -split '\.')[0..1] -join '.'
    } 

    Write-Log "   -> 目前版本: $newVersion"

    if ($newVersion -eq $latestVersion) {
        break
    }

} while ($elapsed -lt $maxWaitSeconds)

# 最終結果判斷
if ($newVersion -eq $latestVersion) {
    Write-Log  "🎉 更新成功！目前版本: $newVersion" "Green"
    # 新增驗證檔案
    New-Item -Path "$flagOkPath" -ItemType File | Out-Null
} else {
    Write-Log "⚠️ 等待超時 ($maxWaitSeconds 秒)，更新版本尚未確認成功，請手動檢查。" "Red"
}

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"


Write-Log "✅ 更新完成於 $timestamp" "Green"
