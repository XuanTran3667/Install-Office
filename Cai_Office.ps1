# Cai_Office.ps1 - Script cai dat Office 2021 Pro Plus tu dong

# Ham cho phep nguoi dung nhan Y/N truoc khi tiep tuc
function Wait-Yes($msg) {
    do {
        $a = Read-Host "$msg (Y/N)"
        if ($a -match '^[Yy]$') { return $true }
        elseif ($a -match '^[Nn]$') {
            Write-Host "❌ Da huy buoc nay." -ForegroundColor Red
            exit
        }
    } while ($true)
}

# Kiem tra quyen Admin chinh xac
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Host "`n❌ Vui long chay PowerShell bang quyen Administrator!" -ForegroundColor Red
    Read-Host "Nhan Enter de thoat..."
    exit
}

Write-Host "`n=== 🚀 BAT DAU CAI OFFICE 2021 PRO PLUS ===`n" -ForegroundColor Cyan

# Buoc 1: Tao thu muc tam
$temp = "$env:TEMP\office_auto"
if (Test-Path $temp) {
    Remove-Item -Recurse -Force $temp
}
New-Item -ItemType Directory -Path $temp | Out-Null
Set-Location $temp

# Buoc 2: Kiem tra WinRAR
$rar = "$env:ProgramFiles\WinRAR\WinRAR.exe"
if (-not (Test-Path $rar)) {
    Write-Host "❌ Khong tim thay WinRAR. Hay cai dat WinRAR truoc." -ForegroundColor Red
    Read-Host "Nhan Enter de thoat..."
    exit
}

# Buoc 3: Xac nhan tai file ISO Office tu Google Drive (su dung fileId)
Wait-Yes "➡️ Ban co muon tai file Office (.img) tu Google Drive?"

function Download-GDriveFile($fileId, $destination) {
    $confirmUrl = "https://docs.google.com/uc?export=download&id=$fileId"
    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $response = Invoke-WebRequest -Uri $confirmUrl -WebSession $session -UseBasicParsing

    if ($response.Content -match 'confirm=([0-9A-Za-z-_]+)') {
        $confirm = $matches[1]
        $downloadUrl = "$confirmUrl&confirm=$confirm"
        Invoke-WebRequest -Uri $downloadUrl -WebSession $session -OutFile $destination -UseBasicParsing
    } else {
        Invoke-WebRequest -Uri $confirmUrl -WebSession $session -OutFile $destination -UseBasicParsing
    }
}

$fileId = "1o4q5VKDpVKifveXv-CTxn5GW9W-omxJz"
$out = "$temp\ProPlusRetail.img"
Write-Host "⬇️ Dang tai file Office (.img)..."
Download-GDriveFile -fileId $fileId -destination $out

if (-not (Test-Path $out)) {
    Write-Host "❌ Loi khi tai file Office." -ForegroundColor Red
    Read-Host "Nhan Enter de thoat..."
    exit
}

# Buoc 4: Giai nen file img bang WinRAR
Wait-Yes "➡️ Giai nen file .img bang WinRAR?"
Write-Host "📦 Dang giai nen..."
Start-Process -FilePath $rar -ArgumentList "x -o+ ProPlusRetail.img *.*" -Wait

# Buoc 5: Kiem tra setup
$setup = "$temp\setup.exe"
$config = "$temp\config.xml"
if (-not (Test-Path $setup) -or -not (Test-Path $config)) {
    Write-Host "❌ Khong tim thay setup.exe hoac config.xml." -ForegroundColor Red
    Read-Host "Nhan Enter de thoat..."
    exit
}

# Buoc 6: Cai dat
Wait-Yes "➡️ Bat dau cai dat Office?"
Write-Host "🚀 Dang tien hanh cai dat..."
Start-Process -FilePath $setup -ArgumentList "/configure config.xml" -Wait

Write-Host "`n✅ Da hoan tat qua trinh cai dat Office!" -ForegroundColor Green
Read-Host "Nhan Enter de ket thuc..."
