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

# Buoc 2: Xac nhan tai file ISO Office
Wait-Yes "➡️ Ban co muon tai file Office .img tu Microsoft?"
$url = "https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-us/ProPlusRetail.img"
$out = "$temp\ProPlusRetail.img"
Write-Host "⬇️ Dang tai file Office (.img)..."
Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing

if (-not (Test-Path $out)) {
    Write-Host "❌ Loi khi tai file Office." -ForegroundColor Red
    Read-Host "Nhan Enter de thoat..."
    exit
}

# Buoc 3: Mount file IMG
Wait-Yes "➡️ Mount file .img?"
$mount = Mount-DiskImage -ImagePath $out -PassThru
$driveLetter = ($mount | Get-Volume).DriveLetter + ":"

# Buoc 4: Kiem tra setup va config.xml
$setup = "$driveLetter\setup.exe"
$config = "$driveLetter\config.xml"
if (-not (Test-Path $setup) -or -not (Test-Path $config)) {
    Write-Host "❌ Khong tim thay setup.exe hoac config.xml trong file IMG." -ForegroundColor Red
    Read-Host "Nhan Enter de thoat..."
    Dismount-DiskImage -ImagePath $out
    exit
}

# Buoc 5: Cai dat
Wait-Yes "➡️ Bat dau cai dat Office?"
Write-Host "🚀 Dang tien hanh cai dat..."
Start-Process -FilePath $setup -ArgumentList "/configure config.xml" -Wait

# Buoc 6: Dismount file IMG
Dismount-DiskImage -ImagePath $out

Write-Host "`n✅ Da hoan tat qua trinh cai dat Office!" -ForegroundColor Green
Read-Host "Nhan Enter de ket thuc..."
