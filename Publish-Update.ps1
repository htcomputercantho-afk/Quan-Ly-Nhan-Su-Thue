param (
    [Parameter(Mandatory=$false)]
    [string]$NewVersion
)

if (-not $NewVersion) {
    $NewVersion = Read-Host "Nhap phien ban moi (Vi du: 1.0.1.0)"
}

if ([string]::IsNullOrWhiteSpace($NewVersion)) {
    Write-Host "Phien ban khong hop le. Huy bo!" -ForegroundColor Red
    exit
}

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "BAT DAU QUY TRINH PHAT HANH BAN CAP NHAT $NewVersion" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan

$csprojPath = "Quan Ly Nhan Su\Quan_Ly_Nhan_Su.csproj"
$xmlPath = "update.xml"
$repoOwner = "htcomputercantho-afk"
$repoName = "Quan-Ly-Nhan-Su-Thue"

# 1. Cap nhat file .csproj
Write-Host "1. Cap nhat phien ban trong .csproj..."
$csprojContent = Get-Content $csprojPath -Raw
$csprojContent = $csprojContent -replace '<Version>.*?</Version>', "<Version>$NewVersion</Version>"
Set-Content -Path $csprojPath -Value $csprojContent -Encoding UTF8

# 2. Cap nhat file update.xml
Write-Host "2. Cap nhat thong tin trong update.xml..."
$xmlContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<item>
    <version>$NewVersion</version>
    <url>https://github.com/$repoOwner/$repoName/releases/download/v$NewVersion/QuanLyNhanSu_v$NewVersion.zip</url>
    <changelog>https://github.com/$repoOwner/$repoName/releases</changelog>
    <mandatory>true</mandatory>
</item>
"@
Set-Content -Path $xmlPath -Value $xmlContent -Encoding UTF8

# 3. Git Add, Commit, Tag, Push
Write-Host "3. Day source code va tag len GitHub de kich hoat Auto Build..."
git add .
git commit -m "Release v$NewVersion"
git tag "v$NewVersion"
git push origin main
git push origin "v$NewVersion"

Write-Host "=============================================" -ForegroundColor Green
Write-Host "HOAN TAT! GitHub Actions dang chay tren server..." -ForegroundColor Green
Write-Host "Ban co the vao tab Actions tren GitHub de xem qua trinh Build & Zip tu dong." -ForegroundColor Green
Write-Host "Ung dung se tu dong tai ban $NewVersion nay!" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Pause
