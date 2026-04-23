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
Write-Host "3. Dang day code va tag len GitHub de tu dong Build..." -ForegroundColor Cyan
git add .
git commit -m "Release v$NewVersion"
git push origin main
git tag "v$NewVersion"
git push origin "v$NewVersion"

Write-Host "=============================================" -ForegroundColor Green
Write-Host "HOAN TAT QUY TRINH!" -ForegroundColor Green
Write-Host "1. Code da duoc day len GitHub." -ForegroundColor White
Write-Host "2. GitHub Actions dang tu dong build ban 'Single EXE' tai tab Actions." -ForegroundColor Yellow
Write-Host "3. Sau vai phut, ban se thay file ZIP san sang trong phan Release!" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green

Pause
