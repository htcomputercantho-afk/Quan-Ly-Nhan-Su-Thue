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
Write-Host "3. Day source code va tag len GitHub..." -ForegroundColor Cyan
git add .
git commit -m "Release v$NewVersion"
git push origin main
git tag "v$NewVersion"
git push origin "v$NewVersion"

# 4. Build ban 'Sieu Sach' tai local (De phong GitHub build ra file rac)
Write-Host "4. Dang build ban 'Sieu Sach' (Single File) tai local..." -ForegroundColor Yellow
$PublishDir = "publish_local"
if (Test-Path $PublishDir) { Remove-Item -Recurse -Force $PublishDir }

dotnet publish "Quan Ly Nhan Su\Quan_Ly_Nhan_Su.csproj" `
    -c Release `
    -r win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfContained=true `
    -p:IncludeAllContentForSelfContained=true `
    -o $PublishDir

if ($LASTEXITCODE -eq 0) {
    Write-Host "Build thanh cong! Dang loc file va nen Zip..." -ForegroundColor Green
    $ZipFile = "QuanLyNhanSu_v$NewVersion.zip"
    if (Test-Path $ZipFile) { Remove-Item $ZipFile }
    
    # CHI LAY NHUNG FILE CAN THIET (EXE, PDB va FONT)
    $FilesToZip = @(
        "$PublishDir\Quan_Ly_Nhan_Su.exe",
        "$PublishDir\Quan_Ly_Nhan_Su.pdb",
        "$PublishDir\LatoFont"
    )
    
    Compress-Archive -Path $FilesToZip -DestinationPath $ZipFile -Force
    
    Write-Host "=============================================" -ForegroundColor Green
    Write-Host "HOAN TAT QUY TRINH!" -ForegroundColor Green
    Write-Host "1. Code da duoc day len GitHub." -ForegroundColor White
    Write-Host "2. File Zip 'SIEU SACH' da duoc tao tai: $ZipFile" -ForegroundColor Yellow
    Write-Host "Dien mao file Zip: Chi co 1 file EXE (270MB), 1 file PDB va thu muc Font!" -ForegroundColor Green
    Write-Host "=============================================" -ForegroundColor Green
} else {
    Write-Host "Co loi trong qua trinh build local!" -ForegroundColor Red
}

Pause
