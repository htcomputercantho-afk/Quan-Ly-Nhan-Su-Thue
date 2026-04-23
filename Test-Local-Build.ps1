# SCRIPT BUILD THU NGHIEM NHANH TAI LOCAL
# Nhiem vu: Tao ra ban Single EXE de kiem tra truoc khi push len GitHub

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "BAT DAU QUY TRINH BUILD TEST LOCAL" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan

$OutDir = "Build_Result_Local"

# Kiem tra va lam sach thu muc cu
if (Test-Path $OutDir) {
    Write-Host "Phat hien ban build cu. Dang don dep thu muc $OutDir..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force $OutDir
}

Write-Host "Dang build app voi cau hinh 'Single EXE'..." -ForegroundColor White

# Chay lenh build than thanh
dotnet publish "Quan Ly Nhan Su\Quan_Ly_Nhan_Su.csproj" `
    -c Release `
    -r win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:DebugType=none `
    -p:DebugSymbols=false `
    -o $OutDir

if ($LASTEXITCODE -eq 0) {
    Write-Host "=============================================" -ForegroundColor Green
    Write-Host "BUILD TEST THANH CONG!" -ForegroundColor Green
    Write-Host "Vi tri file: $OutDir\Quan_Ly_Nhan_Su.exe" -ForegroundColor White
    Write-Host "Ban hay mo thu muc '$OutDir' de kiem tra chat luong file nhe." -ForegroundColor Yellow
    Write-Host "=============================================" -ForegroundColor Green
    
    # Mo thu muc build sau khi xong de ban kiem tra ngay
    ii $OutDir
} else {
    Write-Host "Build that bai! Vui long kiem tra lai code." -ForegroundColor Red
}

Pause
