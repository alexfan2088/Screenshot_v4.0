# 蹇€熻繍琛岃剼鏈?- 鍙互娣诲姞鍒?PowerShell 閰嶇疆鏂囦欢鎴栫洿鎺ヤ娇鐢?# 鏋勫缓骞惰繍琛岄」鐩?function Run-ScreenshotApp {
    Write-Host "姝ｅ湪鏋勫缓椤圭洰..." -ForegroundColor Yellow
    dotnet build Screenshot_v3.0.csproj
    if ($LASTEXITCODE -eq 0) {
        Write-Host "鏋勫缓鎴愬姛锛佹鍦ㄥ惎鍔ㄥ簲鐢ㄧ▼搴?.." -ForegroundColor Green
        Start-Process "bin\Debug\net8.0-windows10.0.19041.0\Screenshot_v3.0.exe"
    } else {
        Write-Host "鏋勫缓澶辫触锛岃妫€鏌ラ敊璇俊鎭? -ForegroundColor Red
    }
}
