# 婕旂ず鑴氭湰锛氬睍绀轰笉鍚屽啓娉曠殑鏁堟灉

Write-Host "`n=== 鏂规硶 1锛氫娇鐢?`$? 鍙橀噺 ===" -ForegroundColor Cyan
Write-Host "dotnet build; if (`$?) { Start-Process '绋嬪簭璺緞' }" -ForegroundColor Yellow
Write-Host "璇存槑锛歚$? 鑷姩鍙橀噺琛ㄧず涓婁竴涓懡浠ゆ槸鍚︽垚鍔? -ForegroundColor Gray

Write-Host "`n=== 鏂规硶 2锛氬垎姝ユ墽琛岋紙鎺ㄨ崘鏂版墜锛?===" -ForegroundColor Cyan
Write-Host @"
dotnet build
if (`$?) {
    Start-Process '绋嬪簭璺緞'
}
"@ -ForegroundColor Yellow
Write-Host "璇存槑锛氭洿鏄撹锛岄€昏緫鏇存竻鏅? -ForegroundColor Gray

Write-Host "`n=== 鏂规硶 3锛氱洿鎺ヨ繍琛岋紙涓嶆鏌ワ級 ===" -ForegroundColor Cyan
Write-Host "dotnet build" -ForegroundColor Yellow
Write-Host "Start-Process '绋嬪簭璺緞'" -ForegroundColor Yellow
Write-Host "璇存槑锛氭棤璁烘瀯寤烘槸鍚︽垚鍔熼兘浼氳繍琛岋紙涓嶆帹鑽愶級" -ForegroundColor Gray

Write-Host "`n=== 鍏抽敭姒傚康 ===" -ForegroundColor Cyan
Write-Host "`$? - 鑷姩鍙橀噺锛岃〃绀轰笂涓€涓懡浠ょ殑鎴愬姛鐘舵€? -ForegroundColor Green
Write-Host "Start-Process - PowerShell cmdlet锛岀敤浜庡惎鍔ㄧ▼搴? -ForegroundColor Green
Write-Host "; - 鍛戒护鍒嗛殧绗︼紝椤哄簭鎵ц澶氫釜鍛戒护" -ForegroundColor Green
