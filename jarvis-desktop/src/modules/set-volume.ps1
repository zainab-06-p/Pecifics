param([int]$Level)

# Simple and reliable method using WScript.Shell SendKeys
$wshShell = New-Object -ComObject WScript.Shell

# Mute/unmute to reset, then set volume using keys
# VolumeDown key (173/0xAD), VolumeUp key (175/0xAF)

# Set to 0 first
for ($i = 0; $i -lt 50; $i++) {
    $wshShell.SendKeys([char]174)  # Volume down
    Start-Sleep -Milliseconds 10
}

# Now increase to target level (each step ~= 2%)
$steps = [Math]::Round($Level / 2)
for ($i = 0; $i -lt $steps; $i++) {
    $wshShell.SendKeys([char]175)  # Volume up
    Start-Sleep -Milliseconds 10
}

Write-Output "SUCCESS:$Level"
