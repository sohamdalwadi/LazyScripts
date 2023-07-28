<#
    Script Name: awake.ps1
    Author: Soham Dalwadi
    Copyright © 2023 Soham Dalwadi. All rights reserved. 

    Description:
    This script is designed to keep Microsoft Teams active and prevent it from going into the "Away" status.
    It periodically simulates user activity to ensure that Teams remains in an active state.

    Requirements:
    - Windows 10
    - PowerShell 5.1 or later

    Usage:
    1. Ensure that Microsoft Teams is running.
    2. Open PowerShell in Windows 10.
    3. Navigate to the folder where the "awake.ps1" script is located.
    4. Execute the script using the following command:
        .\awake.ps1

    Disclaimer:
    This script is provided as-is without any warranty. The author, Soham Dalwadi, and the organization assume 
    no responsibility for any damage caused by using this script.

    © 2023 Soham Dalwadi. All rights reserved. Unauthorized copying or reproduction of this script or any part 
    thereof is prohibited.
#>

Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class User32 {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
"@

function AnimateTimer($seconds, $message) {
    $chars = "-\|/-\|/"
    $totalChars = $chars.Length
    for ($i = 0; $i -lt $seconds; $i++) {
        Write-Host -NoNewline "$message $($i+1) $($chars[$i%$totalChars])`r"
        Start-Sleep -Seconds 1
    }
    Write-Host ""
}

Clear-Host

Echo "Toggling Scroll Lock every 2 minutes. Press CTRL+C to stop."

$WShell = New-Object -com "Wscript.Shell"

$teamsProcess = Get-Process -Name Teams | Select-Object -First 1
$teamsMainWindowHandle = $teamsProcess.MainWindowHandle

$iteration = 1

while ($true) {
  $null = [User32]::SetForegroundWindow($teamsMainWindowHandle)
  if ($?) {
    Echo "Set Teams window to foreground."
  } else {
    Echo "Failed to set Teams window to foreground."
  }

  AnimateTimer 120 "2 minute timer for iteration $($iteration): "
  
  $WShell.sendkeys("{SCROLLLOCK}")
  Write-Host "Scroll Lock toggled ON"
  AnimateTimer 10 "10 second timer for iteration $($iteration): "
  $WShell.sendkeys("{SCROLLLOCK}")
  Write-Host "Scroll Lock toggled OFF"

  $iteration++
  cls
}
  
