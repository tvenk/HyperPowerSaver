# Ensure Administrator Privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Please run this script as Administrator. Right-click and select 'Run as Administrator'." -ForegroundColor Red
    Pause
    Exit
}

# Define Paths
$installDir = "C:\Program Files\HyperPowerSaver"
$mainScriptPath = Join-Path $installDir "HyperPowerSaver.ps1"
$iconPath = Join-Path $installDir "greenbox.ico"

Write-Host "Starting HyperPowerSaver Installation..." -ForegroundColor Cyan

# 1. Clean up old versions
Write-Host "Removing old versions if present..." -ForegroundColor Yellow

# Stop any running instances of HyperPowerSaver
Get-Process | Where-Object { $_.ProcessName -like "*powershell*" -and $_.MainWindowTitle -like "*HyperPowerSaver*" } | Stop-Process -Force -ErrorAction SilentlyContinue

# Stop any instances related to the script's location, just in case
Get-Process | Where-Object { $_.Path -like "$installDir\*" } | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 5  # Wait to ensure all processes have terminated

# Remove old scheduled tasks
Get-ScheduledTask | Where-Object {$_.TaskName -like "*HyperPowerSaver*"} | Unregister-ScheduledTask -Confirm:$false
Write-Host "- Removed old scheduled tasks" -ForegroundColor Green

# Attempt to remove old installation directory after stopping processes
if (Test-Path $installDir) {
    try {
        Remove-Item $installDir -Recurse -Force
        Write-Host "- Removed old installation directory" -ForegroundColor Green
    } catch {
        Write-Host "Error: Could not remove $installDir because it is still in use." -ForegroundColor Red
        Pause
        Exit
    }
}

# 2. Create new installation
Write-Host "`nInstalling new version..." -ForegroundColor Yellow

# Create installation directory
New-Item -ItemType Directory -Force -Path $installDir | Out-Null
Write-Host "- Created installation directory" -ForegroundColor Green

# Create a green box icon
$iconContent = @'
using System;
using System.Drawing;
using System.Drawing.Imaging;

public class IconCreator
{
    public static void CreateGreenBoxIcon(string path)
    {
        using (Bitmap bmp = new Bitmap(16, 16))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.Green);
            bmp.Save(path, ImageFormat.Icon);
        }
    }
}
'@

Add-Type -TypeDefinition $iconContent -ReferencedAssemblies System.Drawing
[IconCreator]::CreateGreenBoxIcon($iconPath)
Write-Host "- Created green box icon" -ForegroundColor Green

# Main script content
$mainScriptContent = @'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Ensure only one instance is running
$mutex = New-Object System.Threading.Mutex($false, "Global\HyperPowerSaverMutex")
if (-not $mutex.WaitOne(0, $false)) {
    [System.Windows.Forms.MessageBox]::Show("HyperPowerSaver is already running.", "HyperPowerSaver", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    Exit
}

# Create a form and hide it
$form = New-Object System.Windows.Forms.Form
$form.WindowState = 'Minimized'
$form.ShowInTaskbar = $false
$form.Hide()

# Define the tray icon with the green box icon
$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Program Files\HyperPowerSaver\greenbox.ico")
$notifyIcon.Visible = $true
$notifyIcon.Text = "HyperPowerSaver"

# Create a menu for the tray icon
$menu = New-Object System.Windows.Forms.ContextMenuStrip
$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Exit")
$menu.Items.Add($exitMenuItem)

$notifyIcon.ContextMenuStrip = $menu

# Define exit functionality
$exitMenuItem.Add_Click({
    # Clean up
    $notifyIcon.Visible = $false
    $notifyIcon.Dispose()
    $form.Close()
    $form.Dispose()
    $mutex.ReleaseMutex()
    [System.Windows.Forms.Application]::Exit()
})

# Function to turn off screen
function Set-ScreenOff {
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class DisplayControl {
        [DllImport("user32.dll")]
        public static extern int SendMessage(int hWnd, int hMsg, int wParam, int lParam);
    }
"@
    [DisplayControl]::SendMessage(-1, 0x0112, 0xF170, 2)
}

# Function to check for user activity (mouse movement)
function Test-UserActivity {
    $currentPosition = [System.Windows.Forms.Cursor]::Position
    Start-Sleep -Milliseconds 100
    $newPosition = [System.Windows.Forms.Cursor]::Position
    return ($currentPosition.X -ne $newPosition.X) -or ($currentPosition.Y -ne $newPosition.Y)
}

# Function to detect video playback (process-based)
function Test-VideoPlayback {
    $audioSessions = Get-Process | Where-Object { $_.MainWindowTitle -ne "" } |
        Where-Object { $_.ProcessName -match "chrome|firefox|edge|vlc|mpv" }
    return $audioSessions.Count -gt 0
}

# Timer and Activity Logic
$lastActivity = Get-Date
$promptShown = $false

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 500
$timer.Add_Tick({
    $currentTime = Get-Date
    $isActive = Test-UserActivity
    $videoPlaying = Test-VideoPlayback

    if ($isActive -or $videoPlaying) {
        $script:lastActivity = Get-Date
        $script:promptShown = $false
    }

    # Check inactivity (only prompt after 30 seconds of inactivity)
    $inactivityTime = ($currentTime - $script:lastActivity).TotalSeconds
    if ($inactivityTime -ge 30 -and !$script:promptShown) {
        $script:promptShown = $true
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you still there? Click Yes to stay active, No to turn off the screen.",
            "HyperPowerSaver",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button1,
            [System.Windows.Forms.MessageBoxOptions]::DefaultDesktopOnly
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            # If the user selects Yes, reset the activity timer and stop reprompting
            $script:lastActivity = Get-Date
            $script:promptShown = $false
        } elseif ($result -eq [System.Windows.Forms.DialogResult]::No) {
            # If the user selects No, turn off the screen
            Set-ScreenOff
            $script:lastActivity = Get-Date
            $script:promptShown = $false
        }
    }
})

$timer.Start()

# Start the application message loop
[System.Windows.Forms.Application]::Run($form)
'@

# Write the main script to file
Set-Content -Path $mainScriptPath -Value $mainScriptContent -Encoding UTF8
Write-Host "- Created main script" -ForegroundColor Green

# Set execution policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force -Scope CurrentUser
Write-Host "- Set execution policy" -ForegroundColor Green

# Create scheduled task
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$mainScriptPath`""
$trigger = New-ScheduledTaskTrigger -AtLogOn
$principal = New-ScheduledTaskPrincipal -UserId (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty UserName) -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable

Register-ScheduledTask -TaskName "HyperPowerSaver" -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force
Write-Host "- Created scheduled task" -ForegroundColor Green

# Start the application
Write-Host "`nStarting HyperPowerSaver..." -ForegroundColor Yellow
Start-ScheduledTask -TaskName "HyperPowerSaver"
Write-Host "- Application started" -ForegroundColor Green

Write-Host "`nInstallation Complete!" -ForegroundColor Cyan
Write-Host "HyperPowerSaver will start automatically when you log in." -ForegroundColor White
Write-Host "Look for the green box icon in your system tray." -ForegroundColor White
Write-Host "`nPress any key to exit..." -ForegroundColor White
Pause
