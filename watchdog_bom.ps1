# -------------------------------------------------------------------
# PEPS BOM Tool - Server Watchdog
# Scheduled by Windows Task Scheduler every 45 minutes.
# Checks if the server is responding; restarts it if not.
#
# NOTE: this file must stay plain ASCII. Windows PowerShell 5.1 can
# misparse a no-BOM UTF-8 .ps1 (e.g. em-dashes, box-drawing chars) when
# invoked via "powershell -File ..." (exactly how Task Scheduler runs
# this), even though the same file looks fine in an editor or via
# Get-Content. That silently breaks the whole script with no output
# anywhere Task Scheduler shows you - do not reintroduce non-ASCII text.
# -------------------------------------------------------------------

$logFile   = "D:\Hari JR. DATA\Development\Bom Tool\bom_server.log"
$workDir   = "D:\Hari JR. DATA\Development\Bom Tool"
$healthUrl = "https://localhost:5020/api/health"
$ts = { (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + " [WATCHDOG]" }

function Write-Log($msg) {
    $line = "$(& $ts)  $msg"
    Add-Content $logFile $line
    Write-Host $line
}

# Use curl.exe (native, bundled since Win10 1803) for the health check
# instead of Invoke-WebRequest. PowerShell's own SSL-trust-callback
# mechanisms (scriptblock or C# delegate) are unreliable for a
# self-signed cert under non-interactive "-File" execution - both were
# tried and both made every health check falsely report the server as
# down, which then triggered an unnecessary restart. curl.exe -k just
# skips cert validation directly, no workaround needed.
function Test-Health {
    $code = & curl.exe -sk -o NUL -w "%{http_code}" --max-time 8 $healthUrl 2>$null
    return ($code -eq "200")
}

# -- Step 1: Check HTTP health endpoint -------------------------------
if (Test-Health) {
    Write-Log "Server is healthy - no action needed."
    exit 0
}

# -- Step 2: Server not responding - recover via run_forever.bat ONLY --
# Never launch python.exe directly from here: a process started by this
# watchdog task would belong to the watchdog's own task instance, and
# Windows Task Scheduler kills a task's child processes when that task
# instance finishes - which happens a few seconds after every run. That
# would silently kill the "recovered" server the moment this script exits.
# run_forever.bat's own loop is the only thing allowed to launch the
# server, so it stays supervised by something that runs indefinitely,
# not a task that exits in seconds.
Write-Log "Server NOT responding - recovering..."

# Kill any python.exe holding port 5020 (by port, not just by command line,
# because zombie processes have a blank CommandLine and would be missed otherwise)
$portPids = @()
netstat -ano | Select-String ":5020\s" | ForEach-Object {
    $parts = ($_.Line -split '\s+').Where({ $_ -ne '' })
    if ($parts.Count -ge 5) { $portPids += [int]$parts[-1] }
}
foreach ($p in ($portPids | Sort-Object -Unique)) {
    Stop-Process -Id $p -Force -ErrorAction SilentlyContinue
    Write-Log "Killed PID $p (held port 5020)"
}

# Also kill any named BOM python process
$zombies = Get-CimInstance Win32_Process | Where-Object {
    $_.Name -eq "python.exe" -and $_.CommandLine -like '*app_v2_1.py*'
}
foreach ($z in $zombies) {
    Stop-Process -Id $z.ProcessId -Force -ErrorAction SilentlyContinue
    Write-Log "Killed zombie PID $($z.ProcessId)"
}

$loopRunning = Get-CimInstance Win32_Process | Where-Object {
    $_.Name -eq "cmd.exe" -and $_.CommandLine -like '*run_forever.bat*'
}
if ($loopRunning) {
    Write-Log "run_forever.bat loop is alive - it will relaunch the server on its own within ~5s."
} else {
    Write-Log "run_forever.bat loop is NOT running - starting it fresh."
    Start-Process -FilePath "$workDir\run_forever.bat" -WorkingDirectory $workDir -WindowStyle Normal
}

Start-Sleep -Seconds 12

# -- Step 3: Verify restart succeeded ----------------------------------
if (Test-Health) {
    $newPid = (Get-CimInstance Win32_Process | Where-Object {
        $_.Name -eq "python.exe" -and $_.CommandLine -like '*app_v2_1.py*'
    } | Select-Object -First 1).ProcessId
    Write-Log "Server restarted successfully - PID $newPid"
} else {
    Write-Log "ERROR: Server failed to start - manual intervention needed!"
}
