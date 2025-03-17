# Script to automatically collect IIS requests information
# v1
# Instructions: Run as administrator in a command window

$source = "IIS Information collection script"
$logName = "Application"
$successEventId = 200
$errorEventId = 400

if (-not [System.Diagnostics.EventLog]::SourceExists($source)) {
    # Register the source for the Application log
    New-EventLog -LogName $logName -Source $source
}

# Download OSDiagTool
$filePath = ".\OSDiagTool-v-threads-cmd.zip"

if (Test-Path $filePath) {
    Write-Host "OSDiagTool already exists. Skipping download."
} else {
	try {
		Invoke-WebRequest -Uri "https://github.com/OutSystems/techsupp-osdiagtool/releases/download/v-threads-cmd/OSDiagTool-v-threads-cmd.zip" -OutFile ".\OSDiagTool-v-threads-cmd.zip"
		Write-Host "OSDiagTool download complete"
	} catch {
		Write-Host "Error occurred downloading OSDiagTool"
		Write-EventLog -LogName $logName -Source $source -EntryType "Warning" -EventId $errorEventId -Message "Error running IIS collection script - error downloading OSDiagTool"
		exit 1
	}
}

$zipPath = ".\OSDiagTool-v-threads-cmd.zip"

# Unzip package
$dateSuffix = (Get-Date -Format "yyyy-MM-dd_HH-mm-ss")
$extractFolder = ".\OSDiagTool-Alarm-$dateSuffix"

if (!(Test-Path $extractFolder)) {
    New-Item -ItemType Directory -Path $extractFolder | Out-Null
}

try {
	Write-Host "Starting zip extraction"
    Expand-Archive -Path $zipPath -DestinationPath $extractFolder -Force
    Write-Host "Unzip of OSDiagTool and extracted to $extractFolder"
} catch {
    Write-Host "Error occurred while unzipping OSDiagTool" -ForegroundColor Red
    Write-Host "Message: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "StackTrace: $($_.Exception.StackTrace)" -ForegroundColor Gray
    Write-EventLog -LogName $logName -Source $source -EntryType "Warning" -EventId $errorEventId -Message "Error running IIS collection script - unable to unzip OSDiagTool"
    exit 1
}

# Run OSDiagTool
try {
    Write-Host "Starting OSDiagTool run"
	Start-Process -FilePath "$extractFolder\OSDiagTool.exe" -ArgumentList "runcmdline" -NoNewWindow -Wait
    Write-Output "Thread dumps collected"
} catch {
    Write-Host "Error running OSDiagTool" -ForegroundColor Red
    Write-Host "Message: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "StackTrace: $($_.Exception.StackTrace)" -ForegroundColor Gray
    Write-EventLog -LogName $logName -Source $source -EntryType "Warning" -EventId $errorEventId -Message "Error running IIS collection script - IIS threads collection failed"
}

# Obtain TCP Dump
$tcpDumpfile = "$extractFolder\TCP-Dump.txt"

try {
    $connections = Get-NetTCPConnection | ForEach-Object {
        $process = Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue
        [PSCustomObject]@{
            LocalAddress  = $_.LocalAddress
            LocalPort     = $_.LocalPort
            RemoteAddress = $_.RemoteAddress
            RemotePort    = $_.RemotePort
            State         = $_.State
            ProcessName   = if ($process) { $process.Name } else { "Unknown" }
        }
    }
    $connections | Format-Table -AutoSize | Out-File -FilePath $tcpDumpfile

    Write-Host "TCP Dump collected"

} catch {
    Write-Host "Error obtaining tcp dump" -ForegroundColor Red
    Write-Host "Message: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "StackTrace: $($_.Exception.StackTrace)" -ForegroundColor Gray
    Write-EventLog -LogName $logName -Source $source -EntryType "Warning" -EventId $errorEventId -Message "Error running IIS collection script - tcp dumps failed"
    exit 1
}

Write-EventLog -LogName $logName -Source $source -EntryType "Information" -EventId $successEventId -Message "Successful collection of IIS threads and TCP dump"
