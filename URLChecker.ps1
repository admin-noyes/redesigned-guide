param(
    [string]$InputFile,
    [string]$OutputFile,
    [int]$MaxThreads = 50,  # Increased default for larger datasets
    [int]$Retries = 3
)

# Read CSV input with Address,URL
$entries = Import-Csv -Path $InputFile

# Use Runspaces for better parallel performance
$runspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
$runspacePool.Open()

$jobs = @()
$jobScript = {
    param($addr, $url, $retries)
    try {
        $trimmedUrl = $url.Trim()
        if ($trimmedUrl -notmatch '^https?://') {
            return @{Address = $addr; Status = "Invalid URL"}
        }
        for ($attempt = 1; $attempt -le $retries; $attempt++) {
            try {
                $response = Invoke-WebRequest -Uri $trimmedUrl -Method Head -TimeoutSec 3 -UseBasicParsing
                return @{Address = $addr; Status = [string]$response.StatusCode}
            } catch {
                if ($attempt -lt $retries) {
                    Start-Sleep -Milliseconds 500
                } else {
                    return @{Address = $addr; Status = "Failed After $retries Attempts"}
                }
            }
        }
    } catch {
        return @{Address = $addr; Status = "Error"}
    }
}

foreach ($entry in $entries) {
    $powershell = [powershell]::Create().AddScript($jobScript).AddArgument($entry.Address).AddArgument($entry.URL).AddArgument($Retries)
    $powershell.RunspacePool = $runspacePool
    $jobs += @{
        PowerShell = $powershell
        Handle = $powershell.BeginInvoke()
        Address = $entry.Address
    }
}

$results = @()
foreach ($job in $jobs) {
    $result = $job.PowerShell.EndInvoke($job.Handle)
    $results += $result
    $job.PowerShell.Dispose()
}

$runspacePool.Close()

$results | Select-Object @{Name='Address';Expression={$_.Address}}, @{Name='Status';Expression={$_.Status}} | Export-Csv -Path $OutputFile -NoTypeInformation