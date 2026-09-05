param([string]$OutputPath = (Join-Path $PSScriptRoot 'OfficeReadiness.json'), [int]$TimeoutSeconds = 30)
$ErrorActionPreference = 'Stop'
$results = @()
foreach ($application in @('Outlook','Excel','PowerPoint','Word')) {
    $progId = $application + '.Application'
    if ($null -eq [type]::GetTypeFromProgID($progId)) {
        $results += [ordered]@{ application=$application;registered=$false;activated=$false;error='APPLICATION_NOT_REGISTERED' }
        continue
    }
    # Isolate COM activation: an unavailable Office server can otherwise hang
    # the entire acceptance runner. This probe creates no document or message.
    $code = "`$ErrorActionPreference='Stop'; try { `$type=[type]::GetTypeFromProgID('$progId'); `$app=[Activator]::CreateInstance(`$type); @{ activated=`$true;version=[string]`$app.Version } | ConvertTo-Json -Compress } catch { @{ activated=`$false;error=`$_.Exception.Message;hresult=`$_.Exception.HResult } | ConvertTo-Json -Compress; exit 1 }"
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $start.Arguments = '-NoProfile -NonInteractive -STA -EncodedCommand ' + [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($code))
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    [void]$process.Start()
    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
        $process.Kill()
        $results += [ordered]@{ application=$application;registered=$true;activated=$false;error='COM_ACTIVATION_TIMEOUT' }
    } else {
        $raw = $process.StandardOutput.ReadToEnd()
        try { $result = $raw | ConvertFrom-Json } catch { $result = @{activated=$false;error='INVALID_PROBE_RESPONSE'} }
        $results += [ordered]@{ application=$application;registered=$true;activated=$result.activated;version=$result.version;error=$result.error;hresult=$result.hresult }
    }
    $process.Dispose()
}
$report = [ordered]@{ schema=1;execution_kind='native_readiness_only';utc=[DateTime]::UtcNow.ToString('O');applications=$results;
    full_acceptance_passed=$false;note='Activation/version checks only. No Office documents, mailbox contents, forty routes, model workflows or layout acceptance were tested.' }
$report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputPath -Encoding utf8
$report | ConvertTo-Json -Depth 6
if (@($results | Where-Object { $_.activated -ne $true }).Count -gt 0) { exit 1 }
