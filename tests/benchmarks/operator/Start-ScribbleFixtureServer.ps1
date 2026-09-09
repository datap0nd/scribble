#requires -Version 5.1
param([string]$AssemblyPath)
. (Join-Path $PSScriptRoot 'TestLab.Common.ps1') -AssemblyPath $AssemblyPath
$session = [Scribble.Testing.TestLab]::Status()
if (-not $session) { throw 'Enable Test Lab first.' }
[void][Scribble.Testing.TestLab]::VerifyKit($session.fixture_root)
$mutex = New-Object Threading.Mutex($false, ('Local\ScribbleFixtureServer-' + $session.session_id))
if (-not $mutex.WaitOne(0)) { $mutex.Dispose(); exit }
$listener = $null
try {
    $pages = @{}
    foreach ($name in @('index.html','operations.html','archive.html')) {
        $pages['/' + $name] = [IO.File]::ReadAllBytes([Scribble.Testing.TestLab]::SafeChild($session.fixture_root, ('inputs/browser/' + $name)))
    }
    $pages['/'] = $pages['/index.html']
    $listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0)
    $listener.Start()
    $receipt = Join-Path ([Scribble.Testing.TestLab]::Root) ('fixture-server-' + $session.session_id + '.json')
    @{session_id=$session.session_id; manifest_sha256=$session.manifest_sha256; port=$listener.LocalEndpoint.Port} |
        ConvertTo-Json | Set-Content -LiteralPath $receipt -Encoding UTF8
    while ($true) {
        $active = [Scribble.Testing.TestLab]::Status()
        if (-not $active -or $active.session_id -ne $session.session_id) { break }
        if (-not $listener.Pending()) { Start-Sleep -Milliseconds 250; continue }
        $client = $listener.AcceptTcpClient()
        try {
            $stream = $client.GetStream(); $stream.ReadTimeout=1500; $stream.WriteTimeout=1500
            $request = New-Object Text.StringBuilder
            while ($request.Length -lt 8192) {
                $b=$stream.ReadByte(); if ($b -lt 0) { break }
                [void]$request.Append([char]$b)
                if ($request.ToString().EndsWith("`r`n`r`n")) { break }
            }
            $line = ($request.ToString() -split "`r`n")[0]
            $match = [regex]::Match($line, '^(GET|HEAD) (/[^ ]*) HTTP/1\.[01]$')
            $ok = $match.Success -and $pages.ContainsKey($match.Groups[2].Value)
            $body = if ($ok) { $pages[$match.Groups[2].Value] } else { [Text.Encoding]::UTF8.GetBytes('Not found') }
            $status = if ($ok) { '200 OK' } else { '404 Not Found' }
            $headers = "HTTP/1.1 $status`r`nContent-Type: text/html; charset=utf-8`r`nContent-Length: $($body.Length)`r`nConnection: close`r`nCache-Control: no-store`r`nX-Scribble-Fixture: $($session.session_id)`r`n`r`n"
            $bytes=[Text.Encoding]::ASCII.GetBytes($headers); $stream.Write($bytes,0,$bytes.Length)
            if (-not $match.Success -or $match.Groups[1].Value -ne 'HEAD') { $stream.Write($body,0,$body.Length) }
        } catch { } finally { $client.Close() }
    }
} finally {
    if ($listener) { $listener.Stop() }
    $mutex.ReleaseMutex(); $mutex.Dispose()
}
