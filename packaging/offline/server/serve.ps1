# KRONOS offline: the same tiny server, for Windows computers with neither Node.js nor Python.
# Uses only PowerShell, which comes with Windows. The launcher passes the app folder in; this
# script does not depend on knowing where it is itself.
param([string]$AppFolder)

$port = 5181

if (-not $AppFolder) {
  $here = $PSScriptRoot
  if (-not $here) { $here = Split-Path -Parent $MyInvocation.MyCommand.Path }
  if (-not $here) { $here = (Get-Location).Path }
  $AppFolder = [System.IO.Path]::Combine($here, '..', 'app')
}
$root = [System.IO.Path]::GetFullPath($AppFolder).TrimEnd('\')
$index = [System.IO.Path]::Combine($root, 'index.html')
if (-not [System.IO.File]::Exists($index)) {
  Write-Host ""
  Write-Host "I can't find the KRONOS app here: $root"
  Write-Host "Make sure you unzipped the whole folder, and start KRONOS from inside it."
  exit 1
}

$types = @{
  '.html'='text/html; charset=utf-8'; '.js'='text/javascript; charset=utf-8'; '.mjs'='text/javascript; charset=utf-8';
  '.css'='text/css; charset=utf-8'; '.svg'='image/svg+xml'; '.png'='image/png'; '.jpg'='image/jpeg'; '.ico'='image/x-icon';
  '.json'='application/json'; '.wasm'='application/wasm'; '.woff'='font/woff'; '.woff2'='font/woff2'
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
try {
  $listener.Start()
} catch {
  Write-Host ""
  Write-Host "Port $port is busy. KRONOS may already be running: open http://localhost:$port/ in your browser."
  exit 1
}

Write-Host ""
Write-Host "KRONOS Offline, version dev"
Write-Host "KRONOS is running at http://localhost:$port/"
Write-Host "App folder: $root"
Write-Host "Keep this window open while you use it. Close it (or press Ctrl+C) to stop."
Write-Host ""

# Internet Explorer can't run KRONOS, and some computers have no default browser at all
# (Start-Process then fails with "No application is associated"). So unless another default
# browser is set, look for Chrome, Edge or Firefox by name.
function Open-Browser([string]$url) {
  $default = $null
  try {
    $default = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice' -ErrorAction Stop).ProgId
  } catch { }
  if ($default -and $default -notlike 'IE.*') {
    try { Start-Process $url -ErrorAction Stop; return $true } catch { }
  }
  $places = @($env:LOCALAPPDATA, ${env:ProgramFiles(x86)}, $env:ProgramFiles) | Where-Object { $_ }
  foreach ($name in @('Google\Chrome\Application\chrome.exe', 'Microsoft\Edge\Application\msedge.exe', 'Mozilla Firefox\firefox.exe')) {
    foreach ($place in $places) {
      $exe = [System.IO.Path]::Combine($place, $name)
      if ([System.IO.File]::Exists($exe)) {
        try { Start-Process -FilePath $exe -ArgumentList $url -ErrorAction Stop; return $true } catch { }
      }
    }
  }
  return $false
}

if (-not (Open-Browser "http://localhost:$port/")) {
  Write-Host "Your browser did not open by itself. Open Chrome, Edge or Firefox and type:  localhost:$port"
  Write-Host "(Internet Explorer can't show KRONOS.)"
  Write-Host ""
}

while ($listener.IsListening) {
  $context = $listener.GetContext()
  $response = $context.Response
  try {
    $relative = [System.Uri]::UnescapeDataString($context.Request.Url.AbsolutePath).TrimStart('/')
    if ($relative -eq '') { $relative = 'index.html' }
    $file = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($root, $relative.Replace('/', '\')))
    $inside = $file.StartsWith($root + '\', [System.StringComparison]::OrdinalIgnoreCase)
    if (-not $inside -or -not [System.IO.File]::Exists($file)) {
      $response.StatusCode = 404
    } else {
      $bytes = [System.IO.File]::ReadAllBytes($file)
      $type = $types[[System.IO.Path]::GetExtension($file).ToLower()]
      if (-not $type) { $type = 'application/octet-stream' }
      $response.StatusCode = 200
      $response.ContentType = $type
      $response.AddHeader('Cache-Control', 'no-cache')
      $response.ContentLength64 = $bytes.Length
      $response.OutputStream.Write($bytes, 0, $bytes.Length)
    }
  } catch {
    Write-Host ("Problem with " + $context.Request.Url.AbsolutePath + ": " + $_.Exception.Message)
    try { $response.StatusCode = 500 } catch { }
  } finally {
    try { $response.Close() } catch { }
  }
}
