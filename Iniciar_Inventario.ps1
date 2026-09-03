param(
  [switch]$NoBrowser
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = [System.IO.Path]::GetFullPath($PSScriptRoot)
$rootPrefix = $root.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
$contentTypes = @{
  '.css' = 'text/css; charset=utf-8'; '.html' = 'text/html; charset=utf-8'; '.js' = 'text/javascript; charset=utf-8'
  '.json' = 'application/json; charset=utf-8'; '.png' = 'image/png'; '.jpg' = 'image/jpeg'; '.jpeg' = 'image/jpeg'
  '.svg' = 'image/svg+xml'; '.webp' = 'image/webp'; '.zip' = 'application/zip'; '.txt' = 'text/plain; charset=utf-8'
}

function Send-Response([System.Net.Sockets.NetworkStream]$Stream, [int]$StatusCode, [string]$ContentType, [byte[]]$Body) {
  $reason = if ($StatusCode -eq 200) { 'OK' } elseif ($StatusCode -eq 404) { 'Not Found' } else { 'Internal Server Error' }
  $header = "HTTP/1.1 $StatusCode $reason`r`nContent-Type: $ContentType`r`nContent-Length: $($Body.Length)`r`nCache-Control: no-store`r`nConnection: close`r`n`r`n"
  $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
  $Stream.Write($headerBytes, 0, $headerBytes.Length)
  if ($Body.Length) { $Stream.Write($Body, 0, $Body.Length) }
}

$listener = $null
$port = $null
foreach ($candidatePort in 4173..4183) {
  $candidate = $null
  try {
    $candidate = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $candidatePort)
    $candidate.Start()
    $listener = $candidate
    $port = $candidatePort
    break
  } catch {
    if ($candidate) { $candidate.Stop() }
  }
}
if (-not $listener) { throw 'No se pudo abrir un puerto local entre 4173 y 4183. Cierra otra instancia de Inventario Legislativo e inténtalo de nuevo.' }

$url = "http://127.0.0.1:$port/"
Write-Host ''
Write-Host 'Inventario Legislativo está listo.' -ForegroundColor Green
Write-Host "Abre: $url"
Write-Host 'Mantén esta ventana abierta mientras uses la aplicación. Para detenerla, presiona Ctrl+C.' -ForegroundColor Yellow
if (-not $NoBrowser) { Start-Process $url }

try {
  while ($true) {
    $client = $listener.AcceptTcpClient()
    $stream = $null
    $reader = $null
    try {
      $stream = $client.GetStream()
      $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::ASCII, $false, 8192, $true)
      $requestLine = $reader.ReadLine()
      while ($reader.ReadLine()) { }
      $requestParts = @($requestLine -split ' ')
      if ($requestParts.Count -lt 2) { throw 'Solicitud HTTP inválida.' }
      $path = ([System.Uri]("http://localhost" + $requestParts[1])).AbsolutePath
      $relative = if ($path -eq '/') { 'index.html' } else { $path.TrimStart('/') }
      $file = [System.IO.Path]::GetFullPath((Join-Path $root $relative))
      if (-not $file.StartsWith($rootPrefix, [System.StringComparison]::OrdinalIgnoreCase) -or -not [System.IO.File]::Exists($file)) {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes('No encontrado')
        Send-Response $stream 404 'text/plain; charset=utf-8' $bytes
        continue
      }
      $extension = [System.IO.Path]::GetExtension($file).ToLowerInvariant()
      $bytes = [System.IO.File]::ReadAllBytes($file)
      $contentType = if ($contentTypes.ContainsKey($extension)) { $contentTypes[$extension] } else { 'application/octet-stream' }
      Send-Response $stream 200 $contentType $bytes
    } catch {
      if ($stream) {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes('Error local al servir el archivo')
        Send-Response $stream 500 'text/plain; charset=utf-8' $bytes
      }
    } finally {
      if ($reader) { $reader.Dispose() }
      if ($client) { $client.Close() }
    }
  }
} finally {
  if ($listener) { $listener.Stop() }
}
