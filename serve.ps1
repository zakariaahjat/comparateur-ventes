$root = 'C:\Users\zakaria01\Downloads\sepraration project'
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add('http://localhost:8642/')
$listener.Start()
Write-Output "Serving $root on http://localhost:8642/"
try {
  while ($listener.IsListening) {
    $ctx = $listener.GetContext()
    $rel = [Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath).TrimStart('/')
    if ($rel -eq '') { $rel = 'Comparateur Ventes.html' }
    $file = Join-Path $root $rel
    if ((Test-Path -LiteralPath $file -PathType Leaf)) {
      $bytes = [IO.File]::ReadAllBytes($file)
      $ext = [IO.Path]::GetExtension($file).ToLower()
      $mime = switch ($ext) {
        '.html' { 'text/html; charset=utf-8' }
        '.js'   { 'application/javascript; charset=utf-8' }
        '.xlsx' { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
        default { 'application/octet-stream' }
      }
      $ctx.Response.ContentType = $mime
      $ctx.Response.ContentLength64 = $bytes.Length
      $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
      $ctx.Response.OutputStream.Close()
    } else {
      $ctx.Response.StatusCode = 404
      $ctx.Response.Close()
    }
  }
} finally { $listener.Stop() }