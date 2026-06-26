$ErrorActionPreference = "Stop"

# --- Credenciales ---
$clasprc = Get-Content "$env:USERPROFILE\.clasprc.json" -Raw | ConvertFrom-Json
$tok = $clasprc.tokens.default
$claspjson = Get-Content "C:\Users\Matias\Agenda_promesas\.clasp.json" -Raw | ConvertFrom-Json
$scriptId = $claspjson.scriptId

# --- 1. Refrescar access token ---
$body = @{
  client_id     = $tok.client_id
  client_secret = $tok.client_secret
  refresh_token = $tok.refresh_token
  grant_type    = "refresh_token"
}
$tokenResp = Invoke-RestMethod -Method Post -Uri "https://oauth2.googleapis.com/token" -Body $body
$access = $tokenResp.access_token
Write-Host "Token refrescado OK"

$headers = @{ Authorization = "Bearer $access" }

# --- 2. Construir payload de contenido ---
$gsFile = Get-ChildItem "C:\Users\Matias\Agenda_promesas" -Filter "*.gs" | Select-Object -First 1
$codigo = [System.IO.File]::ReadAllText($gsFile.FullName)
$manifest = [System.IO.File]::ReadAllText("C:\Users\Matias\Agenda_promesas\appsscript.json")

$payload = @{
  files = @(
    @{ name = "Codigo"; type = "SERVER_JS"; source = $codigo },
    @{ name = "appsscript"; type = "JSON"; source = $manifest }
  )
}
$json = $payload | ConvertTo-Json -Depth 10

# Enviar con bytes UTF-8 explícitos (sin BOM)
$uri = "https://script.googleapis.com/v1/projects/$scriptId/content"
$bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
$req = [System.Net.HttpWebRequest]::Create($uri)
$req.Method = "PUT"
$req.ContentType = "application/json; charset=utf-8"
$req.Headers.Add("Authorization", "Bearer $access")
$req.ContentLength = $bytes.Length
$stream = $req.GetRequestStream()
$stream.Write($bytes, 0, $bytes.Length)
$stream.Close()
try {
  $resp = $req.GetResponse()
  $sr = New-Object System.IO.StreamReader($resp.GetResponseStream())
  $sr.ReadToEnd() | Out-Null
  Write-Host "Contenido subido OK"
} catch {
  $errResp = $_.Exception.Response
  if ($errResp) {
    $sr = New-Object System.IO.StreamReader($errResp.GetResponseStream())
    Write-Host "ERROR subida: $($sr.ReadToEnd())"
  }
  throw
}

# --- 3. Crear nueva versión ---
$verBody = @{ description = "deploy disponibilidad " + (Get-Date -Format "yyyy-MM-dd HH:mm") } | ConvertTo-Json
$verResp = Invoke-RestMethod -Method Post -Uri "https://script.googleapis.com/v1/projects/$scriptId/versions" -Headers $headers -Body $verBody -ContentType "application/json"
$newVer = $verResp.versionNumber
Write-Host "Version creada: $newVer"

# --- 4. Listar deployments y actualizar el web app ---
# Deployment de la cuenta del sistema (promesaschilelosrios) que usa el frontend
$depId = "AKfycbx6f2OJFXaG8t-b_xU8SWkKc7CSK1kvFvHAAkEemJHJfx3XNuRg3J8g97LZUd_N0cuhLw"
Write-Host "Deployment a actualizar: $depId"

$updBody = @{
  deploymentConfig = @{
    scriptId       = $scriptId
    versionNumber  = $newVer
    manifestFileName = "appsscript"
    description    = "get_citas multi-rol"
  }
} | ConvertTo-Json -Depth 10
$updResp = Invoke-RestMethod -Method Put -Uri "https://script.googleapis.com/v1/projects/$scriptId/deployments/$depId" -Headers $headers -Body $updBody -ContentType "application/json"
Write-Host "Deployment actualizado a version $newVer"
Write-Host "LISTO"
