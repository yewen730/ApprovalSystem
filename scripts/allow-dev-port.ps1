# Opens Windows Firewall for the dev server so other devices on the same LAN can connect.
# Right-click PowerShell -> Run as administrator, then from the project folder run:
#   npm run allow-lan
# Or: powershell -ExecutionPolicy Bypass -File scripts/allow-dev-port.ps1

$ErrorActionPreference = "Stop"

$port = 3000
if ($env:PORT -match '^\d+$') {
  $port = [int]$env:PORT
}

$ruleName = "Flowmaster dev (TCP $port)"

$existing = Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue
if ($existing) {
  Write-Host "Firewall rule already exists: $ruleName"
  exit 0
}

try {
  New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort $port | Out-Null
  Write-Host "Created rule: $ruleName (inbound TCP $port)"
  Write-Host "Restart npm run dev if it was already running, then use the http://<your-LAN-IP>:$port URL from other devices."
}
catch {
  Write-Host "Failed to create firewall rule. Run this script from an elevated (Administrator) PowerShell."
  Write-Host $_.Exception.Message
  exit 1
}
