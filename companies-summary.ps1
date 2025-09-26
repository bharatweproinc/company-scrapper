$Email = "rohitkumar12901230@gmail.com"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Mouse Event Support ---
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Mouse {
    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
}
"@

# --- JSON Response Helper ---
function Write-JsonResponse {
    param([bool]$Success, [string]$Message, $Data)
    $json = [PSCustomObject]@{
        success  = $Success
        message  = $Message
        data     = $Data
    } | ConvertTo-Json -Depth 10 -Compress
    Write-Output $json
    exit
}

# --- SendKeys Helper ---
function Send-Keys {
    param([string]$Keys, [int]$Delay = 1)
    [System.Windows.Forms.SendKeys]::SendWait($Keys)
    Start-Sleep -Seconds $Delay
}

# --- Find Chrome Profile by Email ---
function Get-ProfileByEmail {
    param([string]$EmailToMatch)
    $userDataDir = "$env:LOCALAPPDATA\Google\Chrome\User Data"
    $profiles = Get-ChildItem -Path $userDataDir -Directory | Where-Object { $_.Name -match "^(Default|Profile \d+|Profile)$" }
    foreach ($p in $profiles) {
        $prefsFile = Join-Path $p.FullName "Preferences"
        if (Test-Path $prefsFile) {
            try {
                $json = Get-Content $prefsFile -Raw | ConvertFrom-Json
                $profileEmail = $json.account_info.email | Select-Object -First 1
                if ($profileEmail -eq $EmailToMatch) { return $p.Name }
            } catch {}
        }
    }
    return $null
}

# --- Click Screen ---
function Click-Screen([int]$x, [int]$y) {
    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
    [Mouse]::mouse_event(0x02,0,0,0,0) # left down
    [Mouse]::mouse_event(0x04,0,0,0,0) # left up
}

# --- Open Chrome ---
$profile = Get-ProfileByEmail $Email
if (-not $profile) { Write-JsonResponse $false "No Chrome profile found for '$Email'" $null }

$chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$link = "https://product.costar.com/suiteapps/owners/companies/detail/h10ewby/summary"
Start-Process -FilePath $chromePath -ArgumentList "--profile-directory=`"$profile`"", "$link"

$outputFile = "data.json"
$maxPages = 2

for ($i = 0; $i -lt $maxPages; $i++) {

    Write-Output "Waiting for Chrome to open / page load... ($i)"
    Start-Sleep -Seconds 10

    # --- Open DevTools & select all ---
    [System.Windows.Forms.SendKeys]::SendWait("^(+i)")  # Ctrl+Shift+I
    Start-Sleep -Seconds 3
    [System.Windows.Forms.SendKeys]::SendWait("^(a)")    # Ctrl+A
    Start-Sleep -Seconds 2
    [System.Windows.Forms.SendKeys]::SendWait("^(c)")    # Ctrl+C
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait("^(+i)")  # Ctrl+Shift+I close
    Start-Sleep -Seconds 1

    # --- Get HTML from clipboard ---
    $html = $null
    $retry = 0
    while (-not $html -and $retry -lt 5) {
        Start-Sleep -Milliseconds 500
        $html = Get-Clipboard -Format Text
        $retry++
    }
    if (-not $html) { Write-JsonResponse $false "Clipboard empty!" $null }

    # --- Extract label/value pairs ---
    $pattern = '<span class="csg-tui-text.*?csg-ic-fact-label.*?"[^>]*>(.*?)</span>\s*<span class="csg-tui-text.*?csg-ic-fact-value.*?"[^>]*>(.*?)</span>'
    $data = @{}

    [regex]::Matches($html, $pattern) | ForEach-Object {
        $label = $_.Groups[1].Value.Trim()
        $value = $_.Groups[2].Value.Trim()
        if (-not [string]::IsNullOrWhiteSpace($label)) {
            # Remove extra spaces/newlines
            $data[$label] = $value -replace '\s+', ' '
        }
    }

    # --- Save JSON ---
    $data | ConvertTo-Json -Depth 10 | Set-Content -Path $outputFile
    Write-Output "Data saved to $outputFile"

    # Optional: click to next page or break if only 1 page
    break
}

Write-JsonResponse $true "Data extracted successfully" $data
