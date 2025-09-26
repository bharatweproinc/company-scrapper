param(
    [int]$totalRecord = 20
)
$Email = "bharat.weproinc@gmail.com"
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
$link = "https://product.costar.com/tenants/companies"
Start-Process -FilePath $chromePath -ArgumentList "--profile-directory=`"$profile`"", "$link"

$outputFile = "data.json"
$maxPages = 2

$totalPages = $totalRecord / 20;
$finalData = @()
for ($i = 0; $i -lt $totalPages; $i++) {

    Write-Output "Waiting for Chrome to open / page load... ($i)"
    Start-Sleep -Seconds 8

    # --- Open DevTools & select all ---
    [System.Windows.Forms.SendKeys]::SendWait("^(+i)")  # Ctrl+Shift+I to open DevTools
    Start-Sleep -Seconds 3
    [System.Windows.Forms.SendKeys]::SendWait("^(a)")    # Ctrl+A
    Start-Sleep -Seconds 4
    [System.Windows.Forms.SendKeys]::SendWait("^(c)")    # Ctrl+C
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait("^(+i)")  # Ctrl+Shift+I to close DevTools
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

    $htmlWrapped = "<html><body>$html</body></html>"
    $doc = New-Object -ComObject "HTMLFile"
    $doc.IHTMLDocument2_write($htmlWrapped)
    $doc.Close()

    # --- Extract table ---
    $table = $doc.getElementsByTagName("table") | Select-Object -First 1
    if (-not $table) { Write-JsonResponse $false "No table found in HTML" $null }

    $headers = @(
        "Index",
        "Tenant Name",
        "Industry",
        "Territory",
        "HQMarket",
        "Locations",
        "SFOccupied",
        "HighestUseBySF",
        "Employees",
        "Growth",
        "Revenue",
        "Credit Rating",
        "Established",
        "ParentCompany",
        "Website",
        "HQPhone",
        "HQCity",
        "HQState",
        "HQPostalCode",
        "HQCountry",
        "NAICS",
        "SIC"
    )

    # --- Extract rows ---
    $data = @()
    $rows = $table.getElementsByTagName("tr")
    # --- Extract rows safely ---
    foreach ($row in $rows) {
        $cells = $row.getElementsByTagName("td")
        if (-not $cells -or $cells.Length -eq 0) { continue }

        $item = @{}

        for ($j = 0; $j -lt $cells.Length; $j++) {
            $cell = $cells.Item($j)
            if (-not $cell) { continue }

            # Use a default header if index exceeds headers array
            $headerKey = if ($j -lt $headers.Length) { $headers[$j] } else { "Column$j" }

            $aTag = $cell.getElementsByTagName("a")
            if ($aTag.Length -gt 0 -and $aTag.Item(0)) {
                $item[$headerKey] = @{
                    Text = if ($aTag.Item(0).innerText) { $aTag.Item(0).innerText.Trim() } else { "" }
                    Href = if ($aTag.Item(0).href) { $aTag.Item(0).href } else { "" }
                }
            } else {
                $item[$headerKey] = if ($cell.innerText) { $cell.innerText.Trim() } else { "" }
            }
        }

        $data += $item
        $finalData += $item
    }


    # --- Save JSON ---
    if ($i -eq 0) {
        $json = $data | ConvertTo-Json -Depth 10
        Set-Content -Path $outputFile -Value $json -Encoding UTF8
    } else {
        $json = $data | ConvertTo-Json -Depth 10
        $existing = Get-Content $outputFile -Raw | ConvertFrom-Json
        $combined = $existing + $data
        $combined | ConvertTo-Json -Depth 10 | Set-Content -Path $outputFile -Encoding UTF8
    }
    

    Write-Output "Page $i data saved to $outputFile"

    # --- Move to next page if exists ---
    if ($i+1 -lt $totalPages) {
        $screenWidth  = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width
        $screenHeight = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Height
        $nextX = $screenWidth - 520
        $nextY = $screenHeight - 80
        Click-Screen $nextX $nextY
        Start-Sleep -Seconds 3
    }
}

Write-JsonResponse $true "Data extraction completed" $finalData