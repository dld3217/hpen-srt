# bump.ps1 — increment patch version in package-solution.json, SrtDashboard.tsx, and SseRequestForm.tsx
# Usage: .\bump.ps1            (increments patch: 1.0.1.0 → 1.0.2.0)
#        .\bump.ps1 -Minor     (increments minor: 1.0.1.0 → 1.1.0.0)

param([switch]$Minor)

$pkgPath  = "$PSScriptRoot\config\package-solution.json"
$dashPath = "$PSScriptRoot\src\webparts\srtDashboard\components\SrtDashboard.tsx"
$formPath = "$PSScriptRoot\src\webparts\sseRequestForm\components\SseRequestForm.tsx"

# Read current version from package-solution.json
$pkg = Get-Content $pkgPath -Raw | ConvertFrom-Json
$parts = $pkg.solution.version -split '\.'   # ["1","0","1","0"]
$maj = [int]$parts[0]
$min = [int]$parts[1]
$pat = [int]$parts[2]

if ($Minor) {
  $min++
  $pat = 0
} else {
  $pat++
}

$newFull  = "$maj.$min.$pat.0"        # 4-part for package-solution.json
$newShort = "$maj.$min.$pat"          # 3-part for VERSION constant

# Update solution version in package-solution.json (not the feature version inside "features")
# We use ConvertFrom-Json → modify → ConvertTo-Json to avoid regex hitting multiple version fields
$pkg = Get-Content $pkgPath -Raw | ConvertFrom-Json
$pkg.solution.version = $newFull
$pkg | ConvertTo-Json -Depth 10 | Set-Content $pkgPath -Encoding UTF8

# Update VERSION constant in SrtDashboard.tsx
$dash = Get-Content $dashPath -Raw
$dash = $dash -replace "const VERSION = '[\d.]+'", "const VERSION = '$newShort'"
Set-Content $dashPath $dash -NoNewline -Encoding UTF8

# Update VERSION constant in SseRequestForm.tsx
$form = Get-Content $formPath -Raw
$form = $form -replace "const VERSION = '[\d.]+'", "const VERSION = '$newShort'"
Set-Content $formPath $form -NoNewline -Encoding UTF8

Write-Host "Bumped to $newFull" -ForegroundColor Green
Write-Host "  config/package-solution.json  → $newFull"
Write-Host "  SrtDashboard.tsx VERSION       → '$newShort'"
Write-Host "  SseRequestForm.tsx VERSION     → '$newShort'"
