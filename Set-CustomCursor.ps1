# Variables
$sourcePath = "<Path to cursor files>"
$cursorDir = Join-Path -Path $env:windir -ChildPath "<Child Path>"
$schemeName = "Custom Cursor Scheme"

# List of cursor files
$cursorFiles = @(
    "pointer.cur"
    "help.cur"
    "working.ani"
    "busy.ani"
    "precision.cur"
    "handwriting.cur"
    "unavailable.cur"
    "vert.cur"
    "horz.cur"
    "dgn1.cur"
    "dgn2.cur"
    "move.cur"
    "alternate.cur"
    "link.cur"
    "beam.cur"
    "pin.cur"
    "person.cur"
)

# Map of cursor roles to files
$cursorMappings = @{
    "Arrow"          = "pointer.cur"
    "Help"           = "help.cur"
    "AppStarting"    = "working.ani"
    "Wait"           = "busy.ani"
    "Crosshair"      = "precision.cur"
    "IBeam"          = "beam.cur"
    "NWPen"          = "handwriting.cur"
    "No"             = "unavailable.cur"
    "SizeNS"         = "vert.cur"
    "SizeWE"         = "horz.cur"
    "SizeNWSE"       = "dgn1.cur"
    "SizeNESW"       = "dgn2.cur"
    "SizeAll"        = "move.cur"
    "UpArrow"        = "alternate.cur"
    "Hand"           = "link.cur"
    "Person"         = "person.cur"
    "Pin"            = "pin.cur"
}

# Create the cursor directory if it doesn't exist
if (-Not (Test-Path -Path $cursorDir)) {
    New-Item -ItemType Directory -Path $cursorDir -Force
}

# Copy the cursor files to the destination directory
foreach ($file in $cursorFiles) {
    Copy-Item -Path (Join-Path -Path $sourcePath -ChildPath $file) -Destination $cursorDir -Force
}

# Construct the scheme string
$schemeCursors = @()
foreach ($role in @(
    "Arrow",
    "Help",
    "AppStarting",
    "Wait",
    "Crosshair",
    "IBeam",
    "NWPen",
    "No",
    "SizeNS",
    "SizeWE",
    "SizeNWSE",
    "SizeNESW",
    "SizeAll",
    "UpArrow",
    "Hand",
    "Person",
    "Pin"
)) {
    $cursorFile = $cursorMappings[$role]
    $cursorPath = "%SystemRoot%\\Cursors\\<Child Path>\\$cursorFile"
    $schemeCursors += $cursorPath
}

$schemeString = [string]::Join(",", $schemeCursors)

# Add the new scheme to the registry
$schemeRegistryKey = "HKCU:\Control Panel\Cursors\Schemes"
if (-Not (Test-Path $schemeRegistryKey)) {
    New-Item -Path $schemeRegistryKey -Force | Out-Null
}
Set-ItemProperty -Path $schemeRegistryKey -Name $schemeName -Value $schemeString

# Apply the new cursor scheme
$cursorRegistryKey = "HKCU:\Control Panel\Cursors"
Set-ItemProperty -Path $cursorRegistryKey -Name "(Default)" -Value $schemeName

foreach ($role in $cursorMappings.Keys) {
    $cursorFile = $cursorMappings[$role]
    $cursorPath = "$env:SystemRoot\Cursors\<Child Path>\$cursorFile"
    Set-ItemProperty -Path $cursorRegistryKey -Name $role -Value $cursorPath
}

# Refresh cursor settings
$signature = @"
[DllImport("user32.dll", CharSet = CharSet.Auto)]
public static extern IntPtr SystemParametersInfo(uint uiAction, uint uiParam, string pvParam, uint fWinIni);
"@
Add-Type -MemberDefinition $signature -Name "Win32" -Namespace Win32Functions

$SPI_SETCURSORS = 0x0057
$SPIF_UPDATEINIFILE = 0x01
$SPIF_SENDCHANGE = 0x02

[Win32Functions.Win32]::SystemParametersInfo($SPI_SETCURSORS, 0, $null, $SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE) | Out-Null
