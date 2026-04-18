$ErrorActionPreference = "Stop"

$appName = "DocuLoomPrism"
$mainFile = "doculoom_prism.py"
$iconPath = Join-Path $PSScriptRoot "assets\prizm.ico"
$pythonCmd = "python"

Write-Host "Installing build dependency..."
& $pythonCmd -m pip install -r (Join-Path $PSScriptRoot "build_requirements.txt")

Write-Host "Using Python interpreter:"
& $pythonCmd -c "import sys; print(sys.executable); print(sys.version)"

$arguments = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--windowed",
    "--onefile",
    "--hidden-import", "docx",
    "--hidden-import", "pptx",
    "--hidden-import", "pptx.util",
    "--hidden-import", "pypdf",
    "--hidden-import", "PIL.Image",
    "--hidden-import", "PIL.ImageTk",
    "--name", $appName
)

if (Test-Path $iconPath) {
    $arguments += @("--icon", $iconPath)
    Write-Host "Using icon: $iconPath"
} else {
    Write-Host "No icon found at assets\prizm.ico. Building without a custom icon for now."
}

$arguments += (Join-Path $PSScriptRoot $mainFile)

Write-Host "Building $appName ..."
& $pythonCmd @arguments

Write-Host "Build complete. Output:"
Write-Host (Join-Path $PSScriptRoot "dist\$appName.exe")
