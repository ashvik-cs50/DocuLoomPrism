$ErrorActionPreference = "Stop"

$appName = "DocuLoomPrism-Debug"
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
    "--console",
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
}

$arguments += (Join-Path $PSScriptRoot $mainFile)

Write-Host "Building $appName ..."
& $pythonCmd @arguments

Write-Host "Debug build complete. Output:"
Write-Host (Join-Path $PSScriptRoot "dist\$appName.exe")
