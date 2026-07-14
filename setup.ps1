$ErrorActionPreference = "Stop"

Set-Location -LiteralPath $PSScriptRoot

$venvPython = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host "Creating local Python environment in .venv..."

    $created = $false

    if (Get-Command py -ErrorAction SilentlyContinue) {
        try {
            & py -3.12 -m venv .venv
            $created = $true
        }
        catch {
            Write-Host "Python 3.12 was not available through the py launcher; trying default Python 3..."
        }

        if (-not $created) {
            & py -3 -m venv .venv
            $created = $true
        }
    }

    if (-not $created) {
        & python -m venv .venv
    }
}
else {
    Write-Host "Using existing .venv."
}

Write-Host "Installing project dependencies..."
& $venvPython -m pip install -r requirements.txt

Write-Host ""
Write-Host "Setup complete."
Write-Host "In VS Code, open this folder and use the integrated terminal."
Write-Host "Start the local frontend:"
Write-Host ".\.venv\Scripts\python.exe frontend_server.py"
Write-Host ""
Write-Host "Paste test command:"
Write-Host ".\.venv\Scripts\python.exe paste_sender.py --sheet `"Geo TTh`" --row 2 --class-review-file class_review.txt --mode paste-only"
