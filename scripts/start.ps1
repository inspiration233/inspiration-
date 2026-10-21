param([switch]$CheckOnly)
$ErrorActionPreference = 'Stop'
$workspacePath = Split-Path -Parent $PSScriptRoot
$appPath = Join-Path $workspacePath 'app'
$electronPath = Join-Path $appPath 'node_modules\electron\dist\electron.exe'

try {
    if (-not (Test-Path -LiteralPath $electronPath -PathType Leaf)) {
        throw 'Electron is missing. Run "npm ci" in the app folder, then start again.'
    }
    $candidates = [System.Collections.Generic.List[string]]::new()
    if ($env:TOOLBOX_PYTHON) {
        $candidates.Add($env:TOOLBOX_PYTHON)
    }
    foreach ($relativePath in @('.venv\Scripts\python.exe', 'venv\Scripts\python.exe')) {
        $candidatePath = Join-Path $workspacePath $relativePath
        if (Test-Path -LiteralPath $candidatePath -PathType Leaf) {
            $candidates.Add($candidatePath)
        }
    }
    $systemPython = Get-Command python -ErrorAction SilentlyContinue
    if ($systemPython) { $candidates.Add($systemPython.Source) }
    $pythonPath = $null
    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { continue }
        try {
            & $candidate -c "import openpyxl, xlrd; assert tuple(map(int, openpyxl.__version__.split('.')[:2])) >= (3, 1)" 2>$null
        } catch {
            continue
        }
        if ($LASTEXITCODE -eq 0) {
            $pythonPath = $candidate
            break
        }
    }
    if (-not $pythonPath) {
        throw 'Python dependencies are missing. Run "python -m pip install -r requirements.txt", or set TOOLBOX_PYTHON to a Python executable with openpyxl >= 3.1 and xlrd.'
    }
    $env:TOOLBOX_PYTHON = $pythonPath
    $env:PYTHONUTF8 = '1'
    if ($CheckOnly) {
        Write-Output "Ready: $electronPath"
        Write-Output "Python: $pythonPath"
    } else {
        # This is the interactive application window the user is opening.
        Start-Process -FilePath $electronPath -ArgumentList ('"' + $appPath + '"') -WorkingDirectory $appPath -WindowStyle Normal
    }
    exit 0
} catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
