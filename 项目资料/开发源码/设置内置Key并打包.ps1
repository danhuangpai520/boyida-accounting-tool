param(
    [string]$KeyFile = "builtin_api_key.txt"
)

$ErrorActionPreference = "Stop"
$SourceRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent (Split-Path -Parent $SourceRoot)
Set-Location -LiteralPath $SourceRoot

$OutputExeName = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String("5YGa6LSm5omn6KGM5bel5YW3LmV4ZQ=="))
$OutputExe = Join-Path $ProjectRoot $OutputExeName
$BuildStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$VersionRoot = Join-Path $ProjectRoot "项目资料\版本备份"
New-Item -ItemType Directory -Path $VersionRoot -Force | Out-Null

function Copy-IfExists {
    param(
        [string]$Path,
        [string]$Destination
    )
    if (Test-Path -LiteralPath $Path) {
        Copy-Item -LiteralPath $Path -Destination $Destination -Force
        return $true
    }
    return $false
}

function Save-VersionManifest {
    param(
        [string]$Directory,
        [string]$Kind,
        [string]$ExePath
    )
    $Hash = ""
    $Size = 0
    if (Test-Path -LiteralPath $ExePath) {
        $Hash = (Get-FileHash -LiteralPath $ExePath -Algorithm SHA256).Hash
        $Size = (Get-Item -LiteralPath $ExePath).Length
    }
    $Manifest = [ordered]@{
        time = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        kind = $Kind
        exe = (Split-Path -Leaf $ExePath)
        size_bytes = $Size
        sha256 = $Hash
        note = "用于本地回退；不包含明文 API Key。"
    }
    $ManifestPath = Join-Path $Directory "版本说明.json"
    $Manifest | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $ManifestPath -Encoding UTF8
}

$Python = "C:\Users\Administrator\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if (-not (Test-Path -LiteralPath $Python)) {
    $Python = "python"
}

if (Test-Path -LiteralPath $OutputExe) {
    $OldVersionDir = Join-Path $VersionRoot "${BuildStamp}_构建前旧版"
    New-Item -ItemType Directory -Path $OldVersionDir -Force | Out-Null
    Copy-Item -LiteralPath $OutputExe -Destination (Join-Path $OldVersionDir $OutputExeName) -Force
    Save-VersionManifest -Directory $OldVersionDir -Kind "构建前旧版" -ExePath (Join-Path $OldVersionDir $OutputExeName)
}

if (Test-Path -LiteralPath $KeyFile) {
    $ApiKey = (Get-Content -LiteralPath $KeyFile -Raw -Encoding UTF8).Trim()
} elseif (-not [string]::IsNullOrWhiteSpace($env:ZHIPU_API_KEY)) {
    $ApiKey = $env:ZHIPU_API_KEY.Trim()
} else {
    $SavedKeyB64 = & $Python -c "import base64, zhipu_accounting_app as app; key=app.load_saved_api_key().strip(); print(base64.b64encode(key.encode('utf-8')).decode('ascii') if key else '')"
    if (-not [string]::IsNullOrWhiteSpace($SavedKeyB64)) {
        $ApiKey = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($SavedKeyB64.Trim()))
    } else {
        $Secure = Read-Host "Paste ZHIPU API Key (hidden)" -AsSecureString
        $Ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
        try {
            $ApiKey = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($Ptr).Trim()
        } finally {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($Ptr)
        }
    }
}

if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    throw "API Key is empty. Build stopped."
}

$B64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($ApiKey))
$KeyModule = Join-Path $SourceRoot "embedded_default_key.py"
$KeyModuleContent = "BUILTIN_API_KEY_B64 = `"$B64`"`r`n"
[IO.File]::WriteAllText($KeyModule, $KeyModuleContent, [Text.UTF8Encoding]::new($false))

& $Python -m py_compile "zhipu_accounting_app.py"
& $Python "zhipu_accounting_app.py" --self-test
& $Python -m PyInstaller --noconfirm --clean --onefile --windowed --name zhipu_accounting_tool --icon "assets\boyida_truck.ico" --add-data "assets\boyida_truck.png;assets" --add-data "assets\boyida_truck.ico;assets" --add-data "assets\jingzhe_header_line.png;assets" zhipu_accounting_app.py

Copy-Item -LiteralPath "dist\zhipu_accounting_tool.exe" -Destination $OutputExe -Force

[IO.File]::WriteAllText($KeyModule, "BUILTIN_API_KEY_B64 = `"`"`r`n", [Text.UTF8Encoding]::new($false))

$NewVersionDir = Join-Path $VersionRoot "${BuildStamp}_新版"
New-Item -ItemType Directory -Path $NewVersionDir -Force | Out-Null
Copy-Item -LiteralPath $OutputExe -Destination (Join-Path $NewVersionDir $OutputExeName) -Force
Copy-Item -LiteralPath (Join-Path $SourceRoot "zhipu_accounting_app.py") -Destination (Join-Path $NewVersionDir "zhipu_accounting_app.py") -Force
Copy-IfExists -Path (Join-Path $SourceRoot "设置内置Key并打包.ps1") -Destination (Join-Path $NewVersionDir "设置内置Key并打包.ps1") | Out-Null
Copy-IfExists -Path (Join-Path $SourceRoot "zhipu_accounting_tool.spec") -Destination (Join-Path $NewVersionDir "zhipu_accounting_tool.spec") | Out-Null
Copy-IfExists -Path (Join-Path $ProjectRoot "做账执行规范.md") -Destination (Join-Path $NewVersionDir "做账执行规范.md") | Out-Null
Copy-IfExists -Path (Join-Path $ProjectRoot "项目资料\说明与发布\使用说明.md") -Destination (Join-Path $NewVersionDir "使用说明.md") | Out-Null
Save-VersionManifest -Directory $NewVersionDir -Kind "新版" -ExePath (Join-Path $NewVersionDir $OutputExeName)

Write-Host ""
Write-Host "Built EXE with embedded default API Key:"
Write-Host $OutputExe
Write-Host "Version backup:"
Write-Host $NewVersionDir
Write-Host "The temporary key module has been cleared. The EXE still contains the embedded default key."
Write-Host "If $KeyFile is temporary, delete it now."
