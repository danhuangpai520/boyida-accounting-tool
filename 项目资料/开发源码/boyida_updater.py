import base64
import hashlib
import json
import re
import ssl
import subprocess
import time
import urllib.error
import urllib.request
from pathlib import Path


GITHUB_RELEASE_API = "https://api.github.com/repos/danhuangpai520/boyida-accounting-tool/releases/latest"
GITHUB_RELEASES_URL = "https://github.com/danhuangpai520/boyida-accounting-tool/releases"
GITHUB_LATEST_URL = "https://github.com/danhuangpai520/boyida-accounting-tool/releases/latest"
GITHUB_DOWNLOAD_BASE = "https://github.com/danhuangpai520/boyida-accounting-tool/releases/download"
RAW_VERSION_URL = "https://raw.githubusercontent.com/danhuangpai520/boyida-accounting-tool/main/VERSION"


def parse_version_tuple(value: str) -> tuple[int, ...]:
    text = str(value or "").strip().lower()
    text = text[1:] if text.startswith("v") else text
    parts = []
    for token in re.split(r"[.\-_+]", text):
        match = re.match(r"(\d+)", token)
        if match:
            parts.append(int(match.group(1)))
        elif parts:
            break
    return tuple(parts or [0])


def is_newer_version(latest: str, current: str) -> bool:
    left = list(parse_version_tuple(latest))
    right = list(parse_version_tuple(current))
    width = max(len(left), len(right))
    left.extend([0] * (width - len(left)))
    right.extend([0] * (width - len(right)))
    return tuple(left) > tuple(right)


def is_certificate_error(exc: BaseException) -> bool:
    reason = getattr(exc, "reason", None)
    text = f"{exc!r} {reason!r}".lower()
    return (
        isinstance(exc, ssl.SSLError)
        or isinstance(reason, ssl.SSLError)
        or "certificate_verify_failed" in text
        or "certificate verify failed" in text
        or "unable to get local issuer certificate" in text
    )


def _ps_literal(value: str | Path) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def _run_windows_powershell(script: str, timeout: int = 30) -> str:
    prologue = """
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
"""
    encoded = base64.b64encode((prologue + "\n" + script).encode("utf-16le")).decode("ascii")
    completed = subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-EncodedCommand",
            encoded,
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout or "").strip()
        raise RuntimeError(detail or f"PowerShell 更新通道失败，退出码 {completed.returncode}")
    return (completed.stdout or "").strip()


def pick_release_asset(payload: dict) -> dict:
    assets = payload.get("assets") if isinstance(payload, dict) else []
    if not isinstance(assets, list):
        assets = []
    exe_assets = [
        item for item in assets
        if isinstance(item, dict) and str(item.get("name") or "").lower().endswith(".exe")
    ]
    if not exe_assets:
        raise RuntimeError("GitHub 最新版本没有找到 EXE 附件。")

    def score(item: dict) -> int:
        name = str(item.get("name") or "").lower()
        value = 0
        if "boyida" in name or "accounting" in name or "做账" in name:
            value += 20
        if "tool" in name or "工具" in name:
            value += 10
        if "setup" in name or "installer" in name:
            value -= 5
        return value

    return sorted(exe_assets, key=score, reverse=True)[0]


def release_info_from_payload(payload: dict) -> dict:
    tag = str(payload.get("tag_name") or payload.get("name") or "").strip()
    if not tag:
        raise RuntimeError("GitHub 最新版本没有版本号。")
    asset = pick_release_asset(payload)
    download_url = str(asset.get("browser_download_url") or asset.get("url") or "").strip()
    if not download_url:
        raise RuntimeError("GitHub 最新版本附件没有下载地址。")
    digest = str(asset.get("digest") or "").strip().lower()
    sha256 = digest.removeprefix("sha256:") if digest.startswith("sha256:") else ""
    return {
        "tag": tag,
        "name": str(payload.get("name") or tag).strip(),
        "body": str(payload.get("body") or "").strip(),
        "html_url": str(payload.get("html_url") or GITHUB_RELEASES_URL).strip(),
        "asset_name": str(asset.get("name") or "做账执行工具.exe").strip(),
        "download_url": download_url,
        "size": int(asset.get("size") or 0),
        "sha256": sha256.upper(),
    }


def _release_info_from_tag(tag: str, source: str) -> dict:
    asset_name = f"boyida-accounting-tool-{tag}.exe"
    return {
        "tag": tag,
        "name": f"做账执行工具 {tag}",
        "body": "",
        "html_url": f"{GITHUB_RELEASES_URL}/tag/{tag}",
        "asset_name": asset_name,
        "download_url": f"{GITHUB_DOWNLOAD_BASE}/{tag}/{asset_name}",
        "size": 0,
        "sha256": "",
        "source": source,
    }


def fetch_latest_release_info_by_windows_web(timeout: int = 20) -> dict:
    """Use Windows' certificate store when Python/PyInstaller CA roots are unavailable."""
    script = f"""
$headers = @{{ Accept = 'application/vnd.github+json'; 'User-Agent' = 'BoyidaAccountingTool' }}
$payload = Invoke-RestMethod -Uri {_ps_literal(GITHUB_RELEASE_API)} -Headers $headers -TimeoutSec {max(int(timeout), 1)}
$payload | ConvertTo-Json -Depth 12 -Compress
"""
    try:
        raw = _run_windows_powershell(script, timeout=max(int(timeout) + 8, 20))
        info = release_info_from_payload(json.loads(raw))
        info["source"] = "windows_system_certificate"
        return info
    except Exception:
        return fetch_latest_release_info_by_windows_redirect(timeout=timeout)


def fetch_latest_release_info_by_windows_redirect(timeout: int = 20) -> dict:
    script = f"""
$request = [System.Net.HttpWebRequest][System.Net.WebRequest]::Create({_ps_literal(GITHUB_LATEST_URL)})
$request.Method = 'GET'
$request.AllowAutoRedirect = $false
$request.UserAgent = 'BoyidaAccountingTool'
try {{
    $response = $request.GetResponse()
}} catch [System.Net.WebException] {{
    if ($_.Exception.Response) {{
        $response = $_.Exception.Response
    }} else {{
        throw
    }}
}}
try {{
    $location = $response.Headers['Location']
    if (-not $location) {{
        $location = $response.ResponseUri.AbsoluteUri
    }}
    Write-Output $location
}} finally {{
    if ($response) {{ $response.Close() }}
}}
"""
    location = _run_windows_powershell(script, timeout=max(int(timeout) + 8, 20)).strip()
    if location.startswith("/"):
        location = "https://github.com" + location
    match = re.search(r"/releases/tag/([^/?#]+)", location)
    if not match:
        raise RuntimeError("Windows 更新通道无法识别 GitHub 最新版本地址。")
    return _release_info_from_tag(match.group(1), "windows_system_certificate")


def fetch_latest_release_info(current_version: str, timeout: int = 12) -> dict:
    request = urllib.request.Request(
        GITHUB_RELEASE_API,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": f"BoyidaAccountingTool/{current_version}",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        if exc.code == 403:
            try:
                return fetch_latest_release_info_by_redirect(timeout=timeout)
            except Exception:
                return fetch_latest_release_info_by_raw_version(timeout=timeout)
        if exc.code == 404:
            raise RuntimeError("GitHub 更新源不可访问；请确认仓库 Release 是否公开可下载。") from exc
        raise
    except urllib.error.URLError as exc:
        if is_certificate_error(exc):
            return fetch_latest_release_info_by_windows_web(timeout=max(timeout, 20))
        raise
    except ssl.SSLError:
        return fetch_latest_release_info_by_windows_web(timeout=max(timeout, 20))
    info = release_info_from_payload(payload)
    info["source"] = "python_github_api"
    return info


def fetch_latest_release_info_by_redirect(timeout: int = 12) -> dict:
    request = urllib.request.Request(
        GITHUB_LATEST_URL,
        headers={"User-Agent": "BoyidaAccountingTool"},
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            final_url = response.geturl()
    except urllib.error.URLError as exc:
        if is_certificate_error(exc):
            return fetch_latest_release_info_by_windows_redirect(timeout=max(timeout, 20))
        raise
    except ssl.SSLError:
        return fetch_latest_release_info_by_windows_redirect(timeout=max(timeout, 20))
    match = re.search(r"/releases/tag/([^/?#]+)", final_url)
    if not match:
        raise RuntimeError("GitHub 最新版本跳转地址无法识别。")
    return _release_info_from_tag(match.group(1), "python_release_redirect")


def fetch_latest_release_info_by_raw_version(timeout: int = 12) -> dict:
    request = urllib.request.Request(
        RAW_VERSION_URL,
        headers={"User-Agent": "BoyidaAccountingTool"},
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            version = response.read().decode("utf-8").strip()
    except urllib.error.URLError as exc:
        if is_certificate_error(exc):
            return fetch_latest_release_info_by_windows_redirect(timeout=max(timeout, 20))
        return fetch_latest_release_info_by_redirect(timeout=timeout)
    except ssl.SSLError:
        return fetch_latest_release_info_by_windows_redirect(timeout=max(timeout, 20))
    except Exception:
        return fetch_latest_release_info_by_redirect(timeout=timeout)
    tag = version if version.lower().startswith("v") else f"v{version}"
    return _release_info_from_tag(tag, "raw_version")


def download_update_exe(info: dict, target: Path, current_version: str, progress_fn=None) -> str:
    target.parent.mkdir(parents=True, exist_ok=True)
    request = urllib.request.Request(
        info["download_url"],
        headers={"User-Agent": f"BoyidaAccountingTool/{current_version}"},
    )
    digest = hashlib.sha256()
    received = 0
    try:
        with urllib.request.urlopen(request, timeout=30) as response, target.open("wb") as f:
            total = int(response.headers.get("Content-Length") or info.get("size") or 0)
            while True:
                chunk = response.read(1024 * 256)
                if not chunk:
                    break
                f.write(chunk)
                digest.update(chunk)
                received += len(chunk)
                if progress_fn and total:
                    progress_fn(received, total)
    except urllib.error.URLError as exc:
        if is_certificate_error(exc):
            target.unlink(missing_ok=True)
            return download_update_exe_by_windows_web(info, target, current_version, progress_fn=progress_fn)
        raise
    except ssl.SSLError:
        target.unlink(missing_ok=True)
        return download_update_exe_by_windows_web(info, target, current_version, progress_fn=progress_fn)
    actual_sha = digest.hexdigest().upper()
    expected_sha = str(info.get("sha256") or "").strip().upper()
    if expected_sha and actual_sha != expected_sha:
        target.unlink(missing_ok=True)
        raise RuntimeError(f"更新文件校验失败：期望 {expected_sha}，实际 {actual_sha}")
    if target.stat().st_size <= 0:
        target.unlink(missing_ok=True)
        raise RuntimeError("更新文件下载为空。")
    return actual_sha


def download_update_exe_by_windows_web(info: dict, target: Path, current_version: str, progress_fn=None) -> str:
    target.parent.mkdir(parents=True, exist_ok=True)
    if progress_fn:
        progress_fn(1, 4)
    script = f"""
$client = New-Object System.Net.WebClient
$client.Headers.Add('User-Agent', {_ps_literal('BoyidaAccountingTool/' + str(current_version))})
$client.DownloadFile({_ps_literal(info['download_url'])}, {_ps_literal(target)})
"""
    try:
        if progress_fn:
            progress_fn(2, 4)
        _run_windows_powershell(script, timeout=180)
    except Exception:
        target.unlink(missing_ok=True)
        raise
    if progress_fn:
        progress_fn(3, 4)
    if target.stat().st_size <= 0:
        target.unlink(missing_ok=True)
        raise RuntimeError("更新文件下载为空。")
    digest = hashlib.sha256()
    with target.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    actual_sha = digest.hexdigest().upper()
    expected_sha = str(info.get("sha256") or "").strip().upper()
    if expected_sha and actual_sha != expected_sha:
        target.unlink(missing_ok=True)
        raise RuntimeError(f"更新文件校验失败：期望 {expected_sha}，实际 {actual_sha}")
    if progress_fn:
        progress_fn(4, 4)
    return actual_sha


def write_update_launcher(new_exe: Path, current_exe: Path, pid: int, script_dir: Path) -> Path:
    stamp = time.strftime("%Y%m%d_%H%M%S")
    script = script_dir / f"apply_update_{stamp}.cmd"
    backup_dir = current_exe.parent / "更新备份"
    backup_exe = backup_dir / f"{current_exe.stem}_更新前_{stamp}{current_exe.suffix}"
    lines = [
        "@echo off",
        "chcp 65001 >nul",
        f'set "SRC={new_exe}"',
        f'set "DST={current_exe}"',
        f'set "BACKUPDIR={backup_dir}"',
        f'set "BACKUP={backup_exe}"',
        f'set "PID={pid}"',
        "echo 正在等待主程序退出...",
        ":wait_loop",
        'tasklist /FI "PID eq %PID%" | find "%PID%" >nul',
        "if not errorlevel 1 (",
        "  timeout /t 1 /nobreak >nul",
        "  goto wait_loop",
        ")",
        'if not exist "%SRC%" exit /b 2',
        'mkdir "%BACKUPDIR%" 2>nul',
        'if exist "%DST%" copy /Y "%DST%" "%BACKUP%" >nul',
        'copy /Y "%SRC%" "%DST%" >nul',
        "if errorlevel 1 exit /b 3",
        'start "" "%DST%"',
        'del "%~f0"',
    ]
    script.parent.mkdir(parents=True, exist_ok=True)
    script.write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")
    return script
