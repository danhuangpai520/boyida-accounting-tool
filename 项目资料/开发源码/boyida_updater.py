import hashlib
import json
import re
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
            return fetch_latest_release_info_by_raw_version(timeout=timeout)
        if exc.code == 404:
            raise RuntimeError("GitHub 更新源不可访问；请确认仓库 Release 是否对当前电脑公开可下载。") from exc
        raise
    return release_info_from_payload(payload)


def fetch_latest_release_info_by_redirect(timeout: int = 12) -> dict:
    request = urllib.request.Request(
        GITHUB_LATEST_URL,
        headers={"User-Agent": "BoyidaAccountingTool"},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        final_url = response.geturl()
    match = re.search(r"/releases/tag/([^/?#]+)", final_url)
    if not match:
        raise RuntimeError("GitHub 最新版本跳转地址无法识别。")
    tag = match.group(1)
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
    }


def fetch_latest_release_info_by_raw_version(timeout: int = 12) -> dict:
    request = urllib.request.Request(
        RAW_VERSION_URL,
        headers={"User-Agent": "BoyidaAccountingTool"},
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            version = response.read().decode("utf-8").strip()
    except Exception:
        return fetch_latest_release_info_by_redirect(timeout=timeout)
    tag = version if version.lower().startswith("v") else f"v{version}"
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
    }


def download_update_exe(info: dict, target: Path, current_version: str, progress_fn=None) -> str:
    target.parent.mkdir(parents=True, exist_ok=True)
    request = urllib.request.Request(
        info["download_url"],
        headers={"User-Agent": f"BoyidaAccountingTool/{current_version}"},
    )
    digest = hashlib.sha256()
    received = 0
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
    actual_sha = digest.hexdigest().upper()
    expected_sha = str(info.get("sha256") or "").strip().upper()
    if expected_sha and actual_sha != expected_sha:
        target.unlink(missing_ok=True)
        raise RuntimeError(f"更新文件校验失败：期望 {expected_sha}，实际 {actual_sha}")
    if target.stat().st_size <= 0:
        target.unlink(missing_ok=True)
        raise RuntimeError("更新文件下载为空。")
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
