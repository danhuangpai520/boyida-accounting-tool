import csv
import base64
import ctypes
import ctypes.wintypes as wintypes
import gzip
import hashlib
import html
import json
import mimetypes
import os
import platform
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
import zipfile
from collections import defaultdict
from datetime import date
from pathlib import Path
from tkinter import BooleanVar, Canvas, DoubleVar, END, Menu, PhotoImage, StringVar, Text, Tk, Toplevel, filedialog, messagebox, simpledialog, ttk

import xlsxwriter

try:
    from embedded_default_key import BUILTIN_API_KEY_B64
except Exception:
    BUILTIN_API_KEY_B64 = ""


COMPANY_NAME = "保谊达"
APP_TITLE = f"{COMPANY_NAME}车队做账工具"
IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
HEADERS = ["日期", "车号", "驾驶员", "重量", "货物名称", "装货地", "收货地", "备注", "原图"]
OCR_ENGINE_LABELS = {
    "glm": "在线智谱 GLM-OCR（默认）",
    "paddle": "本地 PaddleOCR PP-OCRv5（需安装环境）",
}
OCR_ENGINE_BY_LABEL = {label: key for key, label in OCR_ENGINE_LABELS.items()}
OCR_PROFILE_LABELS = {
    "stable": "稳妥：PP-OCRv5(ch) 标准推理 [CPU/GPU]",
    "fast": "极速：PP-OCRv5(ch)+HPI batch=16 [GPU优先]",
    "trt": "极限：PP-OCRv5(ch)+TensorRT/FP16 [仅GPU]",
}
OCR_PROFILE_BY_LABEL = {label: key for key, label in OCR_PROFILE_LABELS.items()}
OCR_PROFILE_DETAILS = {
    "stable": "同一套 PP-OCRv5 中文 OCR，开启方向分类、图像矫正、文本行方向；CPU/GPU 都可用，作为自动兜底基线。",
    "fast": "同一套 PP-OCRv5 中文 OCR，启用 HPI 高性能推理并提高 batch；GPU 优先，缺 ultra-infer/HPI 时会自动降到稳妥档。",
    "trt": "同一套 PP-OCRv5 中文 OCR，尝试 HPI + TensorRT + FP16；仅适合 GPU 环境，失败会自动降到极速/稳妥档。",
}
DEFAULT_EXCEL_IMAGE_MODE = "external"
EXCEL_IMAGE_MODES = {"external", "embedded"}


class SuppressPopupError(RuntimeError):
    """Task failed after internal retries; keep the details in the progress log."""

DEFAULT_EXCEL_COLUMNS = [
    {"enabled": True, "title": "日期", "source": "日期", "width": 12},
    {"enabled": True, "title": "车号", "source": "车号", "width": 12},
    {"enabled": True, "title": "驾驶员", "source": "驾驶员", "width": 10},
    {"enabled": True, "title": "重量", "source": "重量", "width": 10},
    {"enabled": True, "title": "货物名称", "source": "货物名称", "width": 12},
    {"enabled": True, "title": "装货地", "source": "装货地", "width": 12},
    {"enabled": True, "title": "收货地", "source": "收货地", "width": 12},
    {"enabled": True, "title": "备注", "source": "备注", "width": 30},
    {"enabled": True, "title": "原图", "source": "原图", "width": 18},
]
GLM_OCR_URL = "https://open.bigmodel.cn/api/paas/v4/layout_parsing"
CONFIG_FILE_NAME = "config.json"
GPS_URL = "https://www.zjhzwt.com/"
GPS_SUMMARY_FILE_NAME = "gps_summary.json"
GPS_API_CONFIG_FILE_NAME = "gps_api_config.json"
EXCEL_TEMPLATE_FILE_NAME = "excel_template.json"
GPS_API_DEFAULT_URL = "https://www.zjhzwt.com/gps-web/h5/mgr/car?getTeamsAndCars"
GPS_API_DEFAULT_METHOD = "GET"
LOGO_PNG = Path("assets") / "boyida_truck.png"
LOGO_ICO = Path("assets") / "boyida_truck.ico"
JINGZHE_HEADER_PNG = Path("assets") / "jingzhe_header_line.png"
DEFAULT_GPS_SUMMARY = {
    "company": "杭州保谊达供应链有限公司",
    "total": "18",
    "online": "8",
    "offline": "10",
    "running": "0",
    "stopped": "8",
    "alarm": "0",
    "updated_at": "2026-06-02 03:17",
    "source": "快照",
    "status": "未配置实时接口",
}
THEME = {
    "bg": "#030711",
    "panel": "#07101b",
    "panel_2": "#0b1726",
    "panel_3": "#0e2134",
    "card": "#081421",
    "line": "#1b3a56",
    "line_soft": "#0e263a",
    "text": "#e8f7ff",
    "muted": "#7ca7c4",
    "cyan": "#00d7ff",
    "cyan_dark": "#007fa8",
    "cyan_soft": "#0b4158",
    "orange": "#f2b84b",
    "green": "#3dffad",
    "red": "#ff6178",
    "input": "#060d18",
    "titlebar": "#02050b",
    "log": "#040a13",
    "log_section": "#082339",
    "selected": "#0f3b55",
}


def adaptive_window_metrics(screen_w: int, screen_h: int) -> dict:
    """Calculate a launch size that fits small VM desktops without hiding logs."""
    screen_w = max(int(screen_w or 0), 760)
    screen_h = max(int(screen_h or 0), 560)
    compact = screen_w < 1280 or screen_h < 820

    width = min(1220, max(920, int(screen_w * 0.94)))
    height = min(690, max(610, int(screen_h * 0.78)))
    width = min(width, max(720, screen_w - 24))
    height = min(height, max(540, screen_h - 56))

    min_width = 1080 if not compact else 860
    min_height = 600 if not compact else 560
    min_width = min(min_width, width)
    min_height = min(min_height, height)

    return {
        "width": width,
        "height": height,
        "x": max(0, (screen_w - width) // 2),
        "y": max(0, min((screen_h - height) // 3, 32)),
        "min_width": min_width,
        "min_height": min_height,
        "compact": compact,
    }

DRIVER_BY_PLATE = {
    "沪A19615A": "应云超",
    "沪A25282A": "王从新",
    "沪A26202A": "颜丙龙",
    "浙A05633D": "陈实",
    "浙A06033D": "张华勇",
    "浙A07336D": "孙安足",
    "浙A08165D": "郑小洪",
    "浙A30985D": "文兴俊",
    "浙A31318D": "乔金飞",
    "浙A32233D": "小老表",
    "浙A32236D": "范刚",
    "浙A33875D": "冀明",
    "浙A35228D": "侯泽亮",
    "浙A36185D": "孙时奎",
    "浙A82J38": "何大山",
}


OCR_WORKER = r'''
import argparse
import json
import time
from pathlib import Path

from paddleocr import PaddleOCR


IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}


def load_existing(path):
    if not path.exists():
        return {"engine": {}, "records": []}
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path, data):
    tmp = path.with_suffix(path.suffix + ".tmp")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    tmp.replace(path)


def build_kwargs(profile, device):
    is_cpu = device == "cpu"
    base = {
        "lang": "ch",
        "ocr_version": "PP-OCRv5",
        "device": device,
        "use_doc_orientation_classify": True,
        "use_doc_unwarping": True,
        "use_textline_orientation": True,
        "text_recognition_batch_size": 2 if is_cpu else 8,
        "textline_orientation_batch_size": 2 if is_cpu else 8,
    }
    if not is_cpu and profile in {"fast", "trt"}:
        base["enable_hpi"] = True
        base["text_recognition_batch_size"] = 16
        base["textline_orientation_batch_size"] = 16
    if not is_cpu and profile == "trt":
        base["use_tensorrt"] = True
        base["precision"] = "fp16"
    return base


def init_ocr(profile, device):
    kwargs = build_kwargs(profile, device)
    attempts = [
        kwargs,
        {k: v for k, v in kwargs.items() if k not in {"use_tensorrt", "precision"}},
        {k: v for k, v in kwargs.items() if k not in {"use_tensorrt", "precision", "enable_hpi"}},
        {k: v for k, v in kwargs.items() if k not in {"use_tensorrt", "precision", "enable_hpi", "text_recognition_batch_size", "textline_orientation_batch_size"}},
    ]
    last_exc = None
    for item in attempts:
        try:
            print("ocr_kwargs=" + json.dumps(item, ensure_ascii=False), flush=True)
            return PaddleOCR(**item), item
        except TypeError as exc:
            last_exc = exc
            print("ocr_init_retry=" + repr(exc), flush=True)
    raise last_exc


def normalize_result_json(result):
    payload = result.json
    if callable(payload):
        payload = payload()
    return payload.get("res", payload)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--profile", choices=["stable", "fast", "trt"], default="stable")
    parser.add_argument("--device", choices=["gpu:0", "cpu"], default="gpu:0")
    parser.add_argument("--limit", type=int, default=None)
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    output = Path(args.output)
    images = sorted(p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES)

    data = load_existing(output)
    data.setdefault("records", [])
    done = {record.get("file_name") for record in data["records"]}
    pending = [p for p in images if p.name not in done]
    if args.limit:
        pending = pending[:args.limit]

    print(f"total_images={len(images)} existing={len(done)} pending={len(pending)}", flush=True)
    if not pending:
        save_json(output, data)
        return

    t0 = time.time()
    ocr, used_kwargs = init_ocr(args.profile, args.device)
    data["engine"] = {
        "name": "PaddleOCR",
        "ocr_version": "PP-OCRv5",
        "profile": args.profile,
        "kwargs": used_kwargs,
    }
    print(f"ocr_init_sec={time.time() - t0:.2f}", flush=True)

    for index, image in enumerate(pending, 1):
        started = time.time()
        try:
            result = ocr.predict(str(image))[0]
            payload = normalize_result_json(result)
            rec_texts = payload.get("rec_texts", [])
            rec_scores = payload.get("rec_scores", [])
            angle = (payload.get("doc_preprocessor_res", {}) or {}).get("angle")
            data["records"].append({
                "file_name": image.name,
                "file_path": str(image.resolve()),
                "elapsed_sec": round(time.time() - started, 3),
                "doc_angle": angle,
                "rec_texts": rec_texts,
                "rec_scores": rec_scores,
                "ocr": payload,
            })
            status = f"ok texts={len(rec_texts)}"
        except Exception as exc:
            data["records"].append({
                "file_name": image.name,
                "file_path": str(image.resolve()),
                "elapsed_sec": round(time.time() - started, 3),
                "error": repr(exc),
            })
            status = "error " + repr(exc)
        save_json(output, data)
        print(f"[{index}/{len(pending)}] {image.name} {status} sec={time.time() - started:.2f}", flush=True)


if __name__ == "__main__":
    main()
'''


def compact(text: str) -> str:
    return re.sub(r"\s+", "", text or "").replace("：", ":").upper()


def edit_distance(a: str, b: str) -> int:
    dp = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        prev = dp[0]
        dp[0] = i
        for j, cb in enumerate(b, 1):
            old = dp[j]
            dp[j] = min(dp[j] + 1, dp[j - 1] + 1, prev + (ca != cb))
            prev = old
    return dp[-1]


def nearest_plate(candidate: str) -> str | None:
    if candidate in DRIVER_BY_PLATE:
        return candidate
    distances = sorted((edit_distance(candidate, plate), plate) for plate in DRIVER_BY_PLATE)
    if distances and distances[0][0] <= 1 and (len(distances) == 1 or distances[1][0] > distances[0][0]):
        return distances[0][1]
    return None


def normalize_plate(raw: str) -> tuple[str, str | None]:
    merged = compact(raw).replace("O", "0").replace("I", "1")
    candidates = re.findall(r"[沪浙冀苏][A-Z][0-9A-Z]{5,6}", merged)
    if not candidates:
        return "", None
    candidate = candidates[0]
    plate = nearest_plate(candidate)
    return (plate or candidate), (candidate if plate and plate != candidate else None)


def parse_number(text: str) -> str:
    merged = compact(text)
    for pattern in (r"[ASB]\d{8,}", r"8?2026\d{7,}", r"20?26\d{8,}"):
        match = re.search(pattern, merged)
        if match:
            return match.group(0)
    return ""


def number_date(number: str) -> tuple[int, int, int] | None:
    compact_no = compact(number)
    patterns = (
        r"[ASB](\d{2})(\d{2})(\d{2})\d{2,}",
        r"8?20(\d{2})(\d{2})(\d{2})\d{2,}",
        r"20(\d{2})(\d{2})(\d{2})\d{2,}",
    )
    for pattern in patterns:
        match = re.search(pattern, compact_no)
        if not match:
            continue
        y, mo, d = map(int, match.groups())
        candidate = (2000 + y, mo, d)
        try:
            date(*candidate)
            return candidate
        except ValueError:
            pass
    return None


def parse_dates(texts: list[str], number: str) -> tuple[str, list[str], list[str]]:
    text = " ".join(texts)
    merged = compact(text)
    found: list[tuple[int, int, int]] = []
    for y, m, d in re.findall(r"(20\d{2})[-/.年]?(\d{1,2})[-/.月]?(\d{1,2})", text):
        found.append((int(y), int(m), int(d)))
    for y, m, d in re.findall(r"(20\d{2})(\d{2})(\d{2})", merged):
        found.append((int(y), int(m), int(d)))
    from_number = number_date(number)
    if from_number:
        found.insert(0, from_number)

    valid = []
    for y, m, d in found:
        try:
            date(y, m, d)
        except ValueError:
            continue
        if (y, m, d) not in valid:
            valid.append((y, m, d))

    if not valid:
        return "", [], []
    primary = from_number if from_number in valid else valid[0]
    date_text = f"{primary[0]}/{primary[1]}/{primary[2]}"
    seen = [f"{y}/{m}/{d}" for y, m, d in valid]
    remarks = ["日期冲突需核对"] if len(set(seen)) > 1 else []
    return date_text, seen, remarks


def text_numbers(text: str) -> list[int]:
    value_text = compact(text)
    if re.search(r"20\d{2}[-/.年]?\d{1,2}", value_text):
        return []
    if re.search(r"[沪浙冀苏][A-Z][0-9A-Z]{5,6}", value_text):
        return []
    values = []
    for token in re.findall(r"\d{4,6}", value_text):
        value = int(token)
        if 1000 <= value <= 200000:
            values.append(value)
    return values


def strip_markup(text: str) -> str:
    plain = re.sub(r"<[^>]+>", " ", html.unescape(str(text)))
    return re.sub(r"\s+", " ", plain).strip()


def table_cells(text: str) -> list[str]:
    cells = re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", str(text), flags=re.IGNORECASE | re.DOTALL)
    return [strip_markup(cell) for cell in cells if strip_markup(cell)]


def table_value_after_label(texts: list[str], labels: tuple[str, ...]) -> str:
    label_set = {compact(label) for label in labels}
    stop_labels = {
        "货名", "品名", "规格", "类别", "车号", "毛重", "皮重", "空重", "净重", "结算重量",
        "发货单位", "供货单位", "收货单位", "毛重时间", "皮重时间", "空重时间", "金额", "司磅员",
        "监磅员", "验磅员", "经办人",
    }
    for text in texts:
        cells = table_cells(text)
        for index, cell in enumerate(cells[:-1]):
            key = compact(cell)
            if not any(label in key for label in label_set):
                continue
            value = cells[index + 1].strip()
            value_key = compact(value)
            if value_key and value_key not in stop_labels:
                return value
    return ""


def weight_numbers(text: str) -> list[int]:
    value_text = strip_markup(text)
    values: list[int] = []
    for token in re.findall(r"\d+(?:\.\d+)?", value_text):
        value = float(token)
        if value <= 0:
            continue
        if value < 200 and re.search(r"吨|\\bt\\b", value_text, flags=re.IGNORECASE):
            value *= 1000
        if 1000 <= value <= 200000:
            values.append(int(round(value)))
    return values


def number_near(texts: list[str], index: int, labels: tuple[str, ...], scan_limit: int = 5) -> int | None:
    token_text = strip_markup(texts[index])
    token = compact(token_text)
    stop_words = (
        "货名", "品名", "规格", "类别", "车号", "毛重", "皮重", "空重", "净重", "结算重量",
        "发货单位", "供货单位", "收货单位", "毛重时间", "皮重时间", "空重时间", "金额", "司磅员",
        "监磅员", "验磅员", "经办人",
    )
    for label in labels:
        if label not in token:
            continue
        label_pos = token_text.find(label)
        tail = token_text[label_pos + len(label): label_pos + len(label) + 80] if label_pos >= 0 else token.split(label, 1)[-1][:80]
        stop_positions = [tail.find(stop) for stop in stop_words if stop != label and tail.find(stop) >= 0]
        if stop_positions:
            tail = tail[:min(stop_positions)]
        nums = weight_numbers(tail)
        if nums:
            return nums[0]
        for j in range(index + 1, min(len(texts), index + scan_limit)):
            ahead = compact(texts[j])
            if any(stop in ahead for stop in ("毛重时间", "皮重时间", "空重时间", "金额", "出门", "客户", "仓库", "过磅")):
                continue
            nums = weight_numbers(texts[j])
            if nums:
                return nums[0]
    return None


def parse_weight(record: dict) -> tuple[str, list[str]]:
    texts = [str(t).strip() for t in record.get("rec_texts", []) if str(t).strip()]
    labels = {
        "settlement": ("结算重量", "结算重后", "超系重装"),
        "net": ("净重", "净车", "净星", "净里"),
        "gross": ("毛重", "毛果"),
        "tare": ("皮重", "空重", "免量", "空里"),
    }
    found = {key: [] for key in labels}
    for key, names in labels.items():
        table_value = table_value_after_label(texts, names)
        if table_value:
            found[key].extend(weight_numbers(table_value))
    for i in range(len(texts)):
        for key, names in labels.items():
            value = number_near(texts, i, names, scan_limit=8)
            if value is not None:
                found[key].append(value)

    diff = None
    if found["gross"] and found["tare"] and found["gross"][0] > found["tare"][0]:
        diff = found["gross"][0] - found["tare"][0]

    remarks = []
    value = found["settlement"][0] if found["settlement"] else (found["net"][0] if found["net"] else diff)
    if value is None:
        return "", ["重量模糊"]
    if found["settlement"] and found["net"] and abs(found["settlement"][0] - found["net"][0]) > 20:
        remarks.append("重量冲突需核对")
    if found["settlement"] and diff and abs(found["settlement"][0] - diff) > 20:
        remarks.append("重量冲突需核对")
    if value < 1000 and diff:
        value = diff
        remarks.append("重量模糊")
    return f"{value / 1000:.2f}", remarks


def parse_cargo(texts: list[str]) -> tuple[str, list[str]]:
    known_other = {
        "杂压块": "杂压块",
        "法兰头": "法兰头",
        "冲件料": "冲件料",
        "剪料": "剪料",
        "废铁": "废铁",
        "钢刨花": "钢刨花",
        "数控刨花": "数控刨花",
        "数控创花": "数控刨花",
        "数控刨化": "数控刨花",
        "料头": "料头",
        "废钢": "废钢",
    }
    cargo_text = table_value_after_label(texts, ("货名", "品名"))
    scan_tokens = ([cargo_text] if cargo_text else []) + texts
    for token in scan_tokens:
        c = compact(token)
        if any(word in c for word in ("钢渣", "钠渣", "钾渣", "锅渣", "铆渣", "朝渣")):
            return "钢渣", []
        if "卷子板" in c or ("卷" in c and "板" in c) or any(word in c for word in ("卷予板", "电子板", "客户板", "番子板", "卷子领", "卷子饭", "签予领")):
            return "卷子板", []
        if any(word in c for word in ("钢坯", "钢坏", "钢还", "合金", "钠坯", "销坯", "销址", "合金坯", "合合坯")):
            return "钢坯", []
        for needle, normalized in known_other.items():
            if needle in c:
                return normalized, []

    stop_words = {"车号", "存根", "结算", "类别", "毛重", "发货单位", "供货单位", "空重", "皮重", "收货单位", "净重", "毛重时间", "空重时间", "皮重时间", "金额", "监磅员", "司磅员", "经办人", "仓库", "过磅", "计量单位", "序号", "规格"}
    for i, token in enumerate(texts):
        c = compact(token)
        if "货名" not in c:
            continue
        tail = c.split("货名", 1)[-1]
        if tail and not any(stop in tail for stop in stop_words):
            return token.split("货名", 1)[-1].strip(), []
        for j in range(i + 1, min(len(texts), i + 4)):
            nxt = compact(texts[j])
            if not nxt or any(stop in nxt for stop in stop_words) or re.search(r"\d", nxt):
                continue
            return texts[j].strip(), []
    return "", ["品名待核"]


def classify_route(cargo: str, texts: list[str]) -> tuple[str, str, list[str]]:
    compact_tokens = [compact(text) for text in texts]
    if cargo == "钢渣":
        return "保利", "新登江丰", []
    if cargo == "卷子板":
        return "友谊", "保利", []
    if cargo == "钢坯":
        if any("浙江钧延贸易有限公司过磅码单" in token or "浙江钧延贸易有限公司过磅单" in token for token in compact_tokens):
            return "友谊", "保利", []
        independent_title = any(token in {"过磅码单", "过磅单"} for token in compact_tokens)
        company = any(token == "浙江钧延贸易有限公司" for token in compact_tokens)
        if independent_title and company:
            return "友谊", "新登江丰", []
        return "友谊", "", ["钢坯收货地待核"]
    return "暂定", "暂定", []


def file_hash(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def parse_record(record: dict) -> dict:
    texts = [str(t).strip() for t in record.get("rec_texts", []) if str(t).strip()]
    joined = " ".join(texts)
    remarks: list[str] = []

    number = parse_number(joined)
    date_text, date_values, date_remarks = parse_dates(texts, number)
    remarks.extend(date_remarks)
    if not date_text:
        remarks.append("日期模糊")

    plate, corrected_from = normalize_plate(joined)
    if corrected_from:
        remarks.append(f"车号已由OCR修正:{corrected_from}")
    if not plate:
        remarks.append("缺少车号")
    driver = DRIVER_BY_PLATE.get(plate, "")
    if plate and not driver:
        remarks.append("未匹配到驾驶员")

    weight, weight_remarks = parse_weight(record)
    remarks.extend(weight_remarks)

    cargo, cargo_remarks = parse_cargo(texts)
    remarks.extend(cargo_remarks)
    load_place, receive_place, route_remarks = classify_route(cargo, texts)
    remarks.extend(route_remarks)

    image_path = Path(record.get("file_path", ""))
    return {
        "日期": date_text,
        "车号": plate,
        "驾驶员": driver,
        "重量": weight,
        "货物名称": cargo,
        "装货地": load_place,
        "收货地": receive_place,
        "备注": "；".join(dict.fromkeys(remarks)),
        "原图": record.get("file_name", ""),
        "_file_path": str(image_path),
        "_file_name": record.get("file_name", ""),
        "_number": number,
        "_dates_seen": date_values,
        "_hash": file_hash(image_path) if image_path.exists() else "",
        "_ocr_text": joined,
    }


def duplicate_key(row: dict) -> tuple | None:
    if not all([row["_number"], row["车号"], row["重量"], row["货物名称"]]):
        return None
    return row["_number"], row["车号"], row["重量"], row["货物名称"], tuple(row["_dates_seen"])


def apply_duplicates(rows: list[dict]) -> tuple[list[dict], list[dict]]:
    groups: dict[tuple, list[int]] = defaultdict(list)
    for i, row in enumerate(rows):
        if row["_hash"]:
            groups[("hash", row["_hash"])].append(i)
        key = duplicate_key(row)
        if key:
            groups[("fields", key)].append(i)

    keep_indexes = set(range(len(rows)))
    duplicate_rows = []
    used = set()
    for indexes in groups.values():
        unique = sorted(set(indexes))
        if len(unique) <= 1:
            continue
        keep = unique[0]
        duplicates = [i for i in unique[1:] if i not in used]
        if not duplicates:
            continue
        used.update(duplicates)
        for i in duplicates:
            keep_indexes.discard(i)
            duplicate_rows.append({"重复图片": rows[i]["_file_name"], "对应原图": rows[keep]["_file_name"]})
        note = f"重复{len(duplicates)}张已合并"
        rows[keep]["备注"] = f"{rows[keep]['备注']}；{note}" if rows[keep]["备注"] else note
    return [row for i, row in enumerate(rows) if i in keep_indexes], duplicate_rows


def sortable_date(value: str) -> tuple[int, int, int]:
    match = re.match(r"(\d{4})/(\d{1,2})/(\d{1,2})", value or "")
    if not match:
        return 9999, 12, 31
    return tuple(map(int, match.groups()))


def date_range(rows: list[dict]) -> str:
    dates = [sortable_date(row.get("日期", "")) for row in rows if row.get("日期")]
    dates = [d for d in dates if d[0] != 9999]
    if not dates:
        return "未知日期"
    start, end = min(dates), max(dates)
    if start == end:
        return f"{start[0]}.{start[1]}.{start[2]}"
    return f"{start[0]}.{start[1]}.{start[2]}-{end[0]}.{end[1]}.{end[2]}"


def write_csv(path: Path, rows: list[dict], headers: list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def default_excel_columns() -> list[dict]:
    return [dict(item) for item in DEFAULT_EXCEL_COLUMNS]


def source_from_column_title(title: str) -> str:
    value = str(title or "").strip()
    return value if value in HEADERS else ""


def normalize_excel_columns(columns: object) -> list[dict]:
    source_columns = columns if isinstance(columns, list) else []
    normalized: list[dict] = []
    for item in source_columns:
        if not isinstance(item, dict):
            continue
        source = "" if item.get("source") is None else str(item.get("source", "")).strip()
        title = str(item.get("title") or source or "空白列").strip()
        if not title:
            title = "空白列"
        try:
            width = int(item.get("width", 12))
        except (TypeError, ValueError):
            width = 12
        width = max(6, min(width, 60))
        normalized.append({
            "enabled": bool(item.get("enabled", True)),
            "title": title[:40],
            "source": source[:40],
            "width": width,
        })
    if not normalized or not any(item["enabled"] for item in normalized):
        return default_excel_columns()
    return normalized


def normalize_excel_image_mode(value: object) -> str:
    mode = str(value or "").strip().lower()
    return mode if mode in EXCEL_IMAGE_MODES else DEFAULT_EXCEL_IMAGE_MODE


def excel_template_path() -> Path:
    return app_base_dir() / EXCEL_TEMPLATE_FILE_NAME


def load_excel_template() -> dict:
    path = excel_template_path()
    if not path.exists():
        return {"columns": default_excel_columns(), "image_mode": DEFAULT_EXCEL_IMAGE_MODE}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        columns = data.get("columns", data) if isinstance(data, dict) else data
        image_mode = data.get("image_mode") if isinstance(data, dict) else DEFAULT_EXCEL_IMAGE_MODE
        return {
            "columns": normalize_excel_columns(columns),
            "image_mode": normalize_excel_image_mode(image_mode),
        }
    except Exception:
        return {"columns": default_excel_columns(), "image_mode": DEFAULT_EXCEL_IMAGE_MODE}


def load_excel_columns() -> list[dict]:
    return load_excel_template()["columns"]


def load_excel_image_mode() -> str:
    return load_excel_template()["image_mode"]


def save_excel_template(columns: list[dict], image_mode: str) -> dict:
    normalized = normalize_excel_columns(columns)
    mode = normalize_excel_image_mode(image_mode)
    path = excel_template_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({"columns": normalized, "image_mode": mode}, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"columns": normalized, "image_mode": mode}


def save_excel_columns(columns: list[dict]) -> list[dict]:
    normalized = save_excel_template(columns, DEFAULT_EXCEL_IMAGE_MODE)["columns"]
    return normalized


class DataBlob(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_ubyte)),
    ]


def config_path() -> Path:
    base = os.environ.get("APPDATA")
    root = Path(base) if base else Path.home() / "AppData" / "Roaming"
    return root / "ZhipuAccountingTool" / CONFIG_FILE_NAME


def config_dir() -> Path:
    return config_path().parent


def dpapi_protect(text: str) -> str:
    if os.name != "nt":
        raise RuntimeError("DPAPI only works on Windows")
    data = text.encode("utf-8")
    in_buffer = ctypes.create_string_buffer(data)
    in_blob = DataBlob(len(data), ctypes.cast(in_buffer, ctypes.POINTER(ctypes.c_ubyte)))
    out_blob = DataBlob()
    ok = ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(in_blob),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(out_blob),
    )
    if not ok:
        raise ctypes.WinError()
    try:
        encrypted = ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(out_blob.pbData)
    return base64.b64encode(encrypted).decode("ascii")


def dpapi_unprotect(value: str) -> str:
    if os.name != "nt":
        raise RuntimeError("DPAPI only works on Windows")
    encrypted = base64.b64decode(value.encode("ascii"))
    in_buffer = ctypes.create_string_buffer(encrypted)
    in_blob = DataBlob(len(encrypted), ctypes.cast(in_buffer, ctypes.POINTER(ctypes.c_ubyte)))
    out_blob = DataBlob()
    ok = ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(in_blob),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(out_blob),
    )
    if not ok:
        raise ctypes.WinError()
    try:
        decrypted = ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(out_blob.pbData)
    return decrypted.decode("utf-8")


def load_config() -> dict:
    path = config_path()
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(data: dict) -> None:
    path = config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_saved_api_key() -> str:
    data = load_config()
    if data.get("api_key_dpapi"):
        try:
            return dpapi_unprotect(data["api_key_dpapi"])
        except Exception:
            return ""
    return data.get("api_key_plain", "")


def builtin_api_key() -> str:
    value = (BUILTIN_API_KEY_B64 or "").strip()
    if not value:
        return ""
    try:
        return base64.b64decode(value.encode("ascii")).decode("utf-8").strip()
    except Exception:
        return ""


def initial_api_key() -> tuple[str, str]:
    env_key = os.environ.get("ZHIPU_API_KEY", "").strip()
    if env_key:
        return env_key, "环境变量"
    saved_key = load_saved_api_key().strip()
    if saved_key:
        return saved_key, "本机已保存"
    embedded_key = builtin_api_key()
    if embedded_key:
        return embedded_key, "EXE 内置默认"
    return "", "未填写"


def save_api_key(api_key: str) -> None:
    key = api_key.strip()
    data = load_config()
    if key:
        try:
            data["api_key_dpapi"] = dpapi_protect(key)
            data.pop("api_key_plain", None)
        except Exception:
            data["api_key_plain"] = key
            data.pop("api_key_dpapi", None)
    else:
        data.pop("api_key_dpapi", None)
        data.pop("api_key_plain", None)
    save_config(data)


def delete_saved_api_key() -> None:
    data = load_config()
    data.pop("api_key_dpapi", None)
    data.pop("api_key_plain", None)
    if data:
        save_config(data)
    else:
        path = config_path()
        if path.exists():
            path.unlink()


def _read_saved_secret(data: dict, name: str) -> str:
    encrypted = data.get(f"{name}_dpapi")
    if encrypted:
        try:
            return dpapi_unprotect(encrypted)
        except Exception:
            return ""
    return data.get(f"{name}_plain", "")


def _write_saved_secret(data: dict, name: str, value: str) -> None:
    data.pop(f"{name}_dpapi", None)
    data.pop(f"{name}_plain", None)
    text = value.strip()
    if not text:
        return
    try:
        data[f"{name}_dpapi"] = dpapi_protect(text)
    except Exception:
        data[f"{name}_plain"] = text


def load_saved_gps_api_config() -> dict:
    data = load_config()
    gps = data.get("gps_api")
    if not isinstance(gps, dict):
        return {}
    config: dict = {}
    for key in ("url", "method", "body", "sortField", "company", "timeout"):
        if gps.get(key) not in (None, ""):
            config[key] = gps[key]
    headers = gps.get("headers")
    if isinstance(headers, dict):
        config["headers"] = headers
    config["authorization"] = _read_saved_secret(gps, "authorization")
    config["cookie"] = _read_saved_secret(gps, "cookie")
    return config


def save_gps_api_config(config: dict) -> None:
    data = load_config()
    gps: dict = {}
    for key in ("url", "sortField", "company"):
        value = str(config.get(key) or "").strip()
        if value:
            gps[key] = value
    method = str(config.get("method") or GPS_API_DEFAULT_METHOD).strip().upper()
    gps["method"] = method if method in {"GET", "POST"} else GPS_API_DEFAULT_METHOD
    body = str(config.get("body") or "").strip()
    if body:
        gps["body"] = body
    try:
        gps["timeout"] = int(config.get("timeout") or 20)
    except Exception:
        gps["timeout"] = 20
    headers = config.get("headers")
    if isinstance(headers, dict) and headers:
        gps["headers"] = headers
    _write_saved_secret(gps, "authorization", str(config.get("authorization") or ""))
    _write_saved_secret(gps, "cookie", str(config.get("cookie") or ""))
    data["gps_api"] = gps
    save_config(data)


def delete_saved_gps_api_config() -> None:
    data = load_config()
    data.pop("gps_api", None)
    if data:
        save_config(data)
    else:
        path = config_path()
        if path.exists():
            path.unlink()


PADDLE_ENV_CHECK_SCRIPT = (
    "import importlib.util, sys\n"
    "def p(value): print(value, flush=True)\n"
    "p('python=' + sys.executable)\n"
    "p('version=' + sys.version.replace('\\n', ' '))\n"
    "for name in ['paddleocr', 'paddle']:\n"
    "    p(name + '=' + str(importlib.util.find_spec(name) is not None))\n"
    "try:\n"
    "    import paddle\n"
    "    p('paddle_version=' + getattr(paddle, '__version__', 'unknown'))\n"
    "    p('cuda_compiled=' + str(paddle.is_compiled_with_cuda()))\n"
    "    try:\n"
    "        paddle.utils.run_check()\n"
    "        p('paddle_run_check=True')\n"
    "    except Exception as exc:\n"
    "        p('paddle_run_check=False ' + repr(exc))\n"
    "except Exception as exc:\n"
    "    p('paddle_import_error=' + repr(exc))\n"
    "try:\n"
    "    import paddleocr\n"
    "    p('paddleocr_version=' + getattr(paddleocr, '__version__', 'unknown'))\n"
    "except Exception as exc:\n"
    "    p('paddleocr_import_error=' + repr(exc))\n"
)
PADDLE_INSTALL_DOC_URL = "https://www.paddleocr.ai/main/en/version3.x/paddlepaddle_installation.html"
PADDLE_GPU_CU118_INDEX = "https://www.paddlepaddle.org.cn/packages/stable/cu118/"
PADDLE_GPU_CU126_INDEX = "https://www.paddlepaddle.org.cn/packages/stable/cu126/"
PADDLE_CPU_INDEX = "https://www.paddlepaddle.org.cn/packages/stable/cpu/"
PADDLE_50_SERIES_WHEELS = {
    "3.9": "https://paddle-qa.bj.bcebos.com/paddle-pipeline/Develop-TagBuild-Training-Windows-Gpu-Cuda12.9-Cudnn9.9-Trt10.5-Mkl-Avx-VS2019-SelfBuiltPypiUse/86d658f56ebf3a5a7b2b33ace48f22d10680d311/paddlepaddle_gpu-3.0.0.dev20250717-cp39-cp39-win_amd64.whl",
    "3.10": "https://paddle-qa.bj.bcebos.com/paddle-pipeline/Develop-TagBuild-Training-Windows-Gpu-Cuda12.9-Cudnn9.9-Trt10.5-Mkl-Avx-VS2019-SelfBuiltPypiUse/86d658f56ebf3a5a7b2b33ace48f22d10680d311/paddlepaddle_gpu-3.0.0.dev20250717-cp310-cp310-win_amd64.whl",
    "3.11": "https://paddle-qa.bj.bcebos.com/paddle-pipeline/Develop-TagBuild-Training-Windows-Gpu-Cuda12.9-Cudnn9.9-Trt10.5-Mkl-Avx-VS2019-SelfBuiltPypiUse/86d658f56ebf3a5a7b2b33ace48f22d10680d311/paddlepaddle_gpu-3.0.0.dev20250717-cp311-cp311-win_amd64.whl",
    "3.12": "https://paddle-qa.bj.bcebos.com/paddle-pipeline/Develop-TagBuild-Training-Windows-Gpu-Cuda12.9-Cudnn9.9-Trt10.5-Mkl-Avx-VS2019-SelfBuiltPypiUse/86d658f56ebf3a5a7b2b33ace48f22d10680d311/paddlepaddle_gpu-3.0.0.dev20250717-cp312-cp312-win_amd64.whl",
}


def hidden_startupinfo():
    if os.name != "nt":
        return None
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    return startupinfo


def run_capture(cmd: list[str], timeout: int = 15) -> tuple[int | None, str]:
    try:
        completed = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            startupinfo=hidden_startupinfo(),
        )
        return completed.returncode, completed.stdout
    except FileNotFoundError as exc:
        return None, repr(exc)
    except subprocess.TimeoutExpired as exc:
        output = exc.stdout or ""
        if isinstance(output, bytes):
            output = output.decode("utf-8", errors="replace")
        return None, output + "\nTIMEOUT"
    except Exception as exc:
        return None, repr(exc)


def python_usable(python: str) -> bool:
    code, output = run_capture([python, "-c", "import sys; print(sys.version_info[:2])"], timeout=10)
    return code == 0 and "(" in output


def python_minor_version(python: str) -> str:
    code, output = run_capture([python, "-c", "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"], timeout=10)
    if code == 0:
        return output.strip().splitlines()[-1].strip()
    return ""


def version_at_least(value: str, minimum: str) -> bool:
    def parts(text: str) -> tuple[int, ...]:
        result = []
        for item in re.findall(r"\d+", text):
            result.append(int(item))
        return tuple(result)
    return parts(value) >= parts(minimum)


def detect_hardware() -> dict:
    info = {
        "system": platform.platform(),
        "machine": platform.machine(),
        "gpus": [],
        "nvidia_smi": False,
        "errors": [],
    }
    code, output = run_capture(
        ["nvidia-smi", "--query-gpu=name,memory.total,driver_version", "--format=csv,noheader,nounits"],
        timeout=8,
    )
    if code == 0:
        info["nvidia_smi"] = True
        for line in output.splitlines():
            parts = [part.strip() for part in line.split(",")]
            if len(parts) < 2:
                continue
            try:
                memory_mb = int(float(parts[1]))
            except ValueError:
                memory_mb = 0
            info["gpus"].append({
                "name": parts[0],
                "memory_mb": memory_mb,
                "driver": parts[2] if len(parts) > 2 else "",
                "source": "nvidia-smi",
            })
        return info
    if output:
        info["errors"].append("nvidia-smi: " + output.strip())

    if os.name == "nt":
        ps = (
            "$ErrorActionPreference='SilentlyContinue'; "
            "Get-CimInstance Win32_VideoController | "
            "Select-Object Name,AdapterRAM | ConvertTo-Json -Compress"
        )
        code, output = run_capture(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps], timeout=10)
        if code == 0 and output.strip():
            try:
                payload = json.loads(output)
                devices = payload if isinstance(payload, list) else [payload]
                for item in devices:
                    name = str(item.get("Name", "")).strip()
                    memory = item.get("AdapterRAM") or 0
                    try:
                        memory_mb = int(memory) // (1024 * 1024)
                    except Exception:
                        memory_mb = 0
                    if name:
                        info["gpus"].append({
                            "name": name,
                            "memory_mb": memory_mb,
                            "driver": "",
                            "source": "Win32_VideoController",
                        })
            except Exception as exc:
                info["errors"].append("Win32_VideoController: " + repr(exc))
    return info


def is_nvidia_gpu(gpu: dict) -> bool:
    name = str(gpu.get("name") or "").lower()
    return "nvidia" in name or "geforce" in name or "rtx" in name or "quadro" in name or "tesla" in name


def nvidia_gpus(hardware: dict) -> list[dict]:
    return [gpu for gpu in (hardware.get("gpus") or []) if is_nvidia_gpu(gpu)]


def best_nvidia_gpu(hardware: dict) -> dict:
    gpus = nvidia_gpus(hardware)
    return max(gpus, key=lambda item: item.get("memory_mb") or 0) if gpus else {}


def local_ocr_device_plan(hardware: dict) -> tuple[str, str]:
    gpu = best_nvidia_gpu(hardware)
    if not gpu:
        names = [str(item.get("name") or "").strip() for item in (hardware.get("gpus") or []) if item.get("name")]
        if names:
            return "cpu", "未检测到可用于 Paddle GPU 的 NVIDIA 显卡；检测到：" + " / ".join(names)
        return "cpu", "未检测到 NVIDIA 显卡；将安装并使用 CPU 版 Paddle，本地 OCR 能跑但速度较慢"
    memory_mb = gpu.get("memory_mb") or 0
    memory_text = f"{round(memory_mb / 1024, 1)}GB" if memory_mb else "显存未知"
    return "gpu", f"检测到 NVIDIA 显卡：{gpu.get('name')} / {memory_text} / driver={gpu.get('driver') or '未知'}"


def check_paddle_environment(python: str) -> dict:
    result = {
        "python": python,
        "returncode": None,
        "lines": [],
        "paddleocr": False,
        "paddle": False,
        "cuda_compiled": False,
        "run_check": False,
        "error": "",
    }
    code, output = run_capture([python, "-u", "-c", PADDLE_ENV_CHECK_SCRIPT], timeout=120)
    result["returncode"] = code
    result["lines"] = [line.rstrip() for line in output.splitlines() if line.strip()]
    if code not in (0, None):
        result["error"] = f"exit={code}"
    for line in result["lines"]:
        if line == "paddleocr=True":
            result["paddleocr"] = True
        elif line == "paddle=True":
            result["paddle"] = True
        elif line == "cuda_compiled=True":
            result["cuda_compiled"] = True
        elif line == "paddle_run_check=True":
            result["run_check"] = True
    return result


def strong_local_gpu(hardware: dict) -> tuple[bool, str]:
    gpus = nvidia_gpus(hardware)
    if not gpus:
        return False, "未检测到独立 NVIDIA GPU"
    best = max(gpus, key=lambda item: item.get("memory_mb") or 0)
    name = best.get("name", "")
    memory_mb = best.get("memory_mb") or 0
    is_50 = any(mark in name for mark in ("RTX 50", "5090", "5080", "5070"))
    strong = is_50 or memory_mb >= 10000
    memory_text = f"{round(memory_mb / 1024, 1)}GB" if memory_mb else "显存未知"
    return strong, f"{name} / {memory_text}"


def local_paddle_install_command(python: str, hardware: dict) -> tuple[str, list[str]]:
    best_gpu = best_nvidia_gpu(hardware)
    gpu_name = str(best_gpu.get("name") or "")
    driver = str(best_gpu.get("driver") or "")
    py_version = python_minor_version(python)
    is_50_series = os.name == "nt" and any(mark in gpu_name for mark in ("RTX 50", "5090", "5080", "5070"))
    if is_50_series and py_version in PADDLE_50_SERIES_WHEELS:
        return (
            f"Windows 50 系显卡专用 PaddlePaddle wheel（Python {py_version}）",
            [python, "-m", "pip", "install", "--upgrade", "--prefer-binary", "--timeout", "120", PADDLE_50_SERIES_WHEELS[py_version]],
        )
    if best_gpu:
        if driver and version_at_least(driver, "550.54.14"):
            return (
                "PaddlePaddle GPU cu126 稳定包",
                [python, "-m", "pip", "install", "--upgrade", "--prefer-binary", "--timeout", "120", "paddlepaddle-gpu==3.2.0", "-i", PADDLE_GPU_CU126_INDEX],
            )
        return (
            "PaddlePaddle GPU cu118 稳定包",
            [python, "-m", "pip", "install", "--upgrade", "--prefer-binary", "--timeout", "120", "paddlepaddle-gpu==3.2.0", "-i", PADDLE_GPU_CU118_INDEX],
        )
    return (
        "PaddlePaddle CPU 稳定包",
        [python, "-m", "pip", "install", "--upgrade", "--prefer-binary", "--timeout", "120", "paddlepaddle==3.2.0", "-i", PADDLE_CPU_INDEX],
    )


def build_ocr_recommendation(hardware: dict, paddle_state: dict | None = None, api_key: str = "") -> tuple[str, list[str]]:
    strong_gpu, gpu_summary = strong_local_gpu(hardware)
    local_ready = paddle_ready(paddle_state)
    api_ready = bool(api_key.strip())
    reasons = [f"硬件：{gpu_summary}"]
    if paddle_state:
        reasons.append(
            "Paddle："
            f"paddleocr={paddle_state.get('paddleocr')} "
            f"paddle={paddle_state.get('paddle')} "
            f"cuda={paddle_state.get('cuda_compiled')} "
            f"run_check={paddle_state.get('run_check')}"
        )
    if strong_gpu and local_ready:
        return "推荐用本地 OCR：显卡强，PaddleOCR GPU 环境也已经能跑。", reasons
    if strong_gpu and not local_ready:
        if api_ready:
            return "推荐先用在线 OCR：检测到强显卡，但没找到能直接跑的本地 PaddleOCR GPU 环境。不是没显卡，是 OCR 环境没装好或没选对 Python。", reasons
        return "先填在线 Key，或补本地 OCR：检测到强显卡，但没找到完整 PaddleOCR GPU 环境。", reasons
    if api_ready:
        return "推荐用在线 OCR：不需要装 Python、Paddle 或 CUDA，最适合把 exe 发给别人。", reasons
    return "先填 ZHIPU API Key：普通电脑建议在线 OCR；本地 OCR 只在显卡和 Paddle 环境都通过时推荐。", reasons


def auth_header(api_key: str) -> str:
    key = api_key.strip()
    if key.lower().startswith("bearer "):
        return key
    return "Bearer " + key


def image_to_data_uri(path: Path) -> str:
    mime = mimetypes.guess_type(path.name)[0] or "image/png"
    data = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{data}"


def post_glm_ocr(file_payload: str, api_key: str, timeout: int) -> object:
    body = json.dumps({"model": "glm-ocr", "file": file_payload}, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        GLM_OCR_URL,
        data=body,
        headers={
            "Authorization": auth_header(api_key),
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            response_body = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"GLM-OCR HTTP {exc.code}: {error_body}") from exc
    return json.loads(response_body)


def extract_glm_texts(payload: object) -> tuple[list[str], object]:
    texts: list[str] = []

    def add_text(value: object) -> None:
        if not isinstance(value, str):
            return
        for line in value.replace("\r", "\n").split("\n"):
            line = line.strip()
            if line:
                texts.append(line)

    def walk(value: object) -> None:
        if isinstance(value, dict):
            for key in ("content", "text", "markdown", "md", "md_results"):
                if key in value:
                    add_text(value[key])
            for key in ("data", "result", "results", "layout_details", "pages", "items", "children"):
                if key in value:
                    walk(value[key])
        elif isinstance(value, list):
            for item in value:
                walk(item)
        elif isinstance(value, str):
            add_text(value)

    walk(payload)
    deduped = list(dict.fromkeys(texts))
    return deduped, payload


def call_glm_ocr(image_path: Path, api_key: str, timeout: int = 300) -> dict:
    data_uri = image_to_data_uri(image_path)
    try:
        payload = post_glm_ocr(data_uri, api_key, timeout)
    except RuntimeError as exc:
        # Official docs say base64 is supported; this fallback covers providers that reject data URI prefixes.
        if "GLM-OCR HTTP 400" not in str(exc) and "GLM-OCR HTTP 422" not in str(exc):
            raise
        payload = post_glm_ocr(data_uri.split(",", 1)[-1], api_key, timeout)
    rec_texts, raw_payload = extract_glm_texts(payload)
    return {"rec_texts": rec_texts, "raw": raw_payload}


def input_images(input_dir: Path) -> list[Path]:
    if not input_dir.exists():
        return []
    return sorted(p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES)


def ocr_record_has_text(record: dict) -> bool:
    if record.get("error"):
        return False
    return any(str(text).strip() for text in record.get("rec_texts", []))


def ocr_json_status(raw_json: Path, images: list[Path]) -> dict:
    image_names = {p.name for p in images}
    status = {
        "exists": raw_json.exists(),
        "complete": False,
        "image_count": len(images),
        "ok_count": 0,
        "missing": list(image_names),
        "invalid": [],
        "extra": [],
        "error": "",
    }
    if not raw_json.exists():
        return status
    try:
        data = json.loads(raw_json.read_text(encoding="utf-8"))
    except Exception as exc:
        status["error"] = f"OCR JSON 损坏：{exc}"
        return status
    records = data.get("records", [])
    if not isinstance(records, list):
        status["error"] = "OCR JSON 缺少 records 列表"
        return status

    ok_names: set[str] = set()
    invalid: list[str] = []
    extra: list[str] = []
    for record in records:
        if not isinstance(record, dict):
            continue
        name = str(record.get("file_name") or "")
        if name not in image_names:
            extra.append(name or "未命名记录")
            continue
        if ocr_record_has_text(record):
            ok_names.add(name)
        else:
            invalid.append(name or "未命名记录")
    missing = sorted(image_names - ok_names)
    status.update({
        "ok_count": len(ok_names),
        "missing": missing,
        "invalid": sorted(set(invalid)),
        "extra": sorted(set(extra)),
        "complete": bool(images) and not missing and not invalid and not extra,
    })
    return status


def looks_like_global_ocr_error(message: str) -> bool:
    text = message.lower()
    needles = (
        "http 401",
        "http 403",
        "http 429",
        "http 500",
        "http 502",
        "http 503",
        "http 504",
        "unauthorized",
        "forbidden",
        "timed out",
        "timeout",
        "winerror",
        "name resolution",
        "connection",
        "certificate",
    )
    return any(needle in text for needle in needles)


def run_glm_ocr_batch(input_dir: Path, output_path: Path, api_key: str, log_fn, progress_fn=None) -> None:
    if not api_key.strip():
        raise RuntimeError("在线 GLM-OCR 模式需要填写 ZHIPU API Key。")
    images = sorted(p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES)
    if not images:
        raise RuntimeError("待处理图片中没有可处理图片。")
    image_names = {p.name for p in images}
    if output_path.exists():
        data = json.loads(output_path.read_text(encoding="utf-8"))
    else:
        data = {"engine": {}, "records": []}
    data.setdefault("records", [])
    kept_records = [
        record for record in data["records"]
        if record.get("file_name") in image_names and ocr_record_has_text(record)
    ]
    done = {record.get("file_name") for record in kept_records}
    pending = [p for p in images if p.name not in done]
    data["records"] = kept_records
    data["engine"] = {"name": "GLM-OCR", "model": "glm-ocr", "api_url": GLM_OCR_URL}
    log_fn(f"total_images={len(images)} existing={len(done)} pending={len(pending)}")
    if progress_fn:
        progress_fn(len(done), len(images), "在线 OCR 准备")
    failures: list[str] = []
    for index, image in enumerate(pending, 1):
        started = time.time()
        try:
            result = call_glm_ocr(image, api_key)
            if not result["rec_texts"]:
                raise RuntimeError("GLM-OCR 未返回可识别文字")
            data["records"].append({
                "file_name": image.name,
                "file_path": str(image.resolve()),
                "elapsed_sec": round(time.time() - started, 3),
                "rec_texts": result["rec_texts"],
                "rec_scores": [],
                "ocr": result["raw"],
            })
            status = f"ok texts={len(result['rec_texts'])}"
        except Exception as exc:
            error_text = repr(exc)
            data["records"].append({
                "file_name": image.name,
                "file_path": str(image.resolve()),
                "elapsed_sec": round(time.time() - started, 3),
                "error": error_text,
                "rec_texts": [],
            })
            failures.append(f"{image.name}: {error_text}")
            status = "error " + error_text
        output_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        log_fn(f"[{index}/{len(pending)}] {image.name} {status} sec={time.time() - started:.2f}")
        if progress_fn:
            progress_fn(len(done) + index, len(images), image.name)
        if failures and not done and len(failures) == 1 and looks_like_global_ocr_error(failures[0]):
            break
    status = ocr_json_status(output_path, images)
    if failures or not status["complete"]:
        sample = "；".join(failures[:3]) if failures else "OCR 文本为空或图片未覆盖完整"
        raise RuntimeError(f"OCR 没跑通，已停止生成 Excel。失败 {len(failures)} 张，原因：{sample}")


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd()


def daily_workspace_name(today: date | None = None) -> str:
    today = today or date.today()
    return f"{today:%Y-%m-%d}_保谊达做账表"


def managed_workspace_dir() -> Path:
    path = app_base_dir() / daily_workspace_name()
    path.mkdir(parents=True, exist_ok=True)
    return path


def resource_path(relative: str | Path) -> Path:
    base = Path(getattr(sys, "_MEIPASS", app_base_dir()))
    return base / relative


def managed_input_dir() -> Path:
    path = managed_workspace_dir() / "待处理图片"
    path.mkdir(parents=True, exist_ok=True)
    return path


def managed_output_dir() -> Path:
    path = managed_workspace_dir() / "输出结果"
    path.mkdir(parents=True, exist_ok=True)
    return path


def gps_summary_path() -> Path:
    return config_dir() / GPS_SUMMARY_FILE_NAME


def legacy_gps_summary_path() -> Path:
    return app_base_dir() / GPS_SUMMARY_FILE_NAME


def gps_api_config_path() -> Path:
    return config_dir() / GPS_API_CONFIG_FILE_NAME


def legacy_gps_api_config_path() -> Path:
    return app_base_dir() / GPS_API_CONFIG_FILE_NAME


def load_gps_summary() -> dict:
    for path in (gps_summary_path(), legacy_gps_summary_path()):
        if not path.exists():
            continue
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            summary = {**DEFAULT_GPS_SUMMARY, **data}
            summary.setdefault("source", "快照")
            return summary
        except Exception:
            pass
    return dict(DEFAULT_GPS_SUMMARY)


def save_gps_summary(summary: dict) -> None:
    gps_summary_path().parent.mkdir(parents=True, exist_ok=True)
    gps_summary_path().write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")


def load_file_gps_api_config() -> dict:
    for path in (gps_api_config_path(), legacy_gps_api_config_path()):
        if not path.exists():
            continue
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except Exception:
            pass
    return {}


def load_gps_api_config() -> dict:
    config = load_file_gps_api_config()
    saved = load_saved_gps_api_config()
    if saved:
        config.update(saved)
    return config


def gps_api_has_auth(config: dict) -> bool:
    if config.get("authorization") or config.get("cookie"):
        return True
    headers = config.get("headers")
    if isinstance(headers, dict):
        return any(("auth" in key.lower() or "token" in key.lower()) and bool(value) for key, value in headers.items())
    return False


def gps_api_configured() -> bool:
    return gps_api_has_auth(load_gps_api_config())


def decode_gps_api_payload(payload: object) -> object:
    if not isinstance(payload, dict) or "status" not in payload:
        return payload
    status = payload.get("status")
    if status != 1:
        raise RuntimeError(str(payload.get("result") or payload.get("message") or f"GPS 接口返回 status={status}"))
    result = payload.get("result")
    if payload.get("encry") and isinstance(result, str):
        raw = bytes(ord(ch) & 0xFF for ch in result)
        result = json.loads(gzip.decompress(raw).decode("utf-8"))
    return result


def first_number_from_keys(data: dict, keys: tuple[str, ...]) -> int | None:
    normalized = {str(key).lower(): value for key, value in data.items()}
    for key in keys:
        if key.lower() not in normalized:
            continue
        try:
            return int(float(str(normalized[key.lower()]).strip()))
        except Exception:
            return None
    return None


def gps_count_summary(value: object, company: str) -> dict | None:
    best: dict | None = None
    best_score = 0
    total_keys = ("total", "totalnum", "totalcount", "all", "allnum", "allcount", "carnum", "carcount", "vehiclecount")
    online_keys = ("online", "onlinenum", "onlinecount")
    offline_keys = ("offline", "offlinenum", "offlinecount")
    running_keys = ("running", "runningnum", "runningcount", "run", "runcount")
    stopped_keys = ("stopped", "stoppednum", "stoppedcount", "stop", "stopcount")
    alarm_keys = ("alarm", "alarmnum", "alarmcount", "warn", "warncount")

    def walk(item: object) -> None:
        nonlocal best, best_score
        if isinstance(item, dict):
            total = first_number_from_keys(item, total_keys)
            online = first_number_from_keys(item, online_keys)
            offline = first_number_from_keys(item, offline_keys)
            running = first_number_from_keys(item, running_keys)
            stopped = first_number_from_keys(item, stopped_keys)
            alarm = first_number_from_keys(item, alarm_keys)
            values = {
                "total": total,
                "online": online,
                "offline": offline,
                "running": running,
                "stopped": stopped,
                "alarm": alarm,
            }
            score = sum(1 for value in values.values() if value is not None)
            if score >= 2 and score > best_score:
                best_score = score
                best = values
            for child in item.values():
                walk(child)
        elif isinstance(item, list):
            for child in item:
                walk(child)

    walk(value)
    if not best:
        return None
    total = best.get("total")
    online = best.get("online")
    offline = best.get("offline")
    if total is None and online is not None and offline is not None:
        total = online + offline
    if offline is None and total is not None and online is not None:
        offline = max(total - online, 0)
    return {
        "company": company or DEFAULT_GPS_SUMMARY["company"],
        "total": str(total if total is not None else "-"),
        "online": str(online if online is not None else "-"),
        "offline": str(offline if offline is not None else "-"),
        "running": str(best.get("running") if best.get("running") is not None else "-"),
        "stopped": str(best.get("stopped") if best.get("stopped") is not None else "-"),
        "alarm": str(best.get("alarm") if best.get("alarm") is not None else "0"),
        "updated_at": time.strftime("%Y-%m-%d %H:%M"),
        "source": "实时接口",
        "status": "接口刷新成功",
    }


def gps_state_summary(result: object, company: str) -> dict:
    cars: list[dict] = []
    states: dict = {}
    if isinstance(result, list) and len(result) >= 3 and isinstance(result[1], list):
        cars = [item for item in result[1] if isinstance(item, dict) and item.get("plate")]
        states = result[2] if isinstance(result[2], dict) else {}
    if not cars:
        cars = find_gps_car_list(result)
    if not cars:
        counted = gps_count_summary(result, company)
        if counted:
            return counted
        raise RuntimeError("GPS 接口返回中没有找到车辆列表或统计字段。")

    total = len(cars)
    online = offline = running = stopped = alarm = 0
    for car in cars:
        car_id = str(car.get("id", ""))
        state = states.get(car_id, car.get("state"))
        try:
            state_int = int(state)
        except Exception:
            state_int = 0
        if state_int >= 5:
            online += 1
            if state_int in {7, 8}:
                running += 1
            if 9 <= state_int <= 12:
                stopped += 1
        else:
            offline += 1
        if state_int in {4, 8, 11, 12}:
            alarm += 1

    return {
        "company": company or DEFAULT_GPS_SUMMARY["company"],
        "total": str(total),
        "online": str(online),
        "offline": str(offline),
        "running": str(running),
        "stopped": str(stopped),
        "alarm": str(alarm),
        "updated_at": time.strftime("%Y-%m-%d %H:%M"),
        "source": "实时接口",
        "status": "接口刷新成功",
    }


def find_gps_car_list(value: object) -> list[dict]:
    best: list[dict] = []

    def walk(item: object) -> None:
        nonlocal best
        if isinstance(item, list):
            candidate = [x for x in item if isinstance(x, dict) and x.get("plate")]
            if len(candidate) > len(best):
                best = candidate
            for child in item:
                walk(child)
        elif isinstance(item, dict):
            for child in item.values():
                walk(child)

    walk(value)
    return best


def refresh_gps_summary_from_api() -> dict:
    config = load_gps_api_config()
    if not gps_api_has_auth(config):
        raise RuntimeError("未配置 GPS 授权，当前只能显示 GPS 快照。")
    endpoint = config.get("url") or GPS_API_DEFAULT_URL
    method = str(config.get("method") or "").strip().upper()
    if method not in {"GET", "POST"}:
        method = "POST" if "getTeamAndCarUpdates" in endpoint else GPS_API_DEFAULT_METHOD
    parts = urllib.parse.urlsplit(endpoint)
    query = dict(urllib.parse.parse_qsl(parts.query, keep_blank_values=True))
    if method == "GET":
        query.setdefault("sortField", config.get("sortField", "plate"))
    query["_t"] = str(int(time.time() * 1000))
    url = urllib.parse.urlunsplit((parts.scheme, parts.netloc, parts.path, urllib.parse.urlencode(query), parts.fragment))
    headers = {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json;charset=UTF-8",
        "User-Agent": "BoyidaAccountingTool/1.0",
        "Origin": "https://www.zjhzwt.com",
        "Referer": "https://www.zjhzwt.com/gps-web/main.html",
    }
    headers.update(config.get("headers") or {})
    if config.get("authorization"):
        headers["Authorization"] = config["authorization"]
    if config.get("cookie"):
        headers["Cookie"] = config["cookie"]

    data = None
    if method == "POST":
        body = str(config.get("body") or "").strip()
        data = (body if body else "{}").encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        timeout = int(config.get("timeout", 20))
    except Exception:
        timeout = 20
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        payload = json.loads(resp.read().decode("utf-8"))
    result = decode_gps_api_payload(payload)
    summary = gps_state_summary(result, config.get("company") or DEFAULT_GPS_SUMMARY["company"])
    save_gps_summary(summary)
    return summary


def safe_filename_part(value: object, fallback: str = "空") -> str:
    text = str(value or "").strip()
    text = text.replace("/", "-").replace("\\", "-").replace(":", "-")
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", text)
    text = re.sub(r"\s+", "", text).strip("._ ")
    return (text[:36] or fallback)


def external_image_folder(output: Path) -> Path:
    return output.with_name(f"{output.stem}_原图")


def external_image_filename(row_index: int, row: dict, image_path: Path) -> str:
    suffix = image_path.suffix.lower() if image_path.suffix else ".jpg"
    pieces = [
        f"{row_index:03d}",
        safe_filename_part(row.get("日期"), "未知日期"),
        safe_filename_part(row.get("车号"), "未知车号"),
        safe_filename_part(row.get("驾驶员"), "未知驾驶员"),
        safe_filename_part(image_path.stem or row.get("_file_name"), "原图"),
    ]
    return "_".join(pieces) + suffix


def unique_filename(name: str, used: set[str]) -> str:
    candidate = name
    stem = Path(name).stem
    suffix = Path(name).suffix
    counter = 2
    while candidate.lower() in used:
        candidate = f"{stem}_{counter}{suffix}"
        counter += 1
    used.add(candidate.lower())
    return candidate


def prepare_external_image(row_index: int, row: dict, folder: Path, used_names: set[str]) -> tuple[Path | None, str, str]:
    source = Path(row.get("_file_path", ""))
    if not source.exists():
        return None, "原图缺失", ""
    folder.mkdir(parents=True, exist_ok=True)
    filename = unique_filename(external_image_filename(row_index, row, source), used_names)
    target = folder / filename
    try:
        if source.resolve() != target.resolve():
            shutil.copy2(source, target)
    except FileNotFoundError:
        shutil.copy2(source, target)
    relative = os.path.relpath(target, start=folder.parent).replace(os.sep, "/")
    return target, filename, relative


def parse_raw_ocr(raw_json: Path, output_dir: Path) -> tuple[Path, Path, Path, Path, int, int, int]:
    raw = json.loads(raw_json.read_text(encoding="utf-8"))
    rows = [parse_record(record) for record in raw.get("records", [])]
    rows, duplicate_rows = apply_duplicates(rows)
    rows.sort(key=lambda r: (sortable_date(r["日期"]), r["_file_name"]))

    parsed_json = output_dir / "transport_parsed_records.json"
    preview_csv = output_dir / "运输录入_结构化预览.csv"
    duplicates_csv = output_dir / "重复图片.csv"
    questions_csv = output_dir / "识别疑问_已回填主表.csv"

    questions = [{"原图": row["_file_name"], "疑问": row["备注"], "OCR文本": row["_ocr_text"]} for row in rows if row["备注"]]
    parsed_json.write_text(json.dumps({"records": rows, "duplicates": duplicate_rows}, ensure_ascii=False, indent=2), encoding="utf-8")
    write_csv(preview_csv, rows, HEADERS)
    write_csv(duplicates_csv, duplicate_rows, ["重复图片", "对应原图"])
    write_csv(questions_csv, questions, ["原图", "疑问", "OCR文本"])
    return parsed_json, preview_csv, duplicates_csv, questions_csv, len(rows), len(duplicate_rows), len(questions)


def build_excel(
    parsed_json: Path,
    output_path: Path | None = None,
    columns: list[dict] | None = None,
    image_mode: str | None = None,
) -> tuple[Path, int, int, Path | None]:
    data = json.loads(parsed_json.read_text(encoding="utf-8"))
    rows = data.get("records", [])
    duplicates = data.get("duplicates", [])
    excel_columns = [item for item in normalize_excel_columns(columns or default_excel_columns()) if item["enabled"]]
    if not excel_columns:
        excel_columns = default_excel_columns()
    image_mode = normalize_excel_image_mode(image_mode)
    output = output_path or parsed_json.parent / f"运输录入_{date_range(rows)}.xlsx"
    image_folder = external_image_folder(output) if image_mode == "external" else None

    workbook = xlsxwriter.Workbook(str(output))
    main_ws = workbook.add_worksheet("运输录入")
    dup_ws = workbook.add_worksheet("重复图片")

    header_fmt = workbook.add_format({"font_name": "宋体", "font_size": 10, "bold": True, "align": "center", "valign": "vcenter", "text_wrap": True, "bg_color": "#D9EAF7", "border": 1})
    cell_fmt = workbook.add_format({"font_name": "宋体", "font_size": 10, "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1})
    weight_fmt = workbook.add_format({"font_name": "宋体", "font_size": 10, "align": "center", "valign": "vcenter", "num_format": "0.00", "border": 1})
    link_fmt = workbook.add_format({"font_name": "宋体", "font_size": 10, "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1, "font_color": "#0563C1", "underline": 1})

    for col, item in enumerate(excel_columns):
        main_ws.set_column(col, col, item["width"])
    main_ws.freeze_panes(1, 0)
    main_ws.autofilter(0, 0, len(rows), len(excel_columns) - 1)
    main_ws.set_row(0, 24)
    for col, item in enumerate(excel_columns):
        main_ws.write(0, col, item["title"], header_fmt)

    if image_mode == "embedded" and not hasattr(main_ws, "embed_image"):
        workbook.close()
        raise RuntimeError("当前 XlsxWriter 不支持 worksheet.embed_image()，无法生成 Excel 单元格图片。")

    used_image_names: set[str] = set()
    image_index_rows: list[dict] = []
    for row_idx, row in enumerate(rows, 1):
        main_ws.set_row(row_idx, 86 if image_mode == "embedded" else 24)
        for col, item in enumerate(excel_columns):
            source = item["source"]
            if source == "原图":
                main_ws.write(row_idx, col, "", cell_fmt)
                image_path = Path(row.get("_file_path", ""))
                if image_mode == "embedded" and image_path.exists():
                    main_ws.embed_image(row_idx, col, str(image_path), {"description": row.get("_file_name", "")})
                    main_ws.write_comment(row_idx, col, row.get("_file_name", ""))
                elif image_mode == "external" and image_folder:
                    target, display_name, relative = prepare_external_image(row_idx, row, image_folder, used_image_names)
                    if target:
                        status = main_ws.write_url(row_idx, col, f"external:{relative}", link_fmt, display_name)
                        if status:
                            main_ws.write(row_idx, col, display_name, link_fmt)
                        main_ws.write_comment(row_idx, col, f"原始文件：{row.get('_file_name', '')}\n外部原图：{display_name}")
                        image_index_rows.append({
                            "生成时Excel行号": row_idx + 1,
                            "原图链接文件名": display_name,
                            "原图相对路径": relative,
                            "原始文件名": row.get("_file_name", ""),
                            "日期": row.get("日期", ""),
                            "车号": row.get("车号", ""),
                            "驾驶员": row.get("驾驶员", ""),
                            "重量": row.get("重量", ""),
                        })
                    else:
                        main_ws.write(row_idx, col, display_name, cell_fmt)
                else:
                    main_ws.write(row_idx, col, "原图缺失", cell_fmt)
                continue
            value = row.get(source, "") if source else ""
            if source == "重量" and value:
                try:
                    main_ws.write_number(row_idx, col, float(value), weight_fmt)
                except ValueError:
                    main_ws.write(row_idx, col, value, cell_fmt)
            else:
                main_ws.write(row_idx, col, value, cell_fmt)

    dup_ws.set_column(0, 1, 35)
    dup_ws.set_row(0, 24)
    for col, header in enumerate(["重复图片", "对应原图"]):
        dup_ws.write(0, col, header, header_fmt)
    for row_idx, row in enumerate(duplicates, 1):
        dup_ws.write(row_idx, 0, row.get("重复图片", ""), cell_fmt)
        dup_ws.write(row_idx, 1, row.get("对应原图", ""), cell_fmt)
    dup_ws.autofilter(0, 0, max(len(duplicates), 1), 1)
    dup_ws.freeze_panes(1, 0)
    workbook.close()
    if image_folder and image_index_rows:
        write_csv(
            image_folder / "原图索引.csv",
            image_index_rows,
            ["生成时Excel行号", "原图链接文件名", "原图相对路径", "原始文件名", "日期", "车号", "驾驶员", "重量"],
        )
    return output, len(rows), len(duplicates), image_folder


def find_python_candidates() -> list[str]:
    candidates: list[str] = []
    seen: set[str] = set()
    possible: list[str] = []
    env_python = os.environ.get("ZHIPU_OCR_PYTHON")
    if env_python:
        possible.append(env_python)
    if not getattr(sys, "frozen", False):
        possible.append(sys.executable)
    possible.extend([
        str(Path.home() / ".cache" / "codex-runtimes" / "codex-primary-runtime" / "dependencies" / "python" / "python.exe"),
        r"C:\Python312\python.exe",
        r"C:\Python311\python.exe",
        r"C:\Python310\python.exe",
        r"C:\Python314\python.exe",
        "python",
    ])

    code, output = run_capture(["py", "-0p"], timeout=8)
    if code == 0:
        for line in output.splitlines():
            match = re.search(r"([A-Za-z]:\\.*?python\.exe)", line)
            if match:
                possible.append(match.group(1).strip())

    code, output = run_capture(["where.exe", "python"], timeout=8)
    if code == 0:
        possible.extend(line.strip() for line in output.splitlines() if line.strip())

    home = Path.home()
    common_patterns = [
        home / "AppData" / "Local" / "Programs" / "Python" / "Python*" / "python.exe",
        home / "miniconda3" / "python.exe",
        home / "anaconda3" / "python.exe",
        home / "miniconda3" / "envs" / "*" / "python.exe",
        home / "anaconda3" / "envs" / "*" / "python.exe",
        Path(r"C:\ProgramData\miniconda3\python.exe"),
        Path(r"C:\ProgramData\anaconda3\python.exe"),
        Path(r"C:\ProgramData\miniconda3\envs\*\python.exe"),
        Path(r"C:\ProgramData\anaconda3\envs\*\python.exe"),
        Path(r"C:\Python*\python.exe"),
        Path.cwd() / ".venv" / "Scripts" / "python.exe",
        Path.cwd() / "venv" / "Scripts" / "python.exe",
    ]
    for pattern in common_patterns:
        pattern_text = str(pattern)
        if "*" in pattern_text:
            possible.extend(str(path) for path in Path(pattern.anchor or ".").glob(pattern_text[len(pattern.anchor):]))
        elif pattern.exists():
            possible.append(pattern_text)

    code, output = run_capture(["conda", "env", "list", "--json"], timeout=8)
    if code == 0:
        try:
            for env_path in json.loads(output).get("envs", []):
                possible.append(str(Path(env_path) / "python.exe"))
        except Exception:
            pass

    for candidate in possible:
        resolved = shutil.which(candidate) if candidate == "python" else candidate
        if not resolved:
            continue
        if "WindowsApps" in resolved:
            continue
        if candidate != "python" and not Path(candidate).exists():
            continue
        normalized = str(Path(resolved))
        key = normalized.lower()
        if key not in seen:
            seen.add(key)
            candidates.append(normalized)
    if not candidates:
        candidates.append("python")
    return candidates


def paddle_ready(state: dict | None) -> bool:
    return bool(
        state
        and state.get("paddleocr")
        and state.get("paddle")
        and state.get("cuda_compiled")
        and state.get("run_check")
    )


def paddle_basic_ready(state: dict | None) -> bool:
    return bool(state and state.get("paddleocr") and state.get("paddle"))


def local_ocr_attempt_plan(selected_profile: str, paddle_state: dict | None) -> list[tuple[str, str]]:
    selected = selected_profile if selected_profile in OCR_PROFILE_LABELS else "stable"
    device = "gpu:0" if paddle_ready(paddle_state) else "cpu"
    raw_plan: list[tuple[str, str]]
    if device == "cpu":
        raw_plan = [("stable", "cpu")]
    elif selected == "trt":
        raw_plan = [("trt", "gpu:0"), ("fast", "gpu:0"), ("stable", "gpu:0"), ("stable", "cpu")]
    elif selected == "fast":
        raw_plan = [("fast", "gpu:0"), ("stable", "gpu:0"), ("stable", "cpu")]
    else:
        raw_plan = [("stable", "gpu:0"), ("stable", "cpu")]

    plan: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for item in raw_plan:
        if item not in seen:
            seen.add(item)
            plan.append(item)
    return plan


def local_ocr_attempt_title(profile: str, device: str) -> str:
    names = {
        "stable": "稳妥档",
        "fast": "极速档",
        "trt": "极限档",
    }
    target = "CPU" if device == "cpu" else "GPU"
    return f"{names.get(profile, profile)} / {target}"


def find_best_paddle_environment(preferred_python: str = "", log_fn=None) -> tuple[dict | None, list[dict]]:
    candidates: list[str] = []
    if preferred_python:
        candidates.append(preferred_python)
    candidates.extend(find_python_candidates())
    unique_candidates = []
    for candidate in candidates:
        if candidate and candidate not in unique_candidates:
            unique_candidates.append(candidate)

    checked: list[dict] = []
    best: dict | None = None
    for index, candidate in enumerate(unique_candidates[:12], 1):
        if log_fn:
            log_fn(f"检测 Python {index}/{min(len(unique_candidates), 12)}：{candidate}")
        state = check_paddle_environment(candidate)
        checked.append(state)
        if log_fn:
            log_fn(
                "结果："
                f"paddleocr={state.get('paddleocr')} "
                f"paddle={state.get('paddle')} "
                f"cuda={state.get('cuda_compiled')} "
                f"run_check={state.get('run_check')}"
            )
        if paddle_ready(state):
            return state, checked
        if best is None and paddle_basic_ready(state):
            best = state
    if best is None and checked:
        best = checked[0]
    return best, checked


def self_test() -> int:
    assert DRIVER_BY_PLATE["浙A07336D"] == "孙安足"
    assert "运输录入_2026.5.12-2026.5.31.xlsx".endswith(".xlsx")
    assert normalize_plate("车号 浙A07336D")[0] == "浙A07336D"
    assert parse_dates(["毛重时间 2026-05-31 10:00"], "")[0] == "2026/5/31"
    assert parse_weight({"rec_texts": ["<table><tr><td>净重</td><td>51320 kg</td></tr><tr><td>毛重时间</td><td>2026-5-26</td></tr></table>"]})[0] == "51.32"
    assert parse_weight({"rec_texts": ["<table><tr><td>毛重</td><td>53970</td></tr><tr><td>空重</td><td>40790</td></tr><tr><td>净重</td><td>13180</td></tr><tr><td>结算重量</td><td></td></tr><tr><td>空重时间</td><td>2026-05-23</td></tr></table>"]})[0] == "13.18"
    assert parse_cargo(["<table><tr><td>货名</td><td>朝渣</td></tr></table>"])[0] == "钢渣"
    assert parse_cargo(["<table><tr><td>货名</td><td>客户板</td></tr></table>"])[0] == "卷子板"
    assert auth_header("abc") == "Bearer abc"
    assert auth_header("Bearer abc") == "Bearer abc"
    assert initial_api_key()[1] in {"环境变量", "本机已保存", "EXE 内置默认", "未填写"}
    assert OCR_ENGINE_BY_LABEL[OCR_ENGINE_LABELS["glm"]] == "glm"
    assert OCR_PROFILE_BY_LABEL[OCR_PROFILE_LABELS["fast"]] == "fast"
    assert normalize_excel_columns([])[0]["title"] == "日期"
    assert normalize_excel_image_mode("") == "external"
    assert normalize_excel_image_mode("embedded") == "embedded"
    custom_columns = normalize_excel_columns([{"enabled": True, "title": "客户补充", "source": "", "width": 100}])
    assert custom_columns[0]["source"] == ""
    assert custom_columns[0]["width"] == 60
    small_metrics = adaptive_window_metrics(1024, 768)
    assert small_metrics["width"] <= 1000
    assert small_metrics["height"] <= 712
    assert small_metrics["compact"] is True
    desktop_metrics = adaptive_window_metrics(1920, 1080)
    assert desktop_metrics["width"] == 1220
    assert desktop_metrics["height"] == 690
    assert daily_workspace_name(date(2026, 6, 2)) == "2026-06-02_保谊达做账表"
    gpu_ready_state = {"paddleocr": True, "paddle": True, "cuda_compiled": True, "run_check": True}
    cpu_ready_state = {"paddleocr": True, "paddle": True, "cuda_compiled": False, "run_check": True}
    assert local_ocr_attempt_plan("fast", cpu_ready_state) == [("stable", "cpu")]
    assert local_ocr_attempt_plan("fast", gpu_ready_state)[:2] == [("fast", "gpu:0"), ("stable", "gpu:0")]
    assert local_ocr_attempt_plan("trt", gpu_ready_state)[:3] == [("trt", "gpu:0"), ("fast", "gpu:0"), ("stable", "gpu:0")]
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        image = tmp_dir / "a.jpg"
        image.write_bytes(b"image")
        raw_json = tmp_dir / "paddle_ocr_raw.json"
        raw_json.write_text(json.dumps({"records": [{"file_name": "a.jpg", "error": "fail", "rec_texts": []}]}, ensure_ascii=False), encoding="utf-8")
        failed_status = ocr_json_status(raw_json, [image])
        assert not failed_status["complete"]
        assert failed_status["invalid"] == ["a.jpg"]
        raw_json.write_text(json.dumps({"records": [{"file_name": "a.jpg", "rec_texts": ["车号 浙A07336D"]}]}, ensure_ascii=False), encoding="utf-8")
        ok_status = ocr_json_status(raw_json, [image])
        assert ok_status["complete"]
        raw_json.write_text(json.dumps({"records": [{"file_name": "a.jpg", "rec_texts": ["车号 浙A07336D"]}, {"file_name": "old.jpg", "rec_texts": ["旧图"]}]}, ensure_ascii=False), encoding="utf-8")
        mixed_status = ocr_json_status(raw_json, [image])
        assert not mixed_status["complete"]
        assert mixed_status["extra"] == ["old.jpg"]
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        image_dir = tmp_dir / "待处理图片"
        image_dir.mkdir()
        (image_dir / "fail.jpg").write_bytes(b"image")
        raw_json = tmp_dir / "输出结果" / "paddle_ocr_raw.json"
        raw_json.parent.mkdir()
        original_call_glm = globals()["call_glm_ocr"]
        try:
            globals()["call_glm_ocr"] = lambda _image, _key: {"rec_texts": [], "raw": {}}
            try:
                run_glm_ocr_batch(image_dir, raw_json, "test-key", lambda _msg: None)
                raise AssertionError("empty OCR text must fail")
            except RuntimeError as exc:
                assert "OCR 没跑通" in str(exc)
        finally:
            globals()["call_glm_ocr"] = original_call_glm
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        image = tmp_dir / "单据.jpg"
        image.write_bytes(b"not-an-embedded-image")
        parsed = tmp_dir / "transport_parsed_records.json"
        parsed.write_text(
            json.dumps({
                "records": [{
                    "日期": "2026/5/31",
                    "车号": "浙A07336D",
                    "驾驶员": "孙安足",
                    "重量": "12.34",
                    "货物名称": "卷子板",
                    "装货地": "",
                    "收货地": "",
                    "备注": "",
                    "原图": image.name,
                    "_file_name": image.name,
                    "_file_path": str(image),
                }],
                "duplicates": [],
            }, ensure_ascii=False),
            encoding="utf-8",
        )
        workbook, row_count, dup_count, image_folder = build_excel(parsed, image_mode="external")
        assert workbook.exists()
        assert row_count == 1 and dup_count == 0
        assert image_folder and (image_folder / "原图索引.csv").exists()
        index_header = (image_folder / "原图索引.csv").read_text(encoding="utf-8-sig").splitlines()[0]
        assert "原图相对路径" in index_header and "原始路径" not in index_header
        with zipfile.ZipFile(workbook) as archive:
            rel_names = [name for name in archive.namelist() if name.startswith("xl/worksheets/_rels/sheet1.xml.rels")]
            assert rel_names
            rel_text = archive.read(rel_names[0]).decode("utf-8")
            assert "_原图" in rel_text and "C:" not in rel_text and "file:///" not in rel_text
    if os.name == "nt":
        assert dpapi_unprotect(dpapi_protect("test-key")) == "test-key"
    hardware = {"gpus": [{"name": "NVIDIA GeForce RTX 5090", "memory_mb": 32607}]}
    paddle = {"paddleocr": True, "paddle": True, "cuda_compiled": True, "run_check": True}
    assert build_ocr_recommendation(hardware, paddle, "")[0].startswith("推荐用本地 OCR")
    install_label, install_cmd = local_paddle_install_command(sys.executable, {"gpus": [{"name": "NVIDIA GeForce RTX 5090", "memory_mb": 32607, "driver": "560.00"}]})
    if python_minor_version(sys.executable) in PADDLE_50_SERIES_WHEELS:
        assert "50 系" in install_label and any("paddlepaddle_gpu" in part for part in install_cmd)
    cpu_label, cpu_cmd = local_paddle_install_command(sys.executable, {"gpus": []})
    assert "CPU" in cpu_label and "paddlepaddle==3.2.0" in cpu_cmd
    virtual_label, virtual_cmd = local_paddle_install_command(sys.executable, {"gpus": [{"name": "Microsoft Basic Render Driver", "memory_mb": 0, "driver": ""}]})
    assert "CPU" in virtual_label and "paddlepaddle==3.2.0" in virtual_cmd
    gps_sample = [[], [{"id": "1", "plate": "浙A1"}, {"id": "2", "plate": "浙A2"}, {"id": "3", "plate": "浙A3"}], {"1": 7, "2": 3, "3": 11}]
    gps_summary = gps_state_summary(gps_sample, "测试车队")
    assert gps_summary["total"] == "3"
    assert gps_summary["online"] == "2"
    assert gps_summary["offline"] == "1"
    assert gps_summary["running"] == "1"
    assert gps_summary["stopped"] == "1"
    assert gps_summary["alarm"] == "1"
    return 0


class ExcelTemplateDialog:
    MODE_LABELS = {
        "external": "外部原图链接（轻量，推荐）",
        "embedded": "格内嵌图（文件大）",
    }

    def __init__(self, root: Tk, columns: list[dict], image_mode: str) -> None:
        self.root = root
        self.columns = normalize_excel_columns(columns)
        self.image_mode = StringVar(value=self.MODE_LABELS[normalize_excel_image_mode(image_mode)])
        self.result: dict | None = None
        self.window = Toplevel(root)
        self.window.title("生成表格前确认")
        self.window.geometry("840x570")
        self.window.minsize(800, 520)
        self.window.configure(bg=THEME["bg"])
        self.window.transient(root)
        self.window.protocol("WM_DELETE_WINDOW", self.cancel)
        self._build()
        self.refresh()
        self.window.grab_set()

    def _build(self) -> None:
        outer = ttk.Frame(self.window, padding=14, style="App.TFrame")
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="先确认 Excel 表格列", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="默认是规范九列 + 外部原图链接；特殊批次需要多一列、改列名或调顺序时，在这里调整。",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(6, 12))

        mode_row = ttk.Frame(outer, style="App.TFrame")
        mode_row.pack(fill="x", pady=(0, 10))
        ttk.Label(mode_row, text="原图方式", style="Muted.TLabel").pack(side="left")
        mode_combo = ttk.Combobox(
            mode_row,
            textvariable=self.image_mode,
            values=[self.MODE_LABELS["external"], self.MODE_LABELS["embedded"]],
            state="readonly",
            width=28,
        )
        mode_combo.pack(side="left", padx=(10, 14))
        ttk.Label(mode_row, text="轻量链接会把原图放在 Excel 旁边的文件夹里，表格不卡。", style="Muted.TLabel").pack(side="left")

        table_frame = ttk.Frame(outer, style="App.TFrame")
        table_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(table_frame, columns=("enabled", "title", "width"), show="headings", selectmode="browse")
        for key, text, width in (
            ("enabled", "使用", 70),
            ("title", "输出列名", 300),
            ("width", "列宽", 70),
        ):
            self.tree.heading(key, text=text)
            self.tree.column(key, width=width, anchor="center")
        scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<Button-3>", self.open_column_menu)
        self.tree.bind("<Button-2>", self.open_column_menu)

        edit_row = ttk.Frame(outer, style="App.TFrame")
        edit_row.pack(fill="x", pady=(10, 6))
        edit_buttons = [
            ("启用/停用", self.toggle_enabled),
            ("上移", lambda: self.move_selected(-1)),
            ("下移", lambda: self.move_selected(1)),
            ("改列名", self.rename_selected),
            ("改宽度", self.change_width),
            ("恢复默认", self.reset_default),
        ]
        for index, (text, command) in enumerate(edit_buttons):
            edit_row.columnconfigure(index, weight=1, uniform="excel_edit")
            ttk.Button(edit_row, text=text, command=command, style="Tool.TButton").grid(row=0, column=index, sticky="ew", padx=3)

        bottom = ttk.Frame(outer, style="App.TFrame")
        bottom.pack(fill="x", pady=(8, 0))
        ttk.Button(bottom, text="默认生成", command=self.default_generate, style="Primary.TButton").pack(side="left", ipadx=12)
        ttk.Button(bottom, text="保存并生成", command=self.save_generate, style="Tool.TButton").pack(side="left", padx=8, ipadx=12)
        ttk.Button(bottom, text="取消", command=self.cancel, style="Tool.TButton").pack(side="right", ipadx=12)

    def show(self) -> dict | None:
        self.window.wait_window()
        return self.result

    def selected_image_mode(self) -> str:
        label = self.image_mode.get()
        for mode, text in self.MODE_LABELS.items():
            if text == label:
                return mode
        return DEFAULT_EXCEL_IMAGE_MODE

    def refresh(self, select_index: int | None = None) -> None:
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for index, item in enumerate(self.columns):
            self.tree.insert(
                "",
                END,
                iid=str(index),
                values=("是" if item["enabled"] else "否", item["title"], item["width"]),
            )
        if self.columns:
            if select_index is None:
                select_index = 0
            select_index = max(0, min(select_index, len(self.columns) - 1))
            self.tree.selection_set(str(select_index))
            self.tree.focus(str(select_index))

    def open_column_menu(self, event) -> str:
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self.tree.selection_set(row_id)
            self.tree.focus(row_id)
            row_index = int(row_id)
        else:
            row_index = None

        menu = Menu(
            self.window,
            tearoff=0,
            bg=THEME["panel"],
            fg=THEME["text"],
            activebackground=THEME["panel_2"],
            activeforeground=THEME["cyan"],
            bd=1,
        )
        if row_index is None:
            menu.add_command(label="添加空白列", command=self.add_blank_column)
            menu.add_command(label="恢复默认", command=self.reset_default)
        else:
            menu.add_command(label="启用/停用", command=self.toggle_enabled)
            menu.add_command(label="改列名", command=self.rename_selected)
            menu.add_command(label="改宽度", command=self.change_width)
            menu.add_separator()
            menu.add_command(label="上移", command=lambda: self.move_selected(-1))
            menu.add_command(label="下移", command=lambda: self.move_selected(1))
            menu.add_separator()
            menu.add_command(label="在下方插入空白列", command=lambda: self.add_blank_column(row_index + 1))
            menu.add_command(label="删除这一列", command=self.delete_selected)
            menu.add_separator()
            menu.add_command(label="高级：字段映射", command=self.change_source)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return "break"

    def selected_index(self) -> int | None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning(APP_TITLE, "先选中一列。", parent=self.window)
            return None
        return int(selection[0])

    def validate_columns(self, columns: list[dict]) -> bool:
        enabled = [item for item in normalize_excel_columns(columns) if item["enabled"]]
        if not enabled:
            messagebox.showwarning(APP_TITLE, "至少保留一列。", parent=self.window)
            return False
        if not any(item["source"] == "原图" for item in enabled):
            messagebox.showwarning(APP_TITLE, "必须保留“原图”列，这样最终 Excel 才能追溯原图。", parent=self.window)
            return False
        return True

    def toggle_enabled(self) -> None:
        index = self.selected_index()
        if index is None:
            return
        self.columns[index]["enabled"] = not self.columns[index]["enabled"]
        self.refresh(index)

    def move_selected(self, direction: int) -> None:
        index = self.selected_index()
        if index is None:
            return
        target = index + direction
        if target < 0 or target >= len(self.columns):
            return
        self.columns[index], self.columns[target] = self.columns[target], self.columns[index]
        self.refresh(target)

    def rename_selected(self) -> None:
        index = self.selected_index()
        if index is None:
            return
        value = simpledialog.askstring("改列名", "Excel 里显示的列名：", initialvalue=self.columns[index]["title"], parent=self.window)
        if value is None:
            return
        value = value.strip()
        if not value:
            messagebox.showwarning(APP_TITLE, "列名不能为空。", parent=self.window)
            return
        title = value[:40]
        self.columns[index]["title"] = title
        matched_source = source_from_column_title(title)
        if matched_source:
            self.columns[index]["source"] = matched_source
        self.refresh(index)

    def change_source(self) -> None:
        index = self.selected_index()
        if index is None:
            return
        prompt = "内部字段可填：日期、车号、驾驶员、重量、货物名称、装货地、收货地、备注、原图；留空就是手填空白列。"
        value = simpledialog.askstring("字段映射", prompt, initialvalue=self.columns[index]["source"], parent=self.window)
        if value is None:
            return
        self.columns[index]["source"] = value.strip()[:40]
        self.refresh(index)

    def change_width(self) -> None:
        index = self.selected_index()
        if index is None:
            return
        value = simpledialog.askinteger("改宽度", "列宽，6 到 60：", initialvalue=self.columns[index]["width"], minvalue=6, maxvalue=60, parent=self.window)
        if value is None:
            return
        self.columns[index]["width"] = value
        self.refresh(index)

    def add_blank_column(self, insert_index: int | None = None) -> None:
        value = simpledialog.askstring("加空白列", "新列名：", initialvalue="手填列", parent=self.window)
        if value is None:
            return
        title = value.strip() or "手填列"
        title = title[:40]
        new_column = {"enabled": True, "title": title, "source": source_from_column_title(title), "width": 12}
        if insert_index is None:
            self.columns.append(new_column)
            insert_index = len(self.columns) - 1
        else:
            insert_index = max(0, min(insert_index, len(self.columns)))
            self.columns.insert(insert_index, new_column)
        self.refresh(insert_index)

    def delete_selected(self) -> None:
        index = self.selected_index()
        if index is None:
            return
        removed = self.columns[index]
        remaining = self.columns[:index] + self.columns[index + 1 :]
        if removed["source"] == "原图" and not any(item.get("source") == "原图" and item.get("enabled", True) for item in remaining):
            messagebox.showwarning(APP_TITLE, "原图列必须保留。", parent=self.window)
            return
        self.columns = remaining or default_excel_columns()
        self.refresh(min(index, len(self.columns) - 1))

    def reset_default(self) -> None:
        self.columns = default_excel_columns()
        self.image_mode.set(self.MODE_LABELS[DEFAULT_EXCEL_IMAGE_MODE])
        self.refresh()

    def default_generate(self) -> None:
        self.result = {"action": "generate", "columns": default_excel_columns(), "image_mode": DEFAULT_EXCEL_IMAGE_MODE}
        self.window.destroy()

    def build_saved_template_result(self, action: str) -> dict | None:
        normalized = normalize_excel_columns(self.columns)
        if not self.validate_columns(normalized):
            return None
        try:
            result = save_excel_template(normalized, self.selected_image_mode())
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"保存列模板失败：{exc}", parent=self.window)
            return None
        result["action"] = action
        return result

    def save_generate(self) -> None:
        result = self.build_saved_template_result("generate")
        if result is None:
            return
        self.result = result
        self.window.destroy()

    def cancel(self) -> None:
        self.result = None
        self.window.destroy()


class AccountingApp:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title(APP_TITLE)
        self.root.configure(bg=THEME["bg"])
        self._layout_compact = False
        self._window_width = 1220
        self._window_height = 780
        self._normal_geometry = ""
        self._is_maximized = False
        self._resize_mode = ""
        self._resize_start = (0, 0, 0, 0, 0, 0)
        self._configure_initial_window()
        self.root.overrideredirect(True)
        self.logo_image: PhotoImage | None = None
        self.logo_thumb: PhotoImage | None = None
        self.jingzhe_header_image: PhotoImage | None = None
        self.events: queue.Queue[str] = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self._log_started = False
        self._drag_offset = (0, 0)
        self._progress_started_at = 0.0
        self._last_progress_percent = 0.0
        self._progress_log_started = False

        input_dir = managed_input_dir()
        output_dir = managed_output_dir()
        py_candidates = find_python_candidates()
        key_value, key_source = initial_api_key()
        self.image_dir = StringVar(value=str(input_dir))
        self.work_dir = StringVar(value=str(output_dir))
        self.python_path = StringVar(value=py_candidates[0])
        self.ocr_engine = StringVar(value=OCR_ENGINE_LABELS["glm"])
        self.api_key = StringVar(value=key_value)
        self.api_key_source = StringVar(value=f"Key 来源：{key_source}")
        self.remember_api_key = BooleanVar(value=True)
        self.ocr_profile = StringVar(value=OCR_PROFILE_LABELS["fast"])
        self.ocr_profile_detail = StringVar()
        self.skip_ocr_if_json = BooleanVar(value=True)
        self.raw_json = StringVar(value=str(output_dir / "paddle_ocr_raw.json"))
        self.input_status = StringVar()
        self.output_status = StringVar()
        self.progress_value = DoubleVar(value=0.0)
        self.progress_title = StringVar(value="等待执行")
        self.progress_detail = StringVar(value="进度 0% / 剩余 --")
        self.progress_percent = StringVar(value="0%")
        self.gps_summary = load_gps_summary()
        excel_template = load_excel_template()
        self.excel_columns = excel_template["columns"]
        self.excel_image_mode = excel_template["image_mode"]
        self._refresh_workspace_status()
        self._refresh_ocr_profile_detail()

        self._configure_theme()
        self._set_window_icon()
        self._build_ui()
        self._install_resize_handles()
        self.root.bind("<Map>", self._restore_borderless)
        self.root.after(150, self._drain_events)
        self.root.after(600, self.refresh_gps_on_startup)

    def _configure_initial_window(self) -> None:
        metrics = adaptive_window_metrics(self.root.winfo_screenwidth(), self.root.winfo_screenheight())
        self._layout_compact = bool(metrics["compact"])
        self._window_width = int(metrics["width"])
        self._window_height = int(metrics["height"])
        self.root.geometry(
            f'{metrics["width"]}x{metrics["height"]}+{metrics["x"]}+{metrics["y"]}'
        )
        self.root.minsize(int(metrics["min_width"]), int(metrics["min_height"]))
        self._normal_geometry = self.root.geometry()

    def _set_window_icon(self) -> None:
        ico_path = resource_path(LOGO_ICO)
        png_path = resource_path(LOGO_PNG)
        header_path = resource_path(JINGZHE_HEADER_PNG)
        try:
            if ico_path.exists():
                self.root.iconbitmap(str(ico_path))
        except Exception:
            pass
        try:
            if png_path.exists():
                self.logo_image = PhotoImage(file=str(png_path))
                self.root.iconphoto(True, self.logo_image)
        except Exception:
            pass
        try:
            if header_path.exists():
                self.jingzhe_header_image = PhotoImage(file=str(header_path))
        except Exception:
            self.jingzhe_header_image = None

    def selected_ocr_engine(self) -> str:
        value = self.ocr_engine.get().strip()
        if value in OCR_ENGINE_LABELS:
            return value
        return OCR_ENGINE_BY_LABEL.get(value, "glm")

    def set_ocr_engine(self, engine: str) -> None:
        self.ocr_engine.set(OCR_ENGINE_LABELS.get(engine, OCR_ENGINE_LABELS["glm"]))

    def selected_ocr_profile(self) -> str:
        value = self.ocr_profile.get().strip()
        if value in OCR_PROFILE_LABELS:
            return value
        return OCR_PROFILE_BY_LABEL.get(value, "fast")

    def _refresh_ocr_profile_detail(self, *_args) -> None:
        self.ocr_profile_detail.set(OCR_PROFILE_DETAILS.get(self.selected_ocr_profile(), OCR_PROFILE_DETAILS["fast"]))

    def _start_window_drag(self, event) -> None:
        if self._is_maximized:
            self._toggle_maximize_window()
        self._drag_offset = (event.x_root - self.root.winfo_x(), event.y_root - self.root.winfo_y())

    def _drag_window(self, event) -> None:
        if self._is_maximized:
            return
        x = event.x_root - self._drag_offset[0]
        y = event.y_root - self._drag_offset[1]
        self.root.geometry(f"+{x}+{y}")

    def _start_resize(self, mode: str, event) -> None:
        if self._is_maximized:
            return
        self._resize_mode = mode
        self._resize_start = (
            event.x_root,
            event.y_root,
            self.root.winfo_x(),
            self.root.winfo_y(),
            self.root.winfo_width(),
            self.root.winfo_height(),
        )

    def _perform_resize(self, event) -> None:
        if not self._resize_mode or self._is_maximized:
            return
        start_x, start_y, root_x, root_y, root_w, root_h = self._resize_start
        dx = event.x_root - start_x
        dy = event.y_root - start_y
        min_w, min_h = self.root.minsize()
        new_x, new_y = root_x, root_y
        new_w, new_h = root_w, root_h

        if "e" in self._resize_mode:
            new_w = max(min_w, root_w + dx)
        if "s" in self._resize_mode:
            new_h = max(min_h, root_h + dy)
        if "w" in self._resize_mode:
            new_w = max(min_w, root_w - dx)
            new_x = root_x + (root_w - new_w)
        if "n" in self._resize_mode:
            new_h = max(min_h, root_h - dy)
            new_y = root_y + (root_h - new_h)

        self.root.geometry(f"{int(new_w)}x{int(new_h)}+{int(new_x)}+{int(new_y)}")

    def _finish_resize(self, _event=None) -> None:
        self._resize_mode = ""
        if not self._is_maximized:
            self._normal_geometry = self.root.geometry()

    def _install_resize_handles(self) -> None:
        margin = 6
        cursor_by_mode = {
            "n": "sb_v_double_arrow",
            "s": "sb_v_double_arrow",
            "e": "sb_h_double_arrow",
            "w": "sb_h_double_arrow",
            "ne": "top_right_corner",
            "nw": "top_left_corner",
            "se": "bottom_right_corner",
            "sw": "bottom_left_corner",
        }
        placements = {
            "n": dict(x=margin, y=0, relwidth=1, width=-margin * 2, height=margin),
            "s": dict(x=margin, rely=1, y=-margin, relwidth=1, width=-margin * 2, height=margin),
            "e": dict(relx=1, x=-margin, y=margin, width=margin, relheight=1, height=-margin * 2),
            "w": dict(x=0, y=margin, width=margin, relheight=1, height=-margin * 2),
            "ne": dict(relx=1, x=-margin, y=0, width=margin, height=margin),
            "nw": dict(x=0, y=0, width=margin, height=margin),
            "se": dict(relx=1, rely=1, x=-margin, y=-margin, width=margin, height=margin),
            "sw": dict(rely=1, x=0, y=-margin, width=margin, height=margin),
        }
        for mode, place_args in placements.items():
            handle = ttk.Frame(self.root, style="App.TFrame", cursor=cursor_by_mode.get(mode, "sizing"))
            handle.place(**place_args)
            handle.bind("<ButtonPress-1>", lambda event, item=mode: self._start_resize(item, event))
            handle.bind("<B1-Motion>", self._perform_resize)
            handle.bind("<ButtonRelease-1>", self._finish_resize)

    def _minimize_window(self) -> None:
        self.root.overrideredirect(False)
        self.root.iconify()

    def _toggle_maximize_window(self) -> None:
        if self._is_maximized:
            if self._normal_geometry:
                self.root.geometry(self._normal_geometry)
            self._is_maximized = False
            return
        self._normal_geometry = self.root.geometry()
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        self.root.geometry(f"{screen_w}x{screen_h}+0+0")
        self._is_maximized = True

    def _restore_borderless(self, _event=None) -> None:
        if self.root.state() == "normal":
            self.root.overrideredirect(True)

    def _refresh_workspace_status(self) -> None:
        image_dir = Path(self.image_dir.get())
        work_dir = Path(self.work_dir.get())
        images = input_images(image_dir)
        self.input_status.set(f"待处理图片：{len(images)} 张")
        self.output_status.set(f"批次目录：{work_dir.parent.name}")

    def _configure_theme(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(".", font=("Microsoft YaHei UI", 10))
        style.configure("App.TFrame", background=THEME["bg"])
        style.configure("Panel.TFrame", background=THEME["panel"])
        style.configure("TFrame", background=THEME["bg"])
        style.configure("TLabel", background=THEME["bg"], foreground=THEME["text"])
        style.configure("Panel.TLabel", background=THEME["panel"], foreground=THEME["text"])
        style.configure("Muted.TLabel", background=THEME["bg"], foreground=THEME["muted"])
        style.configure("PanelMuted.TLabel", background=THEME["panel"], foreground=THEME["muted"])
        style.configure("BatchStatus.TLabel", background=THEME["panel"], foreground=THEME["text"], font=("Microsoft YaHei UI", 12, "bold"))
        style.configure("BatchPath.TLabel", background=THEME["panel"], foreground=THEME["cyan"], font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("Title.TLabel", background=THEME["bg"], foreground=THEME["text"], font=("Microsoft YaHei UI", 27, "bold"))
        style.configure("ProgressPercent.TLabel", background=THEME["panel"], foreground=THEME["green"], font=("Consolas", 20, "bold"))
        style.configure("Guide.TLabel", background=THEME["bg"], foreground=THEME["green"], font=("Microsoft YaHei UI", 12, "bold"))
        style.configure("Chip.TLabel", background=THEME["panel_3"], foreground=THEME["cyan"], padding=(10, 4), font=("Microsoft YaHei UI", 9, "bold"))
        style.configure("Titlebar.TFrame", background=THEME["titlebar"])
        style.configure("TitlebarAccent.TFrame", background=THEME["cyan_dark"])
        style.configure("Titlebar.TLabel", background=THEME["titlebar"], foreground=THEME["text"], font=("Microsoft YaHei UI", 10, "bold"))
        style.configure(
            "Chrome.TButton",
            background=THEME["titlebar"],
            foreground=THEME["muted"],
            bordercolor=THEME["titlebar"],
            padding=(10, 3),
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.map(
            "Chrome.TButton",
            background=[("active", THEME["panel_3"]), ("pressed", THEME["panel"])],
            foreground=[("active", THEME["cyan"])],
        )
        style.configure(
            "Close.TButton",
            background=THEME["titlebar"],
            foreground=THEME["muted"],
            bordercolor=THEME["titlebar"],
            padding=(10, 3),
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.map(
            "Close.TButton",
            background=[("active", "#7f1724"), ("pressed", "#420812")],
            foreground=[("active", "#ffffff")],
        )
        style.configure(
            "Panel.TLabelframe",
            background=THEME["panel"],
            foreground=THEME["text"],
            bordercolor=THEME["line"],
            relief="solid",
            padding=10,
        )
        style.configure(
            "Panel.TLabelframe.Label",
            background=THEME["bg"],
            foreground=THEME["cyan"],
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.configure(
            "TEntry",
            fieldbackground=THEME["input"],
            foreground=THEME["text"],
            bordercolor=THEME["line"],
            lightcolor=THEME["cyan_dark"],
            darkcolor=THEME["line"],
            insertcolor=THEME["cyan"],
            padding=5,
        )
        style.map(
            "TEntry",
            fieldbackground=[("focus", THEME["input"])],
            bordercolor=[("focus", THEME["cyan"])],
        )
        style.configure(
            "TCombobox",
            fieldbackground=THEME["input"],
            background=THEME["panel_2"],
            foreground=THEME["text"],
            arrowcolor=THEME["cyan"],
            bordercolor=THEME["line"],
            selectbackground=THEME["panel_2"],
            selectforeground=THEME["text"],
            padding=4,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", THEME["input"])],
            foreground=[("readonly", THEME["text"])],
            bordercolor=[("focus", THEME["cyan"])],
        )
        style.configure(
            "Tool.TButton",
            background=THEME["panel_2"],
            foreground=THEME["text"],
            bordercolor=THEME["line"],
            focusthickness=1,
            focuscolor=THEME["cyan_dark"],
            padding=(12, 7),
        )
        style.map(
            "Tool.TButton",
            background=[("active", THEME["panel_3"]), ("pressed", "#05101c")],
            foreground=[("active", THEME["cyan"])],
            bordercolor=[("focus", THEME["cyan"]), ("active", THEME["cyan_dark"])],
        )
        style.configure(
            "Batch.TButton",
            background=THEME["panel_2"],
            foreground=THEME["text"],
            bordercolor=THEME["line"],
            focusthickness=1,
            focuscolor=THEME["cyan_dark"],
            padding=(12, 10),
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.map(
            "Batch.TButton",
            background=[("active", THEME["panel_3"]), ("pressed", "#05101c")],
            foreground=[("active", THEME["cyan"])],
            bordercolor=[("focus", THEME["cyan"]), ("active", THEME["cyan_dark"])],
        )
        style.configure(
            "Primary.TButton",
            background="#007fa8",
            foreground="#ecfeff",
            bordercolor=THEME["cyan"],
            padding=(16, 7),
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#00a8d4"), ("pressed", "#045a78")],
            foreground=[("active", "#ffffff")],
        )
        style.configure(
            "BatchPrimary.TButton",
            background="#0089a8",
            foreground="#ecfeff",
            bordercolor=THEME["cyan"],
            padding=(16, 12),
            font=("Microsoft YaHei UI", 12, "bold"),
        )
        style.map(
            "BatchPrimary.TButton",
            background=[("active", "#00a8d4"), ("pressed", "#045a78")],
            foreground=[("active", "#ffffff")],
        )
        style.configure(
            "TCheckbutton",
            background=THEME["bg"],
            foreground=THEME["text"],
            indicatorcolor=THEME["input"],
            focuscolor=THEME["cyan_dark"],
            padding=3,
        )
        style.map(
            "TCheckbutton",
            background=[("active", THEME["bg"])],
            foreground=[("active", THEME["cyan"])],
            indicatorcolor=[("selected", THEME["cyan_dark"])],
        )
        style.configure(
            "Treeview",
            background=THEME["log"],
            fieldbackground=THEME["log"],
            foreground=THEME["text"],
            bordercolor=THEME["line"],
            rowheight=25,
        )
        style.configure(
            "Treeview.Heading",
            background=THEME["panel_2"],
            foreground=THEME["cyan"],
            relief="flat",
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.map("Treeview", background=[("selected", THEME["selected"])], foreground=[("selected", "#ffffff")])
        style.configure("Vertical.TScrollbar", background=THEME["panel_2"], troughcolor=THEME["bg"], arrowcolor=THEME["cyan"])
        style.configure(
            "Ops.Horizontal.TProgressbar",
            troughcolor=THEME["input"],
            background=THEME["green"],
            bordercolor=THEME["line"],
            lightcolor=THEME["green"],
            darkcolor=THEME["cyan_dark"],
            thickness=16,
        )

    def _draw_header_panel(self, canvas: Canvas) -> None:
        canvas.delete("all")
        w = max(canvas.winfo_width(), self._window_width - 24, 900)
        h = max(canvas.winfo_height(), 82)
        canvas.create_rectangle(0, 0, w, h, fill=THEME["bg"], outline="")
        canvas.create_rectangle(0, 0, w, h, fill=THEME["log"], outline=THEME["line_soft"])
        canvas.create_line(0, 0, 168, 0, fill=THEME["cyan_dark"], width=2)
        canvas.create_line(w - 220, h - 1, w, h - 1, fill=THEME["green"], width=2)
        for x in range(24, w, 48):
            canvas.create_line(x, 8, x, h - 8, fill=THEME["line_soft"])
        for y in (18, h - 20):
            canvas.create_line(12, y, w - 12, y, fill=THEME["line_soft"])

        if self.jingzhe_header_image is not None:
            image_y = max(0, (h - self.jingzhe_header_image.height()) // 2)
            canvas.create_image(18, image_y, image=self.jingzhe_header_image, anchor="nw")
        else:
            canvas.create_text(30, h // 2, text="重卡线稿", anchor="w", fill=THEME["muted"], font=("Microsoft YaHei UI", 10, "bold"))

        title = f"{COMPANY_NAME} 运输单据做账工具"
        canvas.create_text(w // 2 + 2, h // 2 + 3, text=title, fill="#00111f", font=("Microsoft YaHei UI", 27, "bold"))
        canvas.create_text(w // 2, h // 2 + 1, text=title, fill=THEME["text"], font=("Microsoft YaHei UI", 27, "bold"))

    def _draw_route_panel(self, canvas: Canvas) -> None:
        canvas.delete("all")
        summary = self.gps_summary
        w = max(canvas.winfo_width(), canvas.winfo_reqwidth(), 260)
        h = max(canvas.winfo_height(), canvas.winfo_reqheight(), 118)
        canvas.create_rectangle(0, 0, w, h, fill=THEME["log"], outline=THEME["line"])
        canvas.create_rectangle(2, 2, w - 3, h - 3, outline=THEME["cyan_soft"])
        canvas.create_line(0, 0, 68, 0, fill=THEME["cyan"], width=2)
        canvas.create_line(w - 92, h - 1, w - 1, h - 1, fill=THEME["green"], width=2)
        for x in range(18, w, 32):
            canvas.create_line(x, 8, x, h - 8, fill=THEME["line_soft"])
        for y in range(20, h, 22):
            canvas.create_line(8, y, w - 8, y, fill=THEME["line_soft"])
        title = "GPS 实时" if summary.get("source") == "实时接口" else "GPS 快照"
        canvas.create_text(18, 16, text=title, anchor="w", fill=THEME["cyan"], font=("Microsoft YaHei UI", 10, "bold"))
        canvas.create_text(18, 35, text=summary.get("company", COMPANY_NAME), anchor="w", fill=THEME["muted"], font=("Microsoft YaHei UI", 8))

        stats = [
            ("总车", summary.get("total", "-"), THEME["text"]),
            ("在线", summary.get("online", "-"), THEME["green"]),
            ("离线", summary.get("offline", "-"), THEME["orange"]),
            ("行驶", summary.get("running", "-"), THEME["cyan"]),
            ("停止", summary.get("stopped", "-"), THEME["muted"]),
            ("报警", summary.get("alarm", "-"), THEME["red"]),
        ]
        card_w = max(72, (w - 36) // 3)
        card_h = 26
        for index, (label, value, color) in enumerate(stats):
            col = index % 3
            row = index // 3
            x = 18 + col * card_w
            y = 48 + row * 32
            canvas.create_rectangle(x, y, x + card_w - 10, y + card_h, fill=THEME["card"], outline=THEME["line"])
            canvas.create_line(x, y, x, y + card_h, fill=color, width=2)
            canvas.create_line(x + 4, y + card_h - 1, x + 34, y + card_h - 1, fill=THEME["cyan_soft"])
            canvas.create_text(x + 9, y + card_h // 2, text=label, anchor="w", fill=THEME["muted"], font=("Microsoft YaHei UI", 8))
            canvas.create_text(x + card_w - 23, y + card_h // 2, text=str(value), anchor="e", fill=color, font=("Consolas", 14, "bold"))
        canvas.create_text(
            18,
            h - 10,
            text=f"{summary.get('source', '快照')}：{summary.get('updated_at', '-')}",
            anchor="w",
            fill=THEME["muted"],
            font=("Microsoft YaHei UI", 7, "bold"),
        )

    def _draw_pipeline_panel(self, canvas: Canvas) -> None:
        canvas.delete("all")
        w = max(canvas.winfo_width(), 300)
        h = max(canvas.winfo_height(), 178)
        canvas.create_rectangle(0, 0, w, h, fill=THEME["log"], outline=THEME["line"])
        canvas.create_rectangle(3, 3, w - 4, h - 4, outline=THEME["cyan_soft"])
        canvas.create_line(0, 0, 92, 0, fill=THEME["cyan"], width=2)
        canvas.create_line(w - 130, h - 1, w - 2, h - 1, fill=THEME["green"], width=2)
        for x in range(24, w, 38):
            canvas.create_line(x, 10, x, h - 10, fill=THEME["line_soft"])
        for y in range(28, h, 26):
            canvas.create_line(10, y, w - 10, y, fill=THEME["line_soft"])

        canvas.create_text(18, 20, text="OCR -> Excel 作业管线", anchor="w", fill=THEME["cyan"], font=("Microsoft YaHei UI", 11, "bold"))
        canvas.create_text(18, 42, text="批次处理 / 解析复核 / 交付验收", anchor="w", fill=THEME["muted"], font=("Microsoft YaHei UI", 9))

        stages = [
            ("导入", "IMAGES", THEME["cyan"]),
            ("识别", "OCR", THEME["green"]),
            ("解析", "PARSE", THEME["cyan"]),
            ("复核", "REVIEW", THEME["orange"]),
            ("出表", "EXCEL", THEME["green"]),
        ]
        gap = 10
        margin = 18
        top = 68
        box_h = 58
        usable = w - margin * 2 - gap * (len(stages) - 1)
        box_w = max(44, usable // len(stages))
        for index, (label, code, color) in enumerate(stages):
            x = margin + index * (box_w + gap)
            canvas.create_rectangle(x, top, x + box_w, top + box_h, fill=THEME["card"], outline=THEME["line"])
            canvas.create_line(x, top, x + box_w, top, fill=color, width=2)
            canvas.create_text(x + 9, top + 21, text=label, anchor="w", fill=THEME["text"], font=("Microsoft YaHei UI", 11, "bold"))
            canvas.create_text(x + 9, top + 42, text=code, anchor="w", fill=color, font=("Consolas", 8, "bold"))
            if index < len(stages) - 1:
                ax = x + box_w
                ay = top + box_h // 2
                nx = ax + gap
                canvas.create_line(ax + 3, ay, nx - 5, ay, fill=THEME["cyan_soft"], width=2)
                canvas.create_polygon(nx - 8, ay - 4, nx - 8, ay + 4, nx - 2, ay, fill=THEME["cyan"])

        footer_y = h - 24
        canvas.create_text(18, footer_y, text="整批识别，疑问复核", anchor="w", fill=THEME["muted"], font=("Microsoft YaHei UI", 9))
        canvas.create_text(w - 18, footer_y, text="READY", anchor="e", fill=THEME["green"], font=("Consolas", 11, "bold"))

    def _build_ui(self) -> None:
        chrome = ttk.Frame(self.root, padding=(10, 5), style="Titlebar.TFrame")
        chrome.pack(fill="x")
        chrome.bind("<ButtonPress-1>", self._start_window_drag)
        chrome.bind("<B1-Motion>", self._drag_window)
        title = ttk.Label(chrome, text=APP_TITLE, style="Titlebar.TLabel")
        title.pack(side="left")
        title.bind("<ButtonPress-1>", self._start_window_drag)
        title.bind("<B1-Motion>", self._drag_window)
        ttk.Button(chrome, text="X", command=self.root.destroy, style="Close.TButton", width=3).pack(side="right")
        ttk.Button(chrome, text="□", command=self._toggle_maximize_window, style="Chrome.TButton", width=3).pack(side="right", padx=(4, 0))
        ttk.Button(chrome, text="-", command=self._minimize_window, style="Chrome.TButton", width=3).pack(side="right", padx=(4, 0))
        ttk.Frame(self.root, height=2, style="TitlebarAccent.TFrame").pack(fill="x")

        compact = self._layout_compact
        ultra_compact = self._window_width < 960
        outer = ttk.Frame(self.root, padding=(8 if compact else 12), style="App.TFrame")
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer, style="App.TFrame")
        header.pack(fill="x", pady=(0, 8 if compact else 10))

        header_height = 82 if compact else 92
        self.header_canvas = Canvas(header, height=header_height, bg=THEME["bg"], highlightthickness=0)
        self.header_canvas.pack(fill="x")
        self.header_canvas.bind("<Configure>", lambda _event: self._draw_header_panel(self.header_canvas))
        self._draw_header_panel(self.header_canvas)

        workbench = ttk.Frame(outer, style="App.TFrame")
        workbench.pack(side="top", fill="both", expand=True)
        left_width = 260 if ultra_compact else (292 if compact else 318)
        right_width = left_width
        workbench.columnconfigure(0, minsize=left_width)
        workbench.columnconfigure(1, weight=1)
        workbench.columnconfigure(2, minsize=right_width)
        workbench.rowconfigure(0, weight=1)

        left_panel = ttk.LabelFrame(
            workbench,
            text="批次与操作",
            style="Panel.TLabelframe",
            padding=(10 if compact else 12, 14 if compact else 16, 10 if compact else 12, 8 if compact else 10),
        )
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8 if compact else 10))
        left_panel.columnconfigure(0, weight=1)
        left_panel.rowconfigure(0, weight=1)
        left_panel.rowconfigure(2, weight=1)
        left_content = ttk.Frame(left_panel, style="Panel.TFrame")
        left_content.grid(row=1, column=0, sticky="ew")
        left_content.columnconfigure(0, weight=1)
        ttk.Label(left_content, textvariable=self.input_status, style="BatchStatus.TLabel").grid(row=0, column=0, sticky="ew", pady=(4, 8))
        ttk.Label(left_content, textvariable=self.output_status, style="BatchPath.TLabel").grid(row=1, column=0, sticky="ew", pady=(0, 18))
        ttk.Button(left_content, text="导入图片", command=self._choose_image_dir, style="Batch.TButton").grid(row=2, column=0, sticky="ew", pady=(0, 10), ipady=3)
        ttk.Button(left_content, text="待处理图片", command=self.open_image_dir, style="Batch.TButton").grid(row=3, column=0, sticky="ew", pady=10, ipady=3)

        ttk.Frame(left_content, height=18, style="Panel.TFrame").grid(row=4, column=0, sticky="ew")
        ttk.Button(left_content, text="一键出表", command=self.run_all, style="BatchPrimary.TButton").grid(
            row=5,
            column=0,
            sticky="ew",
            pady=(10, 14),
            ipady=5,
        )
        action_grid = ttk.Frame(left_content, style="Panel.TFrame")
        action_grid.grid(row=6, column=0, sticky="ew")
        action_grid.columnconfigure(0, weight=1)
        action_grid.columnconfigure(1, weight=1)
        action_buttons = [
            ("扫图片", self.scan_batch),
            ("打开结果", self.open_work_dir),
            ("OCR 设置", self.open_ocr_settings),
            ("GPS 设置", self.open_gps_settings),
        ]
        for index, (text, command) in enumerate(action_buttons):
            row = index // 2
            column = index % 2
            ttk.Button(action_grid, text=text, command=command, style="Batch.TButton").grid(
                row=row,
                column=column,
                sticky="ew",
                padx=(0, 4) if column == 0 else (4, 0),
                pady=8,
                ipady=3,
            )

        center_panel = ttk.LabelFrame(workbench, text="流程管线", style="Panel.TLabelframe", padding=(8 if compact else 10, 8 if compact else 10))
        center_panel.grid(row=0, column=1, sticky="nsew", padx=(0, 8 if compact else 10))
        center_panel.rowconfigure(0, weight=0)
        center_panel.rowconfigure(1, weight=1)
        center_panel.columnconfigure(0, weight=1)
        pipeline_width = 300 if ultra_compact else (390 if compact else 470)
        pipeline_height = 116 if ultra_compact else (122 if compact else 132)
        self.pipeline_canvas = Canvas(center_panel, width=pipeline_width, height=pipeline_height, bg=THEME["log"], highlightthickness=0)
        self.pipeline_canvas.grid(row=0, column=0, sticky="ew")
        self.pipeline_canvas.bind("<Configure>", lambda _event: self._draw_pipeline_panel(self.pipeline_canvas))
        self._draw_pipeline_panel(self.pipeline_canvas)

        progress_panel = ttk.Frame(center_panel, style="Panel.TFrame")
        progress_panel.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        progress_panel.columnconfigure(0, weight=1)
        progress_panel.rowconfigure(2, weight=1)
        progress_header = ttk.Frame(progress_panel, style="Panel.TFrame")
        progress_header.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        progress_header.columnconfigure(0, weight=1)
        ttk.Label(progress_header, textvariable=self.progress_title, style="Panel.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(progress_header, textvariable=self.progress_detail, style="PanelMuted.TLabel").grid(row=0, column=1, sticky="e", padx=(12, 8))
        ttk.Label(progress_header, textvariable=self.progress_percent, style="ProgressPercent.TLabel").grid(row=0, column=2, sticky="e")
        ttk.Progressbar(
            progress_panel,
            variable=self.progress_value,
            maximum=100,
            mode="determinate",
            style="Ops.Horizontal.TProgressbar",
        ).grid(row=1, column=0, sticky="ew", pady=(0, 6))
        progress_log_frame = ttk.Frame(progress_panel, style="Panel.TFrame")
        progress_log_frame.grid(row=2, column=0, sticky="nsew")
        progress_log_frame.columnconfigure(0, weight=1)
        progress_log_frame.rowconfigure(0, weight=1)
        self.progress_log_text = Text(
            progress_log_frame,
            wrap="word",
            height=7 if compact else 8,
            bg=THEME["log"],
            fg=THEME["muted"],
            insertbackground=THEME["cyan"],
            selectbackground=THEME["selected"],
            relief="flat",
            borderwidth=0,
            padx=8,
            pady=6,
            state="disabled",
            font=("Consolas", 10),
        )
        self.progress_log_text.tag_configure("section", foreground=THEME["cyan"], font=("Microsoft YaHei UI", 10, "bold"))
        self.progress_log_text.tag_configure("success", foreground=THEME["green"])
        self.progress_log_text.tag_configure("error", foreground="#ffd1d8", background="#3a0d16")
        progress_scroll = ttk.Scrollbar(progress_log_frame, orient="vertical", command=self.progress_log_text.yview)
        self.progress_log_text.configure(yscrollcommand=progress_scroll.set)
        self.progress_log_text.grid(row=0, column=0, sticky="nsew")
        progress_scroll.grid(row=0, column=1, sticky="ns")

        right_panel = ttk.Frame(workbench, style="App.TFrame")
        right_panel.grid(row=0, column=2, sticky="nsew")
        right_panel.columnconfigure(0, weight=1)
        right_panel.rowconfigure(1, weight=1)

        telemetry = ttk.LabelFrame(right_panel, text="车队遥测", style="Panel.TLabelframe", padding=(8, 8))
        telemetry.grid(row=0, column=0, sticky="ew")
        route_width = 260 if ultra_compact else (288 if compact else 310)
        route_height = 124 if compact else 132
        self.route_canvas = Canvas(telemetry, width=route_width, height=route_height, bg=THEME["log"], highlightthickness=0, cursor="hand2")
        self.route_canvas.pack(fill="x")
        self.route_canvas.bind("<Button-1>", lambda _event: self.refresh_gps())
        self.route_canvas.bind("<Configure>", lambda _event: self._draw_route_panel(self.route_canvas))
        self._draw_route_panel(self.route_canvas)

        ocr_panel = ttk.LabelFrame(right_panel, text="OCR 控制", style="Panel.TLabelframe", padding=(8 if compact else 9, 6 if compact else 7))
        ocr_panel.grid(row=1, column=0, sticky="nsew", pady=(8 if compact else 10, 0))
        ocr_panel.columnconfigure(0, weight=1)
        ttk.Label(ocr_panel, text="OCR 引擎", style="PanelMuted.TLabel").grid(row=0, column=0, sticky="w")
        engine_combo = ttk.Combobox(ocr_panel, textvariable=self.ocr_engine, values=list(OCR_ENGINE_LABELS.values()), state="readonly", width=30)
        engine_combo.grid(row=1, column=0, sticky="ew", pady=(3, 6))
        ttk.Label(ocr_panel, text="OCR 档位 / 实际模型", style="PanelMuted.TLabel").grid(row=2, column=0, sticky="w")
        combo = ttk.Combobox(ocr_panel, textvariable=self.ocr_profile, values=list(OCR_PROFILE_LABELS.values()), state="readonly", width=38)
        combo.grid(row=3, column=0, sticky="ew", pady=(3, 4))
        self.ocr_profile.trace_add("write", self._refresh_ocr_profile_detail)
        ttk.Label(ocr_panel, textvariable=self.ocr_profile_detail, style="PanelMuted.TLabel", wraplength=290).grid(row=4, column=0, sticky="ew", pady=(0, 6))
        ttk.Checkbutton(ocr_panel, text="已有 OCR JSON 时跳过整批 OCR", variable=self.skip_ocr_if_json).grid(row=5, column=0, sticky="w", pady=(0, 6))
        ttk.Label(ocr_panel, text="ZHIPU API Key", style="PanelMuted.TLabel").grid(row=6, column=0, sticky="w")
        ttk.Entry(ocr_panel, textvariable=self.api_key, show="*").grid(row=7, column=0, sticky="ew", pady=(3, 5))
        ttk.Checkbutton(ocr_panel, text="记住 Key（仅本机）", variable=self.remember_api_key).grid(row=8, column=0, sticky="w")
        ttk.Label(ocr_panel, textvariable=self.api_key_source, style="PanelMuted.TLabel").grid(row=9, column=0, sticky="w", pady=(5, 0))

        self.log("已启动。放入图片后直接点“一键出表”。")

    def _path_row(self, parent: ttk.Frame, row: int, label: str, var: StringVar, command) -> None:
        ttk.Label(parent, text=label, width=16, style="Panel.TLabel").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        ttk.Button(parent, text="选择", command=command, style="Tool.TButton").grid(row=row, column=2, sticky="e", pady=4)
        parent.columnconfigure(1, weight=1)

    def _open_command_panel(self, title: str, buttons: list[tuple[str, object]], columns: int = 2) -> None:
        window = Toplevel(self.root)
        window.title(title)
        window.configure(bg=THEME["bg"])
        window.transient(self.root)

        body = ttk.Frame(window, padding=14, style="App.TFrame")
        body.pack(fill="both", expand=True)
        for col in range(columns):
            body.columnconfigure(col, weight=1, uniform=f"{title}_buttons")

        for index, (text, command) in enumerate(buttons):
            def run(cmd=command) -> None:
                window.destroy()
                cmd()

            ttk.Button(body, text=text, command=run, style="Tool.TButton").grid(
                row=index // columns,
                column=index % columns,
                sticky="ew",
                padx=5,
                pady=5,
                ipady=3,
            )

        window.resizable(False, False)
        window.update_idletasks()
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - window.winfo_width()) // 2)
        y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - window.winfo_height()) // 2)
        window.geometry(f"+{x}+{y}")

    def open_ocr_settings(self) -> None:
        self._open_command_panel(
            "OCR 设置",
            [
                ("推荐方案", self.recommend_ocr),
                ("方案库", self.show_ocr_strategy_library),
                ("检查环境", self.check_ocr_env),
                ("单独 OCR", self.run_ocr),
                ("装本地 OCR", self.install_local_ocr_env),
                ("官方资料", self.show_ocr_references),
            ],
            columns=3,
        )

    def open_gps_settings(self) -> None:
        self.refresh_gps()
        self._open_command_panel(
            "GPS 设置",
            [
                ("配置授权", self.configure_gps_auth),
                ("刷新 GPS", self.refresh_gps),
                ("GPS 后台", self.open_gps_backend),
            ],
            columns=3,
        )

    def refresh_gps_on_startup(self) -> None:
        if not gps_api_configured():
            return
        if self.worker_thread and self.worker_thread.is_alive():
            self.root.after(1500, self.refresh_gps_on_startup)
            return
        self.run_background("GPS 自动刷新", self._refresh_gps_impl, show_progress=False)

    def configure_gps_auth(self) -> None:
        config = {
            "url": GPS_API_DEFAULT_URL,
            "method": GPS_API_DEFAULT_METHOD,
            "sortField": "plate",
            "company": DEFAULT_GPS_SUMMARY["company"],
            "timeout": 20,
            **load_gps_api_config(),
        }
        window = Toplevel(self.root)
        window.title("GPS 授权")
        window.configure(bg=THEME["bg"])
        window.transient(self.root)

        body = ttk.Frame(window, padding=14, style="App.TFrame")
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)

        url_var = StringVar(value=str(config.get("url") or GPS_API_DEFAULT_URL))
        method_var = StringVar(value=str(config.get("method") or GPS_API_DEFAULT_METHOD).upper())
        company_var = StringVar(value=str(config.get("company") or DEFAULT_GPS_SUMMARY["company"]))
        sort_var = StringVar(value=str(config.get("sortField") or "plate"))
        timeout_var = StringVar(value=str(config.get("timeout") or "20"))
        auth_var = StringVar(value=str(config.get("authorization") or ""))

        def add_entry(row: int, label: str, var: StringVar, show: str = "") -> None:
            ttk.Label(body, text=label, width=14).grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(body, textvariable=var, show=show).grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=4)

        add_entry(0, "接口地址", url_var)
        ttk.Label(body, text="请求方式", width=14).grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(body, textvariable=method_var, values=["GET", "POST"], state="readonly", width=10).grid(row=1, column=1, sticky="w", padx=(8, 0), pady=4)
        add_entry(2, "公司名称", company_var)
        add_entry(3, "排序字段", sort_var)
        add_entry(4, "超时秒数", timeout_var)
        add_entry(5, "Authorization", auth_var, show="*")

        ttk.Label(body, text="Cookie", width=14).grid(row=6, column=0, sticky="nw", pady=4)
        cookie_text = Text(
            body,
            height=5,
            width=64,
            wrap="word",
            bg=THEME["log"],
            fg=THEME["text"],
            insertbackground=THEME["cyan"],
            relief="flat",
            borderwidth=0,
            padx=8,
            pady=6,
            font=("Microsoft YaHei UI", 9),
        )
        cookie_text.grid(row=6, column=1, sticky="ew", padx=(8, 0), pady=4)
        cookie_text.insert("1.0", str(config.get("cookie") or ""))

        ttk.Label(body, text="请求体", width=14).grid(row=7, column=0, sticky="nw", pady=4)
        body_text = Text(
            body,
            height=4,
            width=64,
            wrap="word",
            bg=THEME["log"],
            fg=THEME["text"],
            insertbackground=THEME["cyan"],
            relief="flat",
            borderwidth=0,
            padx=8,
            pady=6,
            font=("Microsoft YaHei UI", 9),
        )
        body_text.grid(row=7, column=1, sticky="ew", padx=(8, 0), pady=4)
        body_text.insert("1.0", str(config.get("body") or ""))

        button_row = ttk.Frame(body, style="App.TFrame")
        button_row.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(12, 0))

        def read_timeout() -> int:
            try:
                return max(5, int(timeout_var.get().strip()))
            except Exception:
                return 20

        def save_auth(refresh_after: bool = False) -> None:
            save_gps_api_config(
                {
                    "url": url_var.get().strip() or GPS_API_DEFAULT_URL,
                    "method": method_var.get().strip().upper() or GPS_API_DEFAULT_METHOD,
                    "sortField": sort_var.get().strip() or "plate",
                    "company": company_var.get().strip() or DEFAULT_GPS_SUMMARY["company"],
                    "timeout": read_timeout(),
                    "authorization": auth_var.get().strip(),
                    "cookie": cookie_text.get("1.0", "end").strip(),
                    "body": body_text.get("1.0", "end").strip(),
                }
            )
            self.log_section("GPS 授权")
            self.log("GPS 授权已保存到本机。", "success")
            window.destroy()
            if refresh_after:
                self.refresh_gps()

        def clear_auth() -> None:
            delete_saved_gps_api_config()
            self.log_section("GPS 授权")
            self.log("GPS 授权已清除。")
            window.destroy()

        ttk.Button(button_row, text="保存并刷新", command=lambda: save_auth(True), style="Primary.TButton").pack(side="left", ipadx=12)
        ttk.Button(button_row, text="保存", command=lambda: save_auth(False), style="Tool.TButton").pack(side="left", padx=8, ipadx=12)
        ttk.Button(button_row, text="清除授权", command=clear_auth, style="Tool.TButton").pack(side="left", ipadx=12)
        ttk.Button(button_row, text="取消", command=window.destroy, style="Tool.TButton").pack(side="right", ipadx=12)

        window.minsize(720, 470)
        window.update_idletasks()
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - window.winfo_width()) // 2)
        y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - window.winfo_height()) // 2)
        window.geometry(f"+{x}+{y}")

    def log(self, message: str, tag: str = "") -> None:
        timestamp = time.strftime("%H:%M:%S")
        self._progress_log(message, tag, timestamp)
        self._log_started = True

    def log_section(self, title: str) -> None:
        self.log(f"【{title}】", "section")

    def _progress_log(self, message: str, tag: str = "", timestamp: str | None = None) -> None:
        widget = getattr(self, "progress_log_text", None)
        if widget is None:
            return
        timestamp = timestamp or time.strftime("%H:%M:%S")
        widget.configure(state="normal")
        if tag == "section" and self._progress_log_started:
            widget.insert(END, "\n")
        compact_message = message.replace("\n", " ").strip()
        widget.insert(END, f"[{timestamp}] {compact_message}\n", tag if tag else None)
        widget.configure(state="disabled")
        widget.see(END)
        self._progress_log_started = True

    def post(self, message: str, tag: str = "") -> None:
        self.events.put((message, tag))

    def post_progress(self, percent: float, title: str = "", detail: str = "") -> None:
        self.events.put(("__progress__", max(0.0, min(100.0, float(percent))), title, detail, time.time()))

    def _format_eta(self, percent: float) -> str:
        if percent <= 0 or percent >= 100 or not self._progress_started_at:
            return "--"
        elapsed = max(time.time() - self._progress_started_at, 0.0)
        remaining = elapsed * (100.0 - percent) / percent
        if remaining < 60:
            return f"{int(max(remaining, 1))} 秒"
        minutes = int(remaining // 60)
        seconds = int(remaining % 60)
        return f"{minutes} 分 {seconds:02d} 秒"

    def _apply_progress(self, percent: float, title: str = "", detail: str = "") -> None:
        percent = max(0.0, min(100.0, float(percent)))
        self._last_progress_percent = percent
        self.progress_value.set(percent)
        self.progress_percent.set(f"{percent:.0f}%")
        if title:
            self.progress_title.set(title)
        eta = self._format_eta(percent)
        elapsed = "--"
        if self._progress_started_at:
            elapsed_seconds = max(time.time() - self._progress_started_at, 0.0)
            elapsed = f"{int(elapsed_seconds)} 秒" if elapsed_seconds < 60 else f"{int(elapsed_seconds // 60)} 分 {int(elapsed_seconds % 60):02d} 秒"
        detail_text = detail.strip()
        suffix = f"已用 {elapsed} / 剩余 {eta}"
        self.progress_detail.set(f"{detail_text}  {suffix}" if detail_text else suffix)

    def _reset_progress(self, title: str) -> None:
        self._progress_started_at = time.time()
        self._last_progress_percent = 0.0
        self._progress_log_started = False
        self.progress_value.set(0.0)
        self.progress_percent.set("0%")
        self.progress_title.set(title)
        self.progress_detail.set("已用 0 秒 / 剩余 --")
        widget = getattr(self, "progress_log_text", None)
        if widget is not None:
            widget.configure(state="normal")
            widget.delete("1.0", END)
            widget.configure(state="disabled")

    def persist_api_key_choice(self) -> None:
        key = self.api_key.get().strip()
        if self.remember_api_key.get() and key:
            save_api_key(key)
            self.root.after(0, self.api_key_source.set, "Key 来源：本机已保存")
            self.post("API Key 已保存到本机；下次打开会自动填好。", "success")
        elif not self.remember_api_key.get():
            delete_saved_api_key()
            source = "EXE 内置默认" if builtin_api_key() else "未填写"
            self.root.after(0, self.api_key_source.set, f"Key 来源：{source}")
            self.post("已关闭记住 Key，并清除本机保存的 Key。")

    def _drain_events(self) -> None:
        try:
            while True:
                item = self.events.get_nowait()
                if isinstance(item, tuple):
                    if item and item[0] == "__progress__":
                        self._apply_progress(item[1], item[2], item[3])
                    else:
                        self.log(item[0], item[1])
                else:
                    self.log(item)
        except queue.Empty:
            pass
        self.root.after(150, self._drain_events)

    def _choose_image_dir(self) -> None:
        path = filedialog.askdirectory(title="导入单据图片所在文件夹")
        if path:
            self.image_dir.set(path)
            output_dir = managed_output_dir()
            self.work_dir.set(str(output_dir))
            self.raw_json.set(str(output_dir / "paddle_ocr_raw.json"))
            self._refresh_workspace_status()

    def _choose_python(self) -> None:
        path = filedialog.askopenfilename(title="选择 python.exe", filetypes=[("Python", "python.exe"), ("Executable", "*.exe"), ("All", "*.*")])
        if path:
            self.python_path.set(path)

    def _choose_raw_json(self) -> None:
        path = filedialog.askopenfilename(title="选择 OCR JSON", filetypes=[("JSON", "*.json"), ("All", "*.*")])
        if path:
            self.raw_json.set(path)

    def open_image_dir(self) -> None:
        self.log_section("待处理图片")
        path = Path(self.image_dir.get()).resolve()
        path.mkdir(parents=True, exist_ok=True)
        self.log(f"打开：{path}")
        os.startfile(str(path))

    def require_idle(self) -> bool:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning(APP_TITLE, "任务正在运行，请等待完成。")
            return False
        return True

    def confirm_excel_template(self) -> dict | None:
        if not self.require_idle():
            return None
        dialog = ExcelTemplateDialog(self.root, self.excel_columns, self.excel_image_mode)
        result = dialog.show()
        if result is not None:
            self.excel_columns = normalize_excel_columns(result.get("columns"))
            self.excel_image_mode = normalize_excel_image_mode(result.get("image_mode"))
        return result

    def scan_batch(self) -> None:
        self.log_section("扫图片")
        self._refresh_workspace_status()
        image_dir = Path(self.image_dir.get())
        work_dir = Path(self.work_dir.get())
        images = input_images(image_dir)
        workspace = work_dir.parent
        status = ocr_json_status(Path(self.raw_json.get()), images)
        status_text = "可用" if status["complete"] else "不可用/需重跑"
        if not status["exists"]:
            status_text = "不存在/需 OCR"
        self.log(f"图片目录：{image_dir}")
        self.log(f"批次总目录：{workspace}")
        self.log(f"输出目录：{work_dir}")
        self.log(f"图片数量：{len(images)}")
        self.log(f"OCR 引擎：{self.ocr_engine.get()}")
        self.log(f"当前 OCR JSON：{status_text}，已识别 {status['ok_count']}/{status['image_count']}")
        for name in ("paddle_ocr_raw.json", "transport_parsed_records.json", "运输录入_结构化预览.csv"):
            self.log(f"{name}：{'存在' if (work_dir / name).exists() else '不存在'}")

    def run_background(self, title: str, fn, show_progress: bool = True) -> None:
        if not self.require_idle():
            return
        self.log_section(title)
        if show_progress:
            self._reset_progress(title)
        def runner():
            self.post("开始")
            if show_progress:
                self.post_progress(1, title, "任务启动")
            try:
                fn()
                if show_progress:
                    self.post_progress(100, "完成", "已完成")
                self.post("完成", "success")
            except Exception as exc:
                error_message = str(exc)
                self._write_error_log(title, traceback.format_exc())
                if show_progress:
                    self.post_progress(self._last_progress_percent, "任务失败", "已停止")
                self.post(f"失败：{error_message}", "error")
                self.post("详细错误已保存到输出结果。")
                if not isinstance(exc, SuppressPopupError):
                    self.root.after(0, lambda msg=error_message: messagebox.showerror(APP_TITLE, msg, parent=self.root))
        self.worker_thread = threading.Thread(target=runner, daemon=True)
        self.worker_thread.start()

    def _write_error_log(self, title: str, details: str) -> None:
        try:
            path = managed_output_dir() / "错误日志.txt"
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open("a", encoding="utf-8") as f:
                f.write(f"\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] {title}\n")
                f.write(details)
                if not details.endswith("\n"):
                    f.write("\n")
        except Exception:
            pass

    def run_ocr(self) -> None:
        self.run_background("运行 OCR", self._run_ocr_impl)

    def recommend_ocr(self) -> None:
        self.run_background("推荐 OCR", self._recommend_ocr_impl)

    def check_ocr_env(self) -> None:
        self.run_background("查环境", self._check_ocr_env_impl)

    def install_local_ocr_env(self) -> None:
        self.run_background("装本地 OCR", self._install_local_ocr_env_impl)

    def show_ocr_strategy_library(self) -> None:
        self.log_section("OCR 方案库")
        self.log("当前可直接调用：在线智谱 GLM-OCR，适合直接发同事使用；本地 PaddleOCR PP-OCRv5，适合你这台 5090 环境配置完整后批量跑。")
        self.log("下一步优先接入：通义 Qwen-OCR / Qwen-VL，用作在线备选和疑问行复核；对中文票据、表格和复杂版式值得抽样对比。")
        self.log("本地增强路线：PaddleOCR PP-StructureV3 / PP-ChatOCRv4 用于表格和关键信息抽取；本地 Qwen2.5-VL / GOT-OCR2 / Surya 用于疑问行复核，不默认整批跑。")
        self.log("海外或 PDF 场景备选：Mistral OCR 适合文档/PDF 转结构化内容；如果以后要接入，先做 20 张样本准确率、漏识率、速度和成本对比。")
        self.log("执行原则：默认不多引擎全量乱跑；先主引擎整批，再对疑问行用备选引擎复核，最后人工确认，不能让模型猜做账字段。")

    def show_ocr_references(self) -> None:
        self.log_section("OCR 官方资料")
        self.log("智谱 GLM-OCR：https://docs.bigmodel.cn/cn/guide/models/vlm/glm-ocr")
        self.log("PaddleOCR PP-OCRv5 / GPU 优化：https://www.paddleocr.ai/main/en/version3.x/pipeline_usage/OCR.html")
        self.log("PaddleOCR 性能 FAQ：https://www.paddleocr.ai/main/en/FAQ.html")
        self.log("通义 Qwen-OCR API：https://www.alibabacloud.com/help/en/model-studio/qwen-vl-ocr-api-reference")
        self.log("Mistral OCR API：https://docs.mistral.ai/capabilities/OCR/basic_ocr")
        self.log("GOT-OCR2 本地模型：https://huggingface.co/docs/transformers/model_doc/got_ocr2")
        self.log("Surya 本地 OCR：https://github.com/datalab-to/surya")

    def _preflight_ocr_for_generation(self) -> tuple[list[Path], dict]:
        self.persist_api_key_choice()
        image_dir = Path(self.image_dir.get()).resolve()
        work_dir = Path(self.work_dir.get()).resolve()
        raw_json = Path(self.raw_json.get()).resolve()
        work_dir.mkdir(parents=True, exist_ok=True)
        images = input_images(image_dir)
        if not images:
            raise RuntimeError(f"待处理图片中没有可处理图片。\n请先把本批次图片放入：{image_dir}")

        status = ocr_json_status(raw_json, images)
        self.post(f"批次总目录：{work_dir.parent}")
        self.post(f"图片预检：{len(images)} 张")
        if status["complete"] and self.skip_ocr_if_json.get():
            self.post(f"OCR JSON 可用：{status['ok_count']}/{status['image_count']}，本次直接进入结构化。", "success")
            return images, status

        if status["exists"]:
            reasons = []
            if status["error"]:
                reasons.append(status["error"])
            if status["missing"]:
                reasons.append(f"缺少 {len(status['missing'])} 张")
            if status["invalid"]:
                reasons.append(f"失败/空文本 {len(status['invalid'])} 张")
            if status["extra"]:
                reasons.append(f"混入旧图片 {len(status['extra'])} 张")
            self.post("OCR JSON 不可用，自动重跑 OCR：" + "，".join(reasons))
            try:
                raw_json.unlink()
            except FileNotFoundError:
                pass

        engine = self.selected_ocr_engine()
        if engine == "glm":
            if not self.api_key.get().strip():
                raise RuntimeError(
                    "在线智谱 GLM-OCR 没有可用 API Key，已停止生成 Excel。\n"
                    "处理办法：点“OCR 设置”填写 ZHIPU API Key，或使用带内置 Key 的新版 exe。"
                )
            self.post("OCR 预检：使用在线智谱 GLM-OCR；本机不需要安装 PaddleOCR。", "section")
            return images, status

        self.post("OCR 预检：当前选择本地 PaddleOCR，开始检查本地环境。")
        paddle_state = self._ensure_local_ocr_ready(install_if_missing=True)
        if paddle_ready(paddle_state):
            self.post("本地 PaddleOCR 环境可用，本次会调用 GPU。", "success")
        else:
            self.post("本地 PaddleOCR 环境可用，本次会调用 CPU。", "success")
        return images, status

    def _recommend_ocr_impl(self) -> None:
        self.persist_api_key_choice()
        hardware = detect_hardware()
        device_plan, device_summary = local_ocr_device_plan(hardware)
        self.post(f"系统：{hardware.get('system')}")
        if hardware.get("gpus"):
            for gpu in hardware["gpus"]:
                memory = gpu.get("memory_mb") or 0
                memory_text = f"{round(memory / 1024, 1)}GB" if memory else "显存未知"
                self.post(f"GPU：{gpu.get('name')} / {memory_text} / driver={gpu.get('driver') or '未知'}")
        else:
            self.post("GPU：未检测到 NVIDIA GPU")
        self.post("本地 OCR 设备判断：" + device_summary)
        if device_plan == "cpu":
            self.post("如果继续选择本地 PaddleOCR，将安装/使用 CPU 版；能跑，但速度比独立显卡慢。", "section")

        self.post("开始扫描本机 Python，查找可用的本地 OCR 环境。")
        paddle_state, checked = find_best_paddle_environment(self.python_path.get(), self.post)
        if paddle_state and paddle_state.get("python"):
            self.root.after(0, self.python_path.set, paddle_state["python"])
            self.post(f"本地 OCR Python 已选中：{paddle_state['python']}")
        recommendation, reasons = build_ocr_recommendation(hardware, paddle_state, self.api_key.get())
        for reason in reasons:
            self.post(reason)
        if paddle_ready(paddle_state):
            self.root.after(0, self.set_ocr_engine, "paddle")
            self.post("已自动切到 paddle：本地 OCR 环境完整。", "success")
        elif paddle_basic_ready(paddle_state):
            if self.selected_ocr_engine() == "paddle":
                self.post("当前保持本地 PaddleOCR 选择；可用 CPU 模式运行，GPU 未通过时不会调用 5090。", "section")
            else:
                self.post("检测到本地 PaddleOCR CPU 环境可用；普通分享仍建议在线 OCR，本地 CPU 会慢。", "section")
        else:
            if self.selected_ocr_engine() == "paddle":
                self.post("当前保持本地 PaddleOCR 选择；环境不完整时，一键出表会先尝试联网安装本地环境。", "section")
            else:
                self.post("已保持在线智谱 GLM-OCR：当前本地 OCR 不完整，在线 OCR 更稳。", "section")
            self.post("说明：5090 只是显卡强；本地 OCR 还需要装好 PaddleOCR + GPU 版 PaddlePaddle，并选对 python.exe。")
            self.post(f"已检查 Python 数量：{len(checked)}")
        self.post(recommendation)

    def _check_ocr_env_impl(self) -> None:
        self.persist_api_key_choice()
        hardware = detect_hardware()
        self.post("先说结论：这里会同时检查在线 OCR 条件和本地 OCR 条件。")
        self.post("在线 OCR：只需要网络 + ZHIPU API Key；不需要本地 Python/Paddle。")
        self.post("在线 Key 状态：" + ("已填写" if self.api_key.get().strip() else "未填写"))
        self.post("本地 OCR：需要 Python + PaddleOCR + PaddlePaddle；有 NVIDIA 且 GPU 检查通过时用 gpu:0，否则可用 CPU 模式。")
        paddle_state, checked = find_best_paddle_environment(self.python_path.get(), self.post)
        if paddle_state and paddle_state.get("python"):
            self.root.after(0, self.python_path.set, paddle_state["python"])
            self.post(f"当前用于本地 OCR 的 Python：{paddle_state['python']}")
        recommendation, reasons = build_ocr_recommendation(hardware, paddle_state, self.api_key.get())
        for reason in reasons:
            self.post(reason)
        device_plan, device_summary = local_ocr_device_plan(hardware)
        self.post("本地设备判断：" + device_summary)
        if device_plan == "cpu":
            self.post("结论：这台机器本地 OCR 会走 CPU，不会假装调用显卡。")
        self.post(recommendation)
        if self.selected_ocr_engine() == "glm":
            self.post("当前选择的是在线智谱 GLM-OCR，所以本地 OCR 环境不完整也不影响在线 OCR。")
            return
        if not paddle_basic_ready(paddle_state):
            self.post("当前选择的是本地 PaddleOCR；环境还不完整，但不会自动切到在线 OCR。")
            self.post("下一步：点“装本地 OCR”，软件会联网补齐 Python / PaddleOCR / PaddlePaddle。", "section")
        elif not paddle_ready(paddle_state):
            self.post("当前本地 PaddleOCR 可用，但 GPU/CUDA 未通过，将以 CPU 模式运行。")

    def _run_streamed_command(self, cmd: list[str], title: str, timeout: int | None = None) -> None:
        self.post(title, "section")
        self.post("运行命令：" + " ".join(cmd))
        env = os.environ.copy()
        env.setdefault("PYTHONUTF8", "1")
        env.setdefault("PIP_DISABLE_PIP_VERSION_CHECK", "1")
        env.setdefault("PIP_DEFAULT_TIMEOUT", "120")
        env.setdefault("PADDLE_PDX_MODEL_SOURCE", "BOS")
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            startupinfo=hidden_startupinfo(),
            env=env,
        )
        assert process.stdout is not None
        started = time.time()
        for line in process.stdout:
            self.post(line.rstrip())
            if timeout and time.time() - started > timeout:
                process.terminate()
                raise RuntimeError(f"{title} 超时，已停止。")
        code = process.wait()
        if code != 0:
            raise RuntimeError(f"{title} 失败，退出码 {code}")

    def _ensure_python_for_local_ocr(self) -> str:
        python = self.python_path.get().strip()
        if python and python_usable(python):
            return python
        for candidate in find_python_candidates():
            if python_usable(candidate):
                self.root.after(0, self.python_path.set, candidate)
                return candidate
        if os.name == "nt":
            winget = shutil.which("winget")
            if winget:
                self.post("未找到可用 Python，开始通过 winget 安装 Python 3.12。", "section")
                self._run_streamed_command(
                    [
                        winget,
                        "install",
                        "--id",
                        "Python.Python.3.12",
                        "-e",
                        "--silent",
                        "--accept-package-agreements",
                        "--accept-source-agreements",
                    ],
                    "安装 Python 3.12",
                    timeout=900,
                )
                for candidate in find_python_candidates():
                    if python_usable(candidate):
                        self.root.after(0, self.python_path.set, candidate)
                        return candidate
        raise RuntimeError("没有找到可用 Python，也无法自动安装。请先安装 Python 3.10-3.12，再点“装本地 OCR”。")

    def _ensure_local_ocr_ready(self, install_if_missing: bool = True) -> dict:
        paddle_state, checked = find_best_paddle_environment(self.python_path.get(), self.post)
        if paddle_basic_ready(paddle_state):
            if paddle_state and paddle_state.get("python"):
                self.root.after(0, self.python_path.set, paddle_state["python"])
            if paddle_ready(paddle_state):
                self.post("本地 OCR GPU 环境完整，将使用 gpu:0。", "success")
            else:
                self.post("本地 OCR CPU 环境可用，将使用 cpu；速度会比 5090 GPU 慢。", "section")
            return paddle_state
        if not install_if_missing:
            raise RuntimeError("本地 PaddleOCR 环境不完整。请选择“装本地 OCR”补齐环境。")
        self.post(f"本地 OCR 环境不完整，已检查 Python 数量：{len(checked)}")
        self.post("保持本地 OCR 选择，开始联网补齐环境；不会自动切到在线 OCR。", "section")
        self._install_local_ocr_env_impl(auto=True)
        paddle_state, _ = find_best_paddle_environment(self.python_path.get(), self.post)
        if paddle_basic_ready(paddle_state):
            if paddle_state and paddle_state.get("python"):
                self.root.after(0, self.python_path.set, paddle_state["python"])
            if paddle_ready(paddle_state):
                self.post("本地 OCR 环境已补齐并通过 GPU 检查，将使用 gpu:0。", "success")
            else:
                self.post("本地 OCR 环境已补齐，GPU 未通过但 CPU 可用，将使用 cpu。", "success")
            return paddle_state
        raise RuntimeError(
            "本地 OCR 自动安装后仍未通过检查，已停止生成。"
            "请查看中间进度日志里的 pip / Paddle 错误；50 系显卡可按官方 50-series wheel 说明手动处理。"
        )

    def _install_local_ocr_env_impl(self, auto: bool = False) -> None:
        hardware = detect_hardware()
        device_plan, device_summary = local_ocr_device_plan(hardware)
        python = self._ensure_python_for_local_ocr()
        self.root.after(0, self.python_path.set, python)
        self.post("本地 OCR 环境只用于 PaddleOCR 模式；在线 GLM-OCR 不需要安装这些依赖。")
        self.post(device_summary)
        if device_plan == "cpu":
            self.post("本次会安装 CPU 版 PaddlePaddle。虚拟机或无 NVIDIA 显卡时可以跑本地 OCR，但不会调用独立显卡。", "section")
        else:
            self.post("本次会优先安装 GPU 版 PaddlePaddle；只有 GPU 检查通过后才会真正使用 gpu:0。", "section")
        self.post(f"本地 OCR Python：{python}")
        self.post_progress(8, "本地 OCR 安装", "准备 Python")
        self._run_streamed_command([python, "-m", "pip", "install", "--upgrade", "--prefer-binary", "--timeout", "120", "pip", "setuptools", "wheel"], "升级 pip 基础工具", timeout=900)
        self.post_progress(24, "本地 OCR 安装", "安装 PaddleOCR")
        self._run_streamed_command([python, "-m", "pip", "install", "--upgrade", "--prefer-binary", "--timeout", "120", "paddleocr"], "安装 PaddleOCR", timeout=1800)
        install_label, paddle_cmd = local_paddle_install_command(python, hardware)
        self.post_progress(55, "本地 OCR 安装", install_label)
        self._run_streamed_command(paddle_cmd, install_label, timeout=2400)
        self.post_progress(85, "本地 OCR 安装", "验证环境")
        self.post("正在验证 PaddleOCR / PaddlePaddle。首次导入可能较慢，请等待；虚拟机无 NVIDIA 时 GPU 未通过属于正常情况。")
        state = check_paddle_environment(python)
        for line in state.get("lines", []):
            self.post(line)
        if not paddle_basic_ready(state):
            self.post(f"官方安装说明：{PADDLE_INSTALL_DOC_URL}")
            webbrowser.open(PADDLE_INSTALL_DOC_URL)
            raise RuntimeError(
                "本地 OCR 依赖已尝试安装，但 PaddleOCR/Paddle 检查仍未通过。"
                "已打开官方安装说明，请按显卡和 Python 版本补齐后重新检查。"
            )
        self.root.after(0, self.set_ocr_engine, "paddle")
        self.post_progress(100 if not auto else 88, "本地 OCR 安装", "验证通过")
        if paddle_ready(state):
            self.post("本地 PaddleOCR GPU 环境安装完成，将使用 gpu:0。", "success")
        else:
            self.post("本地 PaddleOCR CPU 环境安装完成，将使用 cpu；虚拟机/无 NVIDIA 显卡时这是正确结果。", "success")

    def _run_ocr_impl(self, progress_base: float = 0.0, progress_span: float = 100.0) -> None:
        self.persist_api_key_choice()
        image_dir = Path(self.image_dir.get()).resolve()
        work_dir = Path(self.work_dir.get()).resolve()
        raw_json = Path(self.raw_json.get()).resolve()
        work_dir.mkdir(parents=True, exist_ok=True)
        images = input_images(image_dir)
        status = ocr_json_status(raw_json, images)
        if self.skip_ocr_if_json.get() and status["complete"]:
            self.post(f"已存在可用 OCR JSON，按规则跳过整批 OCR：{raw_json}")
            self.post_progress(progress_base + progress_span, "OCR 已完成", f"已识别 {status['ok_count']}/{status['image_count']}")
            return
        if raw_json.exists() and not status["complete"]:
            self.post("已有 OCR JSON 不可用，删除后重新 OCR。")
            raw_json.unlink()
        def progress_fn(done: int, total: int, label: str = "") -> None:
            if total <= 0:
                percent = progress_base
            else:
                percent = progress_base + progress_span * max(0, min(done, total)) / total
            self.post_progress(percent, "OCR 识别中", f"{done}/{total} {label}".strip())

        if self.selected_ocr_engine() == "glm":
            run_glm_ocr_batch(image_dir, raw_json, self.api_key.get().strip(), self.post, progress_fn)
            return
        paddle_state = self._ensure_local_ocr_ready(install_if_missing=True)
        if paddle_state and paddle_state.get("python"):
            self.root.after(0, self.python_path.set, paddle_state["python"])
        if not images:
            raise RuntimeError("待处理图片中没有可处理图片。")

        worker_path = Path(tempfile.gettempdir()) / "zhipu_paddle_ocr_worker.py"
        worker_path.write_text(OCR_WORKER, encoding="utf-8")
        attempts = local_ocr_attempt_plan(self.selected_ocr_profile(), paddle_state)
        if attempts and attempts[0][1] == "cpu":
            self.post("本地 OCR 将使用 CPU 模式；如要调用 5090，请先让 GPU 版 PaddlePaddle 通过检查。")
            self.root.after(0, self.ocr_profile.set, OCR_PROFILE_LABELS["stable"])
        else:
            self.post("本地 OCR 将按机器状态自动尝试档位：" + " -> ".join(local_ocr_attempt_title(p, d) for p, d in attempts))

        failures: list[str] = []
        for index, (profile, device) in enumerate(attempts, 1):
            attempt_title = local_ocr_attempt_title(profile, device)
            if raw_json.exists():
                try:
                    raw_json.unlink()
                except FileNotFoundError:
                    pass
            self.post(f"本地 OCR 尝试 {index}/{len(attempts)}：{attempt_title}", "section")
            self.post_progress(progress_base, "OCR 识别准备", attempt_title)
            try:
                self._run_paddle_ocr_attempt(
                    worker_path,
                    image_dir,
                    raw_json,
                    images,
                    paddle_state,
                    profile,
                    device,
                    progress_fn,
                )
                self.root.after(0, self.ocr_profile.set, OCR_PROFILE_LABELS.get(profile, OCR_PROFILE_LABELS["stable"]))
                self.post(f"本地 OCR 已使用 {attempt_title} 跑通。", "success")
                return
            except Exception as exc:
                message = str(exc)
                failures.append(f"{attempt_title}: {message}")
                self.post(f"{attempt_title} 未跑通：{message}", "error")
                if index < len(attempts):
                    self.post("继续自动切换下一档；不需要手动改设置。", "section")

        raise SuppressPopupError(
            "本地 OCR 自动换档后仍未跑通，已停止生成。"
            "请查看中间日志；可先用在线智谱 OCR 出表，或安装完整本地 Paddle 环境后重试。"
            "失败记录：" + " | ".join(failures[-3:])
        )

    def _run_paddle_ocr_attempt(
        self,
        worker_path: Path,
        image_dir: Path,
        raw_json: Path,
        images: list[Path],
        paddle_state: dict,
        profile: str,
        device: str,
        progress_fn,
    ) -> None:
        cmd = [
            paddle_state["python"],
            str(worker_path),
            "--input-dir",
            str(image_dir),
            "--output",
            str(raw_json),
            "--profile",
            profile,
            "--device",
            device,
        ]
        self.post("运行命令：" + " ".join(cmd))
        env = os.environ.copy()
        env.setdefault("PYTHONUTF8", "1")
        env.setdefault("PADDLE_PDX_MODEL_SOURCE", "BOS")
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace", env=env)
        assert process.stdout is not None
        output_tail: list[str] = []
        for line in process.stdout:
            clean_line = line.rstrip()
            if clean_line:
                output_tail.append(clean_line)
                output_tail = output_tail[-8:]
            self.post(clean_line)
            match = re.search(r"\[(\d+)/(\d+)\]\s+(\S+)", clean_line)
            if match:
                progress_fn(int(match.group(1)), int(match.group(2)), match.group(3))
        code = process.wait()
        if code != 0:
            detail = "；".join(output_tail[-4:]) if output_tail else ""
            raise RuntimeError(f"OCR 进程退出码 {code}" + (f"；{detail}" if detail else ""))
        status = ocr_json_status(raw_json, images)
        if not status["complete"]:
            raise RuntimeError(
                f"结果不完整：已识别 {status['ok_count']}/{status['image_count']}，"
                f"失败/空文本 {len(status['invalid'])}，缺少 {len(status['missing'])}"
            )

    def parse_and_excel(self) -> None:
        template = self.confirm_excel_template()
        if template is None:
            return
        self.run_background("生成 Excel", lambda: self._parse_and_excel_impl(template))

    def _parse_and_excel_impl(self, template: dict | None = None, progress_base: float = 0.0, progress_span: float = 100.0) -> None:
        work_dir = Path(self.work_dir.get()).resolve()
        raw_json = Path(self.raw_json.get()).resolve()
        if not raw_json.exists():
            raise RuntimeError(f"OCR JSON 不存在：{raw_json}")
        images = input_images(Path(self.image_dir.get()).resolve())
        status = ocr_json_status(raw_json, images)
        if not status["complete"]:
            raise RuntimeError(
                f"OCR JSON 不是完整可用结果，禁止生成空白/错误 Excel。"
                f"已识别 {status['ok_count']}/{status['image_count']}，"
                f"失败/空文本 {len(status['invalid'])}，缺少 {len(status['missing'])}，混入旧图片 {len(status['extra'])}。"
            )
        self.post_progress(progress_base + progress_span * 0.15, "结构化解析", "读取 OCR JSON")
        parsed_json, preview_csv, duplicates_csv, questions_csv, row_count, dup_count, question_count = parse_raw_ocr(raw_json, work_dir)
        self.post_progress(progress_base + progress_span * 0.55, "结构化解析", f"解析 {row_count} 行")
        self.post(f"结构化完成：rows={row_count} duplicates={dup_count} questions={question_count}")
        self.post(f"JSON：{parsed_json}")
        self.post(f"CSV：{preview_csv}")
        self.post(f"疑问：{questions_csv}")
        template = template or {"columns": default_excel_columns(), "image_mode": DEFAULT_EXCEL_IMAGE_MODE}
        self.post_progress(progress_base + progress_span * 0.75, "生成 Excel", "写入表格和原图链接")
        output, excel_rows, excel_dups, image_folder = build_excel(
            parsed_json,
            columns=template.get("columns"),
            image_mode=template.get("image_mode"),
        )
        self.post_progress(progress_base + progress_span, "生成 Excel", f"完成 {excel_rows} 行")
        self.post(f"Excel 完成：{output} rows={excel_rows} duplicates={excel_dups}")
        if image_folder:
            self.post(f"原图文件夹：{image_folder}")

    def run_all(self) -> None:
        template = self.confirm_excel_template()
        if template is None:
            return
        self.run_background("一键出表", lambda: self._run_all_impl(template))

    def _run_all_impl(self, template: dict | None = None) -> None:
        self.post_progress(3, "批次预检", "检查图片和 OCR 条件")
        self._preflight_ocr_for_generation()
        self._run_ocr_impl(progress_base=8, progress_span=72)
        self._parse_and_excel_impl(template, progress_base=82, progress_span=16)

    def open_work_dir(self) -> None:
        self.log_section("打开结果")
        path = Path(self.work_dir.get()).resolve()
        path.mkdir(parents=True, exist_ok=True)
        self.log(f"打开：{path}")
        os.startfile(str(path))

    def refresh_gps(self) -> None:
        self.run_background("刷新 GPS", self._refresh_gps_impl, show_progress=False)

    def _refresh_gps_impl(self) -> None:
        try:
            summary = refresh_gps_summary_from_api()
        except Exception as exc:
            summary = load_gps_summary()
            summary["source"] = "快照"
            summary["status"] = str(exc)
            self.gps_summary = summary
            self.root.after(0, lambda: self._draw_route_panel(self.route_canvas))
            raise RuntimeError(f"{exc}；请在“GPS 设置”里配置授权，或重新确认后台登录是否过期。") from exc
        self.gps_summary = summary
        self.root.after(0, lambda: self._draw_route_panel(self.route_canvas))
        self.post(
            f"GPS 已刷新：总车 {summary['total']}，在线 {summary['online']}，离线 {summary['offline']}，行驶 {summary['running']}，停止 {summary['stopped']}，报警 {summary['alarm']}",
            "success",
        )

    def open_gps_backend(self) -> None:
        self.log_section("GPS 后台")
        self.log("打开后台网页；软件面板要更新，请在“GPS 设置”里保存授权后点“刷新 GPS”。")
        if not gps_api_configured():
            self.log("未保存 GPS 授权，当前右上角显示的是快照。", "error")
        webbrowser.open(GPS_URL)

    def run(self) -> None:
        self.root.mainloop()


def _widget_texts(widget) -> list[str]:
    texts: list[str] = []
    try:
        text = str(widget.cget("text"))
        if text:
            texts.append(text)
    except Exception:
        pass
    for child in widget.winfo_children():
        texts.extend(_widget_texts(child))
    return texts


def startup_smoke_test() -> int:
    started = time.time()
    original_startfile = getattr(os, "startfile", None)
    original_webbrowser_open = webbrowser.open
    if original_startfile is not None:
        os.startfile = lambda _path: None
    webbrowser.open = lambda _url: True
    app = None
    try:
        app = AccountingApp()
        app.root.withdraw()
        app.root.update()
        assert managed_input_dir().exists()
        assert managed_output_dir().exists()
        assert Path(app.work_dir.get()).name == "输出结果"
        assert Path(app.work_dir.get()).parent.name == daily_workspace_name()
        assert app.ocr_engine.get() == OCR_ENGINE_LABELS["glm"]
        assert app.ocr_profile.get() == OCR_PROFILE_LABELS["fast"]
        if getattr(sys, "frozen", False):
            key_value, key_source = initial_api_key()
            assert key_value.strip(), key_source

        texts = set(_widget_texts(app.root))
        for label in ("导入图片", "待处理图片", "扫图片", "一键出表", "打开结果", "OCR 设置", "GPS 设置"):
            assert label in texts, label
        assert "□" in texts
        assert "Excel 设置" not in texts
        assert int(app.progress_log_text.cget("height")) <= 8
        assert "执行日志" not in texts
        assert app.pipeline_canvas.winfo_reqheight() <= 168
        assert float(app.progress_value.get()) == 0.0

        app.scan_batch()
        app.open_work_dir()

        app.open_ocr_settings()
        app.root.update()
        ocr_texts = set(_widget_texts(app.root))
        for label in ("推荐方案", "方案库", "检查环境", "单独 OCR", "装本地 OCR", "官方资料"):
            assert label in ocr_texts, label
        for window in app.root.winfo_children():
            try:
                if isinstance(window, Toplevel) and window.title() == "OCR 设置":
                    window.destroy()
            except Exception:
                pass

        app.refresh_gps = lambda: None
        app.open_gps_settings()
        app.root.update()
        gps_texts = set(_widget_texts(app.root))
        for label in ("配置授权", "刷新 GPS", "GPS 后台"):
            assert label in gps_texts, label
        for window in app.root.winfo_children():
            try:
                if isinstance(window, Toplevel) and window.title() == "GPS 设置":
                    window.destroy()
            except Exception:
                pass

        dialog = ExcelTemplateDialog(app.root, app.excel_columns, app.excel_image_mode)
        app.root.update()
        assert tuple(dialog.tree["columns"]) == ("enabled", "title", "width")
        dialog_texts = set(_widget_texts(dialog.window))
        for label in ("默认生成", "保存并生成", "取消", "启用/停用", "改列名", "改宽度"):
            assert label in dialog_texts, label
        assert "读取字段" not in dialog_texts
        dialog.window.destroy()
        app.root.update()
        print(
            "STARTUP_SMOKE_OK "
            f"buttons=5 input={managed_input_dir().exists()} output={managed_output_dir().exists()} "
            f"sec={time.time() - started:.2f}"
        )
        return 0
    except Exception:
        details = traceback.format_exc()
        try:
            error_path = managed_output_dir() / "启动自检错误.txt"
            error_path.write_text(details, encoding="utf-8")
        except Exception:
            pass
        print("STARTUP_SMOKE_FAILED")
        print(details)
        return 2
    finally:
        webbrowser.open = original_webbrowser_open
        if original_startfile is not None:
            os.startfile = original_startfile
        if app is not None:
            try:
                app.root.destroy()
            except Exception:
                pass


def main() -> int:
    if "--self-test" in sys.argv:
        return self_test()
    if "--startup-smoke-test" in sys.argv:
        return startup_smoke_test()
    app = AccountingApp()
    app.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
