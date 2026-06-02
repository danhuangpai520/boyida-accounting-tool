param(
    [string]$Python = "python"
)

Write-Host "Python: $Python"
& $Python --version
& $Python -m pip --version

& $Python -c @"
import importlib.util
import sys

print("python=" + sys.executable)
for name in ["paddleocr", "paddle"]:
    print(f"{name}={importlib.util.find_spec(name) is not None}")

try:
    import paddle
    print("paddle_version=" + getattr(paddle, "__version__", "unknown"))
    print("cuda_compiled=" + str(paddle.is_compiled_with_cuda()))
    try:
        paddle.utils.run_check()
        print("paddle_run_check=True")
    except Exception as exc:
        print("paddle_run_check=False " + repr(exc))
except Exception as exc:
    print("paddle_import_error=" + repr(exc))

try:
    import paddleocr
    print("paddleocr_version=" + getattr(paddleocr, "__version__", "unknown"))
except Exception as exc:
    print("paddleocr_import_error=" + repr(exc))
"@

Write-Host ""
Write-Host "如果 paddleocr=False 或 paddle=False，请先安装 OCR 依赖。"
Write-Host "如果 cuda_compiled=False，说明当前不是 GPU 版 PaddlePaddle。"

