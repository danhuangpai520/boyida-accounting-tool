param(
    [string]$Python = "python"
)

Write-Host "Python: $Python"
& $Python --version
& $Python -m pip install --upgrade --prefer-binary --timeout 120 pip setuptools wheel
& $Python -m pip install --upgrade --prefer-binary --timeout 120 paddleocr xlsxwriter

Write-Host ""
Write-Host "基础依赖安装完成。"
Write-Host "注意：本脚本只装基础依赖。CPU/GPU 版 PaddlePaddle 请优先用软件里的“装本地 OCR”。"
Write-Host "50 系显卡不要随意复制旧 CUDA 命令，先看官方 50-series GPU wheel / cu129 / cu130 说明。"
