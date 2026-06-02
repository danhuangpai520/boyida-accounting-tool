param(
    [string]$Python = "python"
)

Write-Host "Python: $Python"
& $Python --version
& $Python -m pip install --upgrade pip
& $Python -m pip install paddleocr xlsxwriter

Write-Host ""
Write-Host "基础依赖安装完成。"
Write-Host "注意：GPU 版 PaddlePaddle 需要按 PaddleOCR / PaddlePaddle 官方当前说明安装。"
Write-Host "50 系显卡不要随意复制旧 CUDA 命令，先看官方 50-series GPU wheel / cu129 / cu130 说明。"

