#!/bin/bash
# Web 版本启动脚本

cd "$(dirname "$0")"

echo "=================================="
echo "  员工工作量统计系统 - Web 版本"
echo "=================================="
echo ""

# 检查 Python
if command -v python3 &> /dev/null; then
    echo "✓ 启动本地服务器..."
    echo "✓ 访问地址: http://localhost:8000"
    echo ""
    echo "按 Ctrl+C 停止服务器"
    echo ""
    python3 -m http.server 8000
elif command -v python &> /dev/null; then
    echo "✓ 启动本地服务器..."
    echo "✓ 访问地址: http://localhost:8000"
    echo ""
    echo "按 Ctrl+C 停止服务器"
    echo ""
    python -m http.server 8000
else
    echo "❌ 未找到 Python，直接打开 HTML 文件..."
    open index.html
fi

