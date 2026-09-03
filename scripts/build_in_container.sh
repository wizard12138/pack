#!/bin/bash
# ==================================================================
#  财务报销系统 - 银河麒麟 ARM64 构建脚本（start.sh 启动方式）
#
#  改用 start.sh 脚本启动方式，避免 PyInstaller 可执行文件兼容性问题。
#  产物为完整 Python 项目目录，包含 start.sh 启动脚本。
#
#  用法（在 ARM64 机器上）：
#    docker run --rm -v "$PWD":/src -w /src -e PYTHONUTF8=1 rockylinux:8 \
#        bash scripts/build_in_container.sh
#
#  产物：dist/财务报销系统/
# ==================================================================
set -e
cd "$(dirname "$0")/.."

echo "===== 容器环境信息 ====="
uname -m
. /etc/os-release && echo "$PRETTY_NAME"
ldd --version 2>/dev/null | head -1

echo "===== 安装 Python 3.9 与编译工具 ====="
dnf install -y python39 python39-devel gcc make tar gzip
# python39-pip 在部分镜像源中缺失，失败不影响后续 ensurepip 兜底
dnf install -y python39-pip || true

PY=python3.9
"$PY" --version

# 若 pip 未随包提供，用 ensurepip 引导
"$PY" -m pip --version >/dev/null 2>&1 || "$PY" -m ensurepip --upgrade

echo "===== 准备构建虚拟环境 ====="
"$PY" -m venv /build-venv
. /build-venv/bin/activate

pip install --no-cache-dir --upgrade pip
pip install --no-cache-dir -r backend/requirements.txt

echo "===== 清理旧产物 ====="
rm -rf build dist

echo "===== 创建产物目录 ====="
DIST_DIR="dist/财务报销系统"
mkdir -p "$DIST_DIR"

echo "===== 复制源码 ====="
cp -r backend/* "$DIST_DIR/"

echo "===== 复制 Python 虚拟环境 ====="
cp -r /build-venv "$DIST_DIR/.venv"

echo "===== 创建 start.sh 启动脚本 ====="
cat > "$DIST_DIR/start.sh" << 'EOF'
#!/bin/bash
# ==================================================================
#  财务报销系统 - 启动脚本
#
#  用法：./start.sh
#  前提：已安装 Python 3.9+（系统自带或通过包管理器安装）
# ==================================================================
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# 优先使用内置虚拟环境
if [ -d ".venv" ]; then
    echo "使用内置 Python 虚拟环境..."
    source .venv/bin/activate
else
    echo "警告：未找到内置虚拟环境，使用系统 Python"
fi

echo "===== 启动财务报销系统 ====="
echo "访问地址: http://localhost:5000"
echo "按 Ctrl+C 停止服务"
echo ""

PYTHONUTF8=1 python server.py
EOF

chmod +x "$DIST_DIR/start.sh"

echo "===== 创建 README ====="
cat > "$DIST_DIR/README.md" << 'EOF'
# 财务报销系统 - 银河麒麟 ARM64 版本

## 快速启动

```bash
# 1. 解压
tar xzf finance-kylin-arm64.tar.gz
cd 财务报销系统

# 2. 赋予执行权限
chmod +x start.sh

# 3. 启动服务
./start.sh
```

## 系统要求

- 银河麒麟 V10 或 Ubuntu 20.04+ ARM64
- Python 3.9+（系统自带或通过包管理器安装）
- glibc >= 2.28

## 功能说明

- 访问地址：http://localhost:5000
- 健康检查：curl http://127.0.0.1:5000/api/health
- 完全离线运行，无需联网

## 故障排除

### 报错 "GLIBC_2.35 not found"
- 原因：在高 glibc 环境构建
- 解决：使用本构建产物（基于 glibc 2.28）

### 报错 "python: command not found"
- 原因：系统未安装 Python 3.9+
- 解决：`sudo dnf install python39` 或 `sudo apt install python3`

### 中文文件名乱码
- 原因：系统 locale 非 UTF-8
- 解决：设置 `export LANG=zh_CN.UTF-8`
EOF

echo "===== 产物清单 ====="
ls -la "$DIST_DIR/"
echo ""
echo "产物目录: $DIST_DIR"
echo "启动脚本: $DIST_DIR/start.sh"
echo ""
echo "glibc 基线约 2.28，可运行于银河麒麟 ARM64 及 Ubuntu 20.04+ ARM64。"
echo "构建完成。"
