#!/usr/bin/env bash
set -euo pipefail

# 一键构建打包脚本（macOS / Windows）
# 说明：必须在目标系统上执行对应打包（mac 打 mac 安装包，win 打 win 安装包）。

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

# 激活虚拟环境（如果存在）
if [[ -f ".venv/bin/activate" ]]; then
  source .venv/bin/activate
fi

APP_NAME="MD5Tool"
DIST_DIR="$ROOT_DIR/dist"
BUILD_DIR="$ROOT_DIR/build"

python -m pip install -r requirements.txt

UNAME_OUT="$(uname -s)"

if [[ "$UNAME_OUT" == "Darwin" ]]; then
  echo "[macOS] 构建 .app ..."
  python -m PyInstaller \
    --noconfirm \
    --clean \
    --windowed \
    --name "$APP_NAME" \
    app.py

  APP_PATH="$DIST_DIR/$APP_NAME.app"
  if [[ -d "$APP_PATH" ]]; then
    echo "[macOS] 生成 DMG 安装包 ..."
    DMG_PATH="$DIST_DIR/${APP_NAME}.dmg"
    rm -f "$DMG_PATH"
    hdiutil create -volname "$APP_NAME" -srcfolder "$APP_PATH" -ov -format UDZO "$DMG_PATH"
    echo "[macOS] 完成: $DMG_PATH"
  else
    echo "[macOS] 未找到 .app，构建可能失败。"
  fi

elif [[ "$UNAME_OUT" == MINGW* || "$UNAME_OUT" == MSYS* || "$UNAME_OUT" == CYGWIN* ]]; then
  echo "[Windows] 构建 .exe ..."
  python -m PyInstaller \
    --noconfirm \
    --clean \
    --onefile \
    --noconsole \
    --name "$APP_NAME" \
    app.py
  echo "[Windows] 完成: $DIST_DIR/$APP_NAME.exe"

elif [[ "$UNAME_OUT" == "Linux" ]]; then
  echo "[Linux] 构建可执行文件 ..."
  python -m PyInstaller \
    --noconfirm \
    --clean \
    --onefile \
    --name "$APP_NAME" \
    app.py
  echo "[Linux] 完成: $DIST_DIR/$APP_NAME"

else
  echo "不支持的系统: $UNAME_OUT"
  exit 1
fi

# 可选：清理中间文件
# rm -rf "$BUILD_DIR"
