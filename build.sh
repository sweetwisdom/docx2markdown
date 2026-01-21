#!/bin/bash
# 打包 DOCX to Markdown 转换工具

# 清理之前的构建


# 使用 PyInstaller 打包
uv run pyinstaller -F -i icon.ico -w \
  --add-data "icon.ico:." \
  --add-data "src/docx2markdown:docx2markdown" \
  --hidden-import docx2markdown \
  --hidden-import docx2markdown._docx_to_markdown \
  --hidden-import docx2markdown._markdown_to_docx \
  --hidden-import docx \
  --hidden-import lxml \
  --paths src \
  --name "docx2markdown" \
  main.py

start ".\dist"

echo ""
echo "打包完成！"
