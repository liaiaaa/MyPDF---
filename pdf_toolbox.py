# -*- coding: utf-8 -*-
"""
============================================================
 我的专属安全 PDF 工具箱  (MySafePDFToolbox)  v1.3
============================================================
 作者：你本人专属定制（基于 100% 开源库构建）

 安全承诺：
   · 全部功能 100% 本地处理，代码不联网、不上传任何文件
   · 不打包、不加壳、无数字签名伪造 —— 源码完全可见、可审计
   · 不写入注册表、不创建计划任务、不常驻后台

 v1.3 代码模块化拆分：
   · core.py —— 纯本地核心功能（无 GUI）
   · ui.py   —— tkinter 图形界面
   · 本文件  —— 程序入口（保持原启动方式不变）
   · 打包体积、功能与 v1.2 完全一致

 运行方式：
   双击「我的PDF工具箱.lnk」（无黑框），或「启动我的PDF工具箱.bat」。
   也可命令行：python pdf_toolbox.py [可选的pdf文件]
============================================================
"""

# 重新导出核心公共接口（供测试脚本与旧调用兼容）
from core import (
    APP_TITLE, APP_VER, CONFIG_FILE,
    safe_stem, auto_out_path, ensure_dir, fmt_size,
    compress_quality_to_params, load_config, save_config,
    pdf_info, pdf_merge, pdf_split, pdf_compress,
    pdf_encrypt, pdf_decrypt, pdf_to_images, images_to_pdf,
    pdf_to_word, pdf_to_excel, pdf_to_ppt, pdf_extract_images,
)

from ui import main, PdfToolboxApp, FileList, FEATURES


if __name__ == "__main__":
    main()
