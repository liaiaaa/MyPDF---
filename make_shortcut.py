# -*- coding: utf-8 -*-
"""
创建桌面快捷方式（无需任何参数）。
所有路径从脚本自身位置推导，避免 cmd 传参导致中文丢失的问题。
供「部署到新电脑.bat」调用。
"""
import os
import sys
import base64
import subprocess
import ctypes

TOOL_DIR = os.path.dirname(os.path.abspath(__file__))

# 1) pythonw：与本解释器同目录
pydir = os.path.dirname(sys.executable)
pythonw = os.path.join(pydir, "pythonw.exe")
if not os.path.exists(pythonw):
    pythonw = sys.executable

# 2) 真实桌面路径（兼容桌面被重定向 / OneDrive 的情况）
buf = ctypes.create_unicode_buffer(520)
ctypes.windll.shell32.SHGetFolderPathW(None, 0x0010, None, 0, buf)  # CSIDL_DESKTOPDIRECTORY
desktop = buf.value
lnk_path = os.path.join(desktop, "我的PDF工具箱.lnk")

# 3) 构造 PowerShell 脚本（含中文，走 -EncodedCommand 保证 Unicode 无损）
def ps_str(s):
    return s.replace("'", "''")

ps = (
    "$ws=New-Object -ComObject WScript.Shell;"
    "$s=$ws.CreateShortcut('" + ps_str(lnk_path) + "');"
    "$s.TargetPath='" + ps_str(pythonw) + "';"
    "$s.Arguments='pdf_toolbox.py';"
    "$s.WorkingDirectory='" + ps_str(TOOL_DIR) + "';"
    "$s.Save()"
)
encoded = base64.b64encode(ps.encode("utf-16-le")).decode("ascii")
subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                "-EncodedCommand", encoded], check=True)

print("已创建桌面快捷方式：", lnk_path)
print("目标：", pythonw)
