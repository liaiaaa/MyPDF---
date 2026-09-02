# -*- mode: python ; coding: utf-8 -*-
# 我的专属安全 PDF 工具箱 —— 瘦身打包配置
# 仅打包 Windows x64 的 tkdnd 拖拽库；排除无关重型模块
import os

datas = []
binaries = []
hiddenimports = ['tkinterdnd2']

# tkdnd：只保留 Windows x64 两个版本（Tcl8/9），其余平台全部丢弃
_site = r'E:\software\python\Lib\site-packages'
for sub in ('win-x64-tcl9', 'win-x64'):
    src = os.path.join(_site, 'tkinterdnd2', 'tkdnd', sub)
    dst = os.path.join('tkinterdnd2', 'tkdnd', sub)
    if os.path.isdir(src):
        for fn in os.listdir(src):
            binaries.append((os.path.join(src, fn), dst))

a = Analysis(
    ['pdf_toolbox.py'],
    pathex=[r'D:\C\Desktop\cs\workbuddy\updf\MyPDF工具箱'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 体积大且本工具用不到的重型库
        'numpy', 'pandas', 'matplotlib', 'scipy', 'sklearn', 'PIL.ImageQt',
        'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'IPython', 'notebook',
        'jupyter', 'cppyy', 'defusedxml', 'setuptools', 'pkg_resources',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

# 剔除孤儿 OpenSSL DLL（包内无 _ssl/_hashlib 使用它们，省 ~6MB）
a.binaries = [b for b in a.binaries
              if not (str(b[1]).endswith('libcrypto-3.dll')
                      or str(b[1]).endswith('libssl-3.dll'))]

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='MyPDF工具箱',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
