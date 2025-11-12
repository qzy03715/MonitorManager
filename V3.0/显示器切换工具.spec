# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['d:\\code\\Project\\MonitorManager\\V3.0\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('d:\\code\\Project\\MonitorManager\\V3.0\\icon.png', '.'), ('d:\\code\\Project\\MonitorManager\\V3.0\\MultiMonitorTool.exe', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='显示器切换工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['d:\\code\\Project\\MonitorManager\\V3.0\\icon.png'],
)
