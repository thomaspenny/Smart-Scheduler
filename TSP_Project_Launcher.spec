# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['project_launcher.py'],
    pathex=[],
    binaries=[],
    datas=[('help.html', '.')],
    hiddenimports=['postcode_distance_app', 'tsp_clustering_app', 'calendar_organizer_app', 'smart_scheduler_app', 'display_preferences', 'pandas', 'requests', 'shapely', 'sklearn', 'scipy', 'matplotlib', 'win32com.client', 'win32timezone', 'cartopy'],
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
    name='TSP_Project_Launcher',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
