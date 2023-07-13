import sys
import os

block_cipher = None

def get_pulp_path():
    import pulp
    return pulp.__path__[0]

path_main = os.path.dirname(os.path.abspath(sys.argv[2]))

a = Analysis(
    ['frontend.py'],
    pathex=[path_main],
    binaries=[],
    datas=[('matrices_data','matrices_data'), ('config.ini', '.'), ('image.ico', '.')],
    hiddenimports=['pulp,sys'],
    hookspath=['./hooks_dir'],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
a.datas += Tree(get_pulp_path(), prefix='pulp', excludes=["*.pyc"])

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='OneInNine',
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
    icon='image.ico',
)
