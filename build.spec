# build.spec
block_cipher = None

a = Analysis(
    ['app_qt_ui_1.py'],  # Замените на имя вашего главного файла
    pathex=[],
    binaries=[],
    datas=[
        ('credentials.json', '.'),  # Включите необходимые файлы
        ('.env', '.')
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='EquipmentReportGenerator',  # Имя вашего приложения
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Установите True, если хотите видеть консоль
    icon='icon.ico',  # Добавьте иконку, если нужно
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)