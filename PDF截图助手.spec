# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ["_wps_pdf_capture_test.py"],
    pathex=[],
    binaries=[],
    datas=[("使用说明.txt", ".")],
    hiddenimports=[
        "win32com",
        "win32com.client",
        "pythoncom",
        "pywintypes",
        "PIL",
        "PIL.Image",
        "fitz",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "keyboard",
        "torch",
        "torchvision",
        "tensorflow",
        "onnxruntime",
        "pandas",
        "numpy",
        "scipy",
        "matplotlib",
        "sympy",
        "pytest",
        "IPython",
        "jupyter",
        "notebook",
        "sklearn",
        "cv2",
        "transformers",
        "setuptools",
        "pkg_resources",
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="PDF截图助手",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="PDF截图助手",
)
