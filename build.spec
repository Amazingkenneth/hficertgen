# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, copy_metadata

# 1. Gather Streamlit data and metadata
datas = [('gen.py', '.')]
datas += collect_data_files('streamlit')
datas += copy_metadata('streamlit')
datas += collect_data_files('docxtpl')

# 2. Force-include modules used in app.py
# Since PyInstaller doesn't scan app.py, we list its imports here.
hidden_imports = [
    # Your app's direct imports
    'streamlit',
    'pandas',
    'docxtpl',
    'pypdf',
    'pypinyin',
    'docx2pdf',
    'openpyxl',  # Required for pandas to read Excel
    
    # Streamlit internal dependencies often missed
    'streamlit.runtime.scriptrunner.magic_funcs',
    'streamlit.runtime.scriptrunner.script_runner',
    'streamlit.web.cli',
    'pydoc',
]

# 3. Excludes (Clean up truly unused heavy libs)
excluded_modules = [
    'matplotlib', 
    'scipy', 
    'ipython', 
    'notebook', 
    'tkinter', 
    'test', 
    'unittest',
]

a = Analysis(
    ['run.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=['./hooks'],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excluded_modules,
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='CertificateGenerator',
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