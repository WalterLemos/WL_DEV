# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['API_CDA.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tmp'],
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
    name='API_CDA',
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

 solver = recaptchaV2Proxyless()
    #solver = recaptchaV2Proxyon()
    solver.set_verbose(1)
    solver.set_key(chave_api)
    solver.set_website_url(link)
    solver.set_website_key(chave_captcha)
    resposta = solver.solve_and_return_solution()