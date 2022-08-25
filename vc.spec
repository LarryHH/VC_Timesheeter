# -*- mode: python ; coding: utf-8 -*-


block_cipher = None

added_files = [
         ('assets/bg_left.png', 'assets'),
         ('assets/dnd_area.png', 'assets'),
         ('assets/dnd_icon.png', 'assets'),
         ('assets/dnd_icon.jpg', 'assets'),
         ('assets/dnd_icon.svg', 'assets'),
         ('assets/doubts-button_3.png', 'assets'),
         ('assets/entry.png', 'assets'),
         ('assets/github_logo.png', 'assets'),
         ('assets/logo.png', 'assets'),
         ('assets/logo.svg', 'assets'),
         ('fonts/EngraversGothicBT.ttf', 'fonts')
]


a = Analysis(['src\\dnd.py'],
             pathex=[],
             binaries=[],
             datas=added_files,
             hiddenimports=[],
             hookspath=['.'],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,  
          [],
          name='VC_Timesheeter',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None )
