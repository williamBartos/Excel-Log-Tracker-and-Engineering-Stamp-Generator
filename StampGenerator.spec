# -*- mode: python -*-

block_cipher = None


a = Analysis(['StampGenerator.py'],
             pathex=['R:\\Municipal\\Utility\\Working\\01_General\\Programs\\Python\\Stamp & Transmittal Generator\\Test Environment'],
             binaries=None,
             datas=None,
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='StampGenerator',
          debug=False,
          strip=False,
          upx=True,
          console=False )
