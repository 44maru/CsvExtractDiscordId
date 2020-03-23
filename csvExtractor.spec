# -*- mode: python -*-
from kivy_deps import sdl2, glew
block_cipher = None


a = Analysis(['csvExtractor.py'],
             pathex=['.'],
             binaries=[],
             datas=[('ipaexg.ttf', '.'),('csvextractor.kv', '.')],
             hiddenimports=['pkg_resources.py2_warn'],
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
          *[Tree(p) for p in (sdl2.dep_bins + glew.dep_bins)],
          name='csvExtractor',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False )
