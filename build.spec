# -*- mode: python -*-

block_cipher = None


a = Analysis(['vndirectPyQt.py'],
             pathex=['/Users/TriDo/General_Code/vndirect'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
a.datas+= [('chromedriver','/Users/TriDo/General_Code/chromedriver', 'DATA')]
a.datas +=[('testDB.db','/Users/TriDo/General_Code/sqlite3/testDB.db','DATA')]             
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          Tree('./image',prefix ='image'),
          a.zipfiles,
          a.datas,
          name='vndirectPyQt',
          debug=False,
          strip=False,
          upx=True,
          console=False )
app = BUNDLE(exe,
             name='vndirectPyQt.app',
             icon='logo.icns',
             bundle_identifier=None)
