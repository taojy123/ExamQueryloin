# -*- mode: python -*-
a = Analysis([os.path.join(HOMEPATH,'support\\_mountzlib.py'), os.path.join(HOMEPATH,'support\\useUnicode.py'), 'ExamQuery_v2.py'],
             pathex=['E:\\workspace\\GitHub\\ExamQueryloin'])
pyz = PYZ(a.pure)
exe = EXE( pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name=os.path.join('dist', 'ExamQuery_v2.exe'),
          debug=False,
          strip=False,
          upx=True,
          console=True )