# -*- mode: python -*-

block_cipher = None


a = Analysis(['Easy_PPT_Kits_v19.pyw'],
             pathex=['D:\\PythonX\\app\\new'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=['zmq','jsonschema','certifi','tcl','tk','Pywin','isapi','adodbapi','tornado','jinjia2','FixTk','jupyter','spyder','pygments','http'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
path="C:\\Users\\Administrator\\Anaconda3\\envs\\env27\\Lib\\site-packages"

a.datas = [x for x in a.datas if not os.path.dirname(x[1]).startswith("C:\\Users\\Administrator\\Anaconda3\\envs\\env27\\Lib\\site-packages\\guidata-1.7.6-py2.7.egg\\guidata\\images")]

a.datas = [x for x in a.datas if not os.path.dirname(x[1]).startswith("C:\\Users\\Administrator\\Anaconda3\\envs\\env27\\Lib\\site-packages\\guidata-1.7.6-py2.7.egg\\guidata\\locale")]
a.datas = [x for x in a.datas if not os.path.dirname(x[1]).startswith("C:\\Users\\Administrator\\Anaconda3\\envs\\env27\\Lib\\site-packages\\guidata-1.7.6-py2.7.egg\\guidata\\test")]
a.datas = [x for x in a.datas if not os.path.basename(os.path.dirname(x[1]))== "tzdata"]
a.datas = [x for x in a.datas if not "tzdata" in os.path.dirname(x[1])]
a.datas = [x for x in a.datas if not "ttk" in os.path.dirname(x[1])]
a.datas = [x for x in a.datas if not "msgs" in os.path.dirname(x[1])]
a.datas = [x for x in a.datas if not "test" in os.path.dirname(x[1])]

a.datas = [x for x in a.datas if not os.path.dirname(x[1]).startswith("C:\\Users\\Administrator\\Anaconda3\\envs\\env27\\Lib\\site-packages\\tcl\\tk8.5")]
a.datas = [x for x in a.datas if not os.path.normpath(x[1]).startswith("C:\\Users\\Administrator\\Anaconda3\\envs\\env27\\Lib\\site-packages\\future-0.16.0-py2.7.egg\\future\\backports\\test")]
a.datas = [x for x in a.datas if not os.path.dirname(x[1]).startswith("C:\\Users\\Administrator\\Anaconda3\\envs\\env27\\Lib\\site-packages\\guidata-1.7.6-py2.7.egg\\guidata\\images")]

a.datas = [x for x in a.datas if not os.path.normpath(x[1]).endswith("chm")]
a.datas = [x for x in a.datas if not os.path.normpath(x[1]).endswith("png")]
a.datas = [x for x in a.datas if not os.path.normpath(x[1]).endswith("gif")]
a.datas = [x for x in a.datas if not x[0].startswith("IPython")]
a.datas = [x for x in a.datas if not x[0].startswith("isapi")]
a.datas = [x for x in a.datas if not x[0].startswith("adodbapi")]
a.datas = [x for x in a.datas if not x[0].startswith("Demos")]
a.datas = [x for x in a.datas if not x[0].startswith("test")]
a.datas = [x for x in a.datas if not x[0].startswith("requests")]
a.datas = [x for x in a.datas if not x[0].startswith("guidata\\locale")]

a.datas+=Tree('zx_module\\guidata\\images',prefix='guidata\\images',excludes=[],typecode='DATA')


a.binaries = a.binaries - TOC([('htf5.dll', None, None),
 ('scintilla.dll', None, None),
 ('tcl85.dll', None, None),
 ('tcl85.dll', None, None),
 ('tk85.dll', None, None),
 ('_sqlite3', None, None),
 ('_ssl', None, None),
 ('libifcoremd.dll',None,None),
 ('_speedup', None, None),
 ('LIBEAY32.dll', None, None),
 ('split3.dll', None, None),
 ('PIL._imagingtk.pyd', None, None),

  ('SSLEAY32.dll', None, None),
  ('QtPrintSupport', None, None),
 ('_tkinter', None, None),
])
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='Easy_PPT_Kits_v19',
          debug=False,
          strip=False,
          upx=True,
          console=False )
