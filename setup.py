"""
для запуска build переходим в cmd в папку с setup.py и запускаем 'python setup.py build'
"""

from cx_Freeze import setup, Executable
import os

executables = [Executable('main2.py',  #главный файл программы
               targetName='VHO_v1.exe', base = 'Win32GUI'),] #задаем имя файла

excludes = ['concurrent','ctypes','curses','distutils',
            'dateutil','html','json','matplotlib','multiprocessing','PIL',
            'pkg_resources', 'pydoc_data','scipy',]
            # ,,]

includes = ['numpy.core._methods','numpy.lib.format','lxml',] #lxml - не зипуется

#архивируем необходимые модули
zip_include_packages = ['collections', 'encodings', 'importlib',
                        'email','logging','http',]

options = {
      'build_exe': {
            'include_msvcr':True, #добавить файл Microsoft Visual C++
            'includes': includes,
            'excludes': excludes,
            'build_exe': 'VHO',
            'zip_include_packages': zip_include_packages,
            #включить в проект рабочие папки программы
            'include_files':[
                  'antens',
                  'Files',
                  'Project',
                  'setup',
                  'UI',
                  'word_shablons',
            ]
      }
}

setup(name = 'VHO',
      version='0.0.1',
      description='VHO',
      executables=executables,
      options=options)
