"""
PyInstaller 打包脚本
将 Web 应用打包为独立的 Windows 可执行文件（.exe）
解决 Windows 环境下 PyInstaller 的权限和兼容性问题

主要功能：
1. 环境清理：移除用户级 site-packages 避免权限冲突
2. 源码修补：修复 PyInstaller 在 Windows 下的路径解析问题
3. 修补运行时钩子：解决 backports.tarfile 导入错误
4. 打包配置：设置图标、隐藏控制台、包含静态文件等
5. 依赖管理：显式声明所有隐藏导入的模块

作者：keloder
版本：v0.2
"""

import sys
import os
import site
import shutil

site.ENABLE_USER_SITE = False

USER_SITE = os.path.join(os.environ.get('APPDATA', ''), 'Python', 'Python38', 'site-packages')
if os.path.exists(USER_SITE):
    print(f"Removing user site-packages: {USER_SITE}")
    try:
        shutil.rmtree(USER_SITE, ignore_errors=True)
    except Exception:
        pass
if os.path.exists(USER_SITE):
    try:
        os.rename(USER_SITE, USER_SITE + '_bak')
    except Exception:
        pass

PYINSTALLER_BINDEPEND = os.path.join(
    os.path.dirname(site.__file__),
    'site-packages', 'PyInstaller', 'depend', 'bindepend.py'
)

if os.path.exists(PYINSTALLER_BINDEPEND):
    with open(PYINSTALLER_BINDEPEND, 'r', encoding='utf-8') as f:
        content = f.read()

    original_line = "path = pathlib.Path(path).resolve()"
    patched_line = "path = pathlib.Path(path).absolute()"

    if original_line in content:
        content = content.replace(original_line, patched_line)
        with open(PYINSTALLER_BINDEPEND, 'w', encoding='utf-8') as f:
            f.write(content)
        print("Patched PyInstaller bindepend.py: resolve() -> absolute()")
    else:
        print("PyInstaller bindepend.py already patched or different version")

# 修补 PyInstaller 的 pkg_resources 运行时钩子
# 解决 "cannot import name 'tarfile' from 'backports'" 错误
PYINSTALLER_RTH_PKGRES = os.path.join(
    os.path.dirname(site.__file__),
    'site-packages', 'PyInstaller', 'hooks', 'rthooks', 'pyi_rth_pkgres.py'
)

BACKPORTS_FIX = """import sys as _sys, types as _types
try:
    import backports as _bp
except ImportError:
    _bp = _types.ModuleType('backports')
    _bp.__path__ = []
    _bp.__package__ = 'backports'
    _sys.modules['backports'] = _bp
if 'backports.tarfile' not in _sys.modules:
    _tf = _types.ModuleType('backports.tarfile')
    _tf.__package__ = 'backports'
    _tf.__name__ = 'backports.tarfile'
    _sys.modules['backports.tarfile'] = _tf
    setattr(_bp, 'tarfile', _tf)
del _sys, _types, _bp, _tf
"""

if os.path.exists(PYINSTALLER_RTH_PKGRES):
    with open(PYINSTALLER_RTH_PKGRES, 'r', encoding='utf-8') as f:
        rth_content = f.read()

    if 'backports.tarfile' not in rth_content:
        rth_content = BACKPORTS_FIX + '\n' + rth_content
        with open(PYINSTALLER_RTH_PKGRES, 'w', encoding='utf-8') as f:
            f.write(rth_content)
        print("Patched pyi_rth_pkgres.py: added backports.tarfile fix")
    else:
        print("pyi_rth_pkgres.py already patched")
else:
    print(f"WARNING: pyi_rth_pkgres.py not found at {PYINSTALLER_RTH_PKGRES}")

import PyInstaller.__main__

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ICON_PATH = os.path.join(BASE_DIR, 'icon', 'TH.png')

# 获取命令行参数
mode = 'onedir'  # 默认模式
if len(sys.argv) > 1:
    mode = sys.argv[1]

# 构建PyInstaller参数
args = [
    'pyinstaller',
    f'--{mode}',
    '--name=DocFieldReplacer',
    '--icon=' + ICON_PATH,
    '--windowed',
    '--add-data=templates;templates',
    '--add-data=static;static',

    '--hidden-import=flask',
    '--hidden-import=flask.app',
    '--hidden-import=werkzeug',
    '--hidden-import=werkzeug.utils',
    '--hidden-import=docx',
    '--hidden-import=docx.document',
    '--hidden-import=lxml',
    '--hidden-import=lxml.etree',
    '--hidden-import=pystray',
    '--hidden-import=PIL',
    '--hidden-import=PIL.Image',
    '--hidden-import=PIL.ImageDraw',
    '--hidden-import=pkg_resources',

    'web_server.py',
]

print(f"Building in {mode} mode...")
sys.argv = args
PyInstaller.__main__.run()
