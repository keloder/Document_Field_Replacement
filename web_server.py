"""
文档字段批量替换工具 - Web 服务端
基于 Flask 的 Web 应用，提供文档批量字段替换的 HTTP API 接口

功能模块：
1. 替换规则管理（增删改查、导入导出、批量操作）
2. 文件上传与处理（支持 .docx/.doc/.wps 格式）
3. 正向替换与反向还原功能
4. 处理结果下载与批量打包

主要技术特性：
- 使用 python-docx 库进行 Word 文档处理
- 支持保留原始格式的文本替换
- 提供 JSON 和 TXT 格式的规则导入导出
- 自动打开浏览器并隐藏开发服务器警告
- 系统托盘图标支持，可通过任务栏关闭服务

作者：keloder
版本：v0.1
"""

import os
import sys
import atexit
import signal
import threading
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import json
import uuid
import shutil
import logging
import ctypes
import platform

# 抑制 Flask 开发服务器的警告信息，提升用户体验
logging.getLogger('werkzeug').setLevel(logging.ERROR)

from replacer import SmartReplacer
from document_handler import DocumentHandler

# 基础路径配置
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Flask 应用初始化
app = Flask(__name__, static_folder=os.path.join(BASE_DIR, 'static'))

# 用户数据目录配置
def get_user_data_dir():
    """获取用户数据目录，兼容PyInstaller打包环境"""
    # 尝试多种方式获取用户目录
    user_dir_candidates = [
        os.path.expanduser("~"),  # 标准用户目录
        os.path.expandvars("%USERPROFILE%"),  # Windows环境变量
        os.path.expandvars("%HOMEPATH%"),  # Windows环境变量
        os.path.join(os.path.expandvars("%HOMEDRIVE%"), os.path.expandvars("%HOMEPATH%")),  # Windows完整路径
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "user_data")  # 备用目录
    ]
    
    for candidate in user_dir_candidates:
        if candidate and os.path.exists(candidate):
            user_data_dir = os.path.join(candidate, "DocFieldReplacer")
            try:
                os.makedirs(user_data_dir, exist_ok=True)
                # 测试目录是否可写
                test_file = os.path.join(user_data_dir, "test_write.tmp")
                with open(test_file, 'w') as f:
                    f.write("test")
                os.remove(test_file)
                return user_data_dir
            except:
                continue
    
    # 如果所有方法都失败，使用当前目录
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "user_data")

def get_config_file():
    """获取配置文件路径，始终优先使用软件所在目录（便携模式）"""
    app_dir = os.path.dirname(os.path.abspath(__file__))
    app_config_file = os.path.join(app_dir, 'config.json')
    
    # 始终优先使用软件目录的配置文件
    if os.path.exists(app_config_file):
        print(f"使用软件目录配置文件: {app_config_file}")
        return app_config_file
    
    # 如果软件目录配置文件不存在，尝试创建
    try:
        os.makedirs(app_dir, exist_ok=True)
        test_file = os.path.join(app_dir, 'test_write.tmp')
        with open(test_file, 'w') as f:
            f.write("test")
        os.remove(test_file)
        
        # 软件目录可写，创建配置文件
        user_data_dir = get_user_data_dir()
        default_config = {
            'data_dir': app_dir,
            'upload_folder': os.path.join(app_dir, 'uploads'),
            'output_folder': os.path.join(app_dir, 'output'),
            'first_run': True
        }
        
        with open(app_config_file, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, ensure_ascii=False, indent=2)
        print(f"创建软件目录配置文件: {app_config_file}")
        print(f"默认数据目录设置为: {app_dir}")
        return app_config_file
    except Exception as e:
        print(f"软件目录不可写: {e}")
        # 回退到用户数据目录
        user_data_dir = get_user_data_dir()
        user_config_file = os.path.join(user_data_dir, 'config.json')
        print(f"使用用户数据目录配置文件: {user_config_file}")
        return user_config_file

def load_user_config():
    """加载用户配置，兼容PyInstaller打包环境"""
    # 先获取配置文件和用户数据目录
    user_data_dir = get_user_data_dir()
    config_file = get_config_file()
    
    # 优先使用软件目录作为默认数据目录（便携模式）
    app_dir = os.path.dirname(os.path.abspath(__file__))
    
    default_config = {
        'upload_folder': os.path.join(app_dir, 'uploads'),
        'output_folder': os.path.join(app_dir, 'output'),
        'data_dir': app_dir,  # 默认使用软件目录
        'first_run': True  # 首次运行标志
    }
    
    # 确保默认目录存在
    os.makedirs(default_config['upload_folder'], exist_ok=True)
    os.makedirs(default_config['output_folder'], exist_ok=True)
    
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                user_config = json.load(f)
                print(f"加载配置文件: {config_file}")
                print(f"配置内容: {user_config}")
                
                # 验证配置完整性
                config_valid = True
                for key in ['upload_folder', 'output_folder', 'data_dir']:
                    if key not in user_config:
                        print(f"配置项缺失: {key}")
                        config_valid = False
                        break
                
                if config_valid:
                    return user_config, config_file
                else:
                    print("配置文件不完整，使用默认配置")
                    # 修复不完整的配置
                    for key in ['upload_folder', 'output_folder', 'data_dir']:
                        if key not in user_config:
                            user_config[key] = default_config[key]
                    
                    # 保存修复后的配置
                    try:
                        with open(config_file, 'w', encoding='utf-8') as f:
                            json.dump(user_config, f, ensure_ascii=False, indent=2)
                        print("配置文件已修复")
                    except Exception as e:
                        print(f"修复配置文件失败: {e}")
                    
                    return user_config, config_file
        except Exception as e:
            print(f"加载配置文件失败: {e}")
    
    # 创建默认配置
    print(f"创建默认配置文件: {config_file}")
    try:
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, ensure_ascii=False, indent=2)
        print("默认配置文件创建成功")
    except Exception as e:
        print(f"创建配置文件失败: {e}")
    
    return default_config, config_file

# 加载配置
user_config, CONFIG_FILE = load_user_config()
USER_DATA_DIR = user_config['data_dir']

def create_shortcuts():
    """创建桌面和开始菜单快捷方式（仅Windows系统）"""
    if platform.system() != 'Windows':
        print("非Windows系统，跳过创建快捷方式")
        return False
    
    try:
        # 尝试导入winshell和pywin32
        try:
            import winshell
            from win32com.client import Dispatch
        except ImportError:
            print("winshell或pywin32未安装，跳过快捷方式创建")
            return False
        
        # 获取当前可执行文件路径
        if getattr(sys, 'frozen', False):
            # 打包环境
            exe_path = sys.executable
        else:
            # 开发环境
            exe_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'web_server.py')
        
        # 获取程序名称
        app_name = "DocFieldReplacer"
        
        # 创建桌面快捷方式
        desktop = winshell.desktop()
        shortcut_path = os.path.join(desktop, f"{app_name}.lnk")
        
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = exe_path
        shortcut.WorkingDirectory = os.path.dirname(exe_path)
        shortcut.IconLocation = exe_path
        shortcut.Description = "文档字段批量替换工具"
        shortcut.save()
        
        print(f"桌面快捷方式已创建: {shortcut_path}")
        
        # 创建开始菜单快捷方式
        start_menu = winshell.start_menu()
        programs_folder = os.path.join(start_menu, "Programs")
        os.makedirs(programs_folder, exist_ok=True)
        
        start_menu_shortcut = os.path.join(programs_folder, f"{app_name}.lnk")
        shortcut2 = shell.CreateShortCut(start_menu_shortcut)
        shortcut2.Targetpath = exe_path
        shortcut2.WorkingDirectory = os.path.dirname(exe_path)
        shortcut2.IconLocation = exe_path
        shortcut2.Description = "文档字段批量替换工具"
        shortcut2.save()
        
        print(f"开始菜单快捷方式已创建: {start_menu_shortcut}")
        return True
        
    except Exception as e:
        print(f"创建快捷方式失败: {e}")
        return False

def create_shortcuts_fallback():
    """备用方案：使用批处理文件创建快捷方式"""
    if getattr(sys, 'frozen', False):
        exe_path = sys.executable
    else:
        exe_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'web_server.py')
    
    app_name = "DocFieldReplacer"
    
    # 创建批处理文件
    batch_content = f"""
@echo off
setlocal

set "TARGET={exe_path}"
set "WORKDIR={os.path.dirname(exe_path)}"
set "LINKNAME={app_name}"

:: 创建桌面快捷方式
powershell -Command "$s=(New-Object -COM WScript.Shell).CreateShortcut('%USERPROFILE%\\Desktop\\%LINKNAME%.lnk');$s.TargetPath='%TARGET%';$s.WorkingDirectory='%WORKDIR%';$s.Save()"

:: 创建开始菜单快捷方式
powershell -Command "$s=(New-Object -COM WScript.Shell).CreateShortcut('%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\%LINKNAME%.lnk');$s.TargetPath='%TARGET%';$s.WorkingDirectory='%WORKDIR%';$s.Save()"

echo 快捷方式创建完成
pause
"""
    
    batch_file = os.path.join(os.path.dirname(exe_path), "create_shortcuts.bat")
    with open(batch_file, 'w', encoding='utf-8') as f:
        f.write(batch_content)
    
    print(f"备用批处理文件已创建: {batch_file}")
    print("请以管理员权限运行此批处理文件来创建快捷方式")
    
    return False  # 返回False表示需要用户手动操作

# 删除旧的load_user_config函数，已被新的逻辑替代

# 应用配置
user_config, CONFIG_FILE = load_user_config()
app.config['UPLOAD_FOLDER'] = user_config['upload_folder']
app.config['OUTPUT_FOLDER'] = user_config['output_folder']
app.config['DATA_DIR'] = user_config['data_dir']
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB 文件大小限制

# 确保必要的目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# 首次启动检测和快捷方式创建
if user_config.get('first_run', True):
    print("首次启动检测，尝试创建快捷方式...")
    
    # 仅在打包环境下创建快捷方式
    if getattr(sys, 'frozen', False):
        success = create_shortcuts()
        if success:
            print("快捷方式创建成功！")
        else:
            print("快捷方式创建失败，可能需要管理员权限")
    else:
        print("开发环境，跳过快捷方式创建")
    
    # 更新配置，标记为已运行
    user_config['first_run'] = False
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(user_config, f, ensure_ascii=False, indent=2)
        print("首次运行标志已更新")
    except Exception as e:
        print(f"更新配置失败: {e}")

# 全局变量：当前替换规则列表
current_rules = []

# 单实例检测 - 使用文件锁机制
LOCK_FILE = os.path.join(BASE_DIR, 'app.lock')

# 退出状态标志
_app_shutting_down = False


def is_port_in_use(port):
    """检查端口是否被占用"""
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(('127.0.0.1', port))
            return False  # 端口可用
        except OSError:
            return True  # 端口被占用


def check_single_instance():
    """检查是否已有实例在运行"""
    global _app_shutting_down
    import socket
    
    # 首先检查端口是否被占用（最准确的方法）
    if is_port_in_use(5000):
        return False  # 端口被占用，说明有实例在运行
    
    try:
        # 尝试创建锁文件
        if os.path.exists(LOCK_FILE):
            try:
                with open(LOCK_FILE, 'r') as f:
                    pid_str = f.read().strip()
                    if not pid_str:
                        raise ValueError("空PID")
                    pid = int(pid_str)
                # 检查进程是否存在
                try:
                    os.kill(pid, 0)
                    return False  # 进程存在，不允许启动
                except (OSError, ProcessLookupError):
                    # 进程不存在，锁文件无效
                    os.remove(LOCK_FILE)
            except (ValueError, OSError, IOError):
                try:
                    os.remove(LOCK_FILE)
                except:
                    pass
        
        # 创建锁文件
        with open(LOCK_FILE, 'w') as f:
            f.write(str(os.getpid()))
        return True
    except Exception:
        return True


def cleanup_lock_file():
    """清理锁文件"""
    try:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)
    except Exception:
        pass


_tray_icon = None
_tray_lang = 'zh'

TRAY_I18N = {
    'zh': {
        'open_browser': '打开浏览器',
        'stop_service': '关闭服务',
        'exit': '退出',
        'tooltip': '文档字段批量替换工具 - v0.1',
    },
    'en': {
        'open_browser': 'Open Browser',
        'stop_service': 'Stop Service',
        'exit': 'Exit',
        'tooltip': 'Doc Field Replacer - v0.1',
    }
}


def _build_tray_menu(lang):
    """根据语言构建托盘菜单"""
    texts = TRAY_I18N.get(lang, TRAY_I18N['zh'])

    def on_click(icon, item):
        item_text = str(item)
        if item_text in (TRAY_I18N['zh']['open_browser'], TRAY_I18N['en']['open_browser']):
            import webbrowser
            webbrowser.open('http://127.0.0.1:5000')
        elif item_text in (TRAY_I18N['zh']['stop_service'], TRAY_I18N['en']['stop_service']):
            # 停止服务但不退出应用
            try:
                # 发送关闭请求
                import requests
                requests.post('http://127.0.0.1:5000/api/shutdown', timeout=1)
            except:
                pass
        elif item_text in (TRAY_I18N['zh']['exit'], TRAY_I18N['en']['exit']):
            # 完全退出应用
            icon.stop()
            cleanup_lock_file()
            force_exit(0)

    import pystray
    return pystray.Menu(
        pystray.MenuItem(texts['open_browser'], on_click),
        pystray.MenuItem(texts['stop_service'], on_click),
        pystray.MenuItem(texts['exit'], on_click),
    )


def update_tray_language(lang):
    """更新托盘菜单的语言"""
    global _tray_icon, _tray_lang
    _tray_lang = lang
    if _tray_icon is None:
        return
    try:
        texts = TRAY_I18N.get(lang, TRAY_I18N['zh'])
        _tray_icon.menu = _build_tray_menu(lang)
        _tray_icon.title = texts['tooltip']
        _tray_icon.update_menu()
    except Exception:
        pass


def create_system_tray():
    """创建系统托盘图标（Windows）"""
    global _tray_icon, _tray_lang
    try:
        import pystray
        from PIL import Image

        icon_path = os.path.join(BASE_DIR, 'static', 'favicon.png')
        if not os.path.exists(icon_path):
            icon_path = os.path.join(BASE_DIR, 'icon', 'TH.png')
        if not os.path.exists(icon_path):
            print("警告：未找到图标文件，无法创建系统托盘图标")
            return None

        icon_image = Image.open(icon_path)
        _tray_lang = 'zh'
        texts = TRAY_I18N['zh']
        menu = _build_tray_menu('zh')

        _tray_icon = pystray.Icon('DocFieldReplacer', icon_image, texts['tooltip'], menu)

        tray_thread = threading.Thread(target=_tray_icon.run, daemon=True)
        tray_thread.start()

        return _tray_icon
    except ImportError:
        print("警告：pystray 库未安装，无法创建系统托盘图标")
        print("请运行：pip install pystray Pillow")
        return None
    except Exception as e:
        print(f"警告：创建系统托盘图标失败：{e}")
        return None


def force_exit(code=0):
    """强制退出应用，不触发任何信号处理器"""
    global _app_shutting_down
    if _app_shutting_down:
        return
    _app_shutting_down = True
    
    print("正在关闭应用...")
    
    cleanup_lock_file()
    
    # 停止系统托盘图标
    if _tray_icon is not None:
        try:
            _tray_icon.stop()
        except:
            pass
    
    # 停止 Flask 服务器
    try:
        from flask import request
        func = request.environ.get('werkzeug.server.shutdown')
        if func is not None:
            func()
    except:
        pass
    
    # 在PyInstaller环境下，需要更彻底的关闭方法
    import sys
    if getattr(sys, 'frozen', False):
        # 打包环境：使用最强制的方法
        print("PyInstaller环境：强制退出进程")
        
        # 首先尝试终止所有子进程
        try:
            import psutil
            current_process = psutil.Process()
            children = current_process.children(recursive=True)
            for child in children:
                try:
                    child.terminate()
                except:
                    pass
        except:
            pass
        
        # 等待一段时间让子进程退出
        import time
        time.sleep(0.5)
        
        # 强制退出方法序列
        try:
            # 方法1：使用ctypes强制退出（最有效）
            import ctypes
            ctypes.windll.kernel32.ExitProcess(code)
        except:
            try:
                # 方法2：使用os._exit直接退出
                os._exit(code)
            except:
                try:
                    # 方法3：使用sys.exit
                    sys.exit(code)
                except:
                    # 方法4：最后手段
                    import os
                    os.kill(os.getpid(), 9)
    else:
        # 开发环境：正常退出
        print("开发环境：正常退出")
        sys.exit(code)


@app.route('/favicon.ico')
def favicon():
    """提供网页图标"""
    return app.send_static_file('favicon.png')


@app.route('/api/tray-language', methods=['POST'])
def tray_language():
    """更新系统托盘菜单语言"""
    data = request.get_json(silent=True) or {}
    lang = data.get('lang', 'zh')
    update_tray_language(lang)
    return jsonify({'success': True})


@app.route('/api/config', methods=['GET'])
def get_config():
    """获取当前配置"""
    return jsonify({
        'success': True,
        'config': {
            'data_dir': app.config['DATA_DIR'],
            'upload_folder': app.config['UPLOAD_FOLDER'],
            'output_folder': app.config['OUTPUT_FOLDER']
        }
    })


@app.route('/api/config', methods=['POST'])
def update_config():
    """更新配置"""
    global user_config, CONFIG_FILE
    data = request.get_json(silent=True) or {}
    
    # 验证新路径
    new_data_dir = data.get('data_dir', '').strip()
    if not new_data_dir:
        return jsonify({'success': False, 'error': '数据目录不能为空'})
    
    try:
        # 创建新目录
        os.makedirs(new_data_dir, exist_ok=True)
        
        # 更新配置
        new_config = {
            'data_dir': new_data_dir,
            'upload_folder': os.path.join(new_data_dir, 'uploads'),
            'output_folder': os.path.join(new_data_dir, 'output'),
            'first_run': False  # 保持首次运行标志
        }
        
        # 确保配置文件保存在软件目录（便携模式）
        app_dir = os.path.dirname(os.path.abspath(__file__))
        app_config_file = os.path.join(app_dir, 'config.json')
        
        # 保存配置到软件目录
        with open(app_config_file, 'w', encoding='utf-8') as f:
            json.dump(new_config, f, ensure_ascii=False, indent=2)
        
        # 更新全局配置路径
        CONFIG_FILE = app_config_file
        
        # 直接使用新配置，不需要重新加载
        # 更新全局配置路径
        CONFIG_FILE = app_config_file
        
        # 更新应用配置和全局配置（注意键名映射）
        app.config['DATA_DIR'] = new_config['data_dir']
        app.config['UPLOAD_FOLDER'] = new_config['upload_folder']
        app.config['OUTPUT_FOLDER'] = new_config['output_folder']
        user_config = new_config
        
        # 确保目录存在
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
        
        print(f"配置已保存到: {CONFIG_FILE}")
        print(f"新配置内容: {new_config}")
        
        return jsonify({'success': True, 'message': '配置已更新', 'config': new_config})
        
    except Exception as e:
        print(f"更新配置失败: {e}")
        return jsonify({'success': False, 'error': f'更新配置失败: {str(e)}'})


@app.route('/api/clear-data', methods=['POST'])
def clear_data():
    """清除数据目录中的所有文件"""
    try:
        # 获取数据目录
        data_dir = user_config.get('data_dir', USER_DATA_DIR)
        upload_folder = user_config.get('upload_folder', os.path.join(data_dir, 'uploads'))
        output_folder = user_config.get('output_folder', os.path.join(data_dir, 'output'))
        
        # 清除上传目录
        if os.path.exists(upload_folder):
            for filename in os.listdir(upload_folder):
                file_path = os.path.join(upload_folder, filename)
                try:
                    if os.path.isfile(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        import shutil
                        shutil.rmtree(file_path)
                except Exception as e:
                    print(f"删除文件失败 {file_path}: {e}")
        
        # 清除输出目录
        if os.path.exists(output_folder):
            for filename in os.listdir(output_folder):
                file_path = os.path.join(output_folder, filename)
                try:
                    if os.path.isfile(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        import shutil
                        shutil.rmtree(file_path)
                except Exception as e:
                    print(f"删除文件失败 {file_path}: {e}")
        
        # 保留配置文件和目录结构
        os.makedirs(upload_folder, exist_ok=True)
        os.makedirs(output_folder, exist_ok=True)
        
        return jsonify({'success': True, 'message': '数据清除成功'})
        
    except Exception as e:
        return jsonify({'success': False, 'error': f'清除数据失败: {str(e)}'})


@app.route('/static/<path:filename>')
def serve_static(filename):
    """提供静态文件服务"""
    return app.send_static_file(filename)


@app.route('/')
def index():
    """主页面路由，返回前端界面"""
    return render_template('index.html')


@app.route('/api/shutdown', methods=['POST'])
def shutdown():
    """
    优雅关闭应用
    
    注意：此 API 会停止 Flask 服务器，导致所有连接断开
    浏览器页面会显示连接错误，这是正常现象
    """
    def shutdown_server():
        import time
        
        # 立即响应前端
        print("收到关闭请求，开始关闭流程...")
        
        # 停止系统托盘
        if _tray_icon is not None:
            try:
                print("停止系统托盘...")
                _tray_icon.stop()
            except Exception as e:
                print(f"停止系统托盘失败: {e}")
        
        # 清理锁文件
        print("清理锁文件...")
        cleanup_lock_file()
        
        # 给前端一点时间接收响应
        time.sleep(0.5)
        
        # 强制退出
        print("强制退出应用...")
        force_exit(0)
    
    # 立即启动关闭线程
    shutdown_thread = threading.Thread(target=shutdown_server, daemon=False)
    shutdown_thread.start()
    
    return jsonify({'success': True, 'message': '应用正在关闭，请稍后手动关闭浏览器窗口'})


# ==================== 替换规则管理 API ====================


@app.route('/api/rules', methods=['POST'])
def add_rule():
    """
    添加或更新替换规则
    
    请求参数：
    - original: 原文本（必填）
    - replacement: 替换文本
    
    返回值：
    - success: 操作是否成功
    - rules: 更新后的规则列表
    - updated/added: 标识是更新还是新增
    """
    data = request.json
    original = data.get('original', '').strip()
    replacement = data.get('replacement', '').strip()

    # 验证输入
    if not original:
        return jsonify({'success': False, 'error': '原文本不能为空'})

    # 检查是否已存在相同原文本的规则，存在则更新
    for rule in current_rules:
        if rule['original'] == original:
            rule['replacement'] = replacement
            return jsonify({'success': True, 'rules': current_rules, 'updated': True})

    # 新增规则
    current_rules.append({'original': original, 'replacement': replacement})
    return jsonify({'success': True, 'rules': current_rules, 'added': True})


@app.route('/api/rules', methods=['GET'])
def get_rules():
    """获取当前所有替换规则"""
    return jsonify({'rules': current_rules})


@app.route('/api/rules', methods=['PUT'])
def update_rule():
    """
    更新替换规则
    
    请求参数：
    - old_original: 原规则的原文本
    - new_original: 新规则的原文本
    - replacement: 替换文本
    """
    data = request.json
    old_original = data.get('old_original', '').strip()
    new_original = data.get('new_original', '').strip()
    new_replacement = data.get('replacement', '').strip()

    # 查找并更新规则
    for rule in current_rules:
        if rule['original'] == old_original:
            if new_original != old_original:
                rule['original'] = new_original
            rule['replacement'] = new_replacement
            return jsonify({'success': True, 'rules': current_rules})

    return jsonify({'success': False, 'error': '规则不存在'})


@app.route('/api/rules', methods=['DELETE'])
def delete_rule():
    """删除指定原文本的规则"""
    data = request.json
    original = data.get('original', '').strip()
    global current_rules
    current_rules = [r for r in current_rules if r['original'] != original]
    return jsonify({'success': True, 'rules': current_rules})


@app.route('/api/rules/clear', methods=['POST'])
def clear_rules():
    """清空所有替换规则"""
    global current_rules
    current_rules = []
    return jsonify({'success': True})


@app.route('/api/rules/sort', methods=['POST'])
def sort_rules():
    """
    对规则进行排序
    
    排序方式：
    - length: 按原文本长度降序（默认）
    - original: 按原文本字母顺序
    - alpha: 按原文本字母顺序（忽略大小写）
    """
    global current_rules
    sort_by = request.json.get('by', 'length')
    if sort_by == 'length':
        current_rules.sort(key=lambda x: len(x['original']), reverse=True)
    elif sort_by == 'original':
        current_rules.sort(key=lambda x: x['original'])
    elif sort_by == 'alpha':
        current_rules.sort(key=lambda x: x['original'].lower())
    return jsonify({'success': True, 'rules': current_rules})


@app.route('/api/rules/batch', methods=['POST'])
def add_rules_batch():
    """批量添加替换规则"""
    data = request.json
    rules_to_add = data.get('rules', [])
    for r in rules_to_add:
        original = r.get('original', '').strip()
        if original:
            exists = False
            for rule in current_rules:
                if rule['original'] == original:
                    rule['replacement'] = r.get('replacement', '')
                    exists = True
                    break
            if not exists:
                current_rules.append({
                    'original': original,
                    'replacement': r.get('replacement', '')
                })
    return jsonify({'success': True, 'rules': current_rules})


@app.route('/api/rules/batch/delete', methods=['POST'])
def delete_rules_batch():
    """批量删除替换规则"""
    data = request.json
    originals = data.get('originals', [])
    global current_rules
    current_rules = [r for r in current_rules if r['original'] not in originals]
    return jsonify({'success': True, 'rules': current_rules})


@app.route('/api/rules/import', methods=['POST'])
def import_rules():
    """
    从文件导入替换规则
    
    支持格式：
    - JSON: 包含规则数组的文件
    - TXT: 每行格式为 "原文本=替换文本"
    """
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '没有文件'})

    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '没有选择文件'})

    filename = secure_filename(file.filename)
    ext = os.path.splitext(filename)[1].lower()
    temp_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(temp_path)

    imported_count = 0

    try:
        if ext == '.json':
            with open(temp_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                items = data if isinstance(data, list) else data.get('rules', [])
                for item in items:
                    if isinstance(item, dict) and item.get('original'):
                        original = item['original'].strip()
                        replacement = item.get('replacement', '')
                        exists = False
                        for rule in current_rules:
                            if rule['original'] == original:
                                rule['replacement'] = replacement
                                exists = True
                                break
                        if not exists:
                            current_rules.append({'original': original, 'replacement': replacement})
                            imported_count += 1
        elif ext == '.txt':
            with open(temp_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and '=' in line:
                        parts = line.split('=', 1)
                        original = parts[0].strip()
                        replacement = parts[1].strip()
                        exists = False
                        for rule in current_rules:
                            if rule['original'] == original:
                                rule['replacement'] = replacement
                                exists = True
                                break
                        if not exists:
                            current_rules.append({'original': original, 'replacement': replacement})
                            imported_count += 1
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})
    finally:
        os.remove(temp_path)

    return jsonify({'success': True, 'imported': imported_count, 'rules': current_rules})


@app.route('/api/rules/export', methods=['GET'])
def export_rules():
    """
    导出替换规则到文件
    
    支持格式：
    - json: JSON 格式（默认）
    - txt: 文本格式，每行 "原文本=替换文本"
    """
    format_type = request.args.get('format', 'json')

    if format_type == 'json':
        export_data = current_rules.copy()
        filename = 'replace_rules.json'
        temp_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        with open(temp_path, 'w', encoding='utf-8') as f:
            json.dump(export_data, f, ensure_ascii=False, indent=2)
    else:
        lines = []
        for rule in current_rules:
            lines.append(f"{rule['original']}={rule['replacement']}")
        filename = 'replace_rules.txt'
        temp_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        with open(temp_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))

    return send_file(temp_path, as_attachment=True, download_name=filename)


# ==================== 文件上传与处理 API ====================

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {'.docx', '.doc', '.wps'}


@app.route('/api/upload', methods=['POST'])
def upload_files():
    """
    上传文档文件
    
    支持多文件上传，文件会被重命名以避免冲突
    返回上传成功的文件信息列表
    """
    if 'files' not in request.files:
        return jsonify({'success': False, 'error': '没有文件'})

    files = request.files.getlist('files')
    saved_files = []

    for file in files:
        if file.filename:
            filename = secure_filename(file.filename)
            ext = os.path.splitext(filename)[1].lower()
            if ext in ALLOWED_EXTENSIONS:
                # 生成唯一文件名避免冲突
                unique_name = f"{uuid.uuid4().hex[:8]}_{filename}"
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_name)
                file.save(filepath)
                saved_files.append({'name': filename, 'path': filepath, 'size': os.path.getsize(filepath)})

    return jsonify({'success': True, 'files': saved_files})


@app.route('/api/process', methods=['POST'])
def process_files():
    """
    处理上传的文档文件
    
    请求参数：
    - mode: 处理模式（forward: 正向替换, reverse: 反向还原）
    - files: 要处理的文件列表
    - rules: 替换规则列表（可选，不传则使用当前规则）
    
    处理流程：
    1. 验证规则和文件
    2. 初始化替换器
    3. 逐个处理文档
    4. 生成输出文件
    """
    data = request.json
    mode = data.get('mode', 'forward')
    file_ids = data.get('files', [])
    client_rules = data.get('rules', [])

    # 使用客户端传入的规则或当前规则
    if client_rules:
        global current_rules
        current_rules = client_rules

    # 验证输入
    if not current_rules:
        return jsonify({'success': False, 'error': '没有替换规则'})

    if not file_ids:
        return jsonify({'success': False, 'error': '没有选择文件'})

    # 初始化替换器并加载规则
    replacer = SmartReplacer()
    for rule in current_rules:
        replacer.add_rule(rule['original'], rule['replacement'])

    # 根据模式选择替换函数
    def replacer_func(text):
        if mode == 'forward':
            result, _ = replacer.replace(text)
        else:
            result, _ = replacer.reverse_replace(text)
        return result

    results = []
    counter = {}

    # 逐个处理文件
    for file_info in file_ids:
        filepath = file_info.get('path')
        original_name = file_info.get('name', 'unknown')
        if filepath and os.path.exists(filepath):
            try:
                # 打开文档并处理
                doc = DocumentHandler.open_document(filepath)
                count = DocumentHandler.process_document(doc, replacer_func)

                # 生成输出文件名（避免重名）
                base_name = os.path.splitext(original_name)[0]
                ext = os.path.splitext(original_name)[1]
                counter[base_name] = counter.get(base_name, 0) + 1
                seq = counter[base_name]
                output_name = f"BM处理-{seq}{ext}"

                # 保存处理后的文档
                output_path_full = os.path.join(app.config['OUTPUT_FOLDER'], output_name)
                DocumentHandler.save_document(doc, output_path_full)

                results.append({
                    'name': original_name,
                    'success': True,
                    'count': count,
                    'output': output_path_full
                })
            except Exception as e:
                results.append({
                    'name': original_name,
                    'success': False,
                    'error': str(e)
                })

    return jsonify({'success': True, 'results': results})


@app.route('/api/clear', methods=['POST'])
def clear_uploads():
    """清空上传文件夹中的临时文件"""
    for f in os.listdir(app.config['UPLOAD_FOLDER']):
        try:
            os.remove(os.path.join(app.config['UPLOAD_FOLDER'], f))
        except:
            pass
    return jsonify({'success': True})


# ==================== 文件下载 API ====================


@app.route('/api/output-files', methods=['GET'])
def list_output_files():
    """列出输出文件夹中的所有文件"""
    output_dir = app.config['OUTPUT_FOLDER']
    files = []
    if os.path.isdir(output_dir):
        for f in os.listdir(output_dir):
            fp = os.path.join(output_dir, f)
            if os.path.isfile(fp):
                files.append({
                    'name': f,
                    'path': fp,
                    'size': os.path.getsize(fp),
                    'mtime': os.path.getmtime(fp)
                })
    # 按修改时间倒序排列
    files.sort(key=lambda x: x['mtime'], reverse=True)
    return jsonify({'success': True, 'files': files})


@app.route('/api/download-output/<path:filename>', methods=['GET'])
def download_output_file(filename):
    """下载单个输出文件"""
    import re
    # 文件名安全过滤
    safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', filename)
    if not safe_name or '..' in safe_name:
        return jsonify({'success': False, 'error': '无效的文件名'}), 400
    filepath = os.path.join(app.config['OUTPUT_FOLDER'], safe_name)
    if not os.path.exists(filepath) or not os.path.isfile(filepath):
        return jsonify({'success': False, 'error': '文件不存在'}), 404
    return send_file(filepath, as_attachment=True, download_name=safe_name)


@app.route('/api/download-all-output', methods=['GET'])
def download_all_output():
    """将所有输出文件打包为 ZIP 下载"""
    import zipfile
    import io

    output_dir = app.config['OUTPUT_FOLDER']
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        added = False
        if os.path.isdir(output_dir):
            for f in os.listdir(output_dir):
                fp = os.path.join(output_dir, f)
                if os.path.isfile(fp):
                    zf.write(fp, f)
                    added = True

    if not added:
        return jsonify({'success': False, 'error': '没有可下载的输出文件'}), 404

    zip_buffer.seek(0)
    return send_file(
        zip_buffer,
        mimetype='application/zip',
        as_attachment=True,
        download_name='处理结果.zip'
    )


def main():
    """
    应用启动入口
    
    功能：
    - 单实例检测，避免重复启动
    - 自动打开浏览器访问应用
    - 启动 Flask 开发服务器
    - 监听所有网络接口（支持局域网访问）
    """
    # 单实例检测
    if not check_single_instance():
        import sys
        sys.stdout.write("\n" + "="*50 + "\n")
        sys.stdout.write("错误：端口 5000 已被占用！\n")
        sys.stdout.write("\n可能原因：\n")
        sys.stdout.write("  1. 应用已在运行中\n")
        sys.stdout.write("  2. 其他程序占用了端口 5000\n")
        sys.stdout.write("\n解决方法：\n")
        sys.stdout.write("  1. 如果应用已在运行，请关闭后重试\n")
        sys.stdout.write("  2. 使用任务管理器结束占用端口 5000 的进程\n")
        sys.stdout.write("  3. 或在命令行运行: netstat -ano | findstr :5000\n")
        sys.stdout.write("     查找占用端口的进程 PID，然后结束它\n")
        sys.stdout.write("="*50 + "\n")
        sys.stdout.flush()
        force_exit(1)
    
    import webbrowser
    import threading

    def open_browser():
        """延迟 1.5 秒后自动打开浏览器"""
        import time
        time.sleep(1.5)
        webbrowser.open('http://127.0.0.1:5000')

    # 在新线程中打开浏览器，避免阻塞主线程
    threading.Thread(target=open_browser, daemon=True).start()
    
    print("文档字段批量替换工具 v0.1")
    print("作者：keloder")
    print("服务地址：http://127.0.0.1:5000")
    print("按 Ctrl+C 或使用界面关闭按钮退出应用")
    
    # 注册退出时的清理
    import atexit
    atexit.register(cleanup_lock_file)
    
    # 创建系统托盘图标
    tray_icon = create_system_tray()
    if tray_icon:
        print("系统托盘图标已创建，可在任务栏找到应用")
    
    # 启动 Flask 应用
    # debug=True 开启调试模式
    # use_reloader=False 禁用自动重载器（避免子进程干扰）
    # host='0.0.0.0' 允许局域网访问
    try:
        app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)
    except KeyboardInterrupt:
        print("\n应用正在关闭...")
        force_exit(0)

if __name__ == '__main__':
    # 在 Windows 上，PyInstaller 打包后需要特殊处理
    import sys
    import os
    
    # 检查是否是 PyInstaller 打包后的环境
    if getattr(sys, 'frozen', False):
        # 打包后的环境
        try:
            main()
        except SystemExit:
            # 正常退出
            pass
        except Exception as e:
            print(f"应用启动失败: {e}")
            input("按回车键退出...")
            sys.exit(1)
    else:
        # 开发环境
        main()
