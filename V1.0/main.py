import sys
import subprocess
import win32api
import win32con
import os
import threading
import time

# 检查并安装 pywin32
try:
    import win32com.client
except ImportError:
    print("检测到 pywin32 未安装，正在尝试自动安装...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
    print("pywin32 安装完成，请重新运行程序。")
    sys.exit()

from screeninfo import get_monitors
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QLabel, QHBoxLayout, QTextEdit, QCheckBox)
from PyQt6.QtGui import QFont, QIcon
from PyQt6.QtCore import Qt, pyqtSignal, pyqtSlot
import pystray
from PIL import Image, ImageDraw, ImageFont

# --- 核心功能：开机启动项管理者 ---
# 使用 Windows 计划任务实现，这是最现代、最可靠的方式

def resource_path(relative_path):
    """ 获取资源的绝对路径，适用于开发环境和PyInstaller打包后 """
    try:
        # PyInstaller 创建一个临时文件夹，并把路径存储在 _MEIPASS 中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class StartupManager:
    TASK_NAME = "MonitorManagerStartup"
    
    def __init__(self):
        try:
            self.scheduler = win32com.client.Dispatch('Schedule.Service')
            self.scheduler.Connect()
            self.root_folder = self.scheduler.GetFolder('\\')
        except Exception as e:
            print(f"无法连接到任务计划程序: {e}")
            self.scheduler = None

    def is_enabled(self):
        if not self.scheduler: return False
        try:
            self.root_folder.GetTask(self.TASK_NAME)
            return True
        except:
            return False

    def set_startup(self, enable):
        if not self.scheduler: return
        if enable:
            self._create_task()
        else:
            self._delete_task()

    def _create_task(self):
        # 如果任务已存在，先删除
        self._delete_task()

        task_def = self.scheduler.NewTask(0)
        
        # 触发器：用户登录时启动，延迟60秒
        TASK_TRIGGER_LOGON = 9
        trigger = task_def.Triggers.Create(TASK_TRIGGER_LOGON)
        trigger.Id = "LogonTrigger"
        trigger.Delay = "PT60S" # 60秒延迟
        
        # 操作：运行程序
        TASK_ACTION_EXEC = 0
        action = task_def.Actions.Create(TASK_ACTION_EXEC)
        action.ID = "run-monitor-manager"
        
        # 关键：获取当前运行的 exe 路径，或 python 脚本路径
        # 这使得打包后的 exe 文件也能正确设置自启动
        exe_path = sys.executable
        # 如果是通过 python 运行的，路径会是 python.exe，我们需要的是 .py 文件
        # 我们假设打包后，sys.executable 就是我们的 exe
        if "python.exe" in exe_path.lower():
             action.Path = f'"{exe_path}"'
             # __file__ 是当前脚本的绝对路径
             action.Arguments = f'"{os.path.abspath(__file__)}"'
        else:
            action.Path = f'"{exe_path}"'
        
        # 设置工作目录为程序所在目录
        action.WorkingDirectory = f'"{os.path.dirname(os.path.abspath(__file__))}"'

        task_def.RegistrationInfo.Description = "显示器切换工具开机自启动项"
        task_def.RegistrationInfo.Author = "MonitorManager"
        task_def.Settings.Enabled = True
        task_def.Settings.StopIfGoingOnBatteries = False
        task_def.Settings.DisallowStartIfOnBatteries = False
        
        # 注册任务
        TASK_CREATE_OR_UPDATE = 6
        TASK_LOGON_NONE = 0
        self.root_folder.RegisterTaskDefinition(
            self.TASK_NAME,
            task_def,
            TASK_CREATE_OR_UPDATE,
            '', # 无用户
            '', # 无密码
            TASK_LOGON_NONE
        )
        print(f"'{self.TASK_NAME}' 启动项已创建。")

    def _delete_task(self):
        try:
            self.root_folder.DeleteTask(self.TASK_NAME, 0)
            print(f"'{self.TASK_NAME}' 启动项已删除。")
        except Exception:
            pass # 任务不存在时会报错，忽略即可

# 备用图标创建函数 (不变)
def create_dummy_icon(width=64, height=64):
    image = Image.new('RGB', (width, height), color='dodgerblue')
    draw = ImageDraw.Draw(image)
    try: font = ImageFont.truetype("msyh.ttc", 40)
    except IOError: font = ImageFont.load_default()
    draw.text((15, 5), "M", fill="white", font=font)
    return image

# 主窗口类 (有重大修改)
class MonitorApp(QMainWindow):
    toggle_window_signal = pyqtSignal()
    refresh_info_signal = pyqtSignal()
    switch_mode_signal = pyqtSignal(str)
    quit_app_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.startup_manager = StartupManager()
        
        self.toggle_window_signal.connect(self.toggle_visibility)
        self.refresh_info_signal.connect(self.update_display_info)
        self.switch_mode_signal.connect(self.run_displayswitch)
        self.quit_app_signal.connect(QApplication.instance().quit)
        
        icon_path_res = resource_path('icon.png')
        if os.path.exists(icon_path_res): self.setWindowIcon(QIcon(icon_path_res))
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("显示器切换工具")
        self.setGeometry(300, 300, 500, 400) # 稍微加高一点以容纳复选框
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # --- UI 控件 (与之前相同) ---
        self.info_label = QLabel("当前显示器信息:")
        self.info_label.setFont(QFont("Microsoft YaHei", 12))
        self.info_display = QTextEdit()
        self.info_display.setReadOnly(True)
        self.info_display.setFont(QFont("Consolas", 10))
        self.info_display.setFixedHeight(150)
        
        button_layout = QHBoxLayout()
        self.btn_extend = QPushButton("扩展模式")
        self.btn_clone = QPushButton("复制模式")
        self.btn_monitor1 = QPushButton("仅显示器 1")
        self.btn_monitor2 = QPushButton("仅显示器 2")
        button_layout.addWidget(self.btn_extend)
        button_layout.addWidget(self.btn_clone)
        button_layout.addWidget(self.btn_monitor1)
        button_layout.addWidget(self.btn_monitor2)
        
        self.btn_refresh = QPushButton("刷新信息")

        # --- 新增：开机自启动复选框 ---
        self.startup_checkbox = QCheckBox("开机自启动 (60秒后静默启动)")
        self.startup_checkbox.setFont(QFont("Microsoft YaHei", 10))
        # 检查当前状态并设置复选框
        self.startup_checkbox.setChecked(self.startup_manager.is_enabled())
        # 连接信号
        self.startup_checkbox.stateChanged.connect(self.on_startup_checkbox_change)

        # --- 布局 ---
        main_layout.addWidget(self.info_label)
        main_layout.addWidget(self.info_display)
        main_layout.addWidget(self.btn_refresh)
        main_layout.addLayout(button_layout)
        main_layout.addSpacing(20) # 增加一些间距
        main_layout.addWidget(self.startup_checkbox)
        
        # --- 信号连接 (与之前相同) ---
        self.btn_refresh.clicked.connect(self.update_display_info)
        self.btn_extend.clicked.connect(lambda: self.run_displayswitch('/extend'))
        self.btn_clone.clicked.connect(lambda: self.run_displayswitch('/clone'))
        self.btn_monitor1.clicked.connect(lambda: self.run_displayswitch('/internal'))
        self.btn_monitor2.clicked.connect(lambda: self.run_displayswitch('/external'))
        
        self.update_display_info()

    @pyqtSlot(int)
    def on_startup_checkbox_change(self, state):
        enable = state == Qt.CheckState.Checked.value
        self.startup_manager.set_startup(enable)

    def closeEvent(self, event):
        event.ignore()
        self.hide()

    @pyqtSlot()
    def toggle_visibility(self):
        if self.isVisible(): self.hide()
        else: self.show(); self.activateWindow()

    @pyqtSlot()
    def update_display_info(self):
        # 此函数无改动
        self.info_display.clear()
        try:
            monitors, display_mode = get_monitors(), self.get_display_mode()
            info_text = f"检测到 {len(monitors)} 台显示器。\n当前显示模式: {display_mode}\n" + "-" * 30 + "\n"
            for i, m in enumerate(monitors):
                info_text += f"显示器 {i+1}{' (主)' if m.is_primary else ''}: {m.width}x{m.height} at ({m.x},{m.y})\n"
            self.info_display.setText(info_text)
        except Exception as e:
            self.info_display.setText(f"获取信息失败: {e}")

    @pyqtSlot(str)
    def run_displayswitch(self, arg):
        # 此函数无改动
        self.info_display.append(f"\n执行命令: DisplaySwitch.exe {arg}")
        try:
            # 使用 STARTF_USESHOWWINDOW 和 SW_HIDE 来隐藏命令行窗口
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            subprocess.run(f'DisplaySwitch.exe {arg}', check=True, startupinfo=startupinfo, shell=True)
            # 延迟一小段时间后刷新，确保系统已完成切换
            threading.Timer(1.5, self.update_display_info).start()
        except FileNotFoundError:
            self.info_display.append("错误: DisplaySwitch.exe 未找到。请确保程序在系统路径中。")
        except subprocess.CalledProcessError as e:
            self.info_display.append(f"执行命令时出错: {e}")
        except Exception as e:
            self.info_display.append(f"发生未知错误: {e}")
            
    def get_display_mode(self):
        # 此函数无改动
        connected_monitors = len(get_monitors())
        if connected_monitors < 2:
            return "单显示器模式"
        
        # 检查注册表以确定模式 (这是最可靠的方式之一)
        try:
            # 打开注册表项
            key = win32api.RegOpenKeyEx(win32con.HKEY_LOCAL_MACHINE, 
                                      r"SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration", 
                                      0, win32con.KEY_READ)
            # 枚举子项，寻找包含 "Simulated" 的项
            i = 0
            while True:
                try:
                    subkey_name = win32api.RegEnumKey(key, i)
                    if "SIMULATED" in subkey_name.upper(): # 查找第一个活动的配置
                        subkey = win32api.RegOpenKeyEx(key, subkey_name + r"\00", 0, win32con.KEY_READ)
                        # 读取 'Topology' 值
                        # 1 = 内部 (仅显示器1), 2 = 克隆, 4 = 扩展, 8 = 外部 (仅显示器2)
                        topology, _ = win32api.RegQueryValueEx(subkey, "Topology")
                        win32api.RegCloseKey(subkey)
                        win32api.RegCloseKey(key)
                        
                        if topology == 1: return "仅显示器 1 (Internal)"
                        if topology == 2: return "复制 (Clone)"
                        if topology == 4: return "扩展 (Extend)"
                        if topology == 8: return "仅显示器 2 (External)"
                        return "未知"
                    i += 1
                except win32api.error:
                    break # 枚举完成
            win32api.RegCloseKey(key)
        except Exception:
            return "无法读取注册表"
        return "未知"

# --- 系统托盘图标管理 ---
# 这是本次迭代的核心改动之一
class SystemTrayIcon:
    def __init__(self, app_window):
        self.app = app_window
        self.icon = None
        self.last_click_time = 0

        # 加载或创建图标
        icon_path_res = resource_path('icon.png')
        if os.path.exists(icon_path_res):
            self.image = Image.open(icon_path_res)
        else:
            self.image = create_dummy_icon()

    def run(self):
        menu = self._create_menu()
        self.icon = pystray.Icon("MonitorManager", self.image, "显示器切换工具", menu)
        # 关键改动：pystray 的 'run' 会阻塞，所以我们在一个单独的线程中运行它
        threading.Thread(target=self.icon.run, daemon=True).start()

    def _create_menu(self):
        # 关键改动：扁平化菜单结构
        return pystray.Menu(
            # 左键点击行为绑定到 on_clicked 函数
            pystray.MenuItem('显示/隐藏窗口', self.on_clicked, default=True, visible=False),
            # --- 直接展示所有模式 ---
            pystray.MenuItem("扩展模式", lambda: self.app.switch_mode_signal.emit('/extend')),
            pystray.MenuItem("复制模式", lambda: self.app.switch_mode_signal.emit('/clone')),
            pystray.MenuItem("仅显示器 1", lambda: self.app.switch_mode_signal.emit('/internal')),
            pystray.MenuItem("仅显示器 2", lambda: self.app.switch_mode_signal.emit('/external')),
            pystray.Menu.SEPARATOR, # 分隔线
            pystray.MenuItem("刷新信息", lambda: self.app.refresh_info_signal.emit()),
            pystray.Menu.SEPARATOR, # 分隔线
            pystray.MenuItem("退出", self.on_quit)
        )

    def on_clicked(self, icon, item):
        # 关键改动：实现单击与双击的区分
        # 这是一个模拟双击的经典方法
        current_time = time.time()
        # 如果两次点击间隔小于0.3秒，则视为双击
        if current_time - self.last_click_time < 0.3:
            self.app.toggle_window_signal.emit()
            self.last_click_time = 0 # 重置时间，防止三击等行为
        else:
            self.last_click_time = current_time

    def on_quit(self):
        self.icon.stop()
        self.app.quit_app_signal.emit()

# --- 程序主入口 ---
def main():
    # 检查是否已有实例在运行 (可选，但推荐)
    # 此部分可以防止用户意外打开多个程序
    # 如果您不需要此功能，可以注释掉这部分
    try:
        from win32event import CreateMutex
        from win32api import GetLastError
        from winerror import ERROR_ALREADY_EXISTS
        
        handle = CreateMutex(None, 1, "MonitorManager_Mutex_SomeUniqueID")
        if GetLastError() == ERROR_ALREADY_EXISTS:
            print("程序已在运行。")
            return # 直接退出
    except ImportError:
        print("警告: 未找到 pywin32 的部分模块，无法实现单实例检测。")


    app = QApplication(sys.argv)
    main_window = MonitorApp()
    
    # 创建并运行系统托盘图标
    tray = SystemTrayIcon(main_window)
    tray.run()
    
    # 默认不显示主窗口，仅在托盘创建图标
    # main_window.show() 
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
