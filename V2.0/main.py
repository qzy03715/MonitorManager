# MonitorManager V5.0 - 最终版
# 功能:
# - 新增“开机自启动”功能，可通过UI开关控制。
# - 自启动会在开机60秒后静默运行（只显示托盘图标）。
# - 通过修改注册表实现，支持脚本(.py)和打包后(.exe)两种运行方式。
# - 动态识别所有连接的显示器，并在UI上显示分辨率和刷新率。
# - 高级扩展模式：允许用户精确选择两台显示器（主/副）进行扩展。
# - 集成MultiMonitorTool实现可靠切换。
# - 保存/加载配置以恢复原始布局。

import sys
import subprocess
import win32api
import win32con
import time
import logging
import os
import winreg  # 新增：用于操作注册表

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QLabel, QHBoxLayout, QTextEdit,
                             QGridLayout, QMessageBox, QComboBox, QFrame,
                             QCheckBox) # 新增
from PyQt6.QtGui import QFont, QIcon
from PyQt6.QtCore import Qt, pyqtSignal, pyqtSlot
import pystray
from PIL import Image, ImageDraw, ImageFont
import threading

# --- 全局配置 ---
# 配置日志
logging.basicConfig(filename='monitor_manager.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# 工具及配置文件路径
TOOL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'MultiMonitorTool.exe')
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'monitor_config.cfg')

# 注册表项配置
APP_NAME = "MonitorManagerV5"  # 在注册表中的名称
STARTUP_REG_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"

# --- 辅助函数 ---
def create_dummy_icon(width=64, height=64):
    image = Image.new('RGB', (width, height), color='dodgerblue')
    draw = ImageDraw.Draw(image)
    try:
        font = ImageFont.truetype("msyh.ttc", 40)
    except IOError:
        font = ImageFont.load_default()
    draw.text((15, 5), "M", fill="white", font=font)
    return image

# --- 主窗口类 ---
class MonitorApp(QMainWindow):
    # ... (信号定义与之前版本相同)
    toggle_window_signal = pyqtSignal()
    refresh_info_signal = pyqtSignal()
    switch_mode_signal = pyqtSignal(str)
    quit_app_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.monitors = []  
        self.toggle_window_signal.connect(self.toggle_visibility)
        self.refresh_info_signal.connect(self.update_display_info)
        self.switch_mode_signal.connect(self.run_displayswitch_legacy)
        self.quit_app_signal.connect(self.quit_application)
        
        icon_path = 'icon.png'
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("显示器切换工具 V5.0")
        self.setGeometry(300, 300, 750, 650) # 稍微增加高度
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # ... (信息显示、基础控制、单显示器模式、高级扩展模式UI与V4.1完全相同)
        # --- 信息显示区 ---
        self.info_label = QLabel("当前显示器信息:")
        self.info_label.setFont(QFont("Microsoft YaHei", 12))
        self.info_display = QTextEdit()
        self.info_display.setReadOnly(True)
        self.info_display.setFont(QFont("Consolas", 10))
        self.info_display.setFixedHeight(150)
        
        # --- 基础控制区 ---
        static_button_layout = QHBoxLayout()
        self.btn_extend = QPushButton("扩展(所有)")
        self.btn_clone = QPushButton("复制(所有)")
        self.btn_refresh = QPushButton("刷新信息")
        self.btn_save_config = QPushButton("保存当前配置")
        self.btn_load_config = QPushButton("加载保存的配置")
        static_button_layout.addWidget(self.btn_extend)
        static_button_layout.addWidget(self.btn_clone)
        static_button_layout.addWidget(self.btn_refresh)
        static_button_layout.addWidget(self.btn_save_config)
        static_button_layout.addWidget(self.btn_load_config)
        
        # --- 单显示器模式区 ---
        self.dynamic_buttons_label = QLabel("单显示器模式 (仅显示选中的显示器):")
        self.dynamic_buttons_label.setFont(QFont("Microsoft YaHei", 10))
        self.dynamic_buttons_layout = QVBoxLayout()
        
        # --- 高级扩展模式UI ---
        self.advanced_extend_frame = QFrame()
        self.advanced_extend_frame.setFrameShape(QFrame.Shape.StyledPanel)
        advanced_layout = QVBoxLayout(self.advanced_extend_frame)
        advanced_label = QLabel("高级扩展模式 (选择两台，禁用其他):")
        advanced_label.setFont(QFont("Microsoft YaHei", 10, QFont.Weight.Bold))
        selection_layout = QGridLayout()
        selection_layout.addWidget(QLabel("选择主显示器:"), 0, 0)
        self.primary_monitor_combo = QComboBox()
        selection_layout.addWidget(self.primary_monitor_combo, 0, 1)
        selection_layout.addWidget(QLabel("选择扩展副屏:"), 1, 0)
        self.secondary_monitor_combo = QComboBox()
        selection_layout.addWidget(self.secondary_monitor_combo, 1, 1)
        self.btn_apply_extend = QPushButton("应用双屏扩展设置")
        advanced_layout.addWidget(advanced_label)
        advanced_layout.addLayout(selection_layout)
        advanced_layout.addWidget(self.btn_apply_extend)

        # --- 新增：设置区 ---
        settings_layout = QHBoxLayout()
        self.startup_checkbox = QCheckBox("开机后自动启动 (延时60秒静默运行)")
        settings_layout.addWidget(self.startup_checkbox)
        settings_layout.addStretch()

        # --- 主布局 ---
        main_layout.addWidget(self.info_label)
        main_layout.addWidget(self.info_display)
        main_layout.addLayout(static_button_layout)
        main_layout.addWidget(self.create_separator())
        main_layout.addWidget(self.dynamic_buttons_label)
        main_layout.addLayout(self.dynamic_buttons_layout)
        main_layout.addWidget(self.create_separator())
        main_layout.addWidget(self.advanced_extend_frame)
        main_layout.addWidget(self.create_separator())
        main_layout.addLayout(settings_layout) # 添加设置区
        main_layout.addStretch()

        # --- 信号连接 ---
        self.btn_refresh.clicked.connect(self.update_display_info)
        self.btn_extend.clicked.connect(lambda: self.run_displayswitch_legacy('/extend'))
        self.btn_clone.clicked.connect(lambda: self.run_displayswitch_legacy('/clone'))
        self.btn_save_config.clicked.connect(self.save_config)
        self.btn_load_config.clicked.connect(self.load_config)
        self.btn_apply_extend.clicked.connect(self.apply_advanced_extend)
        self.startup_checkbox.stateChanged.connect(self.set_startup_status)
        
        self.update_display_info()
        self.check_startup_status() # 初始化时检查自启动状态

    def create_separator(self):
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        return line

    # ... (从V4.1继承的函数，无需修改)
    def closeEvent(self, event): event.ignore(); self.hide()
    @pyqtSlot()
    def toggle_visibility(self):
        if self.isVisible(): self.hide()
        else: self.show(); self.activateWindow()
    @pyqtSlot()
    def update_display_info(self):
        self.info_display.clear()
        self.monitors = self.get_all_monitors()
        info_text = f"检测到 {len(self.monitors)} 台显示器。\n" + "-" * 60 + "\n"
        for monitor in self.monitors:
            is_primary_str = "(主显示器)" if monitor['is_primary'] else ""
            info_text += f"{monitor['description']} {is_primary_str}\n"
        self.info_display.setText(info_text)
        logging.info(f"--- Display Info Refreshed: Found {len(self.monitors)} monitors ---")
        for monitor in self.monitors: logging.info(f"{monitor['description']}, Primary={monitor['is_primary']}")
        self.update_monitor_controls()
    def get_all_monitors(self):
        monitors_list = []
        try:
            i = 0
            while True:
                device = win32api.EnumDisplayDevices(None, i)
                if device.StateFlags & win32con.DISPLAY_DEVICE_ATTACHED_TO_DESKTOP:
                    settings = win32api.EnumDisplaySettings(device.DeviceName, win32con.ENUM_CURRENT_SETTINGS)
                    is_primary = (settings.Position_x == 0 and settings.Position_y == 0)
                    description = (f"显示器 {i+1}: {settings.PelsWidth}x{settings.PelsHeight} @ {settings.DisplayFrequency}Hz")
                    monitors_list.append({'id': i + 1, 'device': device, 'settings': settings, 'is_primary': is_primary, 'description': description})
                i += 1
        except win32api.error: pass
        return monitors_list
    def update_monitor_controls(self):
        while self.dynamic_buttons_layout.count():
            child = self.dynamic_buttons_layout.takeAt(0)
            if child.widget(): child.widget().deleteLater()
        self.primary_monitor_combo.clear(); self.secondary_monitor_combo.clear()
        for monitor in self.monitors:
            button = QPushButton(f"仅显示 {monitor['description']}")
            button.clicked.connect(lambda checked, num=monitor['id']: self.switch_to_single_display(num))
            self.dynamic_buttons_layout.addWidget(button)
        if len(self.monitors) >= 2:
            self.advanced_extend_frame.show()
            descriptions = [m['description'] for m in self.monitors]
            self.primary_monitor_combo.addItems(descriptions)
            self.secondary_monitor_combo.addItems(descriptions)
            if len(self.monitors) > 1: self.secondary_monitor_combo.setCurrentIndex(1)
        else: self.advanced_extend_frame.hide()
    def apply_advanced_extend(self):
        primary_idx = self.primary_monitor_combo.currentIndex()
        secondary_idx = self.secondary_monitor_combo.currentIndex()
        if primary_idx == secondary_idx:
            QMessageBox.warning(self, "选择错误", "主显示器和扩展副屏不能是同一台显示器！"); return
        primary_monitor_num = self.monitors[primary_idx]['id']; secondary_monitor_num = self.monitors[secondary_idx]['id']
        self.extend_two_monitors(primary_monitor_num, secondary_monitor_num)
    def switch_to_single_display(self, monitor_num):
        logging.info(f"Switching to single display: {monitor_num}")
        self.info_display.append(f"\n[INFO] 切换到仅显示器 {monitor_num}...")
        try:
            cmds = [TOOL_PATH]; [cmds.extend(['/disable', str(m['id'])]) for m in self.monitors if m['id'] != monitor_num]
            cmds.extend(['/enable', str(monitor_num), '/setprimary', str(monitor_num)])
            subprocess.run(cmds, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.info_display.append("[SUCCESS] 切换完成!")
        except Exception as e: self.info_display.append(f"[ERROR] 切换失败: {e}")
        threading.Timer(2.0, self.refresh_info_signal.emit).start()
    def extend_two_monitors(self, primary_num, secondary_num):
        logging.info(f"Extending two monitors: Primary {primary_num}, Secondary {secondary_num}")
        self.info_display.append(f"\n[INFO] 扩展显示器 {primary_num}(主) 和 {secondary_num}(副)...")
        try:
            cmds = [TOOL_PATH]; [cmds.extend(['/disable', str(m['id'])]) for m in self.monitors if m['id'] not in (primary_num, secondary_num)]
            cmds.extend(['/enable', str(primary_num), '/enable', str(secondary_num), '/setprimary', str(primary_num)])
            subprocess.run(cmds, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.run_displayswitch_legacy('/extend')
            self.info_display.append("[SUCCESS] 扩展完成!")
        except Exception as e: self.info_display.append(f"[ERROR] 扩展失败: {e}")
        threading.Timer(2.0, self.refresh_info_signal.emit).start()
    def save_config(self):
        try: subprocess.run([TOOL_PATH, '/SaveConfig', CONFIG_FILE], check=True, creationflags=subprocess.CREATE_NO_WINDOW); self.info_display.append("\n[SUCCESS] 当前配置已保存!")
        except Exception as e: self.info_display.append(f"[ERROR] 保存失败: {e}")
    def load_config(self):
        if not os.path.exists(CONFIG_FILE): QMessageBox.warning(self, "错误", "无已保存的配置文件!"); return
        try: subprocess.run([TOOL_PATH, '/LoadConfig', CONFIG_FILE], check=True, creationflags=subprocess.CREATE_NO_WINDOW); self.info_display.append("\n[SUCCESS] 已加载保存的配置!")
        except Exception as e: self.info_display.append(f"[ERROR] 加载失败: {e}")
        threading.Timer(2.0, self.refresh_info_signal.emit).start()
    @pyqtSlot(str)
    def run_displayswitch_legacy(self, arg):
        """执行Windows内置的DisplaySwitch.exe命令"""
        self.info_display.append(f"\n执行命令: DisplaySwitch.exe {arg}")
        try:
            subprocess.run(['DisplaySwitch.exe', arg], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            threading.Timer(2.0, self.refresh_info_signal.emit).start()
        except Exception as e:
            self.info_display.append(f"命令执行失败: {e}")
            logging.error(f"DisplaySwitch.exe {arg} failed: {e}")

    @pyqtSlot()
    def quit_application(self):
        """退出应用程序"""
        logging.info("Application quit by user.")
        QApplication.instance().quit()

    # --- 新增：注册表操作方法 ---
    @pyqtSlot(int)
    def set_startup_status(self, state):
        """根据复选框状态，启用或禁用开机自启动"""
        if state == Qt.CheckState.Checked.value:
            self.add_to_startup()
        else:
            self.remove_from_startup()

    def add_to_startup(self):
        """添加程序到Windows启动项"""
        try:
            # 获取当前脚本或exe的完整路径
            if getattr(sys, 'frozen', False):
                # 如果是打包后的exe
                app_path = sys.executable
            else:
                # 如果是python脚本
                app_path = os.path.abspath(__file__)
                app_path = f'python.exe "{app_path}"'  # 用python解释器运行
            
            # 构造完整命令：延时60秒后启动，并加上--silent参数
            command = f'cmd /c "timeout /t 60 && {app_path} --silent"'
            
            # 打开注册表
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_REG_KEY, 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, command)
            winreg.CloseKey(key)
            
            self.info_display.append("\n[SUCCESS] 已启用开机自启动 (延时60秒)")
            logging.info("Startup enabled successfully.")
            
        except Exception as e:
            QMessageBox.warning(self, "错误", f"启用自启动失败: {e}")
            self.startup_checkbox.blockSignals(True)
            self.startup_checkbox.setChecked(False)
            self.startup_checkbox.blockSignals(False)
            logging.error(f"Failed to add startup: {e}")

    def remove_from_startup(self):
        """从Windows启动项中移除程序"""
        try:
            # 打开注册表
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_REG_KEY, 0, winreg.KEY_SET_VALUE)
            winreg.DeleteValue(key, APP_NAME)
            winreg.CloseKey(key)
            
            self.info_display.append("\n[SUCCESS] 已禁用开机自启动")
            logging.info("Startup disabled successfully.")
            
        except FileNotFoundError:
            # 注册表项不存在，说明本来就没有启动项
            logging.info("Startup entry not found in registry.")
        except Exception as e:
            QMessageBox.warning(self, "错误", f"禁用自启动失败: {e}")
            self.startup_checkbox.blockSignals(True)
            self.startup_checkbox.setChecked(True)
            self.startup_checkbox.blockSignals(False)
            logging.error(f"Failed to remove startup: {e}")

    def check_startup_status(self):
        """检查程序是否已添加到启动项，初始化复选框状态"""
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_REG_KEY, 0, winreg.KEY_READ)
            value, regtype = winreg.QueryValueEx(key, APP_NAME)
            winreg.CloseKey(key)
            
            # 如果找到了启动项，勾选复选框
            self.startup_checkbox.blockSignals(True)
            self.startup_checkbox.setChecked(True)
            self.startup_checkbox.blockSignals(False)
            logging.info("Startup status: Enabled")
            
        except FileNotFoundError:
            # 注册表项不存在，不勾选复选框
            self.startup_checkbox.blockSignals(True)
            self.startup_checkbox.setChecked(False)
            self.startup_checkbox.blockSignals(False)
            logging.info("Startup status: Disabled")


# --- 托盘图标设置 ---
def setup_tray_icon(main_window):
    """创建系统托盘图标并设置菜单"""
    icon_path = 'icon.png'
    if os.path.exists(icon_path):
        try:
            icon_image = Image.open(icon_path)
        except Exception as e:
            logging.warning(f"Failed to load icon.png: {e}. Using dummy icon.")
            icon_image = create_dummy_icon()
    else:
        logging.info("icon.png not found. Using dummy icon.")
        icon_image = create_dummy_icon()

    def on_show_window(icon, item):
        main_window.toggle_window_signal.emit()

    def on_quit(icon, item):
        icon.stop()
        main_window.quit_app_signal.emit()

    def on_extend(icon, item):
        main_window.switch_mode_signal.emit('/extend')

    def on_clone(icon, item):
        main_window.switch_mode_signal.emit('/clone')

    # 为每个监视器创建菜单项
    def create_single_monitor_items():
        items = []
        for monitor in main_window.monitors:
            def make_handler(monitor_num):
                def handler(icon, item):
                    threading.Thread(target=main_window.switch_to_single_display, args=(monitor_num,)).start()
                return handler
            
            items.append(pystray.MenuItem(
                f'仅 {monitor["description"]}', 
                make_handler(monitor['id'])
            ))
        return items

    # 构建托盘菜单
    menu_items = [
        pystray.MenuItem('显示主窗口', on_show_window),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('扩展模式(所有)', on_extend),
        pystray.MenuItem('复制模式(所有)', on_clone),
        pystray.Menu.SEPARATOR,
    ]
    
    # 添加单显示器菜单项
    if len(main_window.monitors) > 0:
        menu_items.extend(create_single_monitor_items())
        menu_items.append(pystray.Menu.SEPARATOR)
    
    menu_items.append(pystray.MenuItem('退出', on_quit))

    tray_menu = pystray.Menu(*menu_items)
    tray_icon = pystray.Icon("monitor_manager", icon_image, "显示器切换工具 V5.0", tray_menu)

    def run_tray():
        tray_icon.run()

    tray_thread = threading.Thread(target=run_tray, daemon=True)
    tray_thread.start()


# --- 主程序入口 ---
if __name__ == '__main__':
    # 检查MultiMonitorTool是否存在
    if not os.path.exists(TOOL_PATH):
        print("=" * 70)
        print("错误: 缺少 MultiMonitorTool.exe!")
        print("=" * 70)
        print("\n请按照以下步骤操作:")
        print("1. 访问: https://www.nirsoft.net/utils/multi_monitor_tool.html")
        print("2. 下载 MultiMonitorTool (64位版本推荐)")
        print("3. 将 MultiMonitorTool.exe 放在与本脚本相同的目录下")
        print(f"\n当前脚本目录: {os.path.dirname(os.path.abspath(__file__))}")
        print("=" * 70)
        input("\n按任意键退出...")
        sys.exit(1)
    
    logging.info("--- Application Started (V5.0) ---")
    
    # 检查是否为静默启动（从命令行参数判断）
    is_silent = '--silent' in sys.argv
    
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    
    main_window = MonitorApp()
    
    # 设置系统托盘图标
    setup_tray_icon(main_window)
    
    # 如果不是静默启动模式，显示主窗口
    if not is_silent:
        main_window.show()
        logging.info("Main window displayed.")
    else:
        logging.info("Starting in silent mode. Tray icon only.")
    
    sys.exit(app.exec())

