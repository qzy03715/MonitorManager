# MonitorManager V7.0 - 原生API重构版
# - 移除所有外部进程调用，改用原生 Windows CCD API (通过C++ DLL)
# - 实现原子性显示模式切换，大幅减少卡顿和黑屏时间
# - 保留 V6.2 的所有UI功能和用户体验

import sys
import win32api
import win32con
import time
import logging
import os
import winreg
import json
import ctypes
from ctypes import wintypes

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QLabel, QHBoxLayout, QTextEdit,
                             QGridLayout, QMessageBox, QComboBox, QFrame,
                             QCheckBox)
from PyQt6.QtGui import QFont, QIcon
from PyQt6.QtCore import Qt, pyqtSignal, pyqtSlot
import pystray
from PIL import Image, ImageDraw, ImageFont
import threading

# --- 全局配置 ---
logging.basicConfig(filename='monitor_manager.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# 新增：C++ 核心库路径
DLL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'DisplayCore.dll')
APP_NAME = "MonitorManagerV7"
STARTUP_REG_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"

# 方向常量 (保持不变)
ORIENTATION_LANDSCAPE = 0
ORIENTATION_PORTRAIT = 1
ORIENTATION_LANDSCAPE_FLIPPED = 2
ORIENTATION_PORTRAIT_FLIPPED = 3

ORIENTATION_NAMES = {
    ORIENTATION_LANDSCAPE: "横向",
    ORIENTATION_PORTRAIT: "纵向",
    ORIENTATION_LANDSCAPE_FLIPPED: "横向翻转",
    ORIENTATION_PORTRAIT_FLIPPED: "纵向翻转"
}

# --- C++ DLL 接口定义 ---
class DisplayController:
    def __init__(self, dll_path):
        try:
            self.dll = ctypes.CDLL(dll_path)
            # 定义函数原型
            self.dll.SetSingleDisplay.argtypes = [wintypes.LPCWSTR]
            self.dll.SetSingleDisplay.restype = ctypes.c_long
            
            self.dll.SetExtendDisplays.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.c_int, ctypes.c_int]
            self.dll.SetExtendDisplays.restype = ctypes.c_long
            
            self.dll.SetCloneDisplays.argtypes = []
            self.dll.SetCloneDisplays.restype = ctypes.c_long

            self.dll.SetExtendAllDisplays.argtypes = []
            self.dll.SetExtendAllDisplays.restype = ctypes.c_long

            self.is_valid = True
        except Exception as e:
            logging.error(f"Failed to load DisplayCore.dll: {e}")
            self.is_valid = False

    def set_single_display(self, device_name):
        return self.dll.SetSingleDisplay(device_name)

    def set_extend_displays(self, primary_device, secondary_device, primary_orientation, secondary_orientation):
        return self.dll.SetExtendDisplays(primary_device, secondary_device, primary_orientation, secondary_orientation)
    
    def set_clone_displays(self):
        return self.dll.SetCloneDisplays()

    def set_extend_all_displays(self):
        return self.dll.SetExtendAllDisplays()

# --- 主窗口类 (大部分UI代码保持不变) ---
class MonitorApp(QMainWindow):
    # ... (信号定义部分与 V6.2 相同)
    toggle_window_signal = pyqtSignal()
    refresh_info_signal = pyqtSignal()
    quit_app_signal = pyqtSignal()

    def __init__(self, display_controller):
        super().__init__()
        self.controller = display_controller
        self.monitors = []
        self.orientation_config = self.load_orientation_config()
        self.operation_in_progress = False
        
        # ... (信号连接与 V6.2 相同)
        self.toggle_window_signal.connect(self.toggle_visibility)
        self.refresh_info_signal.connect(self.update_display_info)
        self.quit_app_signal.connect(self.quit_application)
        
        icon_path = 'icon.png'
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            
        self.init_ui()

    def init_ui(self):
        # --- UI 布局代码与 V6.2 完全相同，此处省略以节省篇幅 ---
        # --- 您可以直接从 V6.2 复制 init_ui 方法的全部内容 ---
        self.setWindowTitle("显示器切换工具 V7.0 (原生API版)")
        self.setGeometry(300, 300, 800, 650)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 信息显示区
        self.info_label = QLabel("当前显示器信息:")
        self.info_label.setFont(QFont("Microsoft YaHei", 12))
        self.info_display = QTextEdit()
        self.info_display.setReadOnly(True)
        self.info_display.setFont(QFont("Consolas", 10))
        self.info_display.setFixedHeight(150)
        
        # 基础控制区
        static_button_layout = QHBoxLayout()
        self.btn_extend = QPushButton("扩展(所有)")
        self.btn_clone = QPushButton("复制(所有)")
        self.btn_refresh = QPushButton("刷新信息")
        static_button_layout.addWidget(self.btn_extend)
        static_button_layout.addWidget(self.btn_clone)
        static_button_layout.addWidget(self.btn_refresh)
        
        # 单显示器模式区
        self.dynamic_buttons_label = QLabel("单显示器模式 (仅显示选中的显示器):")
        self.dynamic_buttons_label.setFont(QFont("Microsoft YaHei", 10))
        self.dynamic_buttons_layout = QVBoxLayout()
        
        # 高级扩展模式UI
        self.advanced_extend_frame = QFrame()
        self.advanced_extend_frame.setFrameShape(QFrame.Shape.StyledPanel)
        advanced_layout = QVBoxLayout(self.advanced_extend_frame)
        advanced_label = QLabel("高级扩展模式 (选择两台，禁用其他，设置方向):")
        advanced_label.setFont(QFont("Microsoft YaHei", 10, QFont.Weight.Bold))
        selection_layout = QGridLayout()
        selection_layout.addWidget(QLabel("选择主显示器:"), 0, 0)
        self.primary_monitor_combo = QComboBox()
        selection_layout.addWidget(self.primary_monitor_combo, 0, 1)
        selection_layout.addWidget(QLabel("主显示器方向:"), 0, 2)
        self.primary_orientation_combo = QComboBox()
        self.primary_orientation_combo.addItems(list(ORIENTATION_NAMES.values()))
        selection_layout.addWidget(self.primary_orientation_combo, 0, 3)
        selection_layout.addWidget(QLabel("选择扩展副屏:"), 1, 0)
        self.secondary_monitor_combo = QComboBox()
        selection_layout.addWidget(self.secondary_monitor_combo, 1, 1)
        selection_layout.addWidget(QLabel("副显示器方向:"), 1, 2)
        self.secondary_orientation_combo = QComboBox()
        self.secondary_orientation_combo.addItems(list(ORIENTATION_NAMES.values()))
        selection_layout.addWidget(self.secondary_orientation_combo, 1, 3)
        self.btn_apply_extend = QPushButton("应用双屏扩展设置")
        advanced_layout.addWidget(advanced_label)
        advanced_layout.addLayout(selection_layout)
        advanced_layout.addWidget(self.btn_apply_extend)

        # 设置区
        settings_layout = QHBoxLayout()
        self.startup_checkbox = QCheckBox("开机后自动启动 (延时60秒静默运行)")
        settings_layout.addWidget(self.startup_checkbox)
        settings_layout.addStretch()

        # 主布局
        main_layout.addWidget(self.info_label)
        main_layout.addWidget(self.info_display)
        main_layout.addLayout(static_button_layout)
        main_layout.addWidget(self.create_separator())
        main_layout.addWidget(self.dynamic_buttons_label)
        main_layout.addLayout(self.dynamic_buttons_layout)
        main_layout.addWidget(self.create_separator())
        main_layout.addWidget(self.advanced_extend_frame)
        main_layout.addWidget(self.create_separator())
        main_layout.addLayout(settings_layout)
        main_layout.addStretch()

        # --- 信号连接 (已更新) ---
        self.btn_refresh.clicked.connect(self.update_display_info)
        self.btn_extend.clicked.connect(lambda: self.execute_operation(self.controller.set_extend_all_displays))
        self.btn_clone.clicked.connect(lambda: self.execute_operation(self.controller.set_clone_displays))
        self.btn_apply_extend.clicked.connect(self.apply_advanced_extend)
        
        self.startup_checkbox.stateChanged.connect(self.set_startup_status)
        self.primary_monitor_combo.currentIndexChanged.connect(self.load_primary_orientation)
        self.secondary_monitor_combo.currentIndexChanged.connect(self.load_secondary_orientation)
        
        self.update_display_info()
        self.check_startup_status()

    # --- 核心逻辑重构 ---
    def execute_operation(self, operation_func, *args):
        """重构后的核心执行函数"""
        if self.operation_in_progress:
            self.info_display.append("[WARNING] 操作正在进行中，请稍候...\n")
            return
        
        self.operation_in_progress = True
        self.set_buttons_enabled(False)
        self.info_display.append(f"[INFO] 正在应用新的显示配置...")
        QApplication.processEvents() # 确保UI更新

        start_time = time.time()
        try:
            # 直接调用 C++ DLL 中的函数
            result = operation_func(*args)
            duration = time.time() - start_time

            if result == 0: # ERROR_SUCCESS
                self.info_display.append(f"[SUCCESS] ✓ 配置应用成功! (耗时: {duration:.2f} 秒)")
                # 成功后延时刷新，等待系统完全稳定
                threading.Timer(1.0, self.update_display_info).start()
            else:
                self.info_display.append(f"[ERROR] ❌ 配置应用失败! 错误代码: {result}")
        
        except Exception as e:
            self.info_display.append(f"[FATAL] 调用核心库时发生异常: {e}")
        finally:
            self.operation_in_progress = False
            self.set_buttons_enabled(True)

    def switch_to_single_display(self, device_name):
        self.execute_operation(self.controller.set_single_display, device_name)

    def apply_advanced_extend(self):
        primary_idx = self.primary_monitor_combo.currentIndex()
        secondary_idx = self.secondary_monitor_combo.currentIndex()

        if primary_idx == -1 or secondary_idx == -1: return
        if primary_idx == secondary_idx:
            QMessageBox.warning(self, "选择错误", "主显示器和扩展副屏不能是同一台显示器！")
            return
        
        primary_device = self.monitors[primary_idx]['device_name']
        secondary_device = self.monitors[secondary_idx]['device_name']
        primary_orientation = self.primary_orientation_combo.currentIndex()
        secondary_orientation = self.secondary_orientation_combo.currentIndex()
        
        # 保存方向配置
        primary_monitor_id = self.monitors[primary_idx]['id']
        secondary_monitor_id = self.monitors[secondary_idx]['id']
        self.orientation_config[str(primary_monitor_id)] = primary_orientation
        self.orientation_config[str(secondary_monitor_id)] = secondary_orientation
        self.save_orientation_config()
        
        self.execute_operation(self.controller.set_extend_displays, primary_device, secondary_device, primary_orientation, secondary_orientation)

    @pyqtSlot()
    def update_display_info(self):
        # 此函数与 V6.2 基本相同，仅更新描述文本
        self.info_display.clear()
        self.monitors = self.get_all_monitors()
        info_text = f"检测到 {len(self.monitors)} 台显示器。\n" + "-" * 60 + "\n"
        for monitor in self.monitors:
            is_primary_str = "(主显示器)" if monitor['is_primary'] else ""
            info_text += f"{monitor['description']} {is_primary_str}\n"
        self.info_display.setText(info_text)
        self.update_monitor_controls()

    def update_monitor_controls(self):
        # 仅更新 connect 的调用方式
        while self.dynamic_buttons_layout.count():
            child = self.dynamic_buttons_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        
        self.primary_monitor_combo.clear()
        self.secondary_monitor_combo.clear()

        for monitor in self.monitors:
            button = QPushButton(f"仅显示 {monitor['description']}")
            # 注意这里的变化：传递 device_name
            button.clicked.connect(
                lambda checked, name=monitor['device_name']: self.switch_to_single_display(name)
            )
            self.dynamic_buttons_layout.addWidget(button)

        if len(self.monitors) >= 2:
            self.advanced_extend_frame.show()
            descriptions = [m['description'] for m in self.monitors]
            self.primary_monitor_combo.addItems(descriptions)
            self.secondary_monitor_combo.addItems(descriptions)
            if len(self.monitors) > 1:
                self.secondary_monitor_combo.setCurrentIndex(1)
        else:
            self.advanced_extend_frame.hide()

    def get_all_monitors(self):
        # 此函数与 V6.2 完全相同
        monitors_list = []
        try:
            i = 0
            while True:
                device = win32api.EnumDisplayDevices(None, i)
                if device.StateFlags & win32con.DISPLAY_DEVICE_ATTACHED_TO_DESKTOP:
                    settings = win32api.EnumDisplaySettings(device.DeviceName, win32con.ENUM_CURRENT_SETTINGS)
                    is_primary = (settings.Position_x == 0 and settings.Position_y == 0)
                    description = (f"显示器 {i+1} ({device.DeviceString}): "
                                   f"{settings.PelsWidth}x{settings.PelsHeight} @ "
                                   f"{settings.DisplayFrequency}Hz")
                    monitors_list.append({
                        'id': i + 1, 'device_name': device.DeviceName,
                        'is_primary': is_primary, 'description': description
                    })
                i += 1
        except win32api.error: pass
        return monitors_list

    # --- 以下是 V6.2 中未改变的辅助函数 ---
    # create_separator, closeEvent, set_buttons_enabled, toggle_visibility,
    # load_orientation_config, save_orientation_config, load_primary_orientation,
    # load_secondary_orientation, set_startup_status, add_to_startup, 
    # remove_from_startup, check_startup_status, quit_application
    def create_separator(self):
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        return line

    def closeEvent(self, event):
        event.ignore()
        self.hide()

    def set_buttons_enabled(self, enabled):
        self.btn_extend.setEnabled(enabled)
        self.btn_clone.setEnabled(enabled)
        self.btn_refresh.setEnabled(enabled)
        self.btn_apply_extend.setEnabled(enabled)
        self.primary_monitor_combo.setEnabled(enabled)
        self.secondary_monitor_combo.setEnabled(enabled)
        self.primary_orientation_combo.setEnabled(enabled)
        self.secondary_orientation_combo.setEnabled(enabled)
        for i in range(self.dynamic_buttons_layout.count()):
            widget = self.dynamic_buttons_layout.itemAt(i).widget()
            if widget: widget.setEnabled(enabled)

    @pyqtSlot()
    def toggle_visibility(self):
        if self.isVisible(): self.hide()
        else: self.show(); self.activateWindow()

    def load_orientation_config(self):
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'monitor_orientation_config.json')
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f: return json.load(f)
            except: return {}
        return {}

    def save_orientation_config(self):
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'monitor_orientation_config.json')
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(self.orientation_config, f, indent=4)
        except Exception as e: logging.error(f"Failed to save orientation config: {e}")

    def load_primary_orientation(self):
        if self.primary_monitor_combo.currentIndex() >= 0:
            monitor_id = self.monitors[self.primary_monitor_combo.currentIndex()]['id']
            orientation = self.orientation_config.get(str(monitor_id), ORIENTATION_LANDSCAPE)
            self.primary_orientation_combo.setCurrentIndex(orientation)

    def load_secondary_orientation(self):
        if self.secondary_monitor_combo.currentIndex() >= 0:
            monitor_id = self.monitors[self.secondary_monitor_combo.currentIndex()]['id']
            orientation = self.orientation_config.get(str(monitor_id), ORIENTATION_LANDSCAPE)
            self.secondary_orientation_combo.setCurrentIndex(orientation)

    @pyqtSlot()
    def quit_application(self):
        QApplication.instance().quit()
    
    @pyqtSlot(int)
    def set_startup_status(self, state):
        if state == Qt.CheckState.Checked.value: self.add_to_startup()
        else: self.remove_from_startup()

    def add_to_startup(self):
        try:
            app_path = sys.executable if getattr(sys, 'frozen', False) else f'"{sys.executable}" "{os.path.abspath(__file__)}"'
            command = f'{app_path} --silent'
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_REG_KEY, 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, command)
            winreg.CloseKey(key)
        except Exception as e: QMessageBox.warning(self, "错误", f"启用自启动失败: {e}")

    def remove_from_startup(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_REG_KEY, 0, winreg.KEY_SET_VALUE)
            winreg.DeleteValue(key, APP_NAME)
            winreg.CloseKey(key)
        except: pass

    def check_startup_status(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_REG_KEY, 0, winreg.KEY_READ)
            winreg.QueryValueEx(key, APP_NAME)
            winreg.CloseKey(key)
            self.startup_checkbox.setChecked(True)
        except: self.startup_checkbox.setChecked(False)

# --- 托盘图标设置 (与 V6.2 几乎相同，仅更新托盘菜单的调用) ---
def setup_tray_icon(main_window):
    # ... create_dummy_icon 函数与 V6.2 相同
    def create_dummy_icon(width=64, height=64):
        image = Image.new('RGB', (width, height), color='dodgerblue')
        draw = ImageDraw.Draw(image)
        try: font = ImageFont.truetype("msyh.ttc", 40)
        except IOError: font = ImageFont.load_default()
        draw.text((15, 5), "M", fill="white", font=font)
        return image
    
    icon_path = 'icon.png'
    icon_image = Image.open(icon_path) if os.path.exists(icon_path) else create_dummy_icon()

    def on_show_window(icon, item): main_window.toggle_window_signal.emit()
    def on_quit(icon, item): icon.stop(); main_window.quit_app_signal.emit()
    def on_extend(icon, item): main_window.execute_operation(main_window.controller.set_extend_all_displays)
    def on_clone(icon, item): main_window.execute_operation(main_window.controller.set_clone_displays)
    
    def create_single_monitor_items():
        items = []
        for monitor in main_window.monitors:
            handler = lambda name=monitor['device_name']: main_window.switch_to_single_display(name)
            items.append(pystray.MenuItem(f'仅 {monitor["description"]}', handler))
        return items
    
    menu = pystray.Menu(
        pystray.MenuItem('显示主窗口', on_show_window, default=True),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('扩展模式(所有)', on_extend),
        pystray.MenuItem('复制模式(所有)', on_clone),
        pystray.Menu.SEPARATOR,
        *create_single_monitor_items(),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('退出', on_quit)
    )
    
    tray_icon = pystray.Icon("monitor_manager", icon_image, "显示器切换工具 V7.0", menu)
    threading.Thread(target=tray_icon.run, daemon=True).start()


# --- 主程序入口 ---
if __name__ == '__main__':
    # 步骤 1: 立即创建 QApplication 实例。这是所有UI操作的前提。
    app = QApplication(sys.argv)
    
    # 步骤 2: 现在可以安全地进行文件检查和DLL加载，因为如果出错，QMessageBox 可以被正常创建。
    if not os.path.exists(DLL_PATH):
        # 此处调用 QMessageBox 是安全的
        QMessageBox.critical(None, "致命错误", f"核心库 DisplayCore.dll 未找到!\n\n请确保 DisplayCore.dll 与主程序在同一目录下。")
        sys.exit(1)
    display_controller = DisplayController(DLL_PATH)
    if not display_controller.is_valid:
        # 此处调用 QMessageBox 也是安全的
        QMessageBox.critical(None, "致命错误", f"加载 DisplayCore.dll 失败!\n\n请检查DLL是否有效或缺少依赖项(如VC++运行时库)。")
        sys.exit(1)
        
    # 步骤 3: QApplication 已经存在，现在继续执行程序的其余部分。
    app.setQuitOnLastWindowClosed(False)
    
    main_window = MonitorApp(display_controller)
    setup_tray_icon(main_window)
    
    if '--silent' not in sys.argv:
        main_window.show()
    
    # 启动事件循环
    sys.exit(app.exec())