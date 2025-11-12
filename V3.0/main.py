# MonitorManager V6.2 - 防并发 + 双击托盘 + 自动刷新版
# 新增功能:
# - 防止快速多次点击导致的高并发
# - 双击托盘图标显示主窗口
# - 每次切换完成后自动刷新信息

import sys
import subprocess
import win32api
import win32con
import time
import logging
import os
import winreg
import json

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

TOOL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'MultiMonitorTool.exe')
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'monitor_config.cfg')
ORIENTATION_CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'monitor_orientation_config.json')

APP_NAME = "MonitorManagerV6"
STARTUP_REG_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"

# 显示方向常量
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
    toggle_window_signal = pyqtSignal()
    refresh_info_signal = pyqtSignal()
    switch_mode_signal = pyqtSignal(str)
    quit_app_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.monitors = []
        self.monitor_native_resolutions = {}
        self.orientation_config = self.load_orientation_config()
        self.operation_in_progress = False  # 新增：操作锁
        
        self.toggle_window_signal.connect(self.toggle_visibility)
        self.refresh_info_signal.connect(self.update_display_info)
        self.switch_mode_signal.connect(self.handle_switch_mode_signal)
        self.quit_app_signal.connect(self.quit_application)
        
        icon_path = 'icon.png'
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("显示器切换工具 V6.2")
        self.setGeometry(300, 300, 800, 700)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
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

        # --- 设置区 ---
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
        main_layout.addLayout(settings_layout)
        main_layout.addStretch()

        # --- 信号连接 ---
        self.btn_refresh.clicked.connect(self.update_display_info)
        self.btn_extend.clicked.connect(lambda: self.execute_with_lock(self.run_displayswitch_legacy, '/extend'))
        self.btn_clone.clicked.connect(lambda: self.execute_with_lock(self.run_displayswitch_legacy, '/clone'))
        self.btn_save_config.clicked.connect(lambda: self.execute_with_lock(self.save_config))
        self.btn_load_config.clicked.connect(lambda: self.execute_with_lock(self.load_config))
        self.btn_apply_extend.clicked.connect(self.apply_advanced_extend_with_lock)
        self.startup_checkbox.stateChanged.connect(self.set_startup_status)
        
        self.primary_monitor_combo.currentIndexChanged.connect(self.load_primary_orientation)
        self.secondary_monitor_combo.currentIndexChanged.connect(self.load_secondary_orientation)
        
        self.update_display_info()
        self.check_startup_status()

    def create_separator(self):
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        return line

    def closeEvent(self, event):
        event.ignore()
        self.hide()

    # --- 新增：按钮状态管理 ---
    def set_buttons_enabled(self, enabled):
        """启用或禁用所有操作按钮"""
        self.btn_extend.setEnabled(enabled)
        self.btn_clone.setEnabled(enabled)
        self.btn_refresh.setEnabled(enabled)
        self.btn_save_config.setEnabled(enabled)
        self.btn_load_config.setEnabled(enabled)
        self.btn_apply_extend.setEnabled(enabled)
        self.primary_monitor_combo.setEnabled(enabled)
        self.secondary_monitor_combo.setEnabled(enabled)
        self.primary_orientation_combo.setEnabled(enabled)
        self.secondary_orientation_combo.setEnabled(enabled)
        
        # 禁用动态生成的单屏按钮
        for i in range(self.dynamic_buttons_layout.count()):
            widget = self.dynamic_buttons_layout.itemAt(i).widget()
            if widget:
                widget.setEnabled(enabled)

    def execute_with_lock(self, func, *args, **kwargs):
        """带锁执行函数，防止并发，完成后自动刷新"""
        if self.operation_in_progress:
            self.info_display.append("[WARNING] 操作正在进行中，请稍候...\n")
            return
        
        self.operation_in_progress = True
        self.set_buttons_enabled(False)
        self.info_display.append("[INFO] 操作开始...")
        
        try:
            result = func(*args, **kwargs)
            # 操作完成后自动刷新信息
            time.sleep(0.5)
            self.update_display_info()
            self.info_display.append("[INFO] ✓ 操作完成，信息已刷新\n")
            return result
        except Exception as e:
            self.info_display.append(f"[ERROR] 操作异常: {str(e)}")
            import traceback
            self.info_display.append(f"[DEBUG] {traceback.format_exc()}\n")
        finally:
            self.operation_in_progress = False
            self.set_buttons_enabled(True)

    @pyqtSlot()
    def toggle_visibility(self):
        if self.isVisible():
            self.hide()
        else:
            self.show()
            self.activateWindow()

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
        for monitor in self.monitors:
            logging.info(f"{monitor['description']}, Primary={monitor['is_primary']}")
        
        self.update_monitor_controls()

    def get_all_monitors(self):
        """获取所有显示器并保存其原生分辨率"""
        monitors_list = []
        self.monitor_native_resolutions.clear()
        
        try:
            i = 0
            while True:
                device = win32api.EnumDisplayDevices(None, i)
                if device.StateFlags & win32con.DISPLAY_DEVICE_ATTACHED_TO_DESKTOP:
                    settings = win32api.EnumDisplaySettings(device.DeviceName, win32con.ENUM_CURRENT_SETTINGS)
                    is_primary = (settings.Position_x == 0 and settings.Position_y == 0)
                    
                    # 保存原生分辨率（根据当前方向）
                    current_orientation = settings.DisplayOrientation
                    if current_orientation in [ORIENTATION_PORTRAIT, ORIENTATION_PORTRAIT_FLIPPED]:
                        native_width = settings.PelsHeight
                        native_height = settings.PelsWidth
                    else:
                        native_width = settings.PelsWidth
                        native_height = settings.PelsHeight
                    
                    self.monitor_native_resolutions[device.DeviceName] = {
                        'width': native_width,
                        'height': native_height
                    }
                    
                    description = (f"显示器 {i+1}: "
                                   f"{settings.PelsWidth}x{settings.PelsHeight} @ "
                                   f"{settings.DisplayFrequency}Hz")

                    monitors_list.append({
                        'id': i + 1,
                        'device': device,
                        'device_name': device.DeviceName,
                        'settings': settings, 
                        'is_primary': is_primary,
                        'description': description
                    })
                i += 1
        except win32api.error:
            pass
        return monitors_list

    def update_monitor_controls(self):
        while self.dynamic_buttons_layout.count():
            child = self.dynamic_buttons_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        
        self.primary_monitor_combo.clear()
        self.secondary_monitor_combo.clear()

        for monitor in self.monitors:
            button = QPushButton(f"仅显示 {monitor['description']}")
            button.clicked.connect(
                lambda checked, num=monitor['id']: self.execute_with_lock(self.switch_to_single_display, num)
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

    def load_orientation_config(self):
        if os.path.exists(ORIENTATION_CONFIG_FILE):
            try:
                with open(ORIENTATION_CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                logging.info("Orientation config loaded successfully.")
                return config
            except Exception as e:
                logging.error(f"Failed to load orientation config: {e}")
                return {}
        return {}

    def save_orientation_config(self):
        try:
            with open(ORIENTATION_CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.orientation_config, f, indent=4, ensure_ascii=False)
            logging.info("Orientation config saved successfully.")
        except Exception as e:
            logging.error(f"Failed to save orientation config: {e}")

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

    def apply_advanced_extend_with_lock(self):
        """带锁的高级扩展应用"""
        primary_idx = self.primary_monitor_combo.currentIndex()
        secondary_idx = self.secondary_monitor_combo.currentIndex()

        if primary_idx == secondary_idx:
            QMessageBox.warning(self, "选择错误", "主显示器和扩展副屏不能是同一台显示器！")
            return
        
        primary_monitor_num = self.monitors[primary_idx]['id']
        secondary_monitor_num = self.monitors[secondary_idx]['id']
        
        primary_orientation = self.primary_orientation_combo.currentIndex()
        secondary_orientation = self.secondary_orientation_combo.currentIndex()
        
        self.orientation_config[str(primary_monitor_num)] = primary_orientation
        self.orientation_config[str(secondary_monitor_num)] = secondary_orientation
        self.save_orientation_config()
        
        self.execute_with_lock(
            self.extend_two_monitors_with_orientation,
            primary_monitor_num, secondary_monitor_num,
            primary_orientation, secondary_orientation
        )

    def set_monitor_orientation(self, device_name, orientation):
        """设置显示器方向 - 使用保存的原生分辨率"""
        try:
            self.info_display.append(f"[INFO] 开始设置 {device_name} 方向为 {ORIENTATION_NAMES[orientation]}...")
            
            if device_name not in self.monitor_native_resolutions:
                self.info_display.append(f"[ERROR] 未找到 {device_name} 的原生分辨率记录")
                return False
            
            native_res = self.monitor_native_resolutions[device_name]
            native_width = native_res['width']
            native_height = native_res['height']
            
            self.info_display.append(f"[DEBUG] 原生分辨率(横向): {native_width}x{native_height}")
            
            current_settings = win32api.EnumDisplaySettings(device_name, win32con.ENUM_CURRENT_SETTINGS)
            current_settings.DisplayOrientation = orientation
            
            if orientation in [ORIENTATION_PORTRAIT, ORIENTATION_PORTRAIT_FLIPPED]:
                current_settings.PelsWidth = native_height
                current_settings.PelsHeight = native_width
                self.info_display.append(f"[DEBUG] 目标分辨率(纵向): {native_height}x{native_width}")
            else:
                current_settings.PelsWidth = native_width
                current_settings.PelsHeight = native_height
                self.info_display.append(f"[DEBUG] 目标分辨率(横向): {native_width}x{native_height}")
            
            result = win32api.ChangeDisplaySettingsEx(
                device_name,
                current_settings,
                win32con.CDS_UPDATEREGISTRY
            )
            
            self.info_display.append(f"[DEBUG] 返回值: {result}")
            
            if result == win32con.DISP_CHANGE_SUCCESSFUL:
                self.info_display.append(f"[SUCCESS] {device_name} 方向设置成功")
                logging.info(f"Orientation set successfully for {device_name}: {ORIENTATION_NAMES[orientation]}")
                return True
            else:
                self.info_display.append(f"[ERROR] 设置失败，错误代码: {result}")
                return False
                
        except Exception as e:
            self.info_display.append(f"[ERROR] 异常: {str(e)}")
            import traceback
            self.info_display.append(f"[DEBUG] 追踪:\n{traceback.format_exc()}")
            logging.error(f"Error setting orientation for {device_name}: {e}")
            return False

    def extend_two_monitors_with_orientation(self, primary_num, secondary_num, primary_orientation, secondary_orientation):
        logging.info(f"Extending monitors {primary_num} and {secondary_num} with orientations")
        self.info_display.append(f"\n[INFO] 开始配置双屏扩展...")
        self.info_display.append(f"[INFO] 主显示器: 显示器{primary_num} - {ORIENTATION_NAMES[primary_orientation]}")
        self.info_display.append(f"[INFO] 副显示器: 显示器{secondary_num} - {ORIENTATION_NAMES[secondary_orientation]}")
        
        try:
            self.info_display.append(f"\n[INFO] 第一步: 配置显示器启用/禁用...")
            cmds = [TOOL_PATH]
            for m in self.monitors:
                if m['id'] not in (primary_num, secondary_num):
                    cmds.extend(['/disable', str(m['id'])])
            cmds.extend(['/enable', str(primary_num), '/enable', str(secondary_num), '/setprimary', str(primary_num)])
            subprocess.run(cmds, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.info_display.append("[SUCCESS] 显示器已启用")
            
            time.sleep(1.5)
            
            self.info_display.append(f"\n[INFO] 第二步: 设置显示器方向...")
            primary_device = self.monitors[primary_num - 1]['device_name']
            secondary_device = self.monitors[secondary_num - 1]['device_name']
            
            result1 = self.set_monitor_orientation(primary_device, primary_orientation)
            time.sleep(0.8)
            
            result2 = self.set_monitor_orientation(secondary_device, secondary_orientation)
            time.sleep(0.8)
            
            self.info_display.append(f"\n[INFO] 第三步: 应用所有更改...")
            result = win32api.ChangeDisplaySettingsEx(None, None, 0)
            self.info_display.append(f"[DEBUG] 全局更改返回值: {result}")
            
            time.sleep(1)
            
            self.info_display.append(f"\n[INFO] 第四步: 启用扩展模式...")
            try:
                subprocess.run(['DisplaySwitch.exe', '/extend'], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
                self.info_display.append("[SUCCESS] 扩展模式已启用")
            except Exception as e:
                self.info_display.append(f"[WARNING] DisplaySwitch 执行出现警告: {e}")
            
            time.sleep(1)
            
            if result1 and result2:
                self.info_display.append(f"\n[SUCCESS] ✓ 双屏扩展配置完成！")
                logging.info("Extend with orientation successful.")
            else:
                self.info_display.append(f"\n[WARNING] ⚠ 双屏扩展已启用，但部分方向设置可能未完全生效")
            
        except Exception as e:
            self.info_display.append(f"\n[ERROR] 扩展失败: {str(e)}")
            import traceback
            self.info_display.append(f"[DEBUG] 完整追踪:\n{traceback.format_exc()}")
            logging.error(f"Extend with orientation failed: {e}")

    def switch_to_single_display(self, monitor_num):
        logging.info(f"Switching to single display: {monitor_num}")
        self.info_display.append(f"\n[INFO] 切换到仅显示器 {monitor_num}...")
        try:
            cmds = [TOOL_PATH]
            for m in self.monitors:
                if m['id'] != monitor_num:
                    cmds.extend(['/disable', str(m['id'])])
            cmds.extend(['/enable', str(monitor_num), '/setprimary', str(monitor_num)])
            subprocess.run(cmds, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.info_display.append("[SUCCESS] 切换完成!")
        except Exception as e:
            self.info_display.append(f"[ERROR] 切换失败: {e}")

    def save_config(self):
        try:
            subprocess.run([TOOL_PATH, '/SaveConfig', CONFIG_FILE], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.info_display.append("\n[SUCCESS] 当前配置已保存!")
        except Exception as e:
            self.info_display.append(f"[ERROR] 保存失败: {e}")

    def load_config(self):
        if not os.path.exists(CONFIG_FILE):
            QMessageBox.warning(self, "错误", "无已保存的配置文件!")
            return
        try:
            subprocess.run([TOOL_PATH, '/LoadConfig', CONFIG_FILE], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.info_display.append("\n[SUCCESS] 已加载保存的配置!")
        except Exception as e:
            self.info_display.append(f"[ERROR] 加载失败: {e}")

    @pyqtSlot(str)
    def handle_switch_mode_signal(self, arg):
        """处理托盘图标的切换模式信号"""
        threading.Thread(target=lambda: self.execute_with_lock(self.run_displayswitch_legacy, arg)).start()

    def run_displayswitch_legacy(self, arg):
        """执行Windows内置的DisplaySwitch.exe命令"""
        self.info_display.append(f"\n执行命令: DisplaySwitch.exe {arg}")
        try:
            subprocess.run(['DisplaySwitch.exe', arg], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.info_display.append("[INFO] DisplaySwitch 命令执行成功")
        except Exception as e:
            self.info_display.append(f"命令执行失败: {e}")

    @pyqtSlot()
    def quit_application(self):
        logging.info("Application quit by user.")
        QApplication.instance().quit()

    @pyqtSlot(int)
    def set_startup_status(self, state):
        if state == Qt.CheckState.Checked.value:
            self.add_to_startup()
        else:
            self.remove_from_startup()

    def add_to_startup(self):
        try:
            if getattr(sys, 'frozen', False):
                app_path = sys.executable
            else:
                app_path = os.path.abspath(__file__)
                app_path = f'python.exe "{app_path}"'
            
            command = f'cmd /c "timeout /t 60 && {app_path} --silent"'
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_REG_KEY, 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, command)
            winreg.CloseKey(key)
            
            self.info_display.append("\n[SUCCESS] 已启用开机自启动 (延时60秒)")
            
        except Exception as e:
            QMessageBox.warning(self, "错误", f"启用自启动失败: {e}")
            self.startup_checkbox.blockSignals(True)
            self.startup_checkbox.setChecked(False)
            self.startup_checkbox.blockSignals(False)

    def remove_from_startup(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_REG_KEY, 0, winreg.KEY_SET_VALUE)
            winreg.DeleteValue(key, APP_NAME)
            winreg.CloseKey(key)
            self.info_display.append("\n[SUCCESS] 已禁用开机自启动")
        except FileNotFoundError:
            pass
        except Exception as e:
            QMessageBox.warning(self, "错误", f"禁用自启动失败: {e}")

    def check_startup_status(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_REG_KEY, 0, winreg.KEY_READ)
            value, regtype = winreg.QueryValueEx(key, APP_NAME)
            winreg.CloseKey(key)
            self.startup_checkbox.blockSignals(True)
            self.startup_checkbox.setChecked(True)
            self.startup_checkbox.blockSignals(False)
        except FileNotFoundError:
            self.startup_checkbox.blockSignals(True)
            self.startup_checkbox.setChecked(False)
            self.startup_checkbox.blockSignals(False)


# --- 托盘图标设置 ---
def setup_tray_icon(main_window):
    icon_path = 'icon.png'
    if os.path.exists(icon_path):
        try:
            icon_image = Image.open(icon_path)
        except Exception as e:
            icon_image = create_dummy_icon()
    else:
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

    def create_single_monitor_items():
        items = []
        for monitor in main_window.monitors:
            def make_handler(monitor_num):
                def handler(icon, item):
                    threading.Thread(
                        target=lambda: main_window.execute_with_lock(
                            main_window.switch_to_single_display, monitor_num
                        )
                    ).start()
                return handler
            items.append(pystray.MenuItem(f'仅 {monitor["description"]}', make_handler(monitor['id'])))
        return items

    menu_items = [
        pystray.MenuItem('显示主窗口', on_show_window, default=True),  # 新增 default=True 支持双击
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('扩展模式(所有)', on_extend),
        pystray.MenuItem('复制模式(所有)', on_clone),
        pystray.Menu.SEPARATOR,
    ]
    
    if len(main_window.monitors) > 0:
        menu_items.extend(create_single_monitor_items())
        menu_items.append(pystray.Menu.SEPARATOR)
    
    menu_items.append(pystray.MenuItem('退出', on_quit))

    tray_menu = pystray.Menu(*menu_items)
    tray_icon = pystray.Icon("monitor_manager", icon_image, "显示器切换工具 V6.2", tray_menu)

    def run_tray():
        tray_icon.run()

    tray_thread = threading.Thread(target=run_tray, daemon=True)
    tray_thread.start()


# --- 主程序入口 ---
if __name__ == '__main__':
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
    
    logging.info("--- Application Started (V6.2) ---")
    
    is_silent = '--silent' in sys.argv
    
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    
    main_window = MonitorApp()
    setup_tray_icon(main_window)
    
    if not is_silent:
        main_window.show()
        logging.info("Main window displayed.")
    else:
        logging.info("Starting in silent mode. Tray icon only.")
    
    sys.exit(app.exec())

        