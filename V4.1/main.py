# MonitorManager V7.0 - 性能优化版
# 主要优化:
# - 使用 QThread 替代 threading，提升性能
# - 减少不必要的延迟和API调用
# - 批量操作优化
# - 智能缓存机制

import sys
import subprocess
import win32api
import win32con
import logging
import os
import winreg
import json
from typing import Dict, List, Optional

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QLabel, QHBoxLayout, QTextEdit,
                             QGridLayout, QMessageBox, QComboBox, QFrame,
                             QCheckBox, QProgressBar)
from PyQt6.QtGui import QFont, QIcon
from PyQt6.QtCore import Qt, pyqtSignal, pyqtSlot, QThread, QMutex, QMutexLocker
import pystray
from PIL import Image, ImageDraw, ImageFont

# --- 全局配置 ---
logging.basicConfig(
    filename='monitor_manager.log', 
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s'
)

TOOL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'MultiMonitorTool.exe')
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'monitor_config.cfg')
ORIENTATION_CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'monitor_orientation_config.json')

APP_NAME = "MonitorManagerV7"
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


# --- 工作线程类 ---
class MonitorOperationWorker(QThread):
    """异步执行显示器操作的工作线程"""
    operation_started = pyqtSignal(str)
    operation_progress = pyqtSignal(str)
    operation_completed = pyqtSignal(bool, str)
    
    def __init__(self, operation_func, *args, **kwargs):
        super().__init__()
        self.operation_func = operation_func
        self.args = args
        self.kwargs = kwargs
        self.success = False
        self.message = ""
    
    def run(self):
        try:
            self.operation_started.emit(f"开始执行操作...")
            result = self.operation_func(*self.args, **self.kwargs)
            self.success = True
            self.message = "操作成功完成"
            self.operation_completed.emit(True, self.message)
        except Exception as e:
            self.success = False
            self.message = f"操作失败: {str(e)}"
            logging.error(f"Worker error: {e}", exc_info=True)
            self.operation_completed.emit(False, self.message)


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
        self.operation_mutex = QMutex()  # 使用 QMutex 替代布尔锁
        self.current_worker = None
        self.monitors_cache_valid = False
        
        self.toggle_window_signal.connect(self.toggle_visibility)
        self.refresh_info_signal.connect(self.update_display_info)
        self.switch_mode_signal.connect(self.handle_switch_mode_signal)
        self.quit_app_signal.connect(self.quit_application)
        
        icon_path = 'icon.png'
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("显示器切换工具 V7.0 (性能优化版)")
        self.setGeometry(300, 300, 850, 750)
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
        
        # --- 进度条 ---
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("操作进行中...")
        
        # --- 基础控制区 ---
        static_button_layout = QHBoxLayout()
        self.btn_extend = QPushButton("扩展(所有)")
        self.btn_clone = QPushButton("复制(所有)")
        self.btn_refresh = QPushButton("刷新信息")
        self.btn_save_config = QPushButton("保存当前配置")
        self.btn_load_config = QPushButton("加载保存的配置")
        
        # 设置按钮样式
        for btn in [self.btn_extend, self.btn_clone, self.btn_refresh, 
                    self.btn_save_config, self.btn_load_config]:
            btn.setMinimumHeight(35)
        
        static_button_layout.addWidget(self.btn_extend)
        static_button_layout.addWidget(self.btn_clone)
        static_button_layout.addWidget(self.btn_refresh)
        static_button_layout.addWidget(self.btn_save_config)
        static_button_layout.addWidget(self.btn_load_config)
        
        # --- 单显示器模式区 ---
        self.dynamic_buttons_label = QLabel("单显示器模式:")
        self.dynamic_buttons_label.setFont(QFont("Microsoft YaHei", 10))
        self.dynamic_buttons_layout = QVBoxLayout()
        
        # --- 高级扩展模式UI ---
        self.advanced_extend_frame = QFrame()
        self.advanced_extend_frame.setFrameShape(QFrame.Shape.StyledPanel)
        advanced_layout = QVBoxLayout(self.advanced_extend_frame)
        
        advanced_label = QLabel("高级扩展模式:")
        advanced_label.setFont(QFont("Microsoft YaHei", 10, QFont.Weight.Bold))
        
        selection_layout = QGridLayout()
        
        selection_layout.addWidget(QLabel("主显示器:"), 0, 0)
        self.primary_monitor_combo = QComboBox()
        self.primary_monitor_combo.setMinimumHeight(30)
        selection_layout.addWidget(self.primary_monitor_combo, 0, 1)
        
        selection_layout.addWidget(QLabel("主显示器方向:"), 0, 2)
        self.primary_orientation_combo = QComboBox()
        self.primary_orientation_combo.setMinimumHeight(30)
        self.primary_orientation_combo.addItems(list(ORIENTATION_NAMES.values()))
        selection_layout.addWidget(self.primary_orientation_combo, 0, 3)
        
        selection_layout.addWidget(QLabel("扩展副屏:"), 1, 0)
        self.secondary_monitor_combo = QComboBox()
        self.secondary_monitor_combo.setMinimumHeight(30)
        selection_layout.addWidget(self.secondary_monitor_combo, 1, 1)
        
        selection_layout.addWidget(QLabel("副显示器方向:"), 1, 2)
        self.secondary_orientation_combo = QComboBox()
        self.secondary_orientation_combo.setMinimumHeight(30)
        self.secondary_orientation_combo.addItems(list(ORIENTATION_NAMES.values()))
        selection_layout.addWidget(self.secondary_orientation_combo, 1, 3)
        
        self.btn_apply_extend = QPushButton("应用双屏扩展设置")
        self.btn_apply_extend.setMinimumHeight(35)
        
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
        main_layout.addWidget(self.progress_bar)
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
        self.btn_refresh.clicked.connect(self.force_update_display_info)
        self.btn_extend.clicked.connect(lambda: self.execute_async_operation(
            self.run_displayswitch_legacy, '/extend'))
        self.btn_clone.clicked.connect(lambda: self.execute_async_operation(
            self.run_displayswitch_legacy, '/clone'))
        self.btn_save_config.clicked.connect(lambda: self.execute_async_operation(
            self.save_config))
        self.btn_load_config.clicked.connect(lambda: self.execute_async_operation(
            self.load_config))
        self.btn_apply_extend.clicked.connect(self.apply_advanced_extend_async)
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
        
        for i in range(self.dynamic_buttons_layout.count()):
            widget = self.dynamic_buttons_layout.itemAt(i).widget()
            if widget:
                widget.setEnabled(enabled)

    def execute_async_operation(self, func, *args, **kwargs):
        """异步执行操作"""
        locker = QMutexLocker(self.operation_mutex)
        if self.current_worker and self.current_worker.isRunning():
            self.info_display.append("[WARNING] 操作正在进行中，请稍候...\n")
            return
        
        self.set_buttons_enabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # 不确定进度
        
        self.current_worker = MonitorOperationWorker(func, *args, **kwargs)
        self.current_worker.operation_started.connect(self.on_operation_started)
        self.current_worker.operation_progress.connect(self.on_operation_progress)
        self.current_worker.operation_completed.connect(self.on_operation_completed)
        self.current_worker.start()

    @pyqtSlot(str)
    def on_operation_started(self, message):
        self.info_display.append(f"[INFO] {message}")

    @pyqtSlot(str)
    def on_operation_progress(self, message):
        self.info_display.append(f"[PROGRESS] {message}")

    @pyqtSlot(bool, str)
    def on_operation_completed(self, success, message):
        self.progress_bar.setVisible(False)
        self.set_buttons_enabled(True)
        
        if success:
            self.info_display.append(f"[SUCCESS] ✓ {message}")
            # 智能刷新：只在需要时刷新
            self.update_display_info()
        else:
            self.info_display.append(f"[ERROR] ✗ {message}")
        
        self.info_display.append("")  # 空行分隔
        self.current_worker = None

    @pyqtSlot()
    def toggle_visibility(self):
        if self.isVisible():
            self.hide()
        else:
            self.show()
            self.activateWindow()

    @pyqtSlot()
    def force_update_display_info(self):
        """强制刷新显示器信息"""
        self.monitors_cache_valid = False
        self.update_display_info()

    @pyqtSlot()
    def update_display_info(self):
        """更新显示器信息（使用缓存机制）"""
        if not self.monitors_cache_valid:
            self.info_display.clear()
            self.monitors = self.get_all_monitors()
            self.monitors_cache_valid = True
        
        info_text = f"检测到 {len(self.monitors)} 台显示器\n" + "-" * 60 + "\n"
        for monitor in self.monitors:
            is_primary_str = " (主)" if monitor['is_primary'] else ""
            info_text += f"{monitor['description']}{is_primary_str}\n"
        self.info_display.setText(info_text)
        
        logging.info(f"Display info updated: {len(self.monitors)} monitors")
        self.update_monitor_controls()

    def get_all_monitors(self) -> List[Dict]:
        """获取所有显示器信息（优化版）"""
        monitors_list = []
        self.monitor_native_resolutions.clear()
        
        try:
            i = 0
            while True:
                try:
                    device = win32api.EnumDisplayDevices(None, i)
                except:
                    break
                    
                if device.StateFlags & win32con.DISPLAY_DEVICE_ATTACHED_TO_DESKTOP:
                    settings = win32api.EnumDisplaySettings(device.DeviceName, win32con.ENUM_CURRENT_SETTINGS)
                    is_primary = (settings.Position_x == 0 and settings.Position_y == 0)
                    
                    # 计算原生分辨率
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
                    
                    description = (f"显示器{i+1}: {settings.PelsWidth}x{settings.PelsHeight} "
                                 f"@ {settings.DisplayFrequency}Hz")

                    monitors_list.append({
                        'id': i + 1,
                        'device': device,
                        'device_name': device.DeviceName,
                        'settings': settings, 
                        'is_primary': is_primary,
                        'description': description
                    })
                i += 1
        except Exception as e:
            logging.error(f"Error getting monitors: {e}")
        
        return monitors_list

    def update_monitor_controls(self):
        """更新监视器控制界面"""
        # 清空动态按钮
        while self.dynamic_buttons_layout.count():
            child = self.dynamic_buttons_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        
        self.primary_monitor_combo.clear()
        self.secondary_monitor_combo.clear()

        # 创建单显示器按钮
        for monitor in self.monitors:
            button = QPushButton(f"仅显示 {monitor['description']}")
            button.setMinimumHeight(35)
            button.clicked.connect(
                lambda checked, num=monitor['id']: self.execute_async_operation(
                    self.switch_to_single_display, num)
            )
            self.dynamic_buttons_layout.addWidget(button)

        # 更新双显示器选择
        if len(self.monitors) >= 2:
            self.advanced_extend_frame.show()
            descriptions = [m['description'] for m in self.monitors]
            self.primary_monitor_combo.addItems(descriptions)
            self.secondary_monitor_combo.addItems(descriptions)
            if len(self.monitors) > 1:
                self.secondary_monitor_combo.setCurrentIndex(1)
        else:
            self.advanced_extend_frame.hide()

    def load_orientation_config(self) -> Dict:
        """加载方向配置"""
        if os.path.exists(ORIENTATION_CONFIG_FILE):
            try:
                with open(ORIENTATION_CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                logging.error(f"Failed to load orientation config: {e}")
        return {}

    def save_orientation_config(self):
        """保存方向配置"""
        try:
            with open(ORIENTATION_CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.orientation_config, f, indent=4, ensure_ascii=False)
            logging.info("Orientation config saved")
        except Exception as e:
            logging.error(f"Failed to save orientation config: {e}")

    def load_primary_orientation(self):
        """加载主显示器方向"""
        if self.primary_monitor_combo.currentIndex() >= 0:
            monitor_id = self.monitors[self.primary_monitor_combo.currentIndex()]['id']
            orientation = self.orientation_config.get(str(monitor_id), ORIENTATION_LANDSCAPE)
            self.primary_orientation_combo.setCurrentIndex(orientation)

    def load_secondary_orientation(self):
        """加载副显示器方向"""
        if self.secondary_monitor_combo.currentIndex() >= 0:
            monitor_id = self.monitors[self.secondary_monitor_combo.currentIndex()]['id']
            orientation = self.orientation_config.get(str(monitor_id), ORIENTATION_LANDSCAPE)
            self.secondary_orientation_combo.setCurrentIndex(orientation)

    def apply_advanced_extend_async(self):
        """异步应用高级扩展设置"""
        primary_idx = self.primary_monitor_combo.currentIndex()
        secondary_idx = self.secondary_monitor_combo.currentIndex()

        if primary_idx == secondary_idx:
            QMessageBox.warning(self, "选择错误", "主显示器和扩展副屏不能是同一台显示器！")
            return
        
        primary_monitor_num = self.monitors[primary_idx]['id']
        secondary_monitor_num = self.monitors[secondary_idx]['id']
        
        primary_orientation = self.primary_orientation_combo.currentIndex()
        secondary_orientation = self.secondary_orientation_combo.currentIndex()
        
        # 保存方向配置
        self.orientation_config[str(primary_monitor_num)] = primary_orientation
        self.orientation_config[str(secondary_monitor_num)] = secondary_orientation
        self.save_orientation_config()
        
        # 异步执行
        self.execute_async_operation(
            self.extend_two_monitors_with_orientation,
            primary_monitor_num, secondary_monitor_num,
            primary_orientation, secondary_orientation
        )

    def set_monitor_orientation(self, device_name: str, orientation: int) -> bool:
        """设置显示器方向（优化版）"""
        try:
            if device_name not in self.monitor_native_resolutions:
                logging.error(f"Native resolution not found for {device_name}")
                return False
            
            native_res = self.monitor_native_resolutions[device_name]
            native_width = native_res['width']
            native_height = native_res['height']
            
            current_settings = win32api.EnumDisplaySettings(device_name, win32con.ENUM_CURRENT_SETTINGS)
            current_settings.DisplayOrientation = orientation
            
            # 根据方向设置分辨率
            if orientation in [ORIENTATION_PORTRAIT, ORIENTATION_PORTRAIT_FLIPPED]:
                current_settings.PelsWidth = native_height
                current_settings.PelsHeight = native_width
            else:
                current_settings.PelsWidth = native_width
                current_settings.PelsHeight = native_height
            
            # 一次性应用更改（使用 CDS_UPDATEREGISTRY 保存到注册表）
            result = win32api.ChangeDisplaySettingsEx(
                device_name,
                current_settings,
                win32con.CDS_UPDATEREGISTRY | win32con.CDS_NORESET
            )
            
            if result == win32con.DISP_CHANGE_SUCCESSFUL:
                logging.info(f"Orientation set for {device_name}: {ORIENTATION_NAMES[orientation]}")
                return True
            else:
                logging.error(f"Failed to set orientation: error code {result}")
                return False
                
        except Exception as e:
            logging.error(f"Exception in set_monitor_orientation: {e}")
            return False

    def extend_two_monitors_with_orientation(self, primary_num: int, secondary_num: int, 
                                            primary_orientation: int, secondary_orientation: int):
        """扩展双显示器（优化版 - 减少延迟）"""
        logging.info(f"Extending monitors {primary_num} and {secondary_num}")
        
        try:
            # 步骤1：使用 MultiMonitorTool 批量配置
            cmds = [TOOL_PATH]
            for m in self.monitors:
                if m['id'] not in (primary_num, secondary_num):
                    cmds.extend(['/disable', str(m['id'])])
            cmds.extend(['/enable', str(primary_num), '/enable', str(secondary_num), 
                        '/setprimary', str(primary_num)])
            subprocess.run(cmds, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            
            # 步骤2：批量设置方向（使用 NORESET 标志延迟应用）
            primary_device = self.monitors[primary_num - 1]['device_name']
            secondary_device = self.monitors[secondary_num - 1]['device_name']
            
            result1 = self.set_monitor_orientation(primary_device, primary_orientation)
            result2 = self.set_monitor_orientation(secondary_device, secondary_orientation)
            
            # 步骤3：一次性应用所有更改
            win32api.ChangeDisplaySettingsEx(None, None, 0)
            
            # 步骤4：启用扩展模式
            subprocess.run(['DisplaySwitch.exe', '/extend'], 
                         check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            
            if result1 and result2:
                logging.info("Extend with orientation successful")
            else:
                logging.warning("Extend completed but some orientations may not be applied")
            
        except Exception as e:
            logging.error(f"Extend failed: {e}")
            raise

    def switch_to_single_display(self, monitor_num: int):
        """切换到单显示器（优化版）"""
        logging.info(f"Switching to monitor {monitor_num}")
        try:
            cmds = [TOOL_PATH]
            for m in self.monitors:
                if m['id'] != monitor_num:
                    cmds.extend(['/disable', str(m['id'])])
            cmds.extend(['/enable', str(monitor_num), '/setprimary', str(monitor_num)])
            subprocess.run(cmds, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            logging.info(f"Switched to monitor {monitor_num} successfully")
        except Exception as e:
            logging.error(f"Failed to switch: {e}")
            raise

    def save_config(self):
        """保存配置"""
        try:
            subprocess.run([TOOL_PATH, '/SaveConfig', CONFIG_FILE], 
                         check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            logging.info("Config saved")
        except Exception as e:
            logging.error(f"Save config failed: {e}")
            raise

    def load_config(self):
        """加载配置"""
        if not os.path.exists(CONFIG_FILE):
            raise FileNotFoundError("无已保存的配置文件")
        try:
            subprocess.run([TOOL_PATH, '/LoadConfig', CONFIG_FILE], 
                         check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            logging.info("Config loaded")
        except Exception as e:
            logging.error(f"Load config failed: {e}")
            raise

    @pyqtSlot(str)
    def handle_switch_mode_signal(self, arg: str):
        """处理托盘图标的切换模式信号"""
        self.execute_async_operation(self.run_displayswitch_legacy, arg)

    def run_displayswitch_legacy(self, arg: str):
        """执行Windows DisplaySwitch命令"""
        try:
            subprocess.run(['DisplaySwitch.exe', arg], 
                         check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            logging.info(f"DisplaySwitch {arg} executed successfully")
        except Exception as e:
            logging.error(f"DisplaySwitch failed: {e}")
            raise

    @pyqtSlot()
    def quit_application(self):
        """退出应用程序"""
        logging.info("Application quit by user")
        QApplication.instance().quit()

    @pyqtSlot(int)
    def set_startup_status(self, state):
        """设置开机启动状态"""
        if state == Qt.CheckState.Checked.value:
            self.add_to_startup()
        else:
            self.remove_from_startup()

    def add_to_startup(self):
        """添加到开机启动"""
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
            logging.info("Startup enabled")
            
        except Exception as e:
            QMessageBox.warning(self, "错误", f"启用自启动失败: {e}")
            self.startup_checkbox.blockSignals(True)
            self.startup_checkbox.setChecked(False)
            self.startup_checkbox.blockSignals(False)

    def remove_from_startup(self):
        """从开机启动移除"""
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_REG_KEY, 0, winreg.KEY_SET_VALUE)
            winreg.DeleteValue(key, APP_NAME)
            winreg.CloseKey(key)
            self.info_display.append("\n[SUCCESS] 已禁用开机自启动")
            logging.info("Startup disabled")
        except FileNotFoundError:
            pass
        except Exception as e:
            QMessageBox.warning(self, "错误", f"禁用自启动失败: {e}")

    def check_startup_status(self):
        """检查开机启动状态"""
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
    """设置系统托盘图标"""
    icon_path = 'icon.png'
    if os.path.exists(icon_path):
        try:
            icon_image = Image.open(icon_path)
        except:
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
                    main_window.execute_async_operation(
                        main_window.switch_to_single_display, monitor_num
                    )
                return handler
            items.append(pystray.MenuItem(
                f'仅 {monitor["description"]}', 
                make_handler(monitor['id'])
            ))
        return items

    menu_items = [
        pystray.MenuItem('显示主窗口', on_show_window, default=True),
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
    tray_icon = pystray.Icon(
        "monitor_manager", 
        icon_image, 
        "显示器切换工具 V7.0", 
        tray_menu
    )

    def run_tray():
        tray_icon.run()

    from threading import Thread
    tray_thread = Thread(target=run_tray, daemon=True)
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
    
    logging.info("=== Application Started V7.0 (Performance Optimized) ===")
    
    is_silent = '--silent' in sys.argv
    
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    
    # 设置应用程序样式（可选）
    app.setStyle('Fusion')
    
    main_window = MonitorApp()
    setup_tray_icon(main_window)
    
    if not is_silent:
        main_window.show()
        logging.info("Main window displayed")
    else:
        logging.info("Starting in silent mode")
    
    sys.exit(app.exec())
