# 显示器切换工具 V6.2

[**EN**](https://github.com/qzy03715/MonitorManager/blob/main/README.md)

**特别感谢 Gemini、Claude 和 Grok 对本项目做出的宝贵贡献。**

一款强大且易于使用的 Windows 多显示器管理工具。它通过直观的图形界面和系统托盘图标，极大地简化了切换显示模式、旋转屏幕和保存复杂配置等操作。


## ✨ 主要功能

- **一键模式切换**: 快速在 `扩展`、`复制` 和 `仅单屏显示` 模式之间切换。
- **高级屏幕旋转**: 可将任意显示器旋转为 `纵向 (90°)` 或 `翻转 (180°/270°)`, 同时保持其原生分辨率和清晰度。
- **高级双屏设置**: 轻松配置双屏扩展桌面，为每个屏幕独立指定显示方向，并禁用其他所有显示器。
- **配置方案管理**: 保存当前的多显示器布局和方向设置，并可随时一键恢复。
- **系统托盘集成**: 通过便捷的托盘图标访问核心功能。双击图标可显示/隐藏主窗口。
- **开机自启动**: 可将应用设置为随 Windows 自动启动。

## 🚀 如何使用 (终端用户)

1. 访问本仓库的 **Releases 发布页面** [<sup>3</sup>](https://github.com/qzy03715/MonitorManager/releases)。
2. 下载最新的 ZIP 压缩包 (例如 `MonitorManager_v6.2.zip`)。
3. 将压缩包解压到您选择的任意文件夹。
4. **重要提示**: 请确保 `MultiMonitorTool.exe` 与主程序 `显示器切换工具.exe` 在同一目录下。
5. 运行 `显示器切换工具.exe` 启动程序。

## 🛠️ 实现原理 (技术简介)

本应用结合了原生 Windows API 和一个知名的命令行工具，以提供无缝的用户体验。

- **核心控制**: 对于启用/禁用显示器、设置主显示器等基础操作，本应用作为 Nirsoft 出品的强大工具 `MultiMonitorTool.exe` 的图形化前端。这确保了高度的兼容性和稳定性。
- **屏幕旋转与分辨率**: 屏幕旋转功能通过直接调用 **Windows API** (`win32api`) 实现。其核心优势在于，程序启动时会 **缓存每个显示器的原生分辨率**。当您旋转屏幕时 (例如从 1920x1080 切换到纵向模式)，工具会智能地计算出新分辨率 (1080x1920)，而不是让 Windows 降级到模糊的低分辨率。这保证了最佳的显示清晰度。
- **GUI 框架**: 用户界面基于 **PyQt6** 构建，提供了现代化的响应式体验。
- **并发控制**: 内置的操作锁机制可防止多个命令同时运行，确保显示设置在无冲突的情况下被正确应用。

## 🧑‍💻 如何从源码构建

1. **环境准备**:
    - Python 3.9+
    - Git

2. **克隆仓库**:

    ```bash
    git clone https://github.com/qzy03715/MonitorManager.git
    cd MonitorManager
    ```

3. **创建虚拟环境** (推荐):

    ```bash
    python -m venv venv
    # Windows 系统
    venv\Scripts\activate
    # macOS/Linux 系统
    source venv/bin/activate
    ```

4. **安装依赖**:

    ```bash
    pip install PyQt6 pywin32 pystray Pillow PyInstaller
    ```

5. **运行程序**:

    ```bash
    python main.py
    ```

6. **打包为 EXE**:

    ```bash
    python build.py
    ```

    最终的可执行文件将生成在 `dist/` 目录下。

## 📄 开源许可

本项目基于 MIT 许可证 [<sup>4</sup>](LICENSE) 开源。
