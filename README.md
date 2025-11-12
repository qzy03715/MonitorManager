# MonitorManager V6.2

[**中文说明**](https://github.com/qzy03715/MonitorManager/blob/main/README_ZH.md)

**Special thanks to Gemini, Claude, and Grok for their invaluable contributions to this project.**

A powerful and user-friendly tool for managing multi-monitor setups on Windows. It simplifies switching display modes, rotating screens, and saving complex configurations through an intuitive GUI and system tray icon.


## ✨ Key Features

- **One-Click Mode Switching**: Instantly switch between `Extend`, `Duplicate`, and `Single Monitor` modes.
- **Advanced Screen Rotation**: Rotate any monitor to `Portrait (90°)` or `Flipped (180°/270°)`, while preserving its native resolution and clarity.
- **Advanced Dual-Monitor Setup**: Easily configure a two-monitor extended desktop, assigning a specific orientation to each screen while disabling others.
- **Configuration Profiles**: Save your current multi-monitor layout and orientation settings and restore them anytime with a single click.
- **System Tray Integration**: Access key functions from a convenient tray icon. Double-click to show/hide the main window.
- **Auto-Start Option**: Set the application to launch automatically on Windows startup.

## 🚀 How to Use (For End-Users)

1. Navigate to the **Releases Page** [<sup>3</sup>](https://github.com/qzy03715/MonitorManager/releases).
2. Download the latest release ZIP file (e.g., `MonitorManager_v6.2.zip`).
3. Unzip the archive to a folder of your choice.
4. **Important**: Ensure `MultiMonitorTool.exe` is in the same directory as the main executable.
5. Run `显示器切换工具.exe` to start the application.

## 🛠️ How It Works (Technical Insight)

This application combines native Windows APIs with a well-known command-line utility to provide a seamless user experience.

- **Core Control**: It acts as a GUI front-end for Nirsoft's powerful `MultiMonitorTool.exe` for basic operations like enabling/disabling monitors and setting the primary display. This ensures high compatibility and stability.
- **Screen Rotation & Resolution**: Screen rotation is handled by directly calling the **Windows API** (`win32api`). A key feature is its ability to **cache each monitor's native resolution**. When you rotate a screen (e.g., from 1920x1080 to portrait mode), the tool intelligently calculates the new resolution (1080x1920) instead of letting Windows default to a lower, blurry resolution. This guarantees maximum clarity.
- **GUI Framework**: The user interface is built with **PyQt6**, providing a modern and responsive experience.
- **Concurrency Control**: An operation lock prevents multiple commands from running simultaneously, ensuring that display settings are applied correctly without conflicts.

## 🧑‍💻 How to Build from Source

1. **Prerequisites**:
    - Python 3.9+
    - Git

2. **Clone the repository**:

    ```bash
    git clone https://github.com/qzy03715/MonitorManager.git
    cd MonitorManager
    ```

3. **Set up a virtual environment** (recommended):

    ```bash
    python -m venv venv
    # On Windows
    venv\Scripts\activate
    # On macOS/Linux
    source venv/bin/activate
    ```

4. **Install dependencies**:

    ```bash
    pip install PyQt6 pywin32 pystray Pillow PyInstaller
    ```

5. **Run the application**:

    ```bash
    python main.py
    ```

6. **Build the EXE**:

    ```bash
    python build.py
    ```

    The final executable will be located in the `dist/` directory.

## 📄 License

This project is licensed under the MIT License [<sup>4</sup>](LICENSE).
