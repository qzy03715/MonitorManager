import PyInstaller.__main__
import os

# 获取当前脚本所在目录（V3.0目录）
current_dir = os.path.dirname(os.path.abspath(__file__))

# 构建正确的文件路径
icon_path = os.path.join(current_dir, 'icon.png')
tool_path = os.path.join(current_dir, 'MultiMonitorTool.exe')
main_py_path = os.path.join(current_dir, 'main.py')

# 检查文件是否存在
if not os.path.exists(icon_path):
    print(f"错误: 找不到 {icon_path}")
    exit(1)

if not os.path.exists(tool_path):
    print(f"错误: 找不到 {tool_path}")
    exit(1)

if not os.path.exists(main_py_path):
    print(f"错误: 找不到 {main_py_path}")
    exit(1)

print(f"当前目录: {current_dir}")
print(f"图标路径: {icon_path}")
print(f"工具路径: {tool_path}")
print(f"脚本路径: {main_py_path}")

# PyInstaller 打包命令
PyInstaller.__main__.run([
    main_py_path,
    '--name=显示器切换工具',
    '--onefile',
    '--windowed',
    f'--icon={icon_path}',
    f'--add-data={icon_path};.',
    f'--add-data={tool_path};.',
    '--noconsole',
    '--clean',
])

print("\n打包完成！")
print(f"可执行文件位置: {os.path.join(current_dir, 'dist', '显示器切换工具.exe')}")
