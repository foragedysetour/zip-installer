zip-installer — 右键安装压缩包工具（Windows）

功能概述

- 在 Windows 下，为 `.zip`, `.rar`, `.7z` 文件提供右键菜单“从此文件安装”。
- 点击后将压缩包解压到「设置的目标路径\<文件名>」目录下；若目录非空会提示是否覆盖。
- 解压过程中显示 pywebview 的进度窗口；完成后使用 `plyer` 发送通知，并在 Windows 下尝试使用 `win10toast_click` 发送可点击的通知以打开目录。
- 双击运行程序会打开设置界面，可指定目标路径并保存到 `~/.zip-installer_config.json`。

快速开始

1. 安装依赖（推荐在虚拟环境中运行）:

```powershell
python -m pip install -r requirements.txt
```

2. 注册右键菜单（当前为用户级注册，不需要管理员权限）:

```powershell
python zip-installer.py --register
```

这会为当前用户的 `.zip`、`.rar`、`.7z` 注册右键菜单，命令指向当前 Python 解释器和脚本路径。

3. 使用

- 在资源管理器中右键某个压缩包，选择“从此文件安装”。程序会被调用并开始解压。
- 双击 `zip-installer.py` 可以打开设置页面，设置目标路径。

‼️注意

- 解压 7z/rar 依赖 `7z` 命令行工具（7-Zip）。请确保 `7z.exe` 在系统 PATH 中，或者改为打包时将 `zip-installer.exe` 放到包含 `7z.exe` 的环境中。
- 打包成单文件可执行时，请用 `pyinstaller` 打包并在注册时把命令改为指向 exe（脚本中默认使用 `sys.executable` 与脚本路径）。

开发与打包建议

- 构建 exe:

```powershell
pyinstaller -w -F -i logo.ico --hidden-import=plyer.platforms.win.notification zip-installer.py
```

